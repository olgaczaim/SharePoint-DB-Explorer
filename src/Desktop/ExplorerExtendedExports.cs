using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

namespace SharePointExplorer.Desktop
{
    public interface IExplorerExtendedExportController
    {
        DesktopExportSummary ExportAttachments(Node scope,string directory,CancellationToken cancellationToken,Action<DesktopExportProgress> progress);
        DesktopExportSummary ExportMigrationPackage(Node list,string directory,CancellationToken cancellationToken,Action<DesktopExportProgress> progress);
    }
    public sealed partial class ExplorerController : IExplorerExtendedExportController
    {
        public DesktopExportSummary ExportAttachments(Node scope,string directory,CancellationToken cancellationToken,Action<DesktopExportProgress> progress)
        {
            ValidateAttachmentScope(scope);
            if(String.IsNullOrWhiteSpace(directory)) throw new ArgumentException("Choose an export folder.","directory");
            Node selection=SnapshotFile(scope);
            lock(gate)
            {
                EnsureOpen();
                string root=Path.GetFullPath(directory);
                var summary=new DesktopExportSummary {Directory=root};
                if(cancellationToken.IsCancellationRequested){summary.Cancelled=true;return summary;}
                ISharePointAttachmentCatalog catalog=session.Catalog as ISharePointAttachmentCatalog;
                if(catalog==null) throw new NotSupportedException("This source cannot enumerate ordinary-list attachments.");
                var selected=new List<Node>();
                var identities=new Dictionary<Guid,Node>();
                try
                {
                    Node request=SnapshotFile(selection);
                    IEnumerable<Node> documents=selection.Kind==NodeKind.ListItem ? catalog.GetItemAttachments(request) : catalog.EnumerateCurrentListAttachments(request);
                    if(documents==null) throw new InvalidDataException("The source returned no attachment collection.");
                    using(IEnumerator<Node> files=documents.GetEnumerator())
                        while(!cancellationToken.IsCancellationRequested && files.MoveNext())
                        {
                            Node file=files.Current;ValidateAttachmentFile(file,selection);
                            Node existing;
                            if(identities.TryGetValue(file.Id,out existing))
                            {
                                if(existing.AttachmentOwnerId!=file.AttachmentOwnerId || existing.ListItemId!=file.ListItemId || existing.ItemUniqueId!=file.ItemUniqueId)
                                    throw new InvalidDataException("The same attachment document is assigned to different list items.");
                                continue;
                            }
                            Node copy=SnapshotFile(file);identities.Add(copy.Id,copy);
                            if(session.Engine.CanExport(copy)) selected.Add(copy);
                        }
                }
                catch(Exception error){ThrowIfContainsSecret(error);throw;}
                summary.Total=selected.Count;
                if(cancellationToken.IsCancellationRequested){summary.Cancelled=true;return summary;}
                Directory.CreateDirectory(root);
                summary.ReportPath=Path.Combine(root,"attachment-report-"+DateTime.UtcNow.ToString("yyyyMMdd-HHmmss",CultureInfo.InvariantCulture)+"-"+Guid.NewGuid().ToString("N").Substring(0,8)+".csv");
                using(var report=new StreamWriter(new FileStream(summary.ReportPath,FileMode.CreateNew,FileAccess.Write,FileShare.Read),new UTF8Encoding(true)))
                {
                    WriteCsv(report,"Source","SiteId","DocumentId","DocumentPath","Status","Bytes","Sha256","ExportPath","Message","OwnerDocumentId","ItemId","ItemUniqueId");
                    if(progress!=null) progress(new DesktopExportProgress {Completed=0,Total=summary.Total,Document=null,Message=""});
                    foreach(Node document in selected)
                    {
                        if(cancellationToken.IsCancellationRequested){summary.Cancelled=true;break;}
                        var entry=new DesktopExportEntry {Document=document,Status=RecoveryStatus.Success,Message="",Path="",Sha256=""};
                        try
                        {
                            Node owner=selection.Kind==NodeKind.ListItem ? SnapshotFile(selection) : AttachmentOwner(document);
                            PreparedDocument prepared=session.Engine.PrepareItemAttachment(document,owner);
                            Node current=prepared.Document;ValidateAttachmentFile(current,selection);
                            entry.Document=current;
                            string destination=selection.Kind==NodeKind.List ? Path.Combine(root,current.ListItemId.Value.ToString(CultureInfo.InvariantCulture)) : root;
                            ExportResult recovered=session.Exporter.Export(prepared,destination);
                            entry.Path=recovered.Path;entry.Bytes=recovered.Bytes;entry.Sha256=recovered.Sha256;
                        }
                        catch(Exception error){entry.Status=RecoveryErrors.Classify(error);entry.Message=Redact(error.Message,secret);}
                        summary.Entries.Add(entry);
                        if(entry.Status==RecoveryStatus.Success)summary.Success++;
                        else if(entry.Status==RecoveryStatus.Unsupported || entry.Status==RecoveryStatus.Unavailable)summary.Skipped++;
                        else summary.Failed++;
                        WriteCsv(report,session.Catalog.SourceName,entry.Document.SiteId.ToString("D"),entry.Document.Id.ToString("D"),entry.Document.Path,
                            entry.Status.ToString(),entry.Bytes.ToString(CultureInfo.InvariantCulture),entry.Sha256,entry.Path,entry.Message,
                            entry.Document.AttachmentOwnerId.Value.ToString("D"),entry.Document.ListItemId.Value.ToString(CultureInfo.InvariantCulture),entry.Document.ItemUniqueId.Value.ToString("D"));
                        report.Flush();
                        if(progress!=null)progress(new DesktopExportProgress {Completed=summary.Entries.Count,Total=summary.Total,Document=entry.Document,
                            Status=entry.Status,Message=entry.Message,ExportPath=entry.Path});
                    }
                }
                return summary;
            }
        }
        public DesktopExportSummary ExportMigrationPackage(Node list,string directory,CancellationToken cancellationToken,Action<DesktopExportProgress> progress)
        {
            if(list==null) throw new ArgumentNullException("list");
            if((list.Kind!=NodeKind.List && list.Kind!=NodeKind.Library) || list.HistoryVersion!=0 || list.SiteId==Guid.Empty ||
                list.WebId==Guid.Empty || list.ListId==Guid.Empty || list.Id==Guid.Empty)
                throw new ArgumentException("Select a current list or document library with its complete source scope.","list");
            if(String.IsNullOrWhiteSpace(directory)) throw new ArgumentException("Choose an export folder.","directory");
            Node selection=SnapshotFile(list);
            lock(gate)
            {
                EnsureOpen();
                var summary=new DesktopExportSummary {Directory=Path.GetFullPath(directory),Total=1,PackagePath="",ReportPath=""};
                if(cancellationToken.IsCancellationRequested){summary.Cancelled=true;return summary;}
                try
                {
                    var result=new MigrationPackageExporter().Export(session,selection,summary.Directory,cancellationToken,item=>{
                        if(progress!=null)progress(new DesktopExportProgress {Completed=item.Completed,Total=item.Total,Document=item.Document,
                            Status=item.Status,Message=Redact(item.Message,secret),ExportPath=item.ExportPath});
                    });
                    summary.Cancelled=result.Cancelled;
                    if(String.IsNullOrEmpty(result.PackagePath))
                    {
                        if(!result.Cancelled)throw new InvalidDataException("The migration export did not publish a package.");
                        return summary;
                    }
                    if(result.Cancelled)throw new InvalidDataException("A cancelled migration export returned a published package.");
                    summary.PackagePath=result.PackagePath;summary.ReportPath=result.ReportPath;
                    summary.PackageItemCount=result.ItemCount;summary.PackageFileCount=result.FileCount;summary.PackageAttachmentCount=result.AttachmentCount;
                    summary.Success=1;
                    summary.Entries.Add(new DesktopExportEntry {Document=selection,Status=RecoveryStatus.Success,Path=result.PackagePath,Bytes=result.Bytes,Sha256="",Message=""});
                    if(progress!=null)progress(new DesktopExportProgress {Completed=1,Total=1,Document=selection,Status=RecoveryStatus.Success,ExportPath=result.PackagePath,Message=""});
                    return summary;
                }
                catch(Exception error){ThrowIfContainsSecret(error);throw;}
            }
        }
        private static void ValidateAttachmentScope(Node scope)
        {
            if(scope==null)throw new ArgumentNullException("scope");
            if((scope.Kind!=NodeKind.List && scope.Kind!=NodeKind.ListItem) || scope.HistoryVersion!=0 || scope.SiteId==Guid.Empty ||
                scope.WebId==Guid.Empty || scope.ListId==Guid.Empty || scope.Id==Guid.Empty || scope.ListBaseType==1 ||
                (scope.Kind==NodeKind.ListItem && (!scope.ListItemId.HasValue || scope.ListItemId.Value<=0 || !scope.ItemUniqueId.HasValue || scope.ItemUniqueId.Value==Guid.Empty)))
                throw new ArgumentException("Select a current ordinary list or item with its complete source identity.","scope");
        }
        private static void ValidateAttachmentFile(Node file,Node scope)
        {
            if(file==null || file.Kind!=NodeKind.File || file.HistoryVersion!=0 || file.Id==Guid.Empty || file.SiteId!=scope.SiteId ||
                file.WebId!=scope.WebId || file.ListId!=scope.ListId || !file.AttachmentOwnerId.HasValue || file.AttachmentOwnerId.Value==Guid.Empty ||
                file.AttachmentOwnerId.Value==file.Id || !file.ListItemId.HasValue || file.ListItemId.Value<=0 || !file.ItemUniqueId.HasValue ||
                file.ItemUniqueId.Value==Guid.Empty || (scope.Kind==NodeKind.ListItem && (file.AttachmentOwnerId!=scope.Id ||
                    file.ListItemId!=scope.ListItemId || file.ItemUniqueId!=scope.ItemUniqueId)))
                throw new InvalidDataException("The source returned an attachment outside the selected list item scope.");
        }
        private static Node AttachmentOwner(Node file)
        {
            return new Node {Kind=NodeKind.ListItem,SiteId=file.SiteId,WebId=file.WebId,ListId=file.ListId,Id=file.AttachmentOwnerId.Value,
                ListItemId=file.ListItemId,ItemUniqueId=file.ItemUniqueId,ListBaseType=file.ListBaseType,HistoryVersion=0};
        }
    }
}