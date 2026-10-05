using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
namespace SharePointExplorer.Desktop
{
    public interface IExplorerDeletedExportController
    {
        DesktopExportSummary ExportDeletedFiles(IEnumerable<Node> files,string directory,CancellationToken token,Action<DesktopExportProgress> progress);
        DesktopExportSummary ExportDeletedItem(Node item,string directory,CancellationToken token,Action<DesktopExportProgress> progress);
    }
    public sealed partial class ExplorerController : IExplorerDeletedExportController
    {
        private static void ValidateDeletedSelection(Node item,bool fileOnly)
        {
            if(item==null || !item.IsDeleted || item.DeletionTransactionId.Length!=32 || item.SiteId==Guid.Empty || item.WebId==Guid.Empty || item.Id==Guid.Empty ||
                item.HistoryVersion<0 || item.UiVersion<0 || item.InternalVersion<0 || (item.HistoryVersion>0 && item.HistoryVersion!=item.UiVersion) || (fileOnly && item.Kind!=NodeKind.File))
                throw new ArgumentException("Select a retained deleted object with its exact source and version identity.");
            foreach(char value in item.DeletionTransactionId)if(!Uri.IsHexDigit(value))throw new ArgumentException("The deletion transaction identity is invalid.");
        }        private List<Node> GetRecoveryChildren(Node parent)
        {
            ISharePointDeletedCatalog deleted=session.Catalog as ISharePointDeletedCatalog;
            if(parent.Kind==NodeKind.DeletedItems)
            {
                if(deleted==null)throw new NotSupportedException("This source cannot browse retained deleted items.");
                Node site=SnapshotFile(parent);site.Kind=NodeKind.Site;return deleted.GetDeletedItems(site);
            }
            if(parent.IsDeleted)
            {
                if(deleted==null)throw new NotSupportedException("This source cannot browse deleted content.");
                return deleted.GetDeletedChildren(SnapshotFile(parent));
            }
            List<Node> result=session.Catalog.GetChildren(parent);
            if(parent.Kind==NodeKind.Site && deleted!=null)
                result.Add(new Node {Kind=NodeKind.DeletedItems,SiteId=parent.SiteId,WebId=parent.WebId,Id=parent.WebId,Name="Deleted items",Path=(parent.Path??"").TrimEnd('/')+"/Deleted items",HasStream=false});
            return result;
        }
        private List<Node> GetDeletedVersions(Node document)
        {
            lock(gate)
            {
                EnsureOpen();ISharePointDeletedCatalog catalog=session.Catalog as ISharePointDeletedCatalog;
                if(catalog==null)throw new NotSupportedException("This source cannot browse deleted document versions.");
                try{return catalog.GetDeletedFileVersions(SnapshotFile(document));}catch(Exception error){ThrowIfContainsSecret(error);throw;}
            }
        }
        public DesktopExportSummary ExportDeletedFiles(IEnumerable<Node> files,string directory,CancellationToken token,Action<DesktopExportProgress> progress)
        {
            if(files==null)throw new ArgumentNullException("files");
            if(String.IsNullOrWhiteSpace(directory))throw new ArgumentException("Choose an export folder.","directory");
            var selected=new List<Node>();var identities=new HashSet<string>(StringComparer.Ordinal);
            foreach(Node file in files)
            {
                if(file==null || file.Kind!=NodeKind.File || !file.IsDeleted)throw new ArgumentException("Select retained deleted document files.","files");
                ValidateDeletedSelection(file,true);
                if(identities.Add(VersionIdentity(file)+":"+file.DeletionTransactionId))selected.Add(SnapshotFile(file));
            }
            lock(gate)
            {
                EnsureOpen();var result=new DesktopExportSummary {Directory=Path.GetFullPath(directory)};
                selected.RemoveAll(file=>!session.Engine.CanExport(file));result.Total=selected.Count;
                if(token.IsCancellationRequested){result.Cancelled=true;return result;}
                Directory.CreateDirectory(result.Directory);result.ReportPath=Path.Combine(result.Directory,"deleted-export-report-"+Guid.NewGuid().ToString("N")+".csv");
                using(var report=new StreamWriter(new FileStream(result.ReportPath,FileMode.CreateNew,FileAccess.Write,FileShare.Read),new UTF8Encoding(true)))
                {
                    WriteCsv(report,"Source","SiteId","DocumentId","DocumentPath","Status","Bytes","Sha256","ExportPath","Message","DeletionTransactionId","UiVersion","HistoryVersion","Level","InternalVersion");
                    if(progress!=null)progress(new DesktopExportProgress {Total=result.Total,Message=""});
                    foreach(Node file in selected)
                    {
                        if(token.IsCancellationRequested){result.Cancelled=true;break;}
                        var entry=new DesktopExportEntry {Document=file,Status=RecoveryStatus.Success,Path="",Message="",Sha256=""};
                        try{PreparedDocument prepared=session.Engine.PrepareDeleted(file);entry.Document=prepared.Document;ExportResult recovered=session.Exporter.Export(prepared,ExportDirectory(result.Directory,prepared.Document));entry.Path=recovered.Path;entry.Bytes=recovered.Bytes;entry.Sha256=recovered.Sha256;}
                        catch(Exception error){entry.Status=RecoveryErrors.Classify(error);entry.Message=Redact(error.Message,secret);}
                        AddBulkEntry(result,entry,progress);
                        WriteCsv(report,session.Catalog.SourceName,entry.Document.SiteId.ToString("D"),entry.Document.Id.ToString("D"),entry.Document.Path,entry.Status.ToString(),entry.Bytes.ToString(CultureInfo.InvariantCulture),entry.Sha256,entry.Path,entry.Message,entry.Document.DeletionTransactionId,entry.Document.UiVersion.ToString(CultureInfo.InvariantCulture),entry.Document.HistoryVersion.ToString(CultureInfo.InvariantCulture),entry.Document.Level.ToString(CultureInfo.InvariantCulture),entry.Document.InternalVersion.ToString(CultureInfo.InvariantCulture));report.Flush();
                    }
                }
                return result;
            }
        }
        public DesktopExportSummary ExportDeletedItem(Node item,string directory,CancellationToken token,Action<DesktopExportProgress> progress)
        {
            if(item==null)throw new ArgumentNullException("item");ValidateDeletedSelection(item,false);
            if(item.Kind!=NodeKind.ListItem)throw new ArgumentException("Select a deleted ordinary list item.","item");
            if(String.IsNullOrWhiteSpace(directory))throw new ArgumentException("Choose an export folder.","directory");
            Node selection=SnapshotFile(item);
            lock(gate)
            {
                EnsureOpen();var result=new DesktopExportSummary {Directory=Path.GetFullPath(directory),Total=1};
                if(token.IsCancellationRequested){result.Cancelled=true;return result;}
                ISharePointDeletedCatalog catalog=session.Catalog as ISharePointDeletedCatalog;
                if(catalog==null)throw new NotSupportedException("This source cannot recover deleted list metadata.");
                var entry=new DesktopExportEntry {Document=selection,Status=RecoveryStatus.Success,Path="",Message="",Sha256=""};
                try
                {
                    DeletedListItemSnapshot snapshot=catalog.GetDeletedListItem(selection);
                    if(token.IsCancellationRequested){result.Cancelled=true;return result;}
                    ExportResult recovered=new DeletedItemExportService().Export(snapshot,result.Directory);entry.Path=recovered.Path;entry.Bytes=recovered.Bytes;entry.Sha256=recovered.Sha256;
                }
                catch(Exception error){entry.Status=RecoveryErrors.Classify(error);entry.Message=Redact(error.Message,secret);}
                AddBulkEntry(result,entry,progress);return result;
            }
        }
    }
}
