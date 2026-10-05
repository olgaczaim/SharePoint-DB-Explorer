using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
namespace SharePointExplorer.Desktop
{
    public sealed partial class ExplorerController : IExplorerBulkExportController
    {
        public DesktopExportSummary ExportScopeAsZip(Node scope,string archivePath,CancellationToken cancellationToken,Action<DesktopExportProgress> progress)
        {
            if(scope==null)throw new ArgumentNullException("scope");
            if(scope.IsDeleted || (scope.Kind!=NodeKind.Library && scope.Kind!=NodeKind.Folder) || scope.SiteId==Guid.Empty || scope.WebId==Guid.Empty || scope.ListId==Guid.Empty || scope.Id==Guid.Empty || scope.HistoryVersion!=0 || (scope.ListBaseType.HasValue && scope.ListBaseType!=1))
                throw new ArgumentException("Choose a current document library or document folder.","scope");
            if(String.IsNullOrWhiteSpace(archivePath))throw new ArgumentException("Choose a ZIP filename.","archivePath");
            Node selection=SnapshotFile(scope);
            lock(gate)
            {
                EnsureOpen();
                var result=new DesktopExportSummary {Directory=Path.GetDirectoryName(Path.GetFullPath(archivePath)),ArchivePath="",ArchiveReportEntry="",ReportPath=""};
                if(cancellationToken.IsCancellationRequested){result.Cancelled=true;return result;}
                List<Node> files=DiscoverScopeFiles(selection,cancellationToken);
                result.Total=files.Count;
                if(cancellationToken.IsCancellationRequested){result.Cancelled=true;return result;}
                using(var zip=new ValidatedZipArchive(archivePath))
                {
                    if(progress!=null)progress(new DesktopExportProgress {Completed=0,Total=result.Total,Message=""});
                    foreach(Node file in files)
                    {
                        if(cancellationToken.IsCancellationRequested){result.Cancelled=true;break;}
                        var entry=new DesktopExportEntry {Document=file,Status=RecoveryStatus.Success,Message="",Path="",Sha256=""};
                        try
                        {
                            PreparedDocument prepared=selection.Kind==NodeKind.Library ? session.Engine.PrepareLibraryFile(file,selection) : session.Engine.Prepare(file);
                            ValidateScopeDescendant(prepared.Document,selection);
                            if(prepared.Document.ParentId!=file.ParentId || prepared.Document.Path!=file.Path)throw new InvalidDataException("A document moved after folder discovery. Refresh and export again.");
                            entry.Document=prepared.Document;
                            ExportResult recovered=zip.Add(prepared);entry.Path=recovered.Path;entry.Bytes=recovered.Bytes;entry.Sha256=recovered.Sha256;
                        }
                        catch(Exception error)
                        {
                            if(zip.IsFaulted){ThrowIfContainsSecret(error);throw;}
                            entry.Status=RecoveryErrors.Classify(error);entry.Message=Redact(error.Message,secret);
                        }
                        AddBulkEntry(result,entry,progress);
                    }
                    if(result.Cancelled && result.Success==0)return result;
                    result.ArchivePath=zip.Commit(report=>WriteBulkReport(report,result,false),"Cancelled="+result.Cancelled+"\nTotal="+result.Total+"\nExported="+result.Success+"\nFailed="+result.Failed+"\n");
                    result.ArchiveReportEntry=ValidatedZipArchive.ReportEntry;
                }
                return result;
            }
        }
        private List<Node> DiscoverScopeFiles(Node scope,CancellationToken token)
        {
            var files=new Dictionary<Guid,Node>();
            if(scope.Kind==NodeKind.Library)
            {
                ISharePointLibraryCatalog catalog=session.Catalog as ISharePointLibraryCatalog;
                if(catalog==null)throw new NotSupportedException("This source cannot enumerate a complete library.");
                foreach(Node file in catalog.EnumerateCurrentLibraryFiles(SnapshotFile(scope)))
                {
                    if(token.IsCancellationRequested)break;
                    ValidateScopeDescendant(file,scope);AddDiscovered(files,file);
                }
            }
            else
            {
                var pending=new Queue<Node>();pending.Enqueue(scope);
                var visited=new HashSet<Guid>{scope.Id};
                while(pending.Count>0 && !token.IsCancellationRequested)
                {
                    Node parent=pending.Dequeue();List<Node> children=session.Catalog.GetChildren(SnapshotFile(parent));
                    if(children==null)throw new InvalidDataException("The source returned no folder contents.");
                    foreach(Node child in children)
                    {
                        if(token.IsCancellationRequested)break;
                        ValidateScopeDescendant(child,scope);
                        if(child.ParentId!=parent.Id)throw new InvalidDataException("Folder discovery returned an unrelated parent identity.");
                        if(child.Kind==NodeKind.Folder)
                        {
                            if(!visited.Add(child.Id))throw new InvalidDataException("Folder discovery contains cyclic or duplicate ancestry.");
                            pending.Enqueue(SnapshotFile(child));
                        }
                        else if(child.Kind==NodeKind.File)AddDiscovered(files,child);
                        else throw new InvalidDataException("The selected document folder returned an unsupported child kind.");
                    }
                }
            }
            return files.Values.Where(file=>session.Engine.CanExport(file)).OrderBy(file=>file.Path,StringComparer.Ordinal).ThenBy(file=>file.Id).ToList();
        }
        private static void AddDiscovered(Dictionary<Guid,Node> files,Node file)
        {
            Node previous;if(files.TryGetValue(file.Id,out previous))
            {if(previous.ParentId!=file.ParentId || !SameVersionMetadata(previous,file) || VersionIdentity(previous)!=VersionIdentity(file))throw new InvalidDataException("Discovery returned conflicting metadata for one document.");}
            else files.Add(file.Id,SnapshotFile(file));
        }
        private static void ValidateScopeDescendant(Node child,Node scope)
        {
            if(child==null || (child.Kind!=NodeKind.File && child.Kind!=NodeKind.Folder) || child.SiteId!=scope.SiteId || child.WebId!=scope.WebId || child.ListId!=scope.ListId || child.Id==Guid.Empty || child.HistoryVersion!=0)
                throw new InvalidDataException("The source returned content outside the selected library or folder.");
            string root=ScopePath(scope.Path),path=ScopePath(child.Path);
            if(root.Length==0 || !path.StartsWith(root+"/",StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("The source returned a document outside the selected folder path.");
        }
        private static string ScopePath(string path)
        {
            string normalized=(path??"").Replace('\\','/').Trim('/');
            if(normalized.Split('/').Any(part=>part=="." || part=="..") || normalized.Contains(":"))throw new InvalidDataException("The stored scope path is invalid.");
            return normalized;
        }
        public DesktopExportSummary ExportVersions(IEnumerable<Node> documentsOrVersions,string directory,CancellationToken cancellationToken,Action<DesktopExportProgress> progress)
        {
            if(documentsOrVersions==null)throw new ArgumentNullException("documentsOrVersions");
            if(String.IsNullOrWhiteSpace(directory))throw new ArgumentException("Choose an export folder.","directory");
            var input=documentsOrVersions.Select(document=>document==null?null:SnapshotFile(document)).ToList();
            if(input.Count==0 || input.Any(document=>document==null || document.Kind!=NodeKind.File || document.SiteId==Guid.Empty || document.Id==Guid.Empty || document.HistoryVersion<0))throw new ArgumentException("Choose documents or retained document versions.","documentsOrVersions");
            lock(gate)
            {
                EnsureOpen();var result=new DesktopExportSummary {Directory=Path.GetFullPath(directory)};
                if(cancellationToken.IsCancellationRequested){result.Cancelled=true;return result;}
                var selected=new Dictionary<string,Node>(StringComparer.Ordinal);
                foreach(Node document in input)
                {
                    if(cancellationToken.IsCancellationRequested){result.Cancelled=true;return result;}
                    IEnumerable<Node> versions=document.HistoryVersion==0 ? GetFileVersions(document) : new[]{document};
                    foreach(Node version in versions)
                    {
                        string key=VersionIdentity(version);Node previous;
                        if(selected.TryGetValue(key,out previous)){if(!SameVersionMetadata(previous,version))throw new InvalidDataException("Conflicting retained document-version metadata.");}
                        else if(session.Engine.CanExport(version))selected.Add(key,SnapshotFile(version));
                    }
                }
                List<Node> rows=selected.Values.OrderBy(row=>row.Path,StringComparer.Ordinal).ThenBy(row=>row.Id).ThenByDescending(row=>row.UiVersion).ThenByDescending(row=>row.Level).ToList();
                result.Total=rows.Count;Directory.CreateDirectory(result.Directory);
                result.ReportPath=Path.Combine(result.Directory,"version-export-report-"+DateTime.UtcNow.ToString("yyyyMMdd-HHmmss",CultureInfo.InvariantCulture)+"-"+Guid.NewGuid().ToString("N").Substring(0,8)+".csv");
                using(var report=new StreamWriter(new FileStream(result.ReportPath,FileMode.CreateNew,FileAccess.Write,FileShare.Read),new UTF8Encoding(true)))
                {
                    WriteVersionHeader(report);
                    if(progress!=null)progress(new DesktopExportProgress {Completed=0,Total=result.Total,Message=""});
                    foreach(Node row in rows)
                    {
                        if(cancellationToken.IsCancellationRequested){result.Cancelled=true;break;}
                        var entry=new DesktopExportEntry {Document=row,Status=RecoveryStatus.Success,Message="",Path="",Sha256=""};
                        try
                        {
                            PreparedDocument prepared=row.IsDeleted ? session.Engine.PrepareDeleted(row) : session.Engine.PrepareVersion(row);entry.Document=prepared.Document;
                            ExportResult recovered=session.Exporter.Export(prepared,ExportDirectory(result.Directory,prepared.Document),VersionFileName(prepared.Document));
                            entry.Path=recovered.Path;entry.Bytes=recovered.Bytes;entry.Sha256=recovered.Sha256;
                        }
                        catch(Exception error){entry.Status=RecoveryErrors.Classify(error);entry.Message=Redact(error.Message,secret);}
                        AddBulkEntry(result,entry,progress);WriteVersionRow(report,entry);report.Flush();
                    }
                }
                return result;
            }
        }
        private void AddBulkEntry(DesktopExportSummary result,DesktopExportEntry entry,Action<DesktopExportProgress> progress)
        {
            result.Entries.Add(entry);
            if(entry.Status==RecoveryStatus.Success)result.Success++;else if(entry.Status==RecoveryStatus.Unsupported || entry.Status==RecoveryStatus.Unavailable)result.Skipped++;else result.Failed++;
            if(progress!=null)progress(new DesktopExportProgress {Completed=result.Entries.Count,Total=result.Total,Document=entry.Document,Status=entry.Status,Message=entry.Message,ExportPath=entry.Path});
        }
        private void WriteBulkReport(TextWriter report,DesktopExportSummary summary,bool versions)
        {
            if(versions){WriteVersionHeader(report);foreach(DesktopExportEntry entry in summary.Entries)WriteVersionRow(report,entry);return;}
            WriteCsv(report,"Source","SiteId","DocumentId","DocumentPath","Status","Bytes","Sha256","ExportPath","Message");
            foreach(DesktopExportEntry entry in summary.Entries)WriteCsv(report,session.Catalog.SourceName,entry.Document.SiteId.ToString("D"),entry.Document.Id.ToString("D"),entry.Document.Path,entry.Status.ToString(),entry.Bytes.ToString(CultureInfo.InvariantCulture),entry.Sha256,entry.Path,entry.Message);
        }
        private static void WriteVersionHeader(TextWriter report){WriteCsv(report,"Source","SiteId","DocumentId","DocumentPath","Status","Bytes","Sha256","ExportPath","Message","UiVersion","HistoryVersion","Level","InternalVersion","DeletionTransactionId");}
        private void WriteVersionRow(TextWriter report,DesktopExportEntry entry)
        {WriteCsv(report,session.Catalog.SourceName,entry.Document.SiteId.ToString("D"),entry.Document.Id.ToString("D"),entry.Document.Path,entry.Status.ToString(),entry.Bytes.ToString(CultureInfo.InvariantCulture),entry.Sha256,entry.Path,entry.Message,entry.Document.UiVersion.ToString(CultureInfo.InvariantCulture),entry.Document.HistoryVersion.ToString(CultureInfo.InvariantCulture),entry.Document.Level.ToString(CultureInfo.InvariantCulture),entry.Document.InternalVersion.ToString(CultureInfo.InvariantCulture),entry.Document.DeletionTransactionId);}
        public DesktopExportSummary ExportMigrationPackage(Node list,string directory,bool includeHistory,CancellationToken cancellationToken,Action<DesktopExportProgress> progress)
        {
            if(!includeHistory)return ExportMigrationPackage(list,directory,cancellationToken,progress);
            if(list==null)throw new ArgumentNullException("list");
            lock(gate)
            {
                EnsureOpen();var summary=new DesktopExportSummary {Directory=Path.GetFullPath(directory),Total=1};
                var result=new MigrationPackageExporter {IncludeHistory=true}.Export(session,SnapshotFile(list),directory,cancellationToken,p=>{if(progress!=null)progress(new DesktopExportProgress {Completed=p.Completed,Total=p.Total,Document=p.Document,Status=p.Status,Message=Redact(p.Message,secret),ExportPath=p.ExportPath});});
                summary.Cancelled=result.Cancelled;summary.PackagePath=result.PackagePath;summary.ReportPath=result.ReportPath;summary.PackageItemCount=result.ItemCount;summary.PackageFileCount=result.FileCount;summary.PackageAttachmentCount=result.AttachmentCount;
                if(!String.IsNullOrEmpty(result.PackagePath)){summary.Success=1;summary.Entries.Add(new DesktopExportEntry {Document=SnapshotFile(list),Status=RecoveryStatus.Success,Path=result.PackagePath,Bytes=result.Bytes,Message=""});}
                return summary;
            }
        }
    }
}
