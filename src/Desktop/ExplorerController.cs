using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

namespace SharePointExplorer.Desktop
{
    public interface IExplorerController : IDisposable
    {
        string SourceName { get; }
        List<Node> GetRootSites();
        List<Node> GetChildren(Node parent);
        List<Node> GetFileVersions(Node document);
        DesktopExportSummary ExportVersion(Node version, string directory,
            CancellationToken cancellationToken, Action<DesktopExportProgress> progress);
        DesktopExportSummary ExportLibrary(Node library, string directory,
            CancellationToken cancellationToken, Action<DesktopExportProgress> progress);
        DesktopExportSummary ExportFiles(IEnumerable<Node> documents, string directory,
            CancellationToken cancellationToken, Action<DesktopExportProgress> progress);
        DesktopExportSummary ExportFilesAsZip(IEnumerable<Node> documents, string archivePath,
            CancellationToken cancellationToken, Action<DesktopExportProgress> progress);
    }

    public sealed class DesktopExportProgress
    {
        public int Completed, Total;
        public Node Document;
        public RecoveryStatus Status;
        public string Message, ExportPath;
    }

    public sealed class DesktopExportEntry
    {
        public Node Document;
        public RecoveryStatus Status;
        public string Path, Sha256, Message;
        public long Bytes;
    }

    public sealed class DesktopExportSummary
    {
        public int Total, Success, Failed, Skipped;
        public bool Cancelled;
        public string Directory, ReportPath;
        public string ArchivePath, ArchiveReportEntry;
        public string PackagePath;
        public int PackageItemCount,PackageFileCount,PackageAttachmentCount;
        public List<DesktopExportEntry> Entries = new List<DesktopExportEntry>();
    }

    // The desktop coordinates selections and progress; recovery stays in the core.
    public sealed partial class ExplorerController : IExplorerController
    {
        private readonly object gate = new object();
        private readonly RecoverySession session;
        private bool disposed;
        private string secret = String.Empty;

        public ExplorerController(RecoverySession session)
        {
            if (session == null) throw new ArgumentNullException("session");
            this.session = session;
        }

        public static ExplorerController Connect(SqlConnectionOptions options)
        {
            if (options == null) throw new ArgumentNullException("options");
            SqlConnectionOptions snapshot = options.Clone();
            try
            {
                return new ExplorerController(RecoverySession.OpenSql(snapshot)) { secret = snapshot.Password ?? String.Empty };
            }
            catch (Exception error)
            {
                string message = Redact(error.Message, snapshot.Password);
                throw new InvalidOperationException(message);
            }
        }

        public string SourceName
        {
            get { lock (gate) { EnsureOpen(); return session.Catalog.SourceName; } }
        }

        public List<Node> GetRootSites()
        {
            lock (gate)
            {
                EnsureOpen();
                try { return session.Catalog.GetRootSites(); }
                catch (Exception error) { ThrowIfContainsSecret(error); throw; }
            }
        }

        public List<Node> GetChildren(Node parent)
        {
            if (parent == null) throw new ArgumentNullException("parent");
            lock (gate)
            {
                EnsureOpen();
                try { return GetRecoveryChildren(parent); }
                catch (Exception error) { ThrowIfContainsSecret(error); throw; }
            }
        }


        public List<Node> GetFileVersions(Node document)
        {
            if(document==null) throw new ArgumentNullException("document");
            if(document.IsDeleted) return GetDeletedVersions(document);
            if(document.Kind!=NodeKind.File || document.HistoryVersion!=0 || document.SiteId==Guid.Empty || document.Id==Guid.Empty)
                throw new ArgumentException("Select a current document to view its versions.","document");
            Node requested=SnapshotFile(document);
            lock(gate)
            {
                EnsureOpen();
                ISharePointVersionCatalog catalog=session.Catalog as ISharePointVersionCatalog;
                if(catalog==null) throw new NotSupportedException("This source cannot enumerate document versions.");
                try
                {
                    List<Node> rows=catalog.GetFileVersions(requested.SiteId,requested.Id);
                    if(rows==null) throw new InvalidDataException("The source returned no document version collection.");
                    var identities=new Dictionary<string,Node>(StringComparer.Ordinal);
                    foreach(Node row in rows)
                    {
                        if(row==null || row.Kind!=NodeKind.File || row.SiteId!=requested.SiteId || row.Id!=requested.Id ||
                            row.HistoryVersion<0 || row.UiVersion<0 || row.InternalVersion<0 || row.Size<0 ||
                            (row.HistoryVersion>0 && row.HistoryVersion!=row.UiVersion) ||
                            (requested.WebId!=Guid.Empty && row.WebId!=requested.WebId) || (requested.ListId!=Guid.Empty && row.ListId!=requested.ListId))
                            throw new InvalidDataException("The source returned a different document version scope.");
                        Node copy=SnapshotFile(row);string identity=VersionIdentity(copy);Node existing;
                        if(identities.TryGetValue(identity,out existing))
                        {
                            if(!SameVersionMetadata(existing,copy)) throw new InvalidDataException("The source returned conflicting metadata for one document version.");
                        }
                        else identities.Add(identity,copy);
                    }
                    var result=new List<Node>(identities.Values);
                    result.Sort(delegate(Node first,Node second) {
                        int state=(first.HistoryVersion==0?0:1).CompareTo(second.HistoryVersion==0?0:1);
                        if(state!=0) return state;
                        int ui=second.UiVersion.CompareTo(first.UiVersion);if(ui!=0) return ui;
                        int history=second.HistoryVersion.CompareTo(first.HistoryVersion);if(history!=0) return history;
                        int level=second.Level.CompareTo(first.Level);return level!=0?level:second.InternalVersion.CompareTo(first.InternalVersion);
                    });
                    return result;
                }
                catch(Exception error) { ThrowIfContainsSecret(error);throw; }
            }
        }

        public DesktopExportSummary ExportVersion(Node version,string directory,
            CancellationToken cancellationToken,Action<DesktopExportProgress> progress)
        {
            if(version==null) throw new ArgumentNullException("version");
            if(version.Kind!=NodeKind.File || version.HistoryVersion<0 || version.UiVersion<0 || version.InternalVersion<0 || (version.HistoryVersion>0 && version.HistoryVersion!=version.UiVersion) || version.SiteId==Guid.Empty || version.Id==Guid.Empty)
                throw new ArgumentException("Select a document version to export.","version");
            if(String.IsNullOrWhiteSpace(directory)) throw new ArgumentException("Choose an export folder.","directory");
            Node selected=SnapshotFile(version);
            lock(gate)
            {
                EnsureOpen();string root=Path.GetFullPath(directory);
                var summary=new DesktopExportSummary {Total=session.Engine.CanExport(selected)?1:0,Directory=root};
                if(summary.Total==0) return summary;
                if(cancellationToken.IsCancellationRequested) {summary.Cancelled=true;return summary;}
                System.IO.Directory.CreateDirectory(root);
                summary.ReportPath=Path.Combine(root,"version-export-report-"+DateTime.UtcNow.ToString("yyyyMMdd-HHmmss",CultureInfo.InvariantCulture)+"-"+Guid.NewGuid().ToString("N").Substring(0,8)+".csv");
                var entry=new DesktopExportEntry {Document=selected,Status=RecoveryStatus.Success,Message="",Path="",Sha256=""};
                using(var report=new StreamWriter(new FileStream(summary.ReportPath,FileMode.CreateNew,FileAccess.Write,FileShare.Read),new UTF8Encoding(true)))
                {
                    if(selected.IsDeleted)WriteVersionHeader(report);
                    else WriteCsv(report,"Source","SiteId","DocumentId","DocumentPath","Status","Bytes","Sha256","ExportPath","Message","UiVersion","HistoryVersion","Level","InternalVersion");
                    try
                    {
                        PreparedDocument prepared=selected.IsDeleted ? session.Engine.PrepareDeleted(selected) : session.Engine.PrepareVersion(selected);
                        entry.Document=prepared.Document;
                        ExportResult result=session.Exporter.Export(prepared,root,VersionFileName(entry.Document));
                        entry.Path=result.Path;entry.Bytes=result.Bytes;entry.Sha256=result.Sha256;summary.Success=1;
                    }
                    catch(Exception error)
                    {
                        entry.Status=RecoveryErrors.Classify(error);entry.Message=Redact(error.Message,secret);
                        if(entry.Status==RecoveryStatus.Unsupported || entry.Status==RecoveryStatus.Unavailable) summary.Skipped=1;else summary.Failed=1;
                    }
                    summary.Entries.Add(entry);
                    if(selected.IsDeleted)WriteVersionRow(report,entry);
                    else WriteCsv(report,session.Catalog.SourceName,entry.Document.SiteId.ToString("D"),entry.Document.Id.ToString("D"),entry.Document.Path,
                        entry.Status.ToString(),entry.Bytes.ToString(CultureInfo.InvariantCulture),entry.Sha256,entry.Path,entry.Message,
                        entry.Document.UiVersion.ToString(CultureInfo.InvariantCulture),entry.Document.HistoryVersion.ToString(CultureInfo.InvariantCulture),
                        entry.Document.Level.ToString(CultureInfo.InvariantCulture),entry.Document.InternalVersion.ToString(CultureInfo.InvariantCulture));
                }
                if(progress!=null) progress(new DesktopExportProgress {Completed=1,Total=1,Document=entry.Document,Status=entry.Status,Message=entry.Message,ExportPath=entry.Path});
                return summary;
            }
        }

        private static string VersionFileName(Node version)
        {
            string name=DocumentExporter.SafeFileName(version.Name);
            string label=(version.UiVersion/512).ToString(CultureInfo.InvariantCulture)+"."+(version.UiVersion%512).ToString(CultureInfo.InvariantCulture);
            return Path.GetFileNameWithoutExtension(name)+" (v"+label+")"+Path.GetExtension(name);
        }
        private static string VersionIdentity(Node version)
        {
            return version.SiteId.ToString("N")+":"+version.Id.ToString("N")+":"+version.WebId.ToString("N")+":"+version.ListId.ToString("N")+":"+
                version.HistoryVersion+":"+version.UiVersion+":"+version.Level+":"+version.InternalVersion+":"+version.DeletionTransactionId;
        }
        private static bool SameVersionMetadata(Node first,Node second)
        {
            return first.Name==second.Name && first.Path==second.Path && first.Size==second.Size && first.StreamSchema==second.StreamSchema &&
                first.HasStream==second.HasStream && first.Modified==second.Modified && first.ParentId==second.ParentId;
        }
        public DesktopExportSummary ExportFiles(IEnumerable<Node> documents, string directory,
            CancellationToken cancellationToken, Action<DesktopExportProgress> progress)
        {
            if (documents == null) throw new ArgumentNullException("documents");
            if (String.IsNullOrWhiteSpace(directory)) throw new ArgumentException("Choose an export folder.", "directory");
            var selected = new List<Node>();
            var identities = new HashSet<string>(StringComparer.Ordinal);
            foreach (Node document in documents)
            {
                if (document == null || document.Kind != NodeKind.File)
                    throw new ArgumentException("Only document files can be exported.", "documents");
                string identity = document.SiteId.ToString("N") + ":" + document.Id.ToString("N") + ":" + document.DeletionTransactionId;
                if (identities.Add(identity)) selected.Add(SnapshotFile(document));
            }
            if (selected.Count == 0) throw new ArgumentException("Select at least one file.", "documents");
            lock (gate)
            {
                EnsureOpen();
                selected.RemoveAll(document=>!session.Engine.CanExport(document));
                return ExportSelected(selected,directory,cancellationToken,progress,null,false);
            }
        }

        public DesktopExportSummary ExportLibrary(Node library, string directory,
            CancellationToken cancellationToken, Action<DesktopExportProgress> progress)
        {
            if(library == null) throw new ArgumentNullException("library");
            if(String.IsNullOrWhiteSpace(directory)) throw new ArgumentException("Choose an export folder.","directory");
            if(library.Kind != NodeKind.Library || library.HistoryVersion != 0 || library.SiteId == Guid.Empty ||
                library.WebId == Guid.Empty || library.ListId == Guid.Empty || library.Id == Guid.Empty ||
                (library.ListBaseType.HasValue && library.ListBaseType.Value != 1))
                throw new ArgumentException("Select a current document library with its complete source scope.","library");
            var scope = new Node { Kind=library.Kind,SiteId=library.SiteId,WebId=library.WebId,ListId=library.ListId,
                Id=library.Id,HistoryVersion=library.HistoryVersion,Name=library.Name,Path=library.Path,ListBaseType=library.ListBaseType };
            lock(gate)
            {
                EnsureOpen();
                string root=Path.GetFullPath(directory);
                var selected=new List<Node>();
                if(cancellationToken.IsCancellationRequested)
                    return new DesktopExportSummary { Directory=root,Cancelled=true };
                ISharePointLibraryCatalog catalog=session.Catalog as ISharePointLibraryCatalog;
                if(catalog == null) throw new NotSupportedException("This source cannot enumerate a complete document library.");
                try
                {
                    IEnumerable<Node> documents=catalog.EnumerateCurrentLibraryFiles(new Node { Kind=scope.Kind,SiteId=scope.SiteId,
                        WebId=scope.WebId,ListId=scope.ListId,Id=scope.Id,Name=scope.Name,Path=scope.Path });
                    if(documents == null) throw new InvalidDataException("The source returned no library file collection.");
                    var identities=new HashSet<string>(StringComparer.Ordinal);
                    using(IEnumerator<Node> files=documents.GetEnumerator())
                        while(!cancellationToken.IsCancellationRequested && files.MoveNext())
                        {
                            Node file=files.Current;
                            ValidateLibraryFile(file,scope);
                            if(session.Engine.CanExport(file) && identities.Add(file.Id.ToString("N"))) selected.Add(SnapshotFile(file));
                        }
                }
                catch(Exception error) { ThrowIfContainsSecret(error); throw; }
                if(cancellationToken.IsCancellationRequested)
                    return new DesktopExportSummary { Directory=root,Total=selected.Count,Cancelled=true };
                return ExportSelected(selected,root,cancellationToken,progress,scope,true);
            }
        }


        public DesktopExportSummary ExportFilesAsZip(IEnumerable<Node> documents,string archivePath,
            CancellationToken cancellationToken,Action<DesktopExportProgress> progress)
        {
            if(documents==null) throw new ArgumentNullException("documents");
            if(String.IsNullOrWhiteSpace(archivePath)) throw new ArgumentException("Choose a ZIP filename.","archivePath");
            archivePath=Path.GetFullPath(archivePath);
            if(!Path.GetExtension(archivePath).Equals(".zip",StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Choose a filename ending in .zip.","archivePath");
            var selected=new List<Node>();
            var identities=new HashSet<string>(StringComparer.Ordinal);
            foreach(Node document in documents)
            {
                if(document==null || document.Kind!=NodeKind.File)
                    throw new ArgumentException("Only document files can be exported.","documents");
                if(identities.Add(document.SiteId.ToString("N")+":"+document.Id.ToString("N")+":"+document.DeletionTransactionId)) selected.Add(SnapshotFile(document));
            }
            if(selected.Count==0) throw new ArgumentException("Select at least one file.","documents");
            lock(gate)
            {
                EnsureOpen();
                selected.RemoveAll(document=>!session.Engine.CanExport(document));
                var summary=new DesktopExportSummary { Total=selected.Count,Directory=Path.GetDirectoryName(archivePath),
                    ArchivePath="",ArchiveReportEntry="",ReportPath="" };
                if(cancellationToken.IsCancellationRequested) { summary.Cancelled=true; return summary; }
                using(var archive=new ValidatedZipArchive(archivePath))
                {
                    if(progress!=null) progress(new DesktopExportProgress { Completed=0,Total=summary.Total,Document=null,Message="" });
                    foreach(Node document in selected)
                    {
                        if(cancellationToken.IsCancellationRequested) { summary.Cancelled=true; break; }
                        var entry=new DesktopExportEntry { Document=document,Status=RecoveryStatus.Success,Message="",Path="",Sha256="" };
                        try
                        {
                            PreparedDocument prepared=document.IsDeleted ? session.Engine.PrepareDeleted(document) : session.Engine.Prepare(document);
                            entry.Document=prepared.Document;
                            ExportResult result=archive.Add(prepared);
                            entry.Path=result.Path; entry.Bytes=result.Bytes; entry.Sha256=result.Sha256;
                        }
                        catch(Exception error)
                        {
                            // Entry creation/copy errors can damage the ZIP. Recovery
                            // failures occur before entry creation and can safely continue.
                            if(archive.IsFaulted) { ThrowIfContainsSecret(error); throw; }
                            entry.Status=RecoveryErrors.Classify(error);
                            entry.Message=Redact(error.Message,secret);
                        }
                        summary.Entries.Add(entry);
                        if(entry.Status==RecoveryStatus.Success) summary.Success++;
                        else if(entry.Status==RecoveryStatus.Unsupported || entry.Status==RecoveryStatus.Unavailable) summary.Skipped++;
                        else summary.Failed++;
                        if(progress!=null) progress(new DesktopExportProgress { Completed=summary.Entries.Count,Total=summary.Total,
                            Document=entry.Document,Status=entry.Status,Message=entry.Message,ExportPath=entry.Path });
                    }
                    if(summary.Cancelled && summary.Success==0) return summary;
                    string state="Cancelled="+summary.Cancelled+"\nTotal="+summary.Total+"\nProcessed="+summary.Entries.Count+
                        "\nExported="+summary.Success+"\nSkipped="+summary.Skipped+"\nFailed="+summary.Failed+"\n";
                    summary.ArchivePath=archive.Commit(report=>{
                        WriteCsv(report,"Source","SiteId","DocumentId","DocumentPath","Status","Bytes","Sha256","ExportPath","Message");
                        foreach(DesktopExportEntry entry in summary.Entries)
                            WriteCsv(report,session.Catalog.SourceName,entry.Document.SiteId.ToString("D"),entry.Document.Id.ToString("D"),
                                entry.Document.Path,entry.Status.ToString(),entry.Bytes.ToString(CultureInfo.InvariantCulture),
                                entry.Sha256,entry.Path,entry.Message);
                    },state);
                    summary.ArchiveReportEntry=ValidatedZipArchive.ReportEntry;
                }
                return summary;
            }
        }

        private DesktopExportSummary ExportSelected(List<Node> selected,string directory,
            CancellationToken cancellationToken,Action<DesktopExportProgress> progress,Node libraryScope,bool announceTotal)
        {
            string root=Path.GetFullPath(directory);
            var summary=new DesktopExportSummary { Total=selected.Count,Directory=root };
            if(cancellationToken.IsCancellationRequested) { summary.Cancelled=true; return summary; }
            System.IO.Directory.CreateDirectory(root);
            summary.ReportPath=Path.Combine(root,"export-report-"+DateTime.UtcNow.ToString("yyyyMMdd-HHmmss",CultureInfo.InvariantCulture)
                +"-"+Guid.NewGuid().ToString("N").Substring(0,8)+".csv");
            using(var report=new StreamWriter(new FileStream(summary.ReportPath,FileMode.CreateNew,FileAccess.Write,FileShare.Read),new UTF8Encoding(true)))
            {
                WriteCsv(report,"Source","SiteId","DocumentId","DocumentPath","Status","Bytes","Sha256","ExportPath","Message");
                if(announceTotal && progress != null) progress(new DesktopExportProgress { Completed=0,Total=summary.Total,Document=null,Message="" });
                foreach(Node document in selected)
                {
                    if(cancellationToken.IsCancellationRequested) { summary.Cancelled=true; break; }
                    var entry=new DesktopExportEntry { Document=document,Status=RecoveryStatus.Success,Message="",Path="",Sha256="" };
                    try
                    {
                        if(document.HistoryVersion != 0)
                            throw new NotSupportedException("Use the version recovery operation for historical file versions.");
                        ExportResult result;
                        if(libraryScope != null)
                        {
                            // Validate the final recovery metadata, then publish exactly that snapshot.
                            PreparedDocument prepared=session.Engine.PrepareLibraryFile(document,libraryScope);
                            Node current=prepared.Document;
                            ValidateLibraryFile(current,libraryScope);
                            entry.Document=current;
                            result=session.Exporter.Export(prepared,ExportDirectory(root,current));
                        }
                        else
                        {
                            PreparedDocument prepared=document.IsDeleted ? session.Engine.PrepareDeleted(document) : session.Engine.Prepare(document);
                            entry.Document=prepared.Document;
                            result=session.Exporter.Export(prepared,root);
                        }
                        entry.Path=result.Path; entry.Bytes=result.Bytes; entry.Sha256=result.Sha256;
                    }
                    catch(Exception error)
                    {
                        entry.Status=RecoveryErrors.Classify(error);
                        entry.Message=Redact(error.Message,secret);
                    }
                    summary.Entries.Add(entry);
                    if(entry.Status == RecoveryStatus.Success) summary.Success++;
                    else if(entry.Status == RecoveryStatus.Unsupported || entry.Status == RecoveryStatus.Unavailable) summary.Skipped++;
                    else summary.Failed++;
                    WriteCsv(report,session.Catalog.SourceName,entry.Document.SiteId.ToString("D"),entry.Document.Id.ToString("D"),entry.Document.Path,
                        entry.Status.ToString(),entry.Bytes.ToString(CultureInfo.InvariantCulture),entry.Sha256,entry.Path,entry.Message);
                    report.Flush();
                    if(progress != null) progress(new DesktopExportProgress {
                        Completed=summary.Entries.Count,Total=summary.Total,Document=entry.Document,
                        Status=entry.Status,Message=entry.Message,ExportPath=entry.Path });
                }
            }
            return summary;
        }

        private static void ValidateLibraryFile(Node file,Node library)
        {
            if(file == null || file.Kind != NodeKind.File || file.HistoryVersion != 0 || file.Id == Guid.Empty ||
                file.SiteId != library.SiteId || file.WebId != library.WebId || file.ListId != library.ListId)
                throw new InvalidDataException("The source returned a file outside the selected document library.");
        }

        private static Node SnapshotFile(Node file)
        {
            return new Node { Kind=file.Kind,SiteId=file.SiteId,WebId=file.WebId,ListId=file.ListId,Id=file.Id,
                Name=file.Name,Path=file.Path,Size=file.Size,Modified=file.Modified,StreamSchema=file.StreamSchema,
                Level=file.Level,InternalVersion=file.InternalVersion,UiVersion=file.UiVersion,
                HistoryVersion=file.HistoryVersion,ParentId=file.ParentId,HasStream=file.HasStream,ListBaseType=file.ListBaseType,
                ListItemId=file.ListItemId,ItemUniqueId=file.ItemUniqueId,AttachmentOwnerId=file.AttachmentOwnerId,
                Created=file.Created,Title=file.Title,HasAttachments=file.HasAttachments,
                DeletionTransactionId=file.DeletionTransactionId,DeletedAt=file.DeletedAt };
        }

        internal static string ExportDirectory(string root, Node document)
        {
            string result = Path.Combine(root, document.SiteId.ToString("D"));
            string[] parts = (document.Path ?? document.Name ?? "").Split(new char[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
            for (int index = 0; index < parts.Length - 1; index++)
                result = Path.Combine(result, DocumentExporter.SafeFileName(parts[index]));
            return result;
        }

        private static void WriteCsv(TextWriter output, params string[] values)
        {
            for (int index = 0; index < values.Length; index++)
            {
                if (index > 0) output.Write(',');
                output.Write('"');
                output.Write((values[index] ?? "").Replace("\"", "\"\""));
                output.Write('"');
            }
            output.WriteLine();
        }

        public void Dispose()
        {
            lock (gate)
            {
                if (disposed) return;
                disposed = true;
                session.Dispose();
                secret = String.Empty;
            }
        }

        private void ThrowIfContainsSecret(Exception error)
        {
            if (!String.IsNullOrEmpty(secret) && error.Message.IndexOf(secret, StringComparison.Ordinal) >= 0)
                throw new InvalidOperationException(Redact(error.Message, secret));
        }

        private static string Redact(string message, string password)
        {
            return String.IsNullOrEmpty(password) ? message : (message ?? "").Replace(password, "[redacted]");
        }

        private void EnsureOpen() { if (disposed) throw new ObjectDisposedException("ExplorerController"); }
    }
}




