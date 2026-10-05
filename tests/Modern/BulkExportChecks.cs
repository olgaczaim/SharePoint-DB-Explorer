using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using SharePointExplorer.Desktop;
namespace SharePointExplorer.Tests
{
    internal static class BulkExportChecks
    {
        internal static void Run()
        {
            string root=Path.GetFullPath(Path.Combine(".scratch","bulk-export-checks",Guid.NewGuid().ToString("N")));
            var source=new Source();using(var controller=new ExplorerController(new RecoverySession(source,source,DocumentDecoderRegistry.CreateDefault())))
            {
                foreach(Node retained in new[]{source.Library,source.Nested}.Select(MigrationSnapshotCopy.Node))
                {
                    retained.DeletionTransactionId="01";
                    string rejectedPath=Path.Combine(root,"deleted-"+retained.Kind+".zip");
                    bool rejectedDeleted=false;try{controller.ExportScopeAsZip(retained,rejectedPath,CancellationToken.None,null);}catch(ArgumentException){rejectedDeleted=true;}
                    Check(rejectedDeleted && source.ChildrenReads==0 && !File.Exists(rejectedPath),"Current scope ZIP accepted a deleted library/folder or read its content.");
                }
                DesktopExportSummary library=controller.ExportScopeAsZip(source.Library,Path.Combine(root,"library.zip"),CancellationToken.None,null);
                Check(library.Success==3 && library.Total==3,"Whole-library ZIP missed nested content.");
                using(var zip=ZipFile.OpenRead(library.ArchivePath))Check(zip.Entries.Count(entry=>entry.FullName.StartsWith(source.Library.SiteId.ToString("D")+"/",StringComparison.Ordinal))==3 && zip.Entries.Any(entry=>entry.FullName.EndsWith("/Docs/nested/deep/third.txt")),"Library ZIP hierarchy changed.");
                DesktopExportSummary folder=controller.ExportScopeAsZip(source.Nested,Path.Combine(root,"folder.zip"),CancellationToken.None,null);
                Check(folder.Success==2 && folder.Total==2 && source.ChildrenReads==2,"Folder ZIP depended on visible descendants or included sibling files.");
                using(var zip=ZipFile.OpenRead(folder.ArchivePath))Check(!zip.Entries.Any(entry=>entry.Name=="first.txt") && zip.Entries.Any(entry=>entry.Name=="third.txt"),"Folder ZIP violated its exact scope.");
                DesktopExportSummary batch=controller.ExportVersions(new[]{source.First,source.Second,source.First},Path.Combine(root,"versions"),CancellationToken.None,null);
                Check(batch.Total==4 && batch.Success==4 && batch.Entries.Count(entry=>entry.Document.HistoryVersion>0)==2,"Batch version export omitted retained versions or duplicated a selection.");
                Check(batch.Entries.Where(entry=>entry.Document.HistoryVersion>0).All(entry=>File.ReadAllText(entry.Path)=="old "+entry.Document.Name),"Batch version recovery substituted current content.");
                Check(File.ReadAllLines(batch.ReportPath).Length==5 && File.ReadAllText(batch.ReportPath).Contains("HistoryVersion"),"Batch CSV omitted retained version identity.");
                DesktopExportSummary repeated=controller.ExportVersions(new[]{source.First},Path.Combine(root,"versions"),CancellationToken.None,null);
                Check(repeated.Success==2 && repeated.Entries.All(entry=>batch.Entries.All(previous=>previous.Path!=entry.Path)),"Repeated version batch overwrote existing recovery.");
                using(var cancel=new CancellationTokenSource())
                {
                    DesktopExportSummary partial=controller.ExportScopeAsZip(source.Library,Path.Combine(root,"cancelled.zip"),cancel.Token,p=>{if(p.Completed==1)cancel.Cancel();});
                    Check(partial.Cancelled && partial.Success==1,"Scope ZIP cancellation continued exporting.");using(var zip=ZipFile.OpenRead(partial.ArchivePath))Check(zip.GetEntry("summary.txt")!=null,"Cancellation published a broken ZIP.");
                }
                source.WrongParent=true;bool rejected=false;try{controller.ExportScopeAsZip(source.Nested,Path.Combine(root,"invalid.zip"),CancellationToken.None,null);}catch(InvalidDataException){rejected=true;}
                Check(rejected && !File.Exists(Path.Combine(root,"invalid.zip")),"Folder discovery accepted unrelated source ancestry.");
            }
            Console.WriteLine("PASS library/folder ZIP recursion, exact scopes, batch historical bytes, no-overwrite reports and cancellation");
        }
        private static void Check(bool value,string message){if(!value)throw new InvalidDataException(message);}
        private sealed class Source : ISharePointCatalog,ISharePointLibraryCatalog,ISharePointVersionCatalog,IDocumentChunkStore
        {
            internal readonly Node Library,Nested,Deep,First,Second,Third;internal int ChildrenReads;internal bool WrongParent;
            private readonly List<Node> files=new List<Node>();public string SourceName=>"Synthetic bulk fixture";
            internal Source()
            {
                Library=new Node {Kind=NodeKind.Library,SiteId=Guid.NewGuid(),WebId=Guid.NewGuid(),ListId=Guid.NewGuid(),Id=Guid.NewGuid(),Path="Docs",Name="Docs",ListBaseType=1};
                Nested=Make(NodeKind.Folder,"nested",Library.Id,"Docs/nested");Deep=Make(NodeKind.Folder,"deep",Nested.Id,"Docs/nested/deep");
                First=Make(NodeKind.File,"first.txt",Library.Id,"Docs/first.txt");Second=Make(NodeKind.File,"second.txt",Nested.Id,"Docs/nested/second.txt");Third=Make(NodeKind.File,"third.txt",Deep.Id,"Docs/nested/deep/third.txt");files.AddRange(new[]{First,Second,Third});
            }
            private Node Make(NodeKind kind,string name,Guid parent,string path)
            {return new Node {Kind=kind,SiteId=Library.SiteId,WebId=Library.WebId,ListId=Library.ListId,Id=Guid.NewGuid(),ParentId=parent,Path=path,Name=name,ListBaseType=1,StreamSchema=0,HasStream=kind==NodeKind.File,UiVersion=1024,InternalVersion=2,Level=1,Size=Encoding.UTF8.GetByteCount("current "+name),Modified=new DateTime(2024,1,2)};}
            public List<Node> GetChildren(Node parent){ChildrenReads++;var result=(parent.Id==Nested.Id?new[]{Second,Deep}:parent.Id==Deep.Id?new[]{Third}:Array.Empty<Node>()).Select(MigrationSnapshotCopy.Node).ToList();if(WrongParent && result.Count>0)result[0].ParentId=Guid.NewGuid();return result;}
            public IEnumerable<Node> EnumerateCurrentLibraryFiles(Node library)=>files.Select(MigrationSnapshotCopy.Node);
            public Node GetCurrentLibraryFile(Node library,Guid id)=>GetFile(library.SiteId,id);
            public Node GetFile(Guid site,Guid id)=>files.Where(file=>file.SiteId==site && file.Id==id).Select(MigrationSnapshotCopy.Node).SingleOrDefault();
            public List<Node> GetFileVersions(Guid site,Guid id){Node current=GetFile(site,id);Node old=MigrationSnapshotCopy.Node(current);old.UiVersion=old.HistoryVersion=512;old.InternalVersion=513;old.Size=Encoding.UTF8.GetByteCount("old "+old.Name);return new List<Node>{current,old};}
            public Node GetFileVersion(Node selected)=>GetFileVersions(selected.SiteId,selected.Id).SingleOrDefault(file=>file.HistoryVersion==selected.HistoryVersion && file.UiVersion==selected.UiVersion && file.InternalVersion==selected.InternalVersion && file.Level==selected.Level);
            public IList<StoredChunk> ReadChunks(Node document)=>new List<StoredChunk>{new StoredChunk {Partition=0,Content=Encoding.UTF8.GetBytes((document.HistoryVersion>0?"old ":"current ")+document.Name)}};
            public void ValidateSchema(){}public List<string> CheckDatabase()=>new List<string>();public List<Node> GetRootSites()=>new List<Node>();public IEnumerable<Node> EnumerateCurrentFiles(Guid? site)=>files;
        }
    }
}
