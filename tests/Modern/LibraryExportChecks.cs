#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using SharePointExplorer.Desktop;

namespace SharePointExplorer.Tests
{
    internal static class LibraryExportChecks
    {
        public static void Run()
        {
            string root=Path.GetFullPath(Path.Combine(".scratch","library-export-checks",Guid.NewGuid().ToString("N")));
            RecursiveExport(Path.Combine(root,"recursive"));
            Cancellation(Path.Combine(root,"cancel"));
            ScopeAndFreshMetadata(Path.Combine(root,"scope"));
            EmptyAndInput(Path.Combine(root,"empty"));
            Console.WriteLine("PASS complete-library export, nested paths, active scope, reports and cancellation");
        }

        private static void RecursiveExport(string root)
        {
            var catalog=new Catalog();
            Node direct=catalog.Add("stored.aspx","Docs/stored.aspx","direct");
            Node nested=catalog.Add("nested.txt","Docs/deep/more/nested.txt","nested");
            Node ghost=catalog.Add("ghost.aspx","Docs/ghost.aspx","ghost"); ghost.HasStream=false;
            catalog.Add("unsupported.bin","Docs/unsupported.bin","native").StreamSchema=67;
            catalog.Add("broken.bin","Docs/deep/broken.bin","broken").StreamSchema=7;
            Node last=catalog.Add("last.txt","Docs/deep/more/last.txt","last");
            catalog.Discovered.Add(direct); // Duplicate physical/metadata association.
            var store=new Store(catalog);
            using var controller=Controller(catalog,store);
            var progress=new List<DesktopExportProgress>();
            DesktopExportSummary result=controller.ExportLibrary(catalog.Library,root,CancellationToken.None,progress.Add);
            Check(result.Total==4 && result.Success==3 && result.Skipped==0 && result.Failed==1,"Library export omitted nested files, duplicated content, or stopped at a failure.");
            Check(progress.Count==5 && progress[0].Document==null && progress[0].Completed==0 && progress[0].Total==4,"Discovery did not announce a determinate file count.");
            Check(result.Entries.Last().Document.Id==last.Id && File.ReadAllText(result.Entries.Last().Path)=="last","Export did not continue after unsupported/corrupt files.");
            Check(File.Exists(Path.Combine(root,catalog.Library.SiteId.ToString("D"),"Docs","deep","more","nested.txt")),"Nested folders were not preserved.");
            Check(File.ReadAllLines(result.ReportPath).Length==5 && File.ReadAllText(result.ReportPath).Contains("\"Corrupt\""),"Library CSV omitted per-file outcomes.");
            Check(Directory.GetFiles(root,"*.partial",SearchOption.AllDirectories).Length==0,"Failed library export left a partial file.");
            Check(catalog.UnscopedReads==0 && catalog.ScopedReads==4 && store.Reads==4,"Library recovery used an unscoped or repeated metadata lookup.");
            string firstPath=result.Entries.Single(entry=>entry.Document.Id==nested.Id).Path;
            DesktopExportSummary again=controller.ExportLibrary(catalog.Library,root,CancellationToken.None,null!);
            Check(again.Success==3 && File.ReadAllText(firstPath)=="nested" && File.Exists(Path.Combine(Path.GetDirectoryName(firstPath)!,"nested (2).txt")),"Repeated library export overwrote an existing file.");
            Check(direct.Id!=nested.Id,"Fixture IDs unexpectedly collide.");
            Check(result.Entries.Single(entry=>entry.Document.Id==direct.Id).Status==RecoveryStatus.Success &&
                !File.ReadAllText(result.ReportPath).Contains("ghost.aspx") && !File.ReadAllText(result.ReportPath).Contains("unsupported.bin"),
                "A stored ASPX was filtered or a known ineligible document appeared in the report.");
        }

        private static void Cancellation(string root)
        {
            var catalog=new Catalog(); catalog.Add("one.txt","Docs/one.txt","one"); catalog.Add("two.txt","Docs/deep/two.txt","two");
            var store=new Store(catalog);
            using var controller=Controller(catalog,store);
            using var before=new CancellationTokenSource(); before.Cancel();
            DesktopExportSummary cancelled=controller.ExportLibrary(catalog.Library,Path.Combine(root,"before"),before.Token,null!);
            Check(cancelled.Cancelled && catalog.Discoveries==0 && store.Reads==0 && !Directory.Exists(cancelled.Directory),"Pre-cancellation accessed the source or created output.");
            using var discovery=new CancellationTokenSource();
            catalog.OnDiscovery=()=>discovery.Cancel();
            cancelled=controller.ExportLibrary(catalog.Library,Path.Combine(root,"discovery"),discovery.Token,null!);
            Check(cancelled.Cancelled && store.Reads==0 && !Directory.Exists(cancelled.Directory),"Discovery cancellation exported a partial selection.");
            catalog.OnDiscovery=null;
            using var between=new CancellationTokenSource();
            DesktopExportSummary partial=controller.ExportLibrary(catalog.Library,Path.Combine(root,"between"),between.Token,item=>{if(item.Completed==1)between.Cancel();});
            Check(partial.Cancelled && partial.Total==2 && partial.Success==1 && partial.Entries.Count==1 && store.Reads==1,"Cancellation did not stop between library files.");
            Check(File.ReadAllLines(partial.ReportPath).Length==2,"Cancelled library report did not preserve the completed file.");
        }

        private static void ScopeAndFreshMetadata(string root)
        {
            // Discovery is rejected before any output if the provider leaks another scope.
            for(int variant=0;variant<5;variant++)
            {
                var catalog=new Catalog(); Node file=catalog.Add("file.txt","Docs/file.txt","file");
                if(variant==0)file.SiteId=Guid.NewGuid();
                if(variant==1)file.WebId=Guid.NewGuid();
                if(variant==2)file.ListId=Guid.NewGuid();
                if(variant==3)file.Kind=NodeKind.ListItem;
                if(variant==4)file.HistoryVersion=1;
                file.HasStream=false;file.StreamSchema=67;
                var store=new Store(catalog); using var controller=Controller(catalog,store);
                string target=Path.Combine(root,"discovery-"+variant);
                Reject<InvalidDataException>(()=>controller.ExportLibrary(catalog.Library,target,CancellationToken.None,null!));
                Check(store.Reads==0 && !Directory.Exists(target),"Out-of-scope discovery accessed content or created output.");
            }
            // A broken scoped lookup must fail before retrieving bytes.
            for(int variant=0;variant<6;variant++)
            {
                var catalog=new Catalog(); Node file=catalog.Add("file.txt","Docs/file.txt","file");
                Node changed=Copy(file);
                if(variant==0)changed.SiteId=Guid.NewGuid();
                if(variant==1)changed.WebId=Guid.NewGuid();
                if(variant==2)changed.ListId=Guid.NewGuid();
                if(variant==3)changed.Id=Guid.NewGuid();
                if(variant==4)changed.HistoryVersion=1;
                if(variant==5)changed.Kind=NodeKind.Folder;
                catalog.CurrentOverride=_=>changed;
                var store=new Store(catalog); using var controller=Controller(catalog,store);
                DesktopExportSummary result=controller.ExportLibrary(catalog.Library,Path.Combine(root,"current-"+variant),CancellationToken.None,null!);
                Check(result.Failed==1 && result.Entries[0].Status==RecoveryStatus.Corrupt && store.Reads==0,"A wrong current scope reached content retrieval.");
                Check(Directory.GetFiles(result.Directory,"*.partial",SearchOption.AllDirectories).Length==0,"Rejected current scope left partial output.");
            }
            // Inactivated library and moved documents become unavailable at the scoped current lookup.
            for(int variant=0;variant<2;variant++)
            {
                var catalog=new Catalog(); Node file=catalog.Add("file.txt","Docs/file.txt","file");
                catalog.AfterDiscovery=()=>{
                    if(variant==0)catalog.Active=false;
                    else catalog.Current[file.Id].ListId=Guid.NewGuid();
                };
                var store=new Store(catalog); using var controller=Controller(catalog,store);
                DesktopExportSummary result=controller.ExportLibrary(catalog.Library,Path.Combine(root,"inactive-"+variant),CancellationToken.None,null!);
                Check(result.Skipped==1 && result.Entries[0].Status==RecoveryStatus.Unavailable && store.Reads==0,"Inactive or moved source content was recovered.");
            }
            // Fresh unavailability is audited even though the discovery snapshot was eligible.
            for(int variant=0;variant<2;variant++)
            {
                var catalog=new Catalog();Node file=catalog.Add("changed.aspx","Docs/changed.aspx","stored");
                catalog.AfterDiscovery=()=>{if(variant==0)catalog.Current[file.Id].HasStream=false;else catalog.Current[file.Id].StreamSchema=67;};
                var store=new Store(catalog);using var controller=Controller(catalog,store);
                DesktopExportSummary result=controller.ExportLibrary(catalog.Library,Path.Combine(root,"availability-"+variant),CancellationToken.None,null!);
                Check(result.Total==1 && result.Skipped==1 && result.Failed==0 && result.Entries[0].Status==
                    (variant==0 ? RecoveryStatus.Unavailable : RecoveryStatus.Unsupported) && File.ReadAllLines(result.ReportPath).Length==2 &&
                    catalog.ScopedReads==1 && store.Reads==(variant==0 ? 1 : 0),"Fresh unavailable/unsupported library content lost its runtime report.");
            }
            {
                var catalog=new Catalog(); Node file=catalog.Add("old.txt","Docs/old.txt","bytes");
                catalog.AfterDiscovery=()=>{catalog.Current[file.Id].Name="new.txt";catalog.Current[file.Id].Path="Docs/deep/new.txt";};
                var store=new Store(catalog); using var controller=Controller(catalog,store);
                var progress=new List<DesktopExportProgress>();
                DesktopExportSummary result=controller.ExportLibrary(catalog.Library,Path.Combine(root,"renamed"),CancellationToken.None,progress.Add);
                Check(result.Success==1 && result.Entries[0].Document.Name=="new.txt" && result.Entries[0].Document.Path=="Docs/deep/new.txt","Report metadata did not use the prepared current snapshot.");
                Check(result.Entries[0].Path.EndsWith(Path.Combine("deep","new.txt"),StringComparison.Ordinal) && File.ReadAllText(result.ReportPath).Contains("Docs/deep/new.txt") && progress.Last().Document.Name=="new.txt","Export path, report and progress disagree with the prepared snapshot.");
            }
            {
                var catalog=new Catalog(); catalog.Add("one.txt","Docs/one.txt","one");
                Node selected=Copy(catalog.Library);
                catalog.OnDiscovery=()=>{selected.ListId=Guid.NewGuid();selected.WebId=Guid.NewGuid();};
                var store=new Store(catalog); using var controller=Controller(catalog,store);
                Check(controller.ExportLibrary(selected,Path.Combine(root,"frozen"),CancellationToken.None,null!).Success==1,"Caller mutation changed a frozen library scope.");
            }
        }

        private static void EmptyAndInput(string root)
        {
            var catalog=new Catalog(); var store=new Store(catalog); using var controller=Controller(catalog,store);
            var progress=new List<DesktopExportProgress>();
            DesktopExportSummary empty=controller.ExportLibrary(catalog.Library,root,CancellationToken.None,progress.Add);
            Check(empty.Total==0 && empty.Success==0 && !empty.Cancelled && File.ReadAllLines(empty.ReportPath).Length==1 && progress.Count==1 && progress[0].Total==0,"Empty library did not complete with a header-only report.");
            Node ghost=catalog.Add("default.aspx","Docs/default.aspx","template");ghost.HasStream=false;
            catalog.Add("unsupported.bin","Docs/unsupported.bin","native").StreamSchema=67;
            progress.Clear();
            DesktopExportSummary ignored=controller.ExportLibrary(catalog.Library,Path.Combine(root,"ignored"),CancellationToken.None,progress.Add);
            Check(ignored.Total==0 && ignored.Success==0 && ignored.Skipped==0 && ignored.Failed==0 && ignored.Entries.Count==0 &&
                File.ReadAllLines(ignored.ReportPath).Length==1 && progress.Count==1 && progress[0].Total==0 && store.Reads==0 && catalog.ScopedReads==0,
                "A library containing only known ineligible files caused counts, report rows or content lookups.");
            Node list=Copy(catalog.Library);list.Kind=NodeKind.List;
            Reject<ArgumentException>(()=>controller.ExportLibrary(list,root,CancellationToken.None,null!));
            Node wrongRoot=Copy(catalog.Library);wrongRoot.Id=Guid.NewGuid();
            Reject<ContentUnavailableException>(()=>controller.ExportLibrary(wrongRoot,root,CancellationToken.None,null!));
            controller.Dispose();
            Reject<ObjectDisposedException>(()=>controller.ExportLibrary(catalog.Library,root,CancellationToken.None,null!));
        }

        private static ExplorerController Controller(Catalog catalog,Store store)
        {
            return new ExplorerController(new RecoverySession(catalog,store,new DocumentDecoderRegistry(new IDocumentDecoder[] {
                DocumentDecoderRegistry.CreateDefault().Resolve(0),new BrokenDecoder() })));
        }
        private static void Check(bool value,string message){if(!value)throw new Exception(message);}
        private static void Reject<T>(Action action) where T:Exception {try{action();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name);}
        private static Node Copy(Node node)=>new() {Kind=node.Kind,SiteId=node.SiteId,WebId=node.WebId,ListId=node.ListId,Id=node.Id,Name=node.Name,Path=node.Path,
            Size=node.Size,HistoryVersion=node.HistoryVersion,StreamSchema=node.StreamSchema,HasStream=node.HasStream,Modified=node.Modified,Level=node.Level};

        private sealed class Catalog : ISharePointCatalog,ISharePointLibraryCatalog
        {
            internal readonly Node Library=new() {Kind=NodeKind.Library,SiteId=Guid.NewGuid(),WebId=Guid.NewGuid(),ListId=Guid.NewGuid(),Id=Guid.NewGuid(),Name="Docs",Path="Docs"};
            internal readonly List<Node> Discovered=new();
            internal readonly Dictionary<Guid,Node> Current=new();
            internal readonly Dictionary<Guid,byte[]> Content=new();
            internal Action? OnDiscovery,AfterDiscovery;
            internal Func<Guid,Node?>? CurrentOverride;
            internal bool Active=true;
            internal int Discoveries,ScopedReads,UnscopedReads;
            public string SourceName=>"Synthetic library source";
            internal Node Add(string name,string path,string text)
            {
                byte[] bytes=Encoding.UTF8.GetBytes(text);
                var file=new Node {Kind=NodeKind.File,SiteId=Library.SiteId,WebId=Library.WebId,ListId=Library.ListId,Id=Guid.NewGuid(),
                    Name=name,Path=path,Size=bytes.Length,HasStream=true};
                Discovered.Add(file);Current[file.Id]=Copy(file);Content[file.Id]=bytes;return file;
            }
            public IEnumerable<Node> EnumerateCurrentLibraryFiles(Node library)
            {
                Discoveries++;
                if(!Active || library.SiteId!=Library.SiteId || library.WebId!=Library.WebId || library.ListId!=Library.ListId || library.Id!=Library.Id)
                    throw new ContentUnavailableException("Library is inactive or has another root.");
                foreach(Node file in Discovered)
                {
                    // Capture modifications to the discovery snapshot as the current file,
                    // while keeping subsequent rename/move changes separate.
                    Current[file.Id]=Copy(file);
                    OnDiscovery?.Invoke();yield return file;
                }
                AfterDiscovery?.Invoke();
            }
            public Node GetCurrentLibraryFile(Node library,Guid fileId)
            {
                ScopedReads++;
                if(CurrentOverride!=null)return CurrentOverride(fileId)!;
                if(!Active || library.SiteId!=Library.SiteId || library.WebId!=Library.WebId || library.ListId!=Library.ListId || library.Id!=Library.Id)return null!;
                return Current.TryGetValue(fileId,out Node? node) && node.Kind==NodeKind.File && node.HistoryVersion==0 &&
                    node.SiteId==library.SiteId && node.WebId==library.WebId && node.ListId==library.ListId ? Copy(node) : null!;
            }
            public Node GetFile(Guid siteId,Guid fileId){UnscopedReads++;throw new Exception("Library recovery used an unscoped lookup.");}
            public void ValidateSchema(){}
            public List<string> CheckDatabase()=>new();
            public List<Node> GetRootSites()=>new();
            public List<Node> GetChildren(Node parent)=>throw new Exception("Library export depended on expanded tree folders.");
            public IEnumerable<Node> EnumerateCurrentFiles(Guid? siteId)=>throw new Exception("Library export enumerated the entire source.");
        }
        private sealed class Store : IDocumentChunkStore
        {
            private readonly Catalog catalog;internal int Reads;
            internal Store(Catalog catalog){this.catalog=catalog;}
            public IList<StoredChunk> ReadChunks(Node document)
            {
                Reads++;
                if(document.HasStream==false)throw new ContentUnavailableException("Missing template bytes.");
                return new List<StoredChunk> {new() {Partition=0,Content=catalog.Content[document.Id]}};
            }
        }
        private sealed class BrokenDecoder : IDocumentDecoder
        {
            public string Name=>"Synthetic corrupt data";
            public bool Supports(byte schema)=>schema==7;
            public void Write(IList<StoredChunk> chunks,long expected,Stream output)
            {
                output.WriteByte(1);throw new InvalidDataException("Synthetic corrupt library file.");
            }
        }
    }
}
