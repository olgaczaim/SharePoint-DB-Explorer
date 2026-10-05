#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using SharePointExplorer.Desktop;

namespace SharePointExplorer.Tests
{
    internal static class ZipExportChecks
    {
        public static void Run()
        {
            string root=Path.GetFullPath(Path.Combine(".scratch","zip-export-checks",Guid.NewGuid().ToString("N")));
            MixedReadback(Path.Combine(root,"mixed"));
            SafeNamesAndConflicts(Path.Combine(root,"names"));
            Cancellation(Path.Combine(root,"cancel"));
            ScopeAndReports(Path.Combine(root,"scope"));
            FinalizationAndInput(Path.Combine(root,"finalization"));
            Console.WriteLine("PASS ZIP readback, checksums, safe entries, failure cleanup, atomic publication and cancellation");
        }

        private static void MixedReadback(string root)
        {
            var catalog=new Catalog();
            Node plain=catalog.Add("stored.aspx","Docs/stored.aspx","first");
            Node tree=catalog.Add("tree.bin","Docs/deep/tree.bin","betaalpha",66);
            catalog.Chunks[tree.Id]=StorageFixture.Create(Encoding.ASCII.GetBytes("alpha"),Encoding.ASCII.GetBytes("beta"));
            Node native=catalog.Add("unsupported.bin","Docs/unsupported.bin","native",67);
            Node ghost=catalog.Add("ghost.aspx","Docs/ghost.aspx","ghost");ghost.HasStream=false;
            Node broken=catalog.Add("bad.bin","Docs/bad.bin","broken",7);
            // A shredded file whose stored tree references an object absent from storage.
            Node incomplete=catalog.Add("incomplete.bin","Docs/incomplete.bin","betaalpha",66);
            var bad=StorageFixture.Create(Encoding.ASCII.GetBytes("alpha"),Encoding.ASCII.GetBytes("beta"));
            bad.RemoveAt(1);catalog.Chunks[incomplete.Id]=bad;
            Node last=catalog.Add("last.txt","Docs/deep/last.txt","last");
            catalog.Add("not-selected.txt","Docs/not-selected.txt","unselected");
            var store=new Store(catalog);using var controller=Controller(catalog,store);
            string requested=Path.Combine(root,"selection.zip");
            var progress=new List<DesktopExportProgress>();
            DesktopExportSummary result=controller.ExportFilesAsZip(new[] {plain,plain,tree,native,ghost,broken,incomplete,last},requested,CancellationToken.None,item=>{
                Check(!File.Exists(requested),"The final ZIP became visible before finalization.");
                progress.Add(item);
            });
            Check(result.Total==5 && result.Success==3 && result.Skipped==0 && result.Failed==2 && !result.Cancelled,"ZIP mixed outcomes, deduplication or continuation differ.");
            Check(result.ArchivePath==requested && result.ArchiveReportEntry==ValidatedZipArchive.ReportEntry && result.ReportPath=="","ZIP summary does not identify the embedded report.");
            Check(catalog.Lookups==5 && store.Reads==5 && !result.Entries.Any(entry=>entry.Document.Id==native.Id || entry.Document.Id==ghost.Id),
                "Known ineligible ZIP documents caused recovery lookups or reported outcomes.");
            Check(progress.Count==6 && progress[0].Document==null && progress[0].Total==5 && progress.Last().Document.Id==last.Id,"ZIP progress omitted discovery totals or outcomes.");
            using(var archive=ZipFile.OpenRead(result.ArchivePath))
            {
                Check(archive.Entries.Count==5,"Failed or unselected content entered the ZIP.");
                foreach(DesktopExportEntry entry in result.Entries)
                {
                    ZipArchiveEntry? saved=String.IsNullOrEmpty(entry.Path) ? null : archive.GetEntry(entry.Path);
                    if(entry.Status!=RecoveryStatus.Success){Check(saved==null,"A failed document left a ZIP entry.");continue;}
                    Check(saved!=null,"A successful document is absent from ZIP.");
                    byte[] bytes=Read(saved!);
                    Check(bytes.SequenceEqual(catalog.Content[entry.Document.Id]) && bytes.LongLength==entry.Bytes && Digest(bytes)==entry.Sha256,"ZIP readback differs from verified source bytes/count/hash.");
                }
                List<string[]> report=Csv(Text(archive.GetEntry(ValidatedZipArchive.ReportEntry)!));
                Check(report.Count==6 && report[0].Length==9,"Embedded CSV lost report rows or fields.");
                for(int index=0;index<result.Entries.Count;index++)
                {
                    DesktopExportEntry entry=result.Entries[index];
                    Check(report[index+1][3]==entry.Document.Path && report[index+1][4]==entry.Status.ToString() &&
                        report[index+1][6]==entry.Sha256 && report[index+1][7]==entry.Path && report[index+1][8]==entry.Message,"Embedded report does not describe the archived outcomes.");
                }
                Check(report.Single(row=>row[2]==broken.Id.ToString("D"))[8]=="broken, \"row\"\r\nretry","Embedded CSV did not preserve multiline quoted errors.");
                Check(Text(archive.GetEntry(ValidatedZipArchive.SummaryEntry)!).Contains("Cancelled=False\nTotal=5\nProcessed=5"),"ZIP summary omitted complete batch totals.");
            }
            byte[] first=File.ReadAllBytes(result.ArchivePath);
            DesktopExportSummary repeated=controller.ExportFilesAsZip(new[] {plain},requested,CancellationToken.None,null!);
            Check(Path.GetFileName(repeated.ArchivePath)=="selection (2).zip" && first.SequenceEqual(File.ReadAllBytes(result.ArchivePath)),"ZIP export overwrote an existing archive.");
            Check(!Directory.GetFiles(root,"*.partial",SearchOption.AllDirectories).Any(),"ZIP or verified-file temporary artifacts remain.");
        }

        private static void SafeNamesAndConflicts(string root)
        {
            var catalog=new Catalog();
            Node one=catalog.Add("same.txt","Docs/same.txt","one");
            Node two=catalog.Add("SAME.txt","Docs/SAME.txt","two");
            Node unsafeFile=catalog.Add("../../escape?.txt","../CON/C:/../../escape?.txt","safe");
            Node collision=catalog.Add("a.txt","Docs/a.txt","file");
            Node nested=catalog.Add("nested.txt","Docs/a.txt/nested.txt","nested");
            var store=new Store(catalog);using var controller=Controller(catalog,store);
            DesktopExportSummary result=controller.ExportFilesAsZip(new[] {one,two,unsafeFile,collision,nested},Path.Combine(root,"safe.zip"),CancellationToken.None,null!);
            Check(result.Success==5 && result.Entries[1].Path.EndsWith("SAME (2).txt",StringComparison.Ordinal),"Case-insensitive ZIP entry collisions were not renamed.");
            Check(result.Entries[4].Path.Contains("/a.txt (2)/nested.txt"),"A ZIP directory colliding with an existing file was not renamed.");
            Check(result.Entries[2].Path.Contains("/_CON/") && result.Entries[2].Path.Contains("/C_/"),"Reserved/drive directory components were not sanitized.");
            using var archive=ZipFile.OpenRead(result.ArchivePath);
            var paths=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach(ZipArchiveEntry entry in archive.Entries)
            {
                Check(paths.Add(entry.FullName) && !entry.FullName.StartsWith("/") && !entry.FullName.Contains('\\') && !entry.FullName.Contains(':') &&
                    !entry.FullName.Split('/').Any(part=>part=="." || part==".."),"An archive entry is duplicate, rooted or traverses directories.");
                Read(entry);
            }
        }

        private static void Cancellation(string root)
        {
            var catalog=new Catalog();Node first=catalog.Add("one.txt","Docs/one.txt","one");Node second=catalog.Add("two.txt","Docs/two.txt","two");
            var store=new Store(catalog);using var controller=Controller(catalog,store);
            using var pre=new CancellationTokenSource();pre.Cancel();
            string noOutput=Path.Combine(root,"pre","selection.zip");
            DesktopExportSummary cancelled=controller.ExportFilesAsZip(new[] {first,second},noOutput,pre.Token,null!);
            Check(cancelled.Cancelled && cancelled.ArchivePath=="" && store.Reads==0 && !Directory.Exists(Path.GetDirectoryName(noOutput)),"Pre-cancelled ZIP export created output.");
            using var initial=new CancellationTokenSource();
            cancelled=controller.ExportFilesAsZip(new[] {first},Path.Combine(root,"initial","selection.zip"),initial.Token,_=>initial.Cancel());
            Check(cancelled.Cancelled && cancelled.ArchivePath=="" && store.Reads==0 && !Directory.GetFiles(cancelled.Directory).Any(),"Cancellation before a ZIP success published temporary data.");
            using var between=new CancellationTokenSource();
            DesktopExportSummary partial=controller.ExportFilesAsZip(new[] {first,second},Path.Combine(root,"partial","selection.zip"),between.Token,item=>{if(item.Completed==1)between.Cancel();});
            Check(partial.Cancelled && partial.Success==1 && partial.Total==2 && partial.Entries.Count==1 && File.Exists(partial.ArchivePath),"Between-file cancellation did not preserve a valid completed ZIP subset.");
            using(var archive=ZipFile.OpenRead(partial.ArchivePath))
            {
                Check(archive.Entries.Count==3 && Encoding.UTF8.GetString(Read(archive.GetEntry(partial.Entries[0].Path)!))=="one","Cancelled ZIP contains unfinished/missing files.");
                Check(Text(archive.GetEntry(ValidatedZipArchive.SummaryEntry)!).Contains("Cancelled=True\nTotal=2\nProcessed=1"),"Cancelled ZIP does not declare its incomplete selection.");
                Check(Csv(Text(archive.GetEntry(ValidatedZipArchive.ReportEntry)!)).Count==2,"Cancelled ZIP report included unprocessed files.");
            }
            Node changed=catalog.Add("changed.aspx","Docs/changed.aspx","stored");
            catalog.Lookup=id=>{Node current=Copy(catalog.Documents[id]);if(id==changed.Id)current.HasStream=false;return current;};
            using var noSuccess=new CancellationTokenSource();
            cancelled=controller.ExportFilesAsZip(new[] {changed,second},Path.Combine(root,"no-success","selection.zip"),noSuccess.Token,item=>{if(item.Completed==1)noSuccess.Cancel();});
            Check(cancelled.Cancelled && cancelled.Success==0 && cancelled.ArchivePath=="" && !Directory.GetFiles(cancelled.Directory).Any(),"Cancelled report-only selection published a ZIP before any success.");
        }

        private static void ScopeAndReports(string root)
        {
            {
                var catalog=new Catalog();Node ghost=catalog.Add("default.aspx","Docs/default.aspx","template");ghost.HasStream=false;
                Node native=catalog.Add("unsupported.bin","Docs/unsupported.bin","native",67);
                var store=new Store(catalog);using var controller=Controller(catalog,store);
                var progress=new List<DesktopExportProgress>();
                DesktopExportSummary result=controller.ExportFilesAsZip(new[] {ghost,native},Path.Combine(root,"ignored.zip"),CancellationToken.None,progress.Add);
                Check(result.Total==0 && result.Success==0 && result.Skipped==0 && result.Failed==0 && result.Entries.Count==0 &&
                    catalog.Lookups==0 && store.Reads==0 && progress.Count==1 && progress[0].Total==0,"Known ineligible ZIP selection caused counts or retrieval.");
                using var archive=ZipFile.OpenRead(result.ArchivePath);
                Check(archive.Entries.Count==2 && Csv(Text(archive.GetEntry(ValidatedZipArchive.ReportEntry)!)).Count==1 &&
                    Text(archive.GetEntry(ValidatedZipArchive.SummaryEntry)!).Contains("Total=0\nProcessed=0"),"An all-ineligible selection did not produce a clean report-only ZIP.");
            }
            for(int variant=0;variant<2;variant++)
            {
                var catalog=new Catalog();Node file=catalog.Add("changed.aspx","Docs/changed.aspx","stored");Node changed=Copy(file);
                if(variant==0)changed.HasStream=false;else changed.StreamSchema=67;
                catalog.Lookup=_=>changed;
                var store=new Store(catalog);using var controller=Controller(catalog,store);
                DesktopExportSummary result=controller.ExportFilesAsZip(new[] {file},Path.Combine(root,"availability-"+variant+".zip"),CancellationToken.None,null!);
                Check(result.Total==1 && result.Skipped==1 && result.Failed==0 && result.Entries[0].Status==
                    (variant==0 ? RecoveryStatus.Unavailable : RecoveryStatus.Unsupported) && store.Reads==(variant==0 ? 1 : 0),
                    "Fresh unavailable/unsupported ZIP content lost its runtime outcome.");
                using var archive=ZipFile.OpenRead(result.ArchivePath);
                Check(archive.Entries.Count==2 && Csv(Text(archive.GetEntry(ValidatedZipArchive.ReportEntry)!)).Count==2,
                    "Fresh unavailable/unsupported ZIP content lost its detailed audit row.");
            }
            for(int variant=0;variant<4;variant++)
            {
                var catalog=new Catalog();Node file=catalog.Add("selected.txt","Docs/selected.txt","file");
                Node changed=Copy(file);
                if(variant==0)changed.Id=Guid.NewGuid();
                if(variant==1)changed.SiteId=Guid.NewGuid();
                if(variant==2)changed.HistoryVersion=1;
                if(variant==3)changed.Kind=NodeKind.ListItem;
                catalog.Lookup=_=>changed;
                var store=new Store(catalog);using var controller=Controller(catalog,store);
                DesktopExportSummary result=controller.ExportFilesAsZip(new[] {file},Path.Combine(root,"wrong-"+variant+".zip"),CancellationToken.None,null!);
                Check(result.Failed==1 && store.Reads==0,"Wrong current document scope reached ZIP content retrieval.");
                using var archive=ZipFile.OpenRead(result.ArchivePath);
                Check(archive.Entries.Count==2 && Csv(Text(archive.GetEntry(ValidatedZipArchive.ReportEntry)!)).Count==2,"All-failed ZIP lost its report-only result or included file bytes.");
            }
            {
                var catalog=new Catalog();Node file=catalog.Add("old.txt","Docs/old.txt","bytes");
                Node current=Copy(file);current.Name="new.txt";current.Path="Docs/new-folder/new.txt";
                catalog.Lookup=_=>current;
                var store=new Store(catalog);using var controller=Controller(catalog,store);
                DesktopExportSummary result=controller.ExportFilesAsZip(new[] {file},Path.Combine(root,"renamed.zip"),CancellationToken.None,null!);
                Check(catalog.Lookups==1 && result.Entries[0].Document.Name=="new.txt" && result.Entries[0].Path.EndsWith("/new-folder/new.txt"),"ZIP did not use one prepared current metadata snapshot.");
                using var archive=ZipFile.OpenRead(result.ArchivePath);
                Check(Csv(Text(archive.GetEntry(ValidatedZipArchive.ReportEntry)!))[1][3]=="Docs/new-folder/new.txt","ZIP report retained stale discovered metadata.");
            }
        }

        private static void FinalizationAndInput(string root)
        {
            var catalog=new Catalog();Node file=catalog.Add("one.txt","Docs/one.txt","one");var store=new Store(catalog);
            using var session=new RecoverySession(catalog,store,DocumentDecoderRegistry.CreateDefault());
            string failed=Path.Combine(root,"failed.zip");
            using(var archive=new ValidatedZipArchive(failed))
            {
                archive.Add(session.Engine.Prepare(file));
                Reject<IOException>(()=>archive.Commit(_=>throw new IOException("Synthetic report write failure."),""));
            }
            Check(!File.Exists(failed) && !Directory.GetFiles(root,"*.partial").Any(),"ZIP finalization failure published data or retained a temp archive.");
            using var controller=Controller(catalog,store);
            string invalid=Path.Combine(root,"invalid.zip");
            Reject<ArgumentException>(()=>controller.ExportFilesAsZip(Array.Empty<Node>(),invalid,CancellationToken.None,null!));
            Reject<ArgumentException>(()=>controller.ExportFilesAsZip(new[] {new Node {Kind=NodeKind.Library}},invalid,CancellationToken.None,null!));
            Reject<ArgumentException>(()=>controller.ExportFilesAsZip(new[] {file},Path.Combine(root,"invalid.txt"),CancellationToken.None,null!));
            string callbackFault=Path.Combine(root,"callback.zip");
            Reject<InvalidOperationException>(()=>controller.ExportFilesAsZip(new[] {file},callbackFault,CancellationToken.None,item=>{if(item.Completed==1)throw new InvalidOperationException("Synthetic progress failure.");}));
            Check(!File.Exists(callbackFault) && !Directory.GetFiles(root,"*.partial").Any(),"A fatal callback published an incomplete ZIP.");
            controller.Dispose();
            Reject<ObjectDisposedException>(()=>controller.ExportFilesAsZip(new[] {file},invalid,CancellationToken.None,null!));
        }

        private static ExplorerController Controller(Catalog catalog,Store store)=>new(new RecoverySession(catalog,store,
            new DocumentDecoderRegistry(new IDocumentDecoder[] {DocumentDecoderRegistry.CreateDefault().Resolve(0),DocumentDecoderRegistry.CreateDefault().Resolve(66),new BrokenDecoder()})));
        private static void Check(bool value,string message){if(!value)throw new Exception(message);}
        private static void Reject<T>(Action action) where T:Exception {try{action();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name);}
        private static byte[] Read(ZipArchiveEntry entry){using Stream input=entry.Open();using var memory=new MemoryStream();input.CopyTo(memory);return memory.ToArray();}
        private static string Text(ZipArchiveEntry entry){using var reader=new StreamReader(entry.Open(),Encoding.UTF8);return reader.ReadToEnd();}
        private static string Digest(byte[] bytes)=>Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        private static int Find(byte[] haystack,byte[] needle){for(int offset=0;offset<=haystack.Length-needle.Length;offset++)if(haystack.AsSpan(offset,needle.Length).SequenceEqual(needle))return offset;throw new Exception("Fixture payload missing.");}
        private static Node Copy(Node file)=>new() {Kind=file.Kind,SiteId=file.SiteId,Id=file.Id,WebId=file.WebId,ListId=file.ListId,Name=file.Name,Path=file.Path,
            Size=file.Size,Modified=file.Modified,HasStream=file.HasStream,StreamSchema=file.StreamSchema,HistoryVersion=file.HistoryVersion};
        private static List<string[]> Csv(string text)
        {
            var result=new List<string[]>();var row=new List<string>();var field=new StringBuilder();bool quoted=false;
            for(int index=0;index<text.Length;index++)
            {
                char c=text[index];
                if(c=='"'){if(quoted && index+1<text.Length && text[index+1]=='"'){field.Append('"');index++;}else quoted=!quoted;}
                else if(!quoted && c==','){row.Add(field.ToString());field.Clear();}
                else if(!quoted && (c=='\r'||c=='\n')){row.Add(field.ToString());field.Clear();result.Add(row.ToArray());row.Clear();if(c=='\r' && index+1<text.Length && text[index+1]=='\n')index++;}
                else field.Append(c);
            }
            Check(!quoted,"CSV is unterminated.");return result;
        }
        private sealed class Catalog : ISharePointCatalog
        {
            internal readonly Guid Site=Guid.NewGuid();
            internal readonly Dictionary<Guid,Node> Documents=new();
            internal readonly Dictionary<Guid,byte[]> Content=new();
            internal readonly Dictionary<Guid,IList<StoredChunk>> Chunks=new();
            internal Func<Guid,Node>? Lookup;internal int Lookups;
            public string SourceName=>"Synthetic, \"source\"";
            internal Node Add(string name,string path,string text,byte schema=0)
            {
                var bytes=Encoding.UTF8.GetBytes(text);
                var file=new Node {Kind=NodeKind.File,SiteId=Site,Id=Guid.NewGuid(),Name=name,Path=path,Size=bytes.Length,HasStream=true,StreamSchema=schema};
                Documents.Add(file.Id,file);Content.Add(file.Id,bytes);return file;
            }
            public Node GetFile(Guid site,Guid id){Lookups++;return Lookup!=null ? Lookup(id) : Documents.TryGetValue(id,out Node? file) && file.SiteId==site ? Copy(file) : null!;}
            public void ValidateSchema(){}
            public List<string> CheckDatabase()=>new();
            public List<Node> GetRootSites()=>new();
            public List<Node> GetChildren(Node parent)=>new();
            public IEnumerable<Node> EnumerateCurrentFiles(Guid? site)=>throw new Exception("ZIP selection enumerated the whole source.");
        }
        private sealed class Store : IDocumentChunkStore
        {
            private readonly Catalog catalog;internal int Reads;
            internal Store(Catalog catalog){this.catalog=catalog;}
            public IList<StoredChunk> ReadChunks(Node file)
            {
                Reads++;if(file.HasStream==false)throw new ContentUnavailableException("Template bytes unavailable.");
                return catalog.Chunks.TryGetValue(file.Id,out IList<StoredChunk>? chunks) ? chunks : new List<StoredChunk> {new() {Partition=0,Content=catalog.Content[file.Id]}};
            }
        }
        private sealed class BrokenDecoder : IDocumentDecoder
        {
            public string Name=>"Synthetic partial corrupt file";
            public bool Supports(byte schema)=>schema==7;
            public void Write(IList<StoredChunk> chunks,long size,Stream output){output.WriteByte(1);throw new InvalidDataException("broken, \"row\"\r\nretry");}
        }
    }
}
