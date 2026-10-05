#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using SharePointExplorer.Desktop;

namespace SharePointExplorer.Tests
{
    public static class ControllerVersionExportChecks
    {
        public static void Run(string root)
        {
            Directory.CreateDirectory(root);
            MetadataAndScope();
            ExactBytesAndCurrentRoutes(Path.Combine(root,"exact"));
            FailuresAndCancellation(Path.Combine(root,"guards"));
            Console.WriteLine("PASS exact desktop version metadata, historical bytes, names, audit, guards and current-only routes");
        }
        private static void MetadataAndScope()
        {
            var catalog=new Catalog();var store=new Store(catalog);using var controller=Controller(catalog,store);
            catalog.Rows.Add(Copy(catalog.Current));
            List<Node> versions=controller.GetFileVersions(catalog.Current);
            Check(versions.Count==3 && versions[0].HistoryVersion==0 && versions[1].UiVersion==513 && versions[2].UiVersion==512 && store.Reads==0,
                "Version history loaded BLOBs, duplicated current state or lost current-first/descending history order.");
            Check(versions[2].InternalVersion==1025 && versions[2].HistoryVersion==512,
                "Version history treated InternalVersion as the storage history key.");
            versions[2].Name="modified client copy";
            Check(controller.GetFileVersions(catalog.Current).Last().Name=="report.txt","Version metadata leaked caller mutations into the source.");
            Node bad=Copy(catalog.Old);bad.SiteId=Guid.NewGuid();catalog.Rows.Add(bad);
            Reject<InvalidDataException>(()=>controller.GetFileVersions(catalog.Current));catalog.Rows.Remove(bad);
            Node conflicting=Copy(catalog.Old);conflicting.Size++;
            catalog.Rows.Add(conflicting);Reject<InvalidDataException>(()=>controller.GetFileVersions(catalog.Current));catalog.Rows.Remove(conflicting);
            Reject<ArgumentException>(()=>controller.GetFileVersions(catalog.Old));
            var currentOnly=new CurrentOnlyCatalog(catalog);using var legacy=new ExplorerController(new RecoverySession(currentOnly,new Store(catalog),DocumentDecoderRegistry.CreateDefault()));
            Reject<NotSupportedException>(()=>legacy.GetFileVersions(catalog.Current));
        }
        private static void ExactBytesAndCurrentRoutes(string root)
        {
            var catalog=new Catalog();var store=new Store(catalog);using var controller=Controller(catalog,store);
            var progress=new List<DesktopExportProgress>();
            DesktopExportSummary old=controller.ExportVersion(catalog.Old,root,CancellationToken.None,progress.Add);
            Check(old.Success==1 && old.Total==1 && old.Failed==0 && old.Entries.Single().Document.HistoryVersion==512 && store.Last is Node read && read.HistoryVersion==512 && read.InternalVersion==1025,
                "Version recovery did not read the selected UI/history/internal key.");
            DesktopExportEntry entry=old.Entries.Single();byte[] expected=Encoding.UTF8.GetBytes("previous approved content\n");
            Check(File.ReadAllBytes(entry.Path).SequenceEqual(expected) && entry.Bytes==expected.Length && entry.Sha256==Digest(expected),
                "Historical publication does not match independently specified older bytes/hash.");
            Check(Path.GetFileName(entry.Path)=="report (v1.0).txt" && SameDirectory(entry.Path,root) && SameDirectory(old.ReportPath,root) && old.Directory==Path.GetFullPath(root) &&
                progress.Single().Document.HistoryVersion==512 && progress.Single().ExportPath==entry.Path,
                "Version file/report were not published directly in the chosen directory with the expected name, progress and identity.");
            string[] report=File.ReadAllLines(old.ReportPath);string[] fields=report[1].Split(',').Select(value=>value.Trim('"')).ToArray();
            Check(report.Length==2 && report[0].Split(',').Length==13 && fields[9]=="512" && fields[10]=="512" && fields[11]=="1" && fields[12]=="1025" && fields[6]==Digest(expected),
                "Version audit omitted or confused UI/history/level/internal identity and digest.");
            DesktopExportSummary repeated=controller.ExportVersion(catalog.Old,root,CancellationToken.None,null!);
            Check(repeated.Success==1 && repeated.Entries.Single().Path!=entry.Path && SameDirectory(repeated.Entries.Single().Path,root) && SameDirectory(repeated.ReportPath,root) && File.ReadAllBytes(entry.Path).SequenceEqual(expected) && File.ReadAllBytes(repeated.Entries.Single().Path).SequenceEqual(expected),
                "Repeated version export overwrote an existing historical file.");
            DesktopExportSummary current=controller.ExportVersion(catalog.Current,Path.Combine(root,"current-version"),CancellationToken.None,null!);
            Check(current.Success==1 && Path.GetFileName(current.Entries.Single().Path)=="report (v2.0).txt" && SameDirectory(current.Entries.Single().Path,Path.Combine(root,"current-version")) && SameDirectory(current.ReportPath,Path.Combine(root,"current-version")) && File.ReadAllText(current.Entries.Single().Path)=="current document content\n",
                "The explicitly selected current version did not retain its version label and current bytes.");
            int versionLookups=catalog.VersionLookups,currentLookups=catalog.CurrentLookups;
            string ordinaryDirectory=Path.Combine(root,"ordinary");
            DesktopExportSummary ordinary=controller.ExportFiles(new[]{catalog.Current},ordinaryDirectory,CancellationToken.None,null!);
            Check(ordinary.Success==1 && Path.GetFileName(ordinary.Entries.Single().Path)=="report.txt" && SameDirectory(ordinary.Entries.Single().Path,ordinaryDirectory) && SameDirectory(ordinary.ReportPath,ordinaryDirectory) &&
                File.ReadAllText(ordinary.Entries.Single().Path)=="current document content\n" && ordinary.Entries.Single().Sha256==current.Entries.Single().Sha256 &&
                catalog.VersionLookups==versionLookups && catalog.CurrentLookups==currentLookups+1,
                "Ordinary selected-file export did not publish current bytes directly, or reloaded its metadata more than once.");
            string[] ordinaryReport=File.ReadAllLines(ordinary.ReportPath);
            Check(ordinaryReport.Length==2 && ordinaryReport[0].Split(',').Length==9 && ordinaryReport[1].Contains(ordinary.Entries.Single().Path,StringComparison.Ordinal),
                "The selected-file audit lost its nine-column format or direct saved path.");
            DesktopExportSummary archive=controller.ExportFilesAsZip(new[]{catalog.Current},Path.Combine(root,"current.zip"),CancellationToken.None,null!);
            using(ZipArchive zip=ZipFile.OpenRead(archive.ArchivePath))
            using(var reader=new StreamReader(zip.GetEntry(archive.Entries.Single().Path)!.Open(),Encoding.UTF8))
                Check(reader.ReadToEnd()=="current document content\n" && archive.Entries.Single().Path==catalog.Current.SiteId.ToString("D")+"/Docs/Nested/report.txt" && archive.Entries.Single().Document.HistoryVersion==0 && catalog.VersionLookups==versionLookups,
                    "Current selected ZIP export recovered a historical state.");
            int reads=store.Reads;DesktopExportSummary blocked=controller.ExportFiles(new[]{catalog.Old},Path.Combine(root,"ordinary-history"),CancellationToken.None,null!);
            Check(blocked.Skipped==1 && blocked.Success==0 && store.Reads==reads && catalog.VersionLookups==versionLookups,
                "The ordinary selected route began recovering historical nodes.");
        }
        private static void FailuresAndCancellation(string root)
        {
            for(int variant=0;variant<8;variant++)
            {
                var catalog=new Catalog();var store=new Store(catalog);Node wrong=Copy(catalog.Old);
                if(variant==0)wrong.SiteId=Guid.NewGuid();if(variant==1)wrong.Id=Guid.NewGuid();if(variant==2)wrong.WebId=Guid.NewGuid();if(variant==3)wrong.ListId=Guid.NewGuid();
                if(variant==4)wrong.UiVersion++;if(variant==5)wrong.HistoryVersion++;if(variant==6)wrong.Level++;if(variant==7)wrong.InternalVersion++;
                catalog.ResolvedOverride=_=>wrong;using var controller=Controller(catalog,store);
                DesktopExportSummary failed=controller.ExportVersion(catalog.Old,Path.Combine(root,"identity-"+variant),CancellationToken.None,null!);
                Check(failed.Failed==1 && failed.Entries.Single().Status==RecoveryStatus.Corrupt && store.Reads==0,
                    "A different exact version key/scope reached content retrieval: "+variant);
            }
            {
                var catalog=new Catalog();var store=new Store(catalog);using var controller=Controller(catalog,store);using var cancel=new CancellationTokenSource();cancel.Cancel();
                string destination=Path.Combine(root,"pre-cancel");DesktopExportSummary cancelled=controller.ExportVersion(catalog.Old,destination,cancel.Token,null!);
                Check(cancelled.Cancelled && cancelled.Entries.Count==0 && catalog.VersionLookups==0 && store.Reads==0 && !Directory.Exists(destination),
                    "Pre-cancellation queried versions or created published output.");
                catalog.ResolvedOverride=_=>null;DesktopExportSummary missing=controller.ExportVersion(catalog.Old,Path.Combine(root,"missing"),CancellationToken.None,null!);
                Check(missing.Skipped==1 && missing.Entries.Single().Status==RecoveryStatus.Unavailable && store.Reads==0 && catalog.CurrentLookups==0,
                    "An absent historical state fell back to the current document.");
            }
            {
                var catalog=new Catalog();var store=new Store(catalog){NoChunks=true};using var controller=Controller(catalog,store);
                DesktopExportSummary missing=controller.ExportVersion(catalog.Old,Path.Combine(root,"missing-bytes"),CancellationToken.None,null!);
                Check(missing.Skipped==1 && missing.Success==0 && !Directory.GetFiles(root,"report (v1.0)*",SearchOption.AllDirectories).Any(file=>file.Contains("missing-bytes",StringComparison.Ordinal)),
                    "Unavailable historical bytes produced a published version file.");
                Node unsupported=Copy(catalog.Old);unsupported.StreamSchema=67;int lookups=catalog.VersionLookups,reads=store.Reads;
                DesktopExportSummary ignored=controller.ExportVersion(unsupported,Path.Combine(root,"known-native"),CancellationToken.None,null!);
                Check(ignored.Total==0 && ignored.Entries.Count==0 && catalog.VersionLookups==lookups && store.Reads==reads,
                    "A known unsupported historical version reached storage or warning totals.");
            }
        }
        private static ExplorerController Controller(Catalog catalog,Store store)=>new(new RecoverySession(catalog,store,DocumentDecoderRegistry.CreateDefault()));
        private static bool SameDirectory(string file,string directory)=>String.Equals(Path.GetDirectoryName(Path.GetFullPath(file)),Path.GetFullPath(directory),StringComparison.OrdinalIgnoreCase);
        private static string Digest(byte[] bytes)=>Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        private static string Key(Node node)=>node.SiteId+":"+node.Id+":"+node.WebId+":"+node.ListId+":"+node.HistoryVersion+":"+node.UiVersion+":"+node.Level+":"+node.InternalVersion;
        private static Node Copy(Node node)=>new(){Kind=node.Kind,SiteId=node.SiteId,WebId=node.WebId,ListId=node.ListId,Id=node.Id,Name=node.Name,Path=node.Path,Size=node.Size,
            StreamSchema=node.StreamSchema,HasStream=node.HasStream,HistoryVersion=node.HistoryVersion,UiVersion=node.UiVersion,InternalVersion=node.InternalVersion,Level=node.Level,Modified=node.Modified,ParentId=node.ParentId};
        private static void Check(bool value,string message){if(!value)throw new Exception(message);}
        private static void Reject<T>(Action work)where T:Exception{try{work();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name);}
        private sealed class Catalog:ISharePointCatalog,ISharePointVersionCatalog
        {
            internal readonly Node Current,Old,Draft;internal readonly List<Node> Rows=new();internal readonly Dictionary<string,byte[]> Bytes=new();
            internal Func<Node,Node?>? ResolvedOverride;internal int VersionLookups,CurrentLookups;
            public string SourceName=>"Synthetic version source";
            internal Catalog()
            {
                Current=new Node{Kind=NodeKind.File,SiteId=Guid.NewGuid(),WebId=Guid.NewGuid(),ListId=Guid.NewGuid(),Id=Guid.NewGuid(),Name="report.txt",Path="Docs/Nested/report.txt",HasStream=true,UiVersion=1024,InternalVersion=2049,Level=1,Modified=new DateTime(2025,2,3,4,5,6,DateTimeKind.Utc)};
                Add(Current,"current document content\n");Old=Copy(Current);Old.HistoryVersion=512;Old.UiVersion=512;Old.InternalVersion=1025;Old.Modified=Current.Modified.AddDays(-2);Add(Old,"previous approved content\n");
                Draft=Copy(Old);Draft.HistoryVersion=513;Draft.UiVersion=513;Draft.InternalVersion=1026;Draft.Level=2;Add(Draft,"previous draft content\n");
                Rows.Add(Draft);Rows.Add(Old);Rows.Add(Current);
            }
            private void Add(Node node,string text){byte[] data=Encoding.UTF8.GetBytes(text);node.Size=data.Length;Bytes.Add(Key(node),data);}
            public List<Node> GetFileVersions(Guid siteId,Guid fileId)=>siteId==Current.SiteId&&fileId==Current.Id?Rows.Select(Copy).ToList():new();
            public Node GetFileVersion(Node selected)
            {
                VersionLookups++;if(ResolvedOverride!=null)return ResolvedOverride(selected)!;
                Node? match=Rows.FirstOrDefault(row=>Key(row)==Key(selected));return match==null?null!:Copy(match);
            }
            public Node GetFile(Guid siteId,Guid fileId){CurrentLookups++;return siteId==Current.SiteId&&fileId==Current.Id?Copy(Current):null!;}
            public List<Node> GetRootSites()=>new();public List<Node> GetChildren(Node parent)=>new(){Copy(Current)};
            public IEnumerable<Node> EnumerateCurrentFiles(Guid? siteId)=>new[]{Copy(Current)};public void ValidateSchema(){}public List<string> CheckDatabase()=>new();
        }
        private sealed class CurrentOnlyCatalog:ISharePointCatalog
        {
            private readonly Catalog source;internal CurrentOnlyCatalog(Catalog source){this.source=source;}
            public string SourceName=>source.SourceName;public Node GetFile(Guid siteId,Guid fileId)=>source.GetFile(siteId,fileId);
            public List<Node> GetRootSites()=>new();public List<Node> GetChildren(Node parent)=>new();public IEnumerable<Node> EnumerateCurrentFiles(Guid? siteId)=>Array.Empty<Node>();
            public void ValidateSchema(){}public List<string> CheckDatabase()=>new();
        }
        private sealed class Store:IDocumentChunkStore
        {
            private readonly Catalog source;internal int Reads;internal Node? Last;internal bool NoChunks;
            internal Store(Catalog source){this.source=source;}
            public IList<StoredChunk> ReadChunks(Node document)
            {
                Reads++;Last=Copy(document);if(NoChunks)return new List<StoredChunk>();
                return new List<StoredChunk>{new(){Partition=0,Content=(byte[])source.Bytes[Key(document)].Clone()}};
            }
        }
    }
}