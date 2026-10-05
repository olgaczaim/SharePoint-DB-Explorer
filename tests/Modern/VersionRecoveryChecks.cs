#nullable enable
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace SharePointExplorer.Tests
{
    internal static class VersionRecoveryChecks
    {
        public static void Run()
        {
            string root=Path.GetFullPath(Path.Combine(".scratch","version-recovery-checks",Guid.NewGuid().ToString("N")));
            HistoricalMetadata();
            ExactBytesAndPublication(Path.Combine(root,"publication"));
            ScopeAndMutation();
            MissingAndUnsupported(Path.Combine(root,"missing"));
            IntegrityAndLifetime(Path.Combine(root,"integrity"));
            Console.WriteLine("PASS exact-version metadata, historical bytes, immutable identity, no current fallback, safe publication and lifetime");
        }
        private static void HistoricalMetadata()
        {
            using var table=new DataTable();
            table.Columns.Add("UIVersion",typeof(int));table.Columns.Add("InternalVersion",typeof(int));
            table.Columns.Add("Level",typeof(byte));table.Columns.Add("TimeCreated",typeof(DateTime));
            table.Columns.Add("Size",typeof(long));table.Columns.Add("StreamSchema",typeof(byte));table.Columns.Add("HasStream",typeof(bool));
            DateTime saved=new(2020,2,3,4,5,6,DateTimeKind.Utc);
            table.Rows.Add(512,1025,(byte)1,saved,13L,(byte)66,false);
            table.Rows.Add(1024,DBNull.Value,(byte)2,DBNull.Value,7L,(byte)0,DBNull.Value);
            table.Rows.Add(1536,1537,(byte)1,saved,DBNull.Value,(byte)66,true);
            table.Rows.Add(2048,2049,(byte)1,saved,10L,DBNull.Value,true);
            var catalog=new Catalog();Node current=catalog.Current;current.Size=999;current.StreamSchema=0;current.HasStream=true;
            using DataTableReader rows=table.CreateDataReader();
            Check(rows.Read(),"Historical mapper fixture is empty.");
            Node historical=SqlVersionCatalog.ReadHistorical(rows,current);
            Check(historical.HistoryVersion==512 && historical.UiVersion==512 && historical.InternalVersion==1025 && historical.Level==1 &&
                historical.Size==13 && historical.StreamSchema==66 && historical.HasStream==false && historical.Modified==saved,
                "Historical content metadata was replaced by current metadata or confused UI/internal version counters.");
            Check(historical.SiteId==current.SiteId && historical.Id==current.Id && historical.WebId==current.WebId && historical.ListId==current.ListId &&
                historical.Name==current.Name && historical.Path==current.Path,"Historical document identity/location changed during mapping.");
            Check(rows.Read(),"Nullable historical mapper fixture is missing.");
            Node nullable=SqlVersionCatalog.ReadHistorical(rows,current);
            Check(nullable.InternalVersion==0 && nullable.HasStream==null && nullable.Modified==DateTime.MinValue && nullable.Size==7 &&
                nullable.StreamSchema==0 && nullable.Level==2,"Nullable history metadata fell back to the current row.");
            Check(rows.Read(),"Missing historical length fixture is missing.");
            Reject<ContentUnavailableException>(()=>SqlVersionCatalog.ReadHistorical(rows,current));
            Check(rows.Read(),"Missing historical schema fixture is missing.");
            Reject<NotSupportedException>(()=>SqlVersionCatalog.ReadHistorical(rows,current));
        }
        private static void ExactBytesAndPublication(string root)
        {
            var catalog=new Catalog();var store=new Store(catalog);
            using var session=new RecoverySession(catalog,store,DocumentDecoderRegistry.CreateDefault());
            Node selected=Copy(catalog.Historical);
            Check(session.Engine.CanExport(selected),"Metadata eligibility rejected a supported historical version.");
            Reject<NotSupportedException>(()=>session.Engine.Prepare(selected));
            Check(catalog.CurrentReads==0 && catalog.VersionReads==0 && store.Reads.Count==0,"Ordinary export attempted to fetch historical/current content.");
            PreparedDocument prepared=session.Engine.PrepareVersion(selected);
            Check(catalog.CurrentReads==0 && catalog.VersionReads==1 && store.Reads.SequenceEqual(new[] {512}),"Version recovery fetched another history or a current state.");
            // Identity and content metadata are frozen when preparation completes.
            selected.HistoryVersion=0;selected.UiVersion=1024;selected.InternalVersion=1281;selected.Size=1;selected.Name="changed.bin";
            using var output=new MemoryStream();RecoveryResult recovered=prepared.Recover(output);
            Check(output.CanWrite && output.ToArray().SequenceEqual(catalog.Content[512]) && recovered.Sha256==Digest(catalog.Content[512]) &&
                recovered.Bytes==catalog.Content[512].Length && recovered.Document.HistoryVersion==512 && prepared.Document.InternalVersion==1025,
                "Historical recovery returned current bytes, changed its selected state, or closed the caller stream.");
            Directory.CreateDirectory(root);string currentPath=Path.Combine(root,catalog.Current.Name);File.WriteAllBytes(currentPath,catalog.Content[0]);
            ExportResult first=session.Exporter.Export(prepared,root,"document (v1.0).txt");
            ExportResult second=session.Exporter.Export(prepared,root,"document (v1.0).txt");
            Check(Path.GetFileName(first.Path)=="document (v1.0).txt" && Path.GetFileName(second.Path)=="document (v1.0) (2).txt" &&
                File.ReadAllBytes(first.Path).SequenceEqual(catalog.Content[512]) && File.ReadAllBytes(currentPath).SequenceEqual(catalog.Content[0]),
                "Named historical publication changed old bytes or overwrote the current/prior export.");
            ExportResult safe=session.Exporter.Export(prepared,root,"../../CON:version?.txt");
            Check(Path.GetDirectoryName(safe.Path)==root && !Path.GetFileName(safe.Path).Contains(':') && !Path.GetFileName(safe.Path).Contains('?') &&
                File.ReadAllBytes(safe.Path).SequenceEqual(catalog.Content[512]),"An explicit version filename escaped its destination or lost content.");
            Reject<ArgumentException>(()=>session.Exporter.Export(prepared,root," "));
            Check(!Directory.GetFiles(root,"*.partial").Any(),"Version publication left partial files.");
            using var currentOutput=new MemoryStream();
            RecoveryResult current=session.Engine.PrepareVersion(Copy(catalog.Current)).Recover(currentOutput);
            Check(current.Document.HistoryVersion==0 && currentOutput.ToArray().SequenceEqual(catalog.Content[0]),"Explicit current-version recovery returned historical bytes.");
        }
        private static void ScopeAndMutation()
        {
            for(int variant=0;variant<10;variant++)
            {
                var catalog=new Catalog();Node changed=Copy(catalog.Historical);
                if(variant==0)changed.SiteId=Guid.NewGuid();
                if(variant==1)changed.Id=Guid.NewGuid();
                if(variant==2)changed.WebId=Guid.NewGuid();
                if(variant==3)changed.ListId=Guid.NewGuid();
                if(variant==4)changed.HistoryVersion=0;
                if(variant==5)changed.UiVersion=1024;
                if(variant==6)changed.Level=2;
                if(variant==7)changed.InternalVersion++;
                if(variant==8)changed.Kind=NodeKind.Folder;
                if(variant==9)changed.DeletionTransactionId="1234567890ABCDEF1234567890ABCDEF";
                catalog.VersionLookup=_=>changed;
                var store=new Store(catalog);using var session=new RecoverySession(catalog,store,DocumentDecoderRegistry.CreateDefault());
                Reject<InvalidDataException>(()=>session.Engine.PrepareVersion(Copy(catalog.Historical)));
                Check(catalog.VersionReads==1 && catalog.CurrentReads==0 && store.Reads.Count==0,"A different version/scope reached chunk retrieval.");
            }
            {
                var catalog=new Catalog();
                catalog.VersionLookup=query=>{query.HistoryVersion=0;query.UiVersion=catalog.Current.UiVersion;query.InternalVersion=catalog.Current.InternalVersion;return Copy(catalog.Current);};
                var store=new Store(catalog);using var session=new RecoverySession(catalog,store,DocumentDecoderRegistry.CreateDefault());
                Reject<InvalidDataException>(()=>session.Engine.PrepareVersion(Copy(catalog.Historical)));
                Check(store.Reads.Count==0,"A provider mutated its query identity to recover the current version.");
            }
            for(int variant=0;variant<7;variant++)
            {
                var catalog=new Catalog();Node selected=Copy(catalog.Historical);
                if(variant==0)selected.HistoryVersion=-1;
                if(variant==1)selected.UiVersion=1024;
                if(variant==2)selected.InternalVersion=-1;
                if(variant==3)selected.Kind=NodeKind.Library;
                if(variant==4)selected.SiteId=Guid.Empty;
                if(variant==5)selected.Id=Guid.Empty;
                if(variant==6){selected.HistoryVersion=0;selected.UiVersion=-1;}
                var store=new Store(catalog);using var session=new RecoverySession(catalog,store,DocumentDecoderRegistry.CreateDefault());
                Reject<ArgumentException>(()=>session.Engine.PrepareVersion(selected));
                Check(catalog.VersionReads==0 && store.Reads.Count==0,"An invalid selected version caused source reads.");
            }
        }
        private static void MissingAndUnsupported(string root)
        {
            {
                var catalog=new Catalog();catalog.VersionLookup=_=>null;
                var store=new Store(catalog);using var session=new RecoverySession(catalog,store,DocumentDecoderRegistry.CreateDefault());
                Reject<ContentUnavailableException>(()=>session.Engine.PrepareVersion(Copy(catalog.Historical)));
                Check(catalog.CurrentReads==0 && store.Reads.Count==0 && !Directory.Exists(root),"A missing historical row recovered current bytes or created output.");
            }
            {
                var catalog=new Catalog();var store=new Store(catalog);store.MissingHistory=true;
                using var session=new RecoverySession(catalog,store,DocumentDecoderRegistry.CreateDefault());
                Reject<ContentUnavailableException>(()=>session.Engine.PrepareVersion(Copy(catalog.Historical)));
                Check(store.Reads.SequenceEqual(new[] {512}) && catalog.CurrentReads==0,"A missing historical stream fell back to the current map.");
            }
            for(int variant=0;variant<2;variant++)
            {
                var catalog=new Catalog();Node changed=Copy(catalog.Historical);
                if(variant==0)changed.StreamSchema=67;else changed.HasStream=false;
                catalog.VersionLookup=_=>changed;
                var store=new Store(catalog);using var session=new RecoverySession(catalog,store,DocumentDecoderRegistry.CreateDefault());
                if(variant==0)Reject<NotSupportedException>(()=>session.Engine.PrepareVersion(Copy(catalog.Historical)));
                else Reject<ContentUnavailableException>(()=>session.Engine.PrepareVersion(Copy(catalog.Historical)));
                Check(catalog.CurrentReads==0 && store.Reads.Count==(variant==0 ? 0 : 1),"Unsupported/unavailable history fetched replacement current content.");
            }
            {
                var catalog=new Catalog();var store=new Store(catalog);
                using var session=new RecoverySession(new CurrentOnlyCatalog(catalog),store,DocumentDecoderRegistry.CreateDefault());
                Reject<NotSupportedException>(()=>session.Engine.PrepareVersion(Copy(catalog.Historical)));
                Check(catalog.CurrentReads==0 && store.Reads.Count==0,"A source without version support fetched current content.");
            }
        }
        private static void IntegrityAndLifetime(string root)
        {
            {
                var catalog=new Catalog();Node wrong=Copy(catalog.Historical);wrong.Size++;
                catalog.VersionLookup=_=>wrong;
                var store=new Store(catalog);using var session=new RecoverySession(catalog,store,DocumentDecoderRegistry.CreateDefault());
                PreparedDocument prepared=session.Engine.PrepareVersion(Copy(catalog.Historical));
                Reject<InvalidDataException>(()=>session.Exporter.Export(prepared,root,"broken (v1.0).txt"));
                Check(!Directory.GetFiles(root).Any() && store.Reads.SequenceEqual(new[] {512}),"Corrupt history published or retained output.");
            }
            var lifetimeCatalog=new Catalog();var lifetimeStore=new Store(lifetimeCatalog);
            var lifetime=new RecoverySession(lifetimeCatalog,lifetimeStore,DocumentDecoderRegistry.CreateDefault());
            RecoveryEngine engine=lifetime.Engine;PreparedDocument saved=engine.PrepareVersion(Copy(lifetimeCatalog.Historical));
            lifetime.Dispose();
            Reject<ObjectDisposedException>(()=>engine.PrepareVersion(Copy(lifetimeCatalog.Historical)));
            Reject<ObjectDisposedException>(()=>saved.Recover(Stream.Null));
            Reject<ObjectDisposedException>(()=>{_ = saved.Document;});
            Check(lifetimeCatalog.VersionReads==1 && lifetimeStore.Reads.Count==1,"Disposed version recovery accessed the source.");
        }
        private static string Digest(byte[] bytes)=>Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        private static void Check(bool value,string message){if(!value)throw new Exception(message);}
        private static void Reject<T>(Action action) where T:Exception {try{action();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name);}
        private static Node Copy(Node node)=>new() {Kind=node.Kind,SiteId=node.SiteId,Id=node.Id,WebId=node.WebId,ListId=node.ListId,ParentId=node.ParentId,
            Name=node.Name,Path=node.Path,Size=node.Size,Modified=node.Modified,HistoryVersion=node.HistoryVersion,UiVersion=node.UiVersion,
            InternalVersion=node.InternalVersion,Level=node.Level,StreamSchema=node.StreamSchema,HasStream=node.HasStream};
        private sealed class Catalog : ISharePointCatalog,ISharePointVersionCatalog
        {
            internal readonly Node Current=new() {Kind=NodeKind.File,SiteId=Guid.NewGuid(),Id=Guid.NewGuid(),WebId=Guid.NewGuid(),ListId=Guid.NewGuid(),ParentId=Guid.NewGuid(),
                Name="document.txt",Path="Docs/document.txt",HistoryVersion=0,UiVersion=1024,InternalVersion=1281,Level=1,StreamSchema=0,HasStream=true,Size=13};
            internal readonly Node Historical;
            internal readonly Dictionary<int,byte[]> Content=new() {[0]=Encoding.UTF8.GetBytes("current-bytes"),[512]=Encoding.UTF8.GetBytes("history-bytes")};
            internal Func<Node,Node?>? VersionLookup;
            internal int CurrentReads,VersionReads;
            internal Catalog(){Historical=Copy(Current);Historical.HistoryVersion=Historical.UiVersion=512;Historical.InternalVersion=1025;}
            public string SourceName=>"Synthetic exact-version source";
            public List<Node> GetFileVersions(Guid siteId,Guid fileId)=>new() {Copy(Current),Copy(Historical)};
            public Node GetFileVersion(Node selected)
            {
                VersionReads++;
                if(VersionLookup!=null)return VersionLookup(selected)!;
                Node state=selected.HistoryVersion==0 ? Current : Historical;
                return state.SiteId==selected.SiteId && state.Id==selected.Id && state.WebId==selected.WebId && state.ListId==selected.ListId &&
                    state.HistoryVersion==selected.HistoryVersion && state.UiVersion==selected.UiVersion && state.InternalVersion==selected.InternalVersion &&
                    state.Level==selected.Level ? Copy(state) : null!;
            }
            public Node GetFile(Guid siteId,Guid fileId){CurrentReads++;return Copy(Current);}
            public void ValidateSchema(){}
            public List<string> CheckDatabase()=>new();
            public List<Node> GetRootSites()=>new();
            public List<Node> GetChildren(Node parent)=>new();
            public IEnumerable<Node> EnumerateCurrentFiles(Guid? siteId)=>new[] {Copy(Current)};
        }
        private sealed class Store : IDocumentChunkStore
        {
            private readonly Catalog catalog;
            internal readonly List<int> Reads=new();
            internal bool MissingHistory;
            internal Store(Catalog catalog){this.catalog=catalog;}
            public IList<StoredChunk> ReadChunks(Node file)
            {
                Reads.Add(file.HistoryVersion);
                if(file.HasStream==false)throw new ContentUnavailableException("Historical template content is absent.");
                if(file.HistoryVersion>0 && MissingHistory)return new List<StoredChunk>();
                return new List<StoredChunk> {new() {Partition=0,Content=catalog.Content[file.HistoryVersion]}};
            }
        }
        private sealed class CurrentOnlyCatalog : ISharePointCatalog
        {
            private readonly Catalog catalog;internal CurrentOnlyCatalog(Catalog catalog){this.catalog=catalog;}
            public string SourceName=>catalog.SourceName;
            public Node GetFile(Guid siteId,Guid fileId)=>catalog.GetFile(siteId,fileId);
            public void ValidateSchema(){}
            public List<string> CheckDatabase()=>new();
            public List<Node> GetRootSites()=>new();
            public List<Node> GetChildren(Node parent)=>new();
            public IEnumerable<Node> EnumerateCurrentFiles(Guid? siteId)=>new[] {Copy(catalog.Current)};
        }
    }
}