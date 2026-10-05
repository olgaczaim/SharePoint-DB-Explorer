#nullable enable
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using SharePointExplorer;
using SharePointExplorer.Desktop.ViewModels;

namespace SharePointExplorer.Tests
{
    internal static class NativeBlobRecoveryChecks
    {
        public static void Run()
        {
            SqlProjection();
            RoutingAndRecovery();
            PublicationAndZip();
            Console.WriteLine("PASS native resident schemas 1/65, SQL empty/nonempty RBS references, exact bytes/version scope, corruption and ZIP checks");
        }

        private static void SqlProjection()
        {
            byte[] bytes={0,255,13,10,97,0,42};
            foreach(byte[]? reference in new byte[]?[] {null,Array.Empty<byte>()})
            {
                StoredChunk chunk=Row(reference,bytes,bytes.Length);
                Check(chunk.BSN==17 && chunk.StreamId==1 && chunk.Partition==0 && chunk.Type==0 &&
                    chunk.Content.SequenceEqual(bytes),"An absent/empty RBS reference rejected or changed inline native bytes.");
            }
            foreach(byte[] reference in new[] {new byte[] {1},new byte[] {0,0}})
                Reject<ContentUnavailableException>(()=>Row(reference,bytes,bytes.Length));
            Reject<ContentUnavailableException>(()=>Row(Array.Empty<byte>(),null,0,abs:true));
            Reject<InvalidDataException>(()=>Row(Array.Empty<byte>(),null,0));
            Reject<InvalidDataException>(()=>Row(Array.Empty<byte>(),bytes,bytes.Length+1));
            Reject<NotSupportedException>(()=>Row(Array.Empty<byte>(),bytes,bytes.Length,compressed:true));
            Check(Row(Array.Empty<byte>(),Array.Empty<byte>(),0).Content.Length==0,"An empty native inline BLOB was rejected.");
            // Metadata-only files remain unavailable, even though native schema
            // routing now exists. This guard runs before opening a SQL connection.
            var store=new SqlDocumentChunkStore(new SqlRepository("unused-server","unused-database"));
            Reject<ContentUnavailableException>(()=>store.ReadChunks(new Node {StreamSchema=1,HasStream=false}));
        }

        private static StoredChunk Row(byte[]? rbs,byte[]? bytes,int size,bool abs=false,bool compressed=false)
        {
            using var table=new DataTable();
            Type[] types={typeof(long),typeof(long),typeof(byte),typeof(byte),typeof(int),typeof(byte[]),typeof(int),typeof(Guid),typeof(byte[])};
            for(int index=0;index<types.Length;index++)table.Columns.Add("c"+index,types[index]);
            table.Rows.Add(17L,1L,(byte)0,(byte)0,size,(object?)rbs??DBNull.Value,
                compressed ? (object)size : DBNull.Value,abs ? (object)Guid.NewGuid() : DBNull.Value,(object?)bytes??DBNull.Value);
            using DataTableReader reader=table.CreateDataReader();
            Check(reader.Read(),"The native SQL projection fixture is empty.");
            return SqlDocumentChunkStore.ReadChunk(reader);
        }

        private static void RoutingAndRecovery()
        {
            DocumentDecoderRegistry registry=DocumentDecoderRegistry.CreateDefault();
            Check(registry.Supports(1) && registry.Supports(65) && registry.Resolve(1).Name=="Resident plain stream" && registry.Resolve(65).Name=="Resident plain stream","Native schemas 1/65 are not routed to resident recovery.");
            // Readers are chosen from the schema flags; host tagging does not change the reader.
            Check(registry.Resolve(2).Name=="SharePoint generic document tree" && registry.Resolve(64).Name=="Resident plain stream","Schema flags did not select the shredded/plain reader.");
            foreach(byte schema in new byte[] {3,4,8,17,33,67,128,193,255})
                Check(!registry.Supports(schema),"An undefined host/cell layout was enabled: "+schema);
            foreach(byte schema in new byte[] {1,65})
            {
                byte[] expected={0,255,128,0,13,10,42};
                var catalog=new Catalog(expected);catalog.Current.StreamSchema=schema;catalog.Historical.StreamSchema=schema;
                var store=new Store(catalog);
                using var session=new RecoverySession(catalog,store,registry);
                Check(session.Engine.CanExport(catalog.Current),"A supported stored native file is ineligible.");
                var currentRow=new ItemViewModel(catalog.Current);currentRow.IsChecked=true;
                Check(currentRow.CanCheck && currentRow.IsChecked && !new ItemViewModel(catalog.Historical).CanCheck,
                    "Native current files cannot be checked, or history entered the current-file selection.");
                Check(new VersionViewModel(catalog.Current).CanExport && new VersionViewModel(catalog.Historical).CanExport,
                    "Native current or historical versions are unavailable in the version dialog.");
                Node unavailable=Copy(catalog.Current);unavailable.HasStream=false;
                var unavailableRow=new ItemViewModel(unavailable);unavailableRow.IsChecked=true;
                Node unavailableHistory=Copy(catalog.Historical);unavailableHistory.HasStream=false;
                Check(!session.Engine.CanExport(unavailable) && !unavailableRow.CanCheck && !unavailableRow.IsChecked &&
                    !new VersionViewModel(unavailable).CanExport && !new VersionViewModel(unavailableHistory).CanExport,
                    "Native schema routing concealed a metadata-only current or historical file.");
                using var output=new MemoryStream();
                RecoveryResult recovered=session.Engine.Recover(catalog.Current,output);
                Check(output.ToArray().SequenceEqual(expected) && recovered.Sha256==Digest(expected) && recovered.Bytes==expected.Length && output.CanWrite,
                    "Current native binary recovery changed bytes, hash, size or stream ownership.");
                PreparedDocument old=session.Engine.PrepareVersion(catalog.Historical);
                using var older=new MemoryStream();
                RecoveryResult previous=old.Recover(older);
                Check(previous.Document.HistoryVersion==512 && older.ToArray().SequenceEqual(catalog.PreviousBytes) &&
                    previous.Sha256==Digest(catalog.PreviousBytes) && previous.Sha256!=recovered.Sha256,
                    "Native history recovery fell back to current bytes.");
                IDocumentDecoder decoder=registry.Resolve(schema);
                using var untouched=new MemoryStream();
                Reject<InvalidDataException>(()=>decoder.Write(new List<StoredChunk> {Chunk(expected)},expected.Length+1,untouched));
                Check(untouched.Length==0,"Native corrupt declared length wrote output bytes.");
                Reject<NotSupportedException>(()=>decoder.Write(new List<StoredChunk> {Chunk(expected),Chunk(expected)},expected.Length*2,untouched));
                StoredChunk otherPartition=Chunk(expected);otherPartition.Partition=1;
                Reject<NotSupportedException>(()=>decoder.Write(new List<StoredChunk> {otherPartition},expected.Length,untouched));
                Reject<ContentUnavailableException>(()=>decoder.Write(Array.Empty<StoredChunk>(),1,untouched));
                decoder.Write(Array.Empty<StoredChunk>(),0,untouched);
                Check(untouched.Length==0,"Empty native file recovery invented bytes.");
            }
        }

        private static void PublicationAndZip()
        {
            string root=Path.GetFullPath(Path.Combine(".scratch","native-blob-checks",Guid.NewGuid().ToString("N")));
            byte[] bytes={0,250,42,13,10,0};
            var catalog=new Catalog(bytes);catalog.Current.StreamSchema=65;catalog.Historical.StreamSchema=65;
            using var session=new RecoverySession(catalog,new Store(catalog),DocumentDecoderRegistry.CreateDefault());
            PreparedDocument prepared=session.Engine.Prepare(catalog.Current);
            Directory.CreateDirectory(root);
            string existing=Path.Combine(root,catalog.Current.Name);
            File.WriteAllText(existing,"preserve existing");
            ExportResult file=session.Exporter.Export(prepared,root);
            Check(file.Path!=existing && File.ReadAllText(existing)=="preserve existing" &&
                File.ReadAllBytes(file.Path).SequenceEqual(bytes) && file.Sha256==Digest(bytes),
                "Native file publication overwrote an existing file or changed bytes.");
            string archivePath;
            ExportResult entry;
            using(var zip=new ValidatedZipArchive(Path.Combine(root,"native.zip")))
            {
                entry=zip.Add(prepared);
                archivePath=zip.Commit(writer=>writer.WriteLine("DocumentId,Sha256\r\n"+catalog.Current.Id+","+entry.Sha256),"One verified native file");
            }
            using(ZipArchive zip=ZipFile.OpenRead(archivePath))
            {
                ZipArchiveEntry native=zip.GetEntry(entry.Path)??throw new Exception("Native ZIP entry is absent.");
                using Stream input=native.Open();using var restored=new MemoryStream();input.CopyTo(restored);
                Check(restored.ToArray().SequenceEqual(bytes) && Digest(restored.ToArray())==entry.Sha256 &&
                    zip.GetEntry(ValidatedZipArchive.ReportEntry)!=null,"Native ZIP content or report integrity differs.");
            }
            catalog.Current.Size++;
            Reject<InvalidDataException>(()=>session.Exporter.Export(catalog.Current,root));
            Check(!Directory.GetFiles(root,"*.partial").Any(),"Native corruption left a temporary export.");
        }

        private static StoredChunk Chunk(byte[] bytes)=>new() {BSN=17,StreamId=1,Partition=0,Type=0,Content=bytes};
        private static string Digest(byte[] bytes)=>Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        private static void Check(bool value,string message){if(!value)throw new Exception(message);}
        private static void Reject<T>(Action action) where T:Exception{try{action();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name);}
        private static Node Copy(Node source)=>new() {Kind=source.Kind,SiteId=source.SiteId,WebId=source.WebId,ListId=source.ListId,Id=source.Id,
            Name=source.Name,Path=source.Path,ParentId=source.ParentId,Size=source.Size,StreamSchema=source.StreamSchema,HasStream=source.HasStream,
            Level=source.Level,UiVersion=source.UiVersion,InternalVersion=source.InternalVersion,HistoryVersion=source.HistoryVersion,Modified=source.Modified};
        private sealed class Catalog : ISharePointCatalog,ISharePointVersionCatalog
        {
            internal readonly byte[] Bytes,PreviousBytes={3,2,1,0};
            internal readonly Node Current,Historical;
            internal Catalog(byte[] bytes)
            {
                Bytes=bytes;Current=new Node {Kind=NodeKind.File,SiteId=Guid.NewGuid(),WebId=Guid.NewGuid(),ListId=Guid.NewGuid(),Id=Guid.NewGuid(),
                    Name="resident.aspx",Path="Documents/resident.aspx",Size=bytes.Length,StreamSchema=1,HasStream=true,Level=1,UiVersion=1024,
                    InternalVersion=1025,Modified=new DateTime(2020,1,2,3,4,5,DateTimeKind.Utc)};
                Historical=Copy(Current);Historical.UiVersion=512;Historical.HistoryVersion=512;Historical.InternalVersion=513;Historical.Size=PreviousBytes.Length;
            }
            public string SourceName=>"Native resident fixture";
            public Node GetFile(Guid siteId,Guid fileId)=>Copy(Current);
            public Node GetFileVersion(Node selected)=>Copy(selected.HistoryVersion==0?Current:Historical);
            public List<Node> GetFileVersions(Guid siteId,Guid fileId)=>new() {Copy(Current),Copy(Historical)};
            public void ValidateSchema(){}
            public List<string> CheckDatabase()=>new();
            public List<Node> GetRootSites()=>new();
            public List<Node> GetChildren(Node node)=>new();
            public IEnumerable<Node> EnumerateCurrentFiles(Guid? siteId)=>new[] {Copy(Current)};
        }
        private sealed class Store : IDocumentChunkStore
        {
            private readonly Catalog catalog;
            internal Store(Catalog catalog){this.catalog=catalog;}
            public IList<StoredChunk> ReadChunks(Node file)=>new List<StoredChunk> {Chunk(file.HistoryVersion==0?catalog.Bytes:catalog.PreviousBytes)};
        }
    }
}
