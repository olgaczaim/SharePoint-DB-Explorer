using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using SharePointExplorer;

namespace SharePointExplorer.Tests
{
    public static class RecoveryEngineChecks
    {
        public static void Run()
        {
            TestVerificationAndScope();
            TestRoutingAndAvailability();
            TestSessionLifetime();
            TestPublication();
            Console.WriteLine("PASS recovery provider contracts, scope, routing, checksums and atomic publication");
        }
        private static void TestVerificationAndScope()
        {
            byte[] expected=Encoding.ASCII.GetBytes("betaalpha");
            var catalog=new FakeCatalog(Document(66,expected.Length));
            var store=new FakeStore(StorageFixture.Create(Encoding.ASCII.GetBytes("alpha"),Encoding.ASCII.GetBytes("beta")));
            var engine=new RecoveryEngine(catalog,store,DocumentDecoderRegistry.CreateDefault());
            Node selected=Document(77,1234);
            RecoveryResult verified=engine.Verify(selected);
            Check(verified.Bytes==expected.Length && verified.Sha256==Digest(expected),"Verify did not reconstruct/hash the authoritative source content.");
            Check(verified.Document.StreamSchema==66 && verified.Document.Size==9 && !String.IsNullOrEmpty(verified.Decoder),"Verify returned the caller's stale document metadata.");
            PreparedDocument prepared=engine.Prepare(selected);
            prepared.Document.Size=123;
            catalog.Document.Size=321;
            using(var output=new MemoryStream())
            {
                RecoveryResult recovered=prepared.Recover(output);
                Check(recovered.Bytes==9 && recovered.Sha256==verified.Sha256 && Equal(output.ToArray(),expected),"Prepared document metadata changed after preparation.");
                Check(output.CanWrite,"Recovery closed the caller's output stream.");
            }
            catalog.Document.Size=9;
            catalog.Document.Id=Guid.NewGuid();
            int reads=store.Reads;
            Throws<InvalidDataException>(delegate { engine.Prepare(selected); },"a catalog returning a different document");
            Check(store.Reads==reads,"A different document scope reached the chunk store.");
            catalog.Document=Document(66,9);
            catalog.Document.SiteId=Guid.NewGuid();
            Throws<InvalidDataException>(delegate { engine.Prepare(selected); },"a catalog returning a different site collection");
            Check(store.Reads==reads,"A different site scope reached the chunk store.");
            selected.HistoryVersion=1;
            int lookups=catalog.Lookups;
            Throws<NotSupportedException>(delegate { engine.Prepare(selected); },"historical scope silently replaced by current content");
            Check(catalog.Lookups==lookups && store.Reads==reads,"Historical selection reached the current catalog/storage provider.");
        }
        private static void TestRoutingAndAvailability()
        {
            var registry=DocumentDecoderRegistry.CreateDefault();
            Check(registry.Resolve(0).Supports(0) && registry.Resolve(66).Supports(66),"Default decoder routing is incomplete.");
            foreach(byte schema in new byte[] {67,77,255})
            {
                var catalog=new FakeCatalog(Document(schema,4));
                var store=new FakeStore(new List<StoredChunk>());
                var engine=new RecoveryEngine(catalog,store,registry);
                Exception error=Capture(delegate { engine.Verify(catalog.Document); });
                Check(error is NotSupportedException && store.Reads==0,"Unsupported content was read before schema routing.");
                Check(error.Message.Contains("Document storage schema "+schema),"An unverified storage schema lost its explanatory error.");
            }
            var plainCatalog=new FakeCatalog(Document(0,3));
            var missingStore=new FakeStore(new List<StoredChunk>());
            var missingEngine=new RecoveryEngine(plainCatalog,missingStore,registry);
            Exception missing=Capture(delegate { missingEngine.Verify(plainCatalog.Document); });
            Check(missing is ContentUnavailableException && RecoveryErrors.Classify(missing)==RecoveryStatus.Unavailable,"Missing document bytes were not classified as unavailable.");
            var genericCatalog=new FakeCatalog(Document(66,3));
            var rawOnly=new FakeStore(new List<StoredChunk> { Chunk(new byte[] {1,2,3}) });
            Exception missingMain=Capture(delegate { new RecoveryEngine(genericCatalog,rawOnly,registry).Verify(genericCatalog.Document); });
            Check(missingMain is ContentUnavailableException,"A generic document without its primary stream was not classified as unavailable.");
            missingStore.Failure=new ContentUnavailableException("A template or external content provider is needed.");
            Check(RecoveryErrors.Classify(Capture(delegate { missingEngine.Verify(plainCatalog.Document); }))==RecoveryStatus.Unavailable,"Provider availability errors lost their status.");
            byte[] plain=new byte[] {1,2,3};
            var resident=new FakeStore(new List<StoredChunk> { Chunk(plain) });
            RecoveryResult plainResult=new RecoveryEngine(plainCatalog,resident,registry).Verify(plainCatalog.Document);
            Check(plainResult.Bytes==3 && plainResult.Sha256==Digest(plain),"Resident plain stream recovery did not use the common checksum path.");
            var emptyCatalog=new FakeCatalog(Document(0,0));
            RecoveryResult empty=new RecoveryEngine(emptyCatalog,new FakeStore(new List<StoredChunk>()),registry).Verify(emptyCatalog.Document);
            Check(empty.Bytes==0 && empty.Sha256==Digest(new byte[0]),"A resident empty file was not verified correctly.");
            var customCatalog=new FakeCatalog(Document(7,3));
            var customRegistry=new DocumentDecoderRegistry(new IDocumentDecoder[] {new TestDecoder(7,plain,false)});
            var customEngine=new RecoveryEngine(customCatalog,resident,customRegistry);
            Check(customEngine.Verify(customCatalog.Document).Sha256==Digest(plain),"A registered content decoder was not reusable by the engine.");
            var shortRegistry=new DocumentDecoderRegistry(new IDocumentDecoder[] {new TestDecoder(7,new byte[] {1},false)});
            Throws<InvalidDataException>(delegate {new RecoveryEngine(customCatalog,resident,shortRegistry).Verify(customCatalog.Document);},"a decoder returning too few bytes");
            var longRegistry=new DocumentDecoderRegistry(new IDocumentDecoder[] {new TestDecoder(7,new byte[] {1,2,3,4},false)});
            using(var output=new MemoryStream())
            {
                Throws<InvalidDataException>(delegate {new RecoveryEngine(customCatalog,resident,longRegistry).Recover(customCatalog.Document,output);},"a decoder exceeding its byte limit");
                Check(output.Length==0 && output.CanWrite,"Excess output was written or the caller's stream was closed.");
            }
            var ambiguous=new DocumentDecoderRegistry(new IDocumentDecoder[] {new TestDecoder(7,plain,false),new TestDecoder(7,plain,false)});
            Throws<InvalidOperationException>(delegate {ambiguous.Resolve(7);},"ambiguous decoder registration");
        }
        private static void TestSessionLifetime()
        {
            var catalog=new FakeCatalog(Document(0,3));
            var store=new FakeStore(new List<StoredChunk> {Chunk(new byte[] {1,2,3})});
            var session=new RecoverySession(catalog,store,DocumentDecoderRegistry.CreateDefault());
            RecoveryEngine engine=session.Engine;
            PreparedDocument prepared=engine.Prepare(catalog.Document);
            DocumentExporter exporter=session.Exporter;
            session.Dispose(); session.Dispose();
            Check(catalog.Disposals==1 && store.Disposals==1,"Session resource ownership/disposal is inconsistent.");
            Throws<ObjectDisposedException>(delegate {var ignored=session.Catalog;},"catalog use after session disposal");
            Throws<ObjectDisposedException>(delegate {engine.Verify(catalog.Document);},"a retained engine after session disposal");
            Throws<ObjectDisposedException>(delegate {prepared.Recover(Stream.Null);},"a prepared document after session disposal");
            Throws<ObjectDisposedException>(delegate {exporter.Export(catalog.Document,"unused");},"a retained exporter after session disposal");
        }
        private static void TestPublication()
        {
            string root=Path.GetFullPath(Path.Combine(Environment.CurrentDirectory,".scratch","engine-checks",Guid.NewGuid().ToString("N")));
            byte[] expected=Encoding.ASCII.GetBytes("file");
            var catalog=new FakeCatalog(Document(0,expected.Length));
            var store=new FakeStore(new List<StoredChunk> {Chunk(expected)});
            var exporter=new DocumentExporter(new RecoveryEngine(catalog,store,DocumentDecoderRegistry.CreateDefault()));
            string directory=Path.Combine(root,"complete");
            ExportResult first=exporter.Export(catalog.Document,directory);
            ExportResult second=exporter.Export(catalog.Document,directory);
            Check(first.Path!=second.Path && Path.GetFileName(second.Path)=="capture (2).bin","Repeated recovery replaced an existing export.");
            Check(Equal(File.ReadAllBytes(first.Path),expected) && Equal(File.ReadAllBytes(second.Path),expected),"File publication changed recovered content.");
            Check(Directory.GetFiles(directory,"*.partial").Length==0,"Successful recovery left a temporary file.");
            var failureCatalog=new FakeCatalog(Document(7,expected.Length));
            var failureRegistry=new DocumentDecoderRegistry(new IDocumentDecoder[] {new TestDecoder(7,new byte[] {1,2},true)});
            var failing=new DocumentExporter(new RecoveryEngine(failureCatalog,store,failureRegistry));
            string incomplete=Path.Combine(root,"incomplete");
            Throws<InvalidDataException>(delegate {failing.Export(failureCatalog.Document,incomplete);},"a decoder failure after partial output");
            Check(Directory.Exists(incomplete) && Directory.GetFiles(incomplete).Length==0,"Failed recovery published or retained a partial file.");
        }
        private static Node Document(byte schema,long size)
        {
            return new Node {Kind=NodeKind.File,SiteId=new Guid("11111111-1111-1111-1111-111111111111"),Id=new Guid("22222222-2222-2222-2222-222222222222"),Name="capture.bin",Path="library/capture.bin",Size=size,StreamSchema=schema,Level=1,HasStream=true};
        }
        private static StoredChunk Chunk(byte[] bytes) {return new StoredChunk {BSN=1,StreamId=2,Partition=0,Type=0,Content=bytes};}
        private static string Digest(byte[] bytes) {using(var hash=SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-","").ToLowerInvariant();}
        private static bool Equal(byte[] a,byte[] b) {if(a.Length!=b.Length)return false;for(int i=0;i<a.Length;i++)if(a[i]!=b[i])return false;return true;}
        private static void Check(bool condition,string message) {if(!condition)throw new Exception(message);}
        private static Exception Capture(Action action) {try{action();}catch(Exception error){return error;}throw new Exception("The operation unexpectedly succeeded.");}
        private static void Throws<T>(Action action,string reason) where T:Exception {Exception error=Capture(action);if(!(error is T))throw new Exception("Wrong error for "+reason+": "+error.GetType().Name,error);}
        private sealed class FakeCatalog : ISharePointCatalog,IDisposable
        {
            internal Node Document;
            internal int Lookups,Disposals;
            internal FakeCatalog(Node document) {Document=document;}
            public string SourceName {get{return "Memory restored content";}}
            public void ValidateSchema() { }
            public List<string> CheckDatabase() {return new List<string>();}
            public List<Node> GetRootSites() {return new List<Node>();}
            public List<Node> GetChildren(Node parent) {return new List<Node>();}
            public Node GetFile(Guid siteId,Guid fileId) {Lookups++;return Document;}
            public IEnumerable<Node> EnumerateCurrentFiles(Guid? siteId) {return new Node[] {Document};}
            public void Dispose() {Disposals++;}
        }
        private sealed class FakeStore : IDocumentChunkStore,IDisposable
        {
            private readonly IList<StoredChunk> chunks;
            internal int Reads,Disposals;
            internal Exception Failure;
            internal FakeStore(IList<StoredChunk> chunks) {this.chunks=chunks;}
            public IList<StoredChunk> ReadChunks(Node document) {Reads++;if(Failure!=null)throw Failure;return chunks;}
            public void Dispose() {Disposals++;}
        }
        private sealed class TestDecoder : IDocumentDecoder
        {
            private readonly byte schema;
            private readonly byte[] content;
            private readonly bool fail;
            internal TestDecoder(byte schema,byte[] content,bool fail) {this.schema=schema;this.content=content;this.fail=fail;}
            public string Name {get{return "Test codec";}}
            public bool Supports(byte value) {return value==schema;}
            public void Write(IList<StoredChunk> chunks,long expectedSize,Stream output) {output.Write(content,0,content.Length);if(fail)throw new InvalidDataException("Injected decoder failure.");}
        }
    }
}