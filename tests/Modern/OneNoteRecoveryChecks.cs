using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using SharePointExplorer;

namespace SharePointExplorer.Tests
{
    internal static class OneNoteRecoveryChecks
    {
        private const string G="12345678-2222-3333-4444-555555555555";
        private static readonly byte[] Opaque=Encoding.Unicode.GetBytes("Recovered OneNote section: sample text.");
        public static void Run()
        {
            foreach(bool toc in new[]{false,true})
            {
                List<StoredChunk> chunks=Fixture(toc);
                byte[] result=StoredDocumentReader.ReconstructOneNote(chunks,toc);
                Check(new Guid(result.Take(16).ToArray())==Guid.Parse("7b5c52e4-d88c-4da7-aeb1-5378d02996d3"),"OneNote file-type header changed.");
                Check(new Guid(result.Skip(48).Take(16).ToArray())==Guid.Parse("638de92f-a6d4-4bc1-9a36-b3fc2511a5b7"),"OneNote server format header changed.");
                Check(result[^2]==0xeb && result[^1]==1,"OneNote packaging footer is absent.");
                WireCheck(result);
                Check(Contains(result,Opaque) && result.SequenceEqual(StoredDocumentReader.ReconstructOneNote(chunks,toc)),"OneNote opaque contents changed or repeated recovery is nondeterministic.");
                var catalog=new Catalog(toc,chunks);
                using var session=new RecoverySession(catalog,catalog,DocumentDecoderRegistry.CreateDefault());
                Check(session.Engine.CanExport(catalog.Current),"OneNote is not selectable for recovery.");
                PreparedDocument prepared=session.Engine.Prepare(catalog.Current);
                Check(prepared.OutputSize==result.Length && prepared.Document.Size==12345,"OneNote output/source size identities were mixed.");
                using var output=new MemoryStream();RecoveryResult recovered=prepared.Recover(output);
                Check(recovered.Bytes==result.Length && recovered.Document.Size==12345 && recovered.Sha256==Hash(result) && output.ToArray().SequenceEqual(result),"OneNote exact publication length/hash changed.");
                string publication=Path.GetFullPath(Path.Combine(".scratch","onenote-unit",Guid.NewGuid().ToString("N")));
                ExportResult published=session.Exporter.Export(prepared,publication);
                Check(File.ReadAllBytes(published.Path).SequenceEqual(result) && published.Bytes==prepared.OutputSize,"OneNote standalone publication incorrectly used logical source length.");
                using(var archive=new ValidatedZipArchive(Path.Combine(publication,"note.zip")))
                {
                    ExportResult entry=archive.Add(prepared);string zipPath=archive.Commit(w=>w.WriteLine(entry.Sha256),"OneNote package");
                    using var zip=System.IO.Compression.ZipFile.OpenRead(zipPath);
                    using Stream stream=zip.GetEntry(entry.Path).Open();using var restored=new MemoryStream();stream.CopyTo(restored);
                    Check(restored.ToArray().SequenceEqual(result) && entry.Bytes==result.Length,"OneNote ZIP publication incorrectly used logical source length.");
                }
                PreparedDocument old=session.Engine.PrepareVersion(catalog.Old);
                using var previous=new MemoryStream();old.Recover(previous);
                Check(previous.ToArray().SequenceEqual(result) && old.Document.HistoryVersion==512,"OneNote historical recovery chose current metadata.");
                foreach(string bad in new[]{"missing-base","cycle","missing-group","missing-object","missing-cell","wrong-schema","duplicate-partition","orphan-history"})
                    Reject(()=>StoredDocumentReader.ReconstructOneNote(Fixture(toc,bad),toc));
                Check(StoredDocumentReader.ReconstructOneNote(Fixture(toc,"fragmented"),toc).SequenceEqual(result),"A fragmented OneNote object group was not reassembled into the same package.");
                Reject(()=>StoredDocumentReader.ReconstructOneNote(Fixture(toc,"blob-object"),toc));
                Reject(()=>StoredDocumentReader.ReconstructOneNote(chunks,!toc));
                byte[] fragment=chunks[0].Content.Take(chunks[0].Content.Length-50).ToArray();
                Reject(()=>StoredDocumentReader.ReconstructOneNote(new[]{new StoredChunk{Partition=0,Type=11,StreamId=1,Content=fragment}},toc));
                var fakePayload=Join(Opaque,Packet(20,3,Element(1,Id(25))),Index(false),Opaque);
                byte[] fake=StoredDocumentReader.ReconstructOneNote(Fixture(toc,payload:fakePayload),toc);
                Check(Contains(fake,fakePayload),"OneNote opaque payload was scanned or truncated when it contained frame/index signatures.");
            }
            Console.WriteLine("PASS native OneNote section/notebook server packages: opaque/JCID partitions, inherited revisions, reassembled fragmented groups, exact synthesized size/hash and incomplete/cyclic/corrupt/BLOB-backed graph rejection");
        }
        public static void RunSql(SqlConnectionOptions options,SharePointExplorer.Desktop.ExplorerController controller)
        {
            Guid site=Guid.Parse("cbd6e0be-6ee8-4b9d-9b04-d9031327831b");
            string root=Path.GetFullPath(Path.Combine(".scratch","onenote-integration",Guid.NewGuid().ToString("N")));
            var files=new[]{new Node{Kind=NodeKind.File,SiteId=site,Id=Guid.Parse("95d26910-78e9-49a2-9b7b-a53665837681")},new Node{Kind=NodeKind.File,SiteId=site,Id=Guid.Parse("cf1649f3-c3ab-49bd-a506-fafef52f1b33")}};
            var saved=controller.ExportFiles(files,root,System.Threading.CancellationToken.None,null);
            Check(saved.Success==2 && saved.Failed==0 && saved.Skipped==0,"SQL native OneNote file recovery failed.");
            foreach(var entry in saved.Entries)
            {
                byte[] bytes=File.ReadAllBytes(entry.Path);WireCheck(bytes);
                Check(Hash(bytes)==entry.Sha256 && bytes.Length==entry.Bytes && bytes.Length!=entry.Document.Size,"SQL native OneNote package measurement/checksum was incorrect.");
                if(entry.Document.StreamSchema==82)Check(Contains(bytes,Encoding.Unicode.GetBytes("User1")),"Recovered OneNote properties were missing from the known section.");
            }
            var zip=controller.ExportFilesAsZip(files,Path.Combine(root,"notebook.zip"),System.Threading.CancellationToken.None,null);
            Check(zip.Success==2 && zip.Entries.All(e=>saved.Entries.Any(s=>s.Document.Id==e.Document.Id && s.Sha256==e.Sha256 && s.Bytes==e.Bytes)),"ZIP OneNote content differs from verified standalone recovery.");
            Console.WriteLine("PASS actualSQL native OneNote section and notebook index in ordinary export/ZIP: "+root);
        }
        // Independent public wire-header reader checks every compound boundary,
        // payload length and element-type sequence without the product writer.
        private static void WireCheck(byte[] bytes)
        {
            int p=68;var ends=new Stack<int>();int elementCount=0;
            while(p<bytes.Length)
            {
                int first=bytes[p],kind=first&3;
                if(kind==1 || kind==3)
                {
                    int type=kind==1?bytes[p++]>>2:BitConverter.ToUInt16(bytes,p)>>2;if(kind==3)p+=2;
                    Check(ends.Count>0 && ends.Pop()==type,"OneNote wire compound footer mismatch.");continue;
                }
                uint header;
                if(kind==0){header=BitConverter.ToUInt16(bytes,p);p+=2;}
                else if(kind==2){header=BitConverter.ToUInt32(bytes,p);p+=4;}
                else throw new Exception("Invalid wire header.");
                int typeId=(int)((header>>3)&(kind==0?63U:16383U));long size=header>>(kind==0?9:17);
                if(kind==2 && size==32767)size=(long)ReadUInt(bytes,ref p);
                Check(size>=0 && p+size<=bytes.Length,"OneNote wire payload is truncated.");
                if(typeId==1){int data=p;ReadWireId(bytes,ref data);Check(bytes[data++]==128,"Invalid wire serial number marker.");data+=24;ulong element=ReadUInt(bytes,ref data);Check(element>=1 && element<=5 && data==p+size,"Invalid wire data element descriptor.");elementCount++;}
                if((header&4)!=0)ends.Push(typeId);p=checked(p+(int)size);
            }
            Check(ends.Count==0 && elementCount>=6,"OneNote wire package is incomplete.");
        }
        private static ulong ReadUInt(byte[] b,ref int p)
        {byte first=b[p];if(first==0){p++;return 0;}if(first==128){p++;ulong v=BitConverter.ToUInt64(b,p);p+=8;return v;}int zero=0;while((first&(1<<zero))==0)zero++;int width=zero+1;ulong n=0;for(int i=0;i<width;i++)n|=(ulong)b[p++]<<(i*8);return n>>width;}
        private static void ReadWireId(byte[] b,ref int p)
        {byte x=b[p];if(x==0){p++;return;}int width=(x&7)==4?1:(x&63)==32?2:(x&127)==64?3:x==128?5:throw new Exception("Invalid wire identifier.");p+=width+16;}
        private static List<StoredChunk> Fixture(bool toc,string bad="",byte[] payload=null)
        {
            int basis=bad=="missing-base"?27:10;
            byte[] root=Element(3,Join(Id(13),Id(7)));
            byte[] cells=bad=="missing-cell"?Join(U(1),Id(11),Id(28)):U(0);
            byte[] objectBytes=Join(Id(7),U(1),U(bad=="missing-object"?1UL:0UL),bad=="missing-object"?Id(29):Array.Empty<byte>(),cells,Binary(payload??Opaque),U(0));
            byte[] jcid=Element(2,Join(Id(7),U(4),U(0),U(0),Binary(new byte[]{1,0,2,0}),U(0)));
            byte[] objects=Join(jcid,Element(2,objectBytes),bad=="duplicate-partition"?Element(2,objectBytes):Array.Empty<byte>());
            byte[] current=Packet(3,4,Join(Element(1,Join(Id(9),Id(basis))),Element(2,Array.Empty<byte>()),Element(4,U(0),Element(5,Id(bad=="missing-group"?26:5)))));
            byte[] old=Packet(4,4,Join(Element(1,Join(Id(10),bad=="cycle"?Id(9):new byte[]{0})),Element(2,Array.Empty<byte>(),root),Element(4,U(0),Element(5,Id(6)))));
            byte[] schema=new Guid((toc^(bad=="wrong-schema"))?"e4dbfd38-e5c7-408b-a8a1-0e7b421e1f5f":"1f937cb4-b26f-445f-b9f8-17e20160e461").ToByteArray();
            byte[] manifest=Packet(1,2,Join(Element(1,schema),Element(2,Array.Empty<byte>(),Element(3,Join(Id(12),Id(11),Id(14))))));
            byte[] cell=Packet(2,3,Element(1,Id(9)));
            byte[] group=Packet(5,5,Element(1,Array.Empty<byte>(),objects));
            byte[] baseGroup=Packet(6,5,Element(1,Array.Empty<byte>(),Join(jcid,Element(2,Join(Id(7),U(1),U(0),U(0),Binary(Opaque),U(0))))));
            byte[] orphan=bad=="orphan-history"?Packet(15,4,Join(Element(1,Join(Id(16),new byte[]{0})),Element(2,Array.Empty<byte>(),root),Element(4,U(0),Element(5,Id(26))))):Array.Empty<byte>();
            if(bad=="fragmented")
            {
                // The object group is stored as two fragments of its group identity.
                byte[] fragmentA=Join(Atom(5,Join(Id(5),BitConverter.GetBytes(73))),Element(2,Array.Empty<byte>(),FragmentObject(3,jcid)),new byte[]{0});
                byte[] fragmentB=Join(Atom(5,Join(Id(5),BitConverter.GetBytes(73))),Element(2,Array.Empty<byte>(),FragmentObject(3,Element(2,objectBytes))),new byte[]{0});
                group=Join(Packet(30,11,fragmentA,5),Packet(31,11,fragmentB,5));
            }
            if(bad=="blob-object")
            {
                byte[] odb=Element(3,Join(Id(7),U(1),U(0),U(0),Id(40),U(0)));
                byte[] piece=Element(1,Join(Id(40),BitConverter.GetBytes(73),U((ulong)Opaque.Length),U(0),U((ulong)Opaque.Length),Binary(Opaque)));
                group=Join(Packet(5,5,Element(1,Array.Empty<byte>(),Join(jcid,odb))),Packet(41,12,piece));
            }
            byte[] dictionary=Join(Atom(17,new Guid(G).ToByteArray()),new byte[]{5},Element(1,Join(U(2),new Guid(G).ToByteArray())),new byte[]{0,0});
            byte[] blobs=Join(new byte[]{9},manifest,cell,current,old,group,baseGroup,orphan,Index(bad=="orphan-history"),new byte[]{0});
            return new(){new StoredChunk{Partition=0,Type=11,StreamId=1,Content=Join(Atom(5,new byte[]{0}),dictionary,blobs,new byte[]{0})}};
        }
        // Converts a group state element (2) into the fragment layout (3), which
        // adds a change frequency after the partition.
        private static byte[] FragmentObject(int kind,byte[] state)
        {
            int p=1,length=BitConverter.ToUInt16(state,p)>>5;p+=2;
            byte[] item=state.Skip(p).Take(length).ToArray();
            int idLength=17;
            byte[] partition=item.Skip(idLength).Take(1).ToArray();
            byte[] rest=item.Skip(idLength+1).ToArray();
            rest=rest.Take(rest.Length-1).ToArray();
            return Element(kind,Join(item.Take(idLength).ToArray(),partition,U(0),rest));
        }
        private static byte[] Index(bool orphan)
        {
            byte[] Map(int kind,int key,int target,int second=0){byte[] pointer=Atom(5,Join(Id(target),U(1),U((ulong)kind)));return Element(kind==2?5:4,Join(Id(key),kind==2?Id(second):Array.Empty<byte>(),Binary(pointer),U(1)));}
            byte[] entries=Join(Atom(5,new byte[]{0}),Atom(9,BitConverter.GetBytes(4)),Map(1,12,1),Map(2,11,2,14),Map(3,9,3),Map(3,10,4),orphan?Map(3,16,15):Array.Empty<byte>(),new byte[]{0,0});
            return Join(Atom(13,Join(new byte[]{12},Guid.Parse("fb737e23-2247-4ec8-8d39-bda7f137215e").ToByteArray(),U(1),U(1),U(0),U(0),Binary(entries))),new byte[]{0});
        }
        // A contained moveable blob holding one data element. The serial number
        // (73) is the compressed form referencing table slot 2.
        private static byte[] Packet(int id,int type,byte[] fields,int blob=0)
        {byte[] frame=Join(new byte[]{0xba,0x32,0x4e,0x2e,5,0x24,0,1,13,13,0},Atom(5,Join(Id(id),U((ulong)type),BitConverter.GetBytes(73))),new byte[]{0,9},fields,new byte[]{0,0,0});return Join(Atom(21,Join(Id(blob==0?id:blob),U(0),U(0),U(0),U(0),Binary(frame),U(0),U(0),U(0),U(0))),new byte[]{0});}
        private static byte[] Element(int type,byte[] payload,byte[] children=null)
        {return Join(payload.Length==0?new[]{(byte)((type<<2)|1)}:Atom((type<<2)|1,payload),children??Array.Empty<byte>(),new byte[]{0});}
        private static byte[] Atom(int marker,byte[] payload){byte[] header=payload.Length<2048?BitConverter.GetBytes((ushort)((payload.Length<<5)|4)):BitConverter.GetBytes((uint)((payload.Length<<11)|2));return Join(new[]{(byte)marker},header,payload);}
        private static byte[] Id(int value){return Join(new[]{(byte)((value<<3)|4)},new Guid(G).ToByteArray());}
        private static byte[] Binary(byte[] b){return Join(U((ulong)b.Length),b);}
        private static byte[] U(ulong n){if(n==0)return new byte[]{0};for(int w=1;w<=7;w++)if(n<(1UL<<(w*7))){ulong v=(n<<w)|(1UL<<(w-1));return Enumerable.Range(0,w).Select(i=>(byte)(v>>(i*8))).ToArray();}return Join(new byte[]{128},BitConverter.GetBytes(n));}
        private static byte[] Join(params byte[][] parts){return parts.SelectMany(p=>p).ToArray();}
        private static bool Contains(byte[] haystack,byte[] needle){for(int p=0;p<=haystack.Length-needle.Length;p++)if(haystack.AsSpan(p,needle.Length).SequenceEqual(needle))return true;return false;}
        private static string Hash(byte[] b){return Convert.ToHexString(SHA256.HashData(b)).ToLowerInvariant();}
        private static void Reject(Action action){try{action();throw new Exception("Invalid OneNote graph was accepted.");}catch(InvalidDataException){}catch(ContentUnavailableException){}catch(NotSupportedException){}}
        private static void Check(bool condition,string message){if(!condition)throw new Exception(message);}
        private sealed class Catalog:ISharePointCatalog,ISharePointVersionCatalog,IDocumentChunkStore
        {
            internal Node Current,Old;private readonly List<StoredChunk> chunks;
            internal Catalog(bool toc,List<StoredChunk> chunks){this.chunks=chunks;Current=new Node{Kind=NodeKind.File,SiteId=Guid.NewGuid(),Id=Guid.NewGuid(),Size=12345,StreamSchema=(byte)(toc?98:82),HasStream=true,UiVersion=1024,InternalVersion=1025,Level=1,Name=toc?"index.onetoc2":"section.one"};Old=MigrationSnapshotCopy.Node(Current);Old.HistoryVersion=512;Old.UiVersion=512;Old.InternalVersion=513;}
            public string SourceName=>"OneNote fixture";public void ValidateSchema(){}public List<string> CheckDatabase()=>new();public List<Node> GetRootSites()=>new();public List<Node> GetChildren(Node n)=>new();public Node GetFile(Guid s,Guid d)=>MigrationSnapshotCopy.Node(Current);public IEnumerable<Node> EnumerateCurrentFiles(Guid? s)=>new[]{Current};public List<Node> GetFileVersions(Guid s,Guid d)=>new(){Current,Old};public Node GetFileVersion(Node n)=>MigrationSnapshotCopy.Node(n.HistoryVersion==0?Current:Old);public IList<StoredChunk> ReadChunks(Node n)=>chunks;
        }
    }
}