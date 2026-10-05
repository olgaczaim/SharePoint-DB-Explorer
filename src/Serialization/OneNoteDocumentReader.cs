using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace SharePointExplorer
{
    // SharePoint's persisted host wrapper is expanded into the public MS-ONESTORE
    // 2.8 server packaging format. Object/property bytes remain opaque; their
    // partition, references, current cells and inherited revisions are preserved.
    internal static partial class StoredDocumentReader
    {
        private sealed class NoteObject
        {
            internal string Id;
            internal ulong Partition;
            internal byte[] Data;
            internal readonly List<string> Refs=new List<string>();
            internal readonly List<string[]> Cells=new List<string[]>();
        }
        private sealed class NoteElement
        {
            internal string Id, Revision, Basis;
            internal int Type;
            internal Guid Schema;
            internal readonly List<string[]> Roots=new List<string[]>();
            internal readonly List<string> Groups=new List<string>();
            internal readonly List<NoteObject> Objects=new List<NoteObject>();
        }
        private sealed class NoteMapping
        {
            internal int Kind;
            internal string First, Second, Target;
        }
        internal static byte[] ReconstructOneNote(IList<StoredChunk> chunks,bool toc)
        {
            if(chunks==null)throw new ArgumentNullException("chunks");
            // Structured parsing: opaque property data is never searched for
            // identifiers or signatures. Fragmented groups are reassembled.
            ShreddedStore store=ShreddedStore.Load(chunks);
            var elements=new SortedDictionary<string,NoteElement>(StringComparer.Ordinal);
            foreach(string id in store.ElementIds(PersistedStorage.StorageManifestType))
            {
                ShreddedStore.StorageManifest manifest=store.Storage(id);
                var element=new NoteElement{Id=id,Type=2,Schema=manifest.Schema};
                element.Roots.AddRange(manifest.Roots);
                elements.Add(id,element);
            }
            foreach(string id in store.ElementIds(PersistedStorage.CellManifestType))
                elements.Add(id,new NoteElement{Id=id,Type=3,Revision=store.CellRevision(id)});
            foreach(string id in store.ElementIds(PersistedStorage.RevisionManifestType))
            {
                ShreddedStore.RevisionManifest revision=store.Revision(id);
                var element=new NoteElement{Id=id,Type=4,Revision=revision.Id,Basis=revision.Basis};
                element.Roots.AddRange(revision.Roots);
                element.Groups.AddRange(revision.Groups);
                elements.Add(id,element);
            }
            foreach(string id in store.ElementIds(PersistedStorage.ObjectGroupType).Concat(store.FragmentedGroupIds()))
            {
                var element=new NoteElement{Id=id,Type=5};
                ShreddedStore.ObjectSet objects=store.GroupObjects(id);
                foreach(string key in objects.Keys)
                {
                    ShreddedStore.ObjectEntry entry=objects.Entries[key];
                    if(entry.Error!=null)throw entry.Error;
                    ShreddedStore.StoredObject stored=entry.Value;
                    if(stored.FromBlob)throw new NotSupportedException("OneNote object data BLOBs require a BLOB-preserving package serializer, which this version does not provide.");
                    var item=new NoteObject{Id=stored.Id,Partition=stored.Partition,Data=new byte[stored.Count]};
                    Buffer.BlockCopy(stored.Buffer,stored.Offset,item.Data,0,stored.Count);
                    item.Refs.AddRange(stored.References);
                    item.Cells.AddRange(stored.Cells);
                    element.Objects.Add(item);
                }
                elements.Add(id,element);
            }
            if(elements.Count>100000)throw new InvalidDataException("Too many OneNote elements.");
            var mappings=store.Mappings.Select(m=>new NoteMapping{Kind=m.Kind,First=m.First,Second=m.Second,Target=m.Target}).ToList();
            if(mappings.Count==0)throw new ContentUnavailableException("The current OneNote storage index is absent.");
            long sourceBytes=store.SourceBytes;
            ValidateNoteGraph(elements,mappings,toc);
            // Deterministic package identities make repeated recoveries comparable.
            // They identify this synthesized package, not an original desktop file.
            byte[] fingerprint;
            using(var hash=SHA256.Create())
            {
                foreach(var e in elements)
                {
                    byte[] bytes=SerializeNoteElement(e.Value,Guid.Empty,1);
                    hash.TransformBlock(bytes,0,bytes.Length,null,0);
                }
                byte[] map=Encoding.UTF8.GetBytes(MappingKey(mappings));
                hash.TransformFinalBlock(map,0,map.Length);fingerprint=hash.Hash;
            }
            var fileId=new Guid(fingerprint.Take(16).ToArray());
            var serialId=new Guid(fingerprint.Skip(16).Take(16).ToArray());
            long outputBudget=checked(sourceBytes*32+1048576);
            using(var output=new MemoryStream())
            using(var writer=new BinaryWriter(output,Encoding.UTF8,true))
            {
                writer.Write(new Guid("7b5c52e4-d88c-4da7-aeb1-5378d02996d3").ToByteArray());
                writer.Write(fileId.ToByteArray());writer.Write(serialId.ToByteArray());
                writer.Write(new Guid("638de92f-a6d4-4bc1-9a36-b3fc2511a5b7").ToByteArray());writer.Write((uint)0);
                byte[] packaging=NotePayload(w=>{WriteNoteId(w,IndexRoot);w.Write(NoteSchema(toc).ToByteArray());});
                NoteStart(writer,0x7a,true,packaging);NoteStart(writer,0x15,true,new byte[]{0});
                var serials=new Dictionary<string,ulong>(StringComparer.Ordinal);ulong sequence=1;
                foreach(var item in elements)serials.Add(item.Key,sequence++);
                byte[] indexPayload=NotePayload(w=>{WriteNoteId(w,IndexRoot);WriteNoteSerial(w,serialId,sequence);NoteUInt(w,1);});
                NoteStart(writer,1,true,indexPayload);
                foreach(NoteMapping mapping in mappings.OrderBy(m=>m.Kind).ThenBy(m=>m.First,StringComparer.Ordinal).ThenBy(m=>m.Second,StringComparer.Ordinal))
                {
                    byte[] payload=NotePayload(w=>{
                        if(mapping.Kind==2){WriteNoteId(w,mapping.First);WriteNoteId(w,mapping.Second);}
                        else if(mapping.Kind==3)WriteNoteId(w,mapping.First);
                        WriteNoteId(w,mapping.Target);WriteNoteSerial(w,serialId,serials[mapping.Target]);
                    });
                    NoteStart(writer,mapping.Kind==1?0x11:mapping.Kind==2?0x0e:0x0d,false,payload);
                }
                NoteEnd(writer,1);
                foreach(var item in elements)
                {
                    byte[] bytes=SerializeNoteElement(item.Value,serialId,serials[item.Key]);
                    if(bytes.LongLength>outputBudget-output.Length)throw new InvalidDataException("The synthesized OneNote package exceeds its reconstruction budget.");
                    writer.Write(bytes);
                }
                NoteEnd(writer,0x15);NoteEnd(writer,0x7a);
                if(output.Length>checked(sourceBytes*32+1048576))throw new InvalidDataException("The synthesized OneNote package exceeds its reconstruction budget.");
                return output.ToArray();
            }
        }
        private static Guid NoteSchema(bool toc) {return new Guid(toc?"e4dbfd38-e5c7-408b-a8a1-0e7b421e1f5f":"1f937cb4-b26f-445f-b9f8-17e20160e461");}
        private static string MappingKey(IEnumerable<NoteMapping> maps)
        {return String.Join("\n",maps.Select(m=>m.Kind+"|"+m.First+"|"+m.Second+"|"+m.Target).OrderBy(s=>s,StringComparer.Ordinal));}
        private static void ValidateNoteGraph(SortedDictionary<string,NoteElement> elements,List<NoteMapping> maps,bool toc)
        {
            NoteElement Fetch(string id,int type){NoteElement e;if(id==null || !elements.TryGetValue(id,out e))throw new ContentUnavailableException("A required OneNote data element is absent: "+id);if(e.Type!=type)throw new InvalidDataException("A OneNote index points to an unexpected element type.");return e;}
            var revisions=new Dictionary<string,NoteElement>(StringComparer.Ordinal);var cells=new Dictionary<string,NoteElement>(StringComparer.Ordinal);
            NoteMapping manifest=maps.SingleOrDefault(m=>m.Kind==1);if(manifest==null)throw new InvalidDataException("OneNote has no unique current storage manifest.");
            NoteElement storage=Fetch(manifest.Target,2);if(storage.Schema!=NoteSchema(toc))throw new InvalidDataException("OneNote storage schema disagrees with the file type.");
            foreach(NoteMapping m in maps)
            {
                if(m.Kind==3){NoteElement e=Fetch(m.Target,4);if(e.Revision!=m.First)throw new InvalidDataException("A OneNote revision mapping disagrees with its manifest.");revisions.Add(m.First,e);}
                else if(m.Kind==2)cells.Add(m.First+"|"+m.Second,Fetch(m.Target,3));
            }
            foreach(string[] root in storage.Roots)if(root[0]==null || !cells.ContainsKey(root[1]+"|"+root[2]))throw new ContentUnavailableException("A current OneNote root cell is absent.");
            foreach(string startRevision in revisions.Keys.Concat(cells.Values.Select(c=>c.Revision)).Distinct(StringComparer.Ordinal))
            {
                var chain=new List<NoteElement>();var visited=new HashSet<string>(StringComparer.Ordinal);string revision=startRevision;
                while(revision!=null)
                {
                    if(chain.Count>256 || !visited.Add(revision))throw new InvalidDataException("OneNote revision inheritance is cyclic or too deep.");
                    NoteElement r;if(!revisions.TryGetValue(revision,out r))throw new ContentUnavailableException("A required inherited OneNote revision is absent.");chain.Add(r);revision=r.Basis;
                }
                var objects=new Dictionary<string,NoteObject>(StringComparer.Ordinal);var roots=new Dictionary<string,string>(StringComparer.Ordinal);
                chain.Reverse();foreach(NoteElement r in chain)
                {
                    foreach(string[] root in r.Roots){if(root[0]==null || root[1]==null)throw new InvalidDataException("Invalid OneNote revision root.");roots[root[0]]=root[1];}
                    var inRevision=new HashSet<string>(StringComparer.Ordinal);
                    foreach(string groupId in r.Groups)foreach(NoteObject obj in Fetch(groupId,5).Objects)
                    {string key=obj.Id+"|"+obj.Partition;if(!inRevision.Add(key))throw new InvalidDataException("A OneNote revision has conflicting object partitions.");objects[key]=obj;}
                }
                foreach(string root in roots.Values)if(!objects.ContainsKey(root+"|1"))throw new ContentUnavailableException("A OneNote root object is absent from its retained revision chain.");
                foreach(NoteObject obj in objects.Values)
                {
                    foreach(string id in obj.Refs)if(id!=null && !objects.ContainsKey(id+"|1"))throw new ContentUnavailableException("A OneNote referenced object is absent from its retained revision chain.");
                    foreach(string[] pair in obj.Cells)if(!cells.ContainsKey(pair[0]+"|"+pair[1]))throw new ContentUnavailableException("A OneNote referenced cell is absent.");
                }
            }
        }
        private static byte[] SerializeNoteElement(NoteElement e,Guid serial,ulong number)
        {
            using(var output=new MemoryStream())using(var w=new BinaryWriter(output,Encoding.UTF8,true))
            {
                NoteStart(w,1,true,NotePayload(p=>{WriteNoteId(p,e.Id);WriteNoteSerial(p,serial,number);NoteUInt(p,(ulong)e.Type);}));
                if(e.Type==2)
                {NoteStart(w,0x0c,false,e.Schema.ToByteArray());foreach(string[] root in e.Roots)NoteStart(w,7,false,NotePayload(p=>{foreach(string id in root)WriteNoteId(p,id);}));}
                else if(e.Type==3)NoteStart(w,0x0b,false,NotePayload(p=>WriteNoteId(p,e.Revision)));
                else if(e.Type==4)
                {NoteStart(w,0x1a,false,NotePayload(p=>{WriteNoteId(p,e.Revision);WriteNoteId(p,e.Basis);}));foreach(string[] root in e.Roots)NoteStart(w,0x0a,false,NotePayload(p=>{foreach(string id in root)WriteNoteId(p,id);}));foreach(string group in e.Groups)NoteStart(w,0x19,false,NotePayload(p=>WriteNoteId(p,group)));}
                else if(e.Type==5)
                {
                    NoteStart(w,0x1d,true,Array.Empty<byte>());
                    foreach(NoteObject obj in e.Objects)NoteStart(w,0x18,false,NotePayload(p=>{WriteNoteId(p,obj.Id);NoteUInt(p,obj.Partition);NoteUInt(p,(ulong)obj.Data.Length);NoteUInt(p,(ulong)obj.Refs.Count);NoteUInt(p,(ulong)obj.Cells.Count);}));
                    NoteEnd(w,0x1d);NoteStart(w,0x1e,true,Array.Empty<byte>());
                    foreach(NoteObject obj in e.Objects)NoteStart(w,0x16,false,NotePayload(p=>{NoteUInt(p,(ulong)obj.Refs.Count);foreach(string id in obj.Refs)WriteNoteId(p,id);NoteUInt(p,(ulong)obj.Cells.Count);foreach(string[] cell in obj.Cells){WriteNoteId(p,cell[0]);WriteNoteId(p,cell[1]);}NoteUInt(p,(ulong)obj.Data.Length);p.Write(obj.Data);}));
                    NoteEnd(w,0x1e);
                }
                NoteEnd(w,1);return output.ToArray();
            }
        }
        private static byte[] NotePayload(Action<BinaryWriter> action){using(var s=new MemoryStream())using(var w=new BinaryWriter(s,Encoding.UTF8,true)){action(w);return s.ToArray();}}
        private static void NoteStart(BinaryWriter w,int type,bool compound,byte[] payload)
        {
            uint flags=compound?4U:0U;
            if(type<64 && payload.Length<128)w.Write((ushort)(((uint)payload.Length<<9)|((uint)type<<3)|flags));
            else {uint length=(uint)Math.Min(payload.Length,32767);w.Write((length<<17)|((uint)type<<3)|flags|2U);if(length==32767)NoteUInt(w,(ulong)payload.Length);}
            w.Write(payload);
        }
        private static void NoteEnd(BinaryWriter w,int type){if(type<64)w.Write((byte)((type<<2)|1));else w.Write((ushort)((type<<2)|3));}
        private static void NoteUInt(BinaryWriter w,ulong value)
        {
            if(value==0){w.Write((byte)0);return;}
            for(int width=1;width<=7;width++)if(value<(1UL<<(width*7)))
            {ulong encoded=(value<<width)|(1UL<<(width-1));for(int i=0;i<width;i++)w.Write((byte)(encoded>>(8*i)));return;}
            w.Write((byte)0x80);w.Write(value);
        }
        private static void WriteNoteId(BinaryWriter w,string id)
        {
            if(id==null){w.Write((byte)0);return;}
            string[] parts=id.Split(':');if(parts.Length!=3 || parts[0]!="G")throw new InvalidDataException("OneNote identity was not expanded.");
            uint value=UInt32.Parse(parts[2],CultureInfo.InvariantCulture);Guid guid=Guid.ParseExact(parts[1],"N");
            if(value<32)w.Write((byte)((value<<3)|4));
            else if(value<1024)w.Write((ushort)((value<<6)|32));
            else if(value<131072){uint encoded=(value<<7)|64;w.Write((byte)encoded);w.Write((byte)(encoded>>8));w.Write((byte)(encoded>>16));}
            else {w.Write((byte)128);w.Write(value);}
            w.Write(guid.ToByteArray());
        }
        private static void WriteNoteSerial(BinaryWriter w,Guid guid,ulong value)
        {w.Write((byte)128);w.Write(guid.ToByteArray());w.Write(value);}
    }
}