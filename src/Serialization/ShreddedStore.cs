using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SharePointExplorer
{
    // The cell model of one stored document state: storage index mappings,
    // manifests, revisions and object groups. Rows are read in order, the main
    // row first, and every data element is parsed as it is read. Only broken
    // structure is damage; a damaged row is used up to the
    // damage, the rest of it is skipped and the skip is logged. Each parsed object
    // updates independent generic state using its embedded sequence number.
    // Object data is inline, split into partial object fragments, or held in
    // object data BLOB fragments.
    internal sealed class ShreddedStore
    {
        internal const string IndexRoot="G:fb737e2322474ec88d39bda7f137215e:1";
        // Storage index key of the storage manifest entry.
        private const string ManifestKey="G:4d9eb5304da44f43b859b31cc0700665:1";
        internal const int ManifestMapping=1, CellMapping=2, RevisionMapping=3;

        internal sealed class Mapping
        {
            internal int Kind;
            internal string First, Second, Target;
        }
        internal sealed class StorageManifest
        {
            internal Guid Schema;
            internal readonly List<string[]> Roots=new List<string[]>();
        }
        internal sealed class RevisionManifest
        {
            internal string Id, Basis;
            internal readonly List<string[]> Roots=new List<string[]>();
            internal readonly List<string> Groups=new List<string>();
        }
        internal sealed class StoredObject
        {
            internal string Id;
            internal ulong Partition;
            internal List<string> References;
            internal List<string[]> Cells;
            internal byte[] Buffer;
            internal int Offset, Count;
            internal bool FromBlob;
        }
        private sealed class Piece
        {
            internal ulong Size, Start;
            internal byte[] Buffer;
            internal int Offset, Count;
        }
        private sealed class PartialObject
        {
            internal StoredObject Header;
            internal ulong Size;
            internal readonly List<Piece> Pieces=new List<Piece>();
        }
        private sealed class BlobObject
        {
            internal StoredObject Header;
            internal string BlobId;
        }

        private readonly Dictionary<string,List<PersistedStorage.DataElement>> elements=new Dictionary<string,List<PersistedStorage.DataElement>>(StringComparer.Ordinal);
        private readonly Dictionary<string,List<PersistedStorage.DataElement>> fragments=new Dictionary<string,List<PersistedStorage.DataElement>>(StringComparer.Ordinal);
        private readonly Dictionary<string,List<Piece>> blobs=new Dictionary<string,List<Piece>>(StringComparer.Ordinal);
        internal GenericDocumentStore GenericObjects { get; private set; } = new GenericDocumentStore();
        private readonly List<string[]> roots=new List<string[]>();
        private readonly List<Mapping> mappings=new List<Mapping>();
        private Dictionary<ulong,Guid> globalTable;
        internal IList<Mapping> Mappings { get { return mappings.AsReadOnly(); } }
        internal long SourceBytes { get; private set; }
        // Rows whose remainder was skipped after structural damage.
        internal int DamagedRows { get; private set; }

        // Reads every primary-partition row: the first main row (stream 1) first,
        // then the other rows as supplied; further main rows are ignored. The
        // first GUID table read serves rows that declare none.
        internal static ShreddedStore Load(IList<StoredChunk> chunks)
        {
            if(chunks==null) throw new ArgumentNullException("chunks");
            var rows=new List<StoredChunk>();
            foreach(StoredChunk chunk in chunks)
            {
                if(chunk==null) throw new InvalidDataException("A storage chunk descriptor is missing.");
                if(chunk.Partition==0) rows.Add(chunk);
            }
            StoredChunk main=rows.FirstOrDefault(row=>row.StreamId==1);
            if(main==null) throw new InvalidDataException("The main persisted stream is absent from the backup.");
            var store=new ShreddedStore();
            store.ReadRow(main);
            foreach(StoredChunk row in rows)
                if(row.StreamId!=1) store.ReadRow(row);
            foreach(StoredChunk row in rows)
                if(row.Content!=null) store.SourceBytes=checked(store.SourceBytes+row.Content.LongLength);
            return store;
        }

        // A damaged row is used up to the damage; the rest of it is skipped, the
        // skip is logged and loading continues with the next row.
        private void ReadRow(StoredChunk row)
        {
            PersistedStorage.HostBlob host=PersistedStorage.ReadHostBlob(row.Content,globalTable);
            if(globalTable==null) globalTable=host.FirstTable;
            InvalidDataException damage=host.Damage;
            try
            {
                foreach(PersistedStorage.ContainedBlob blob in host.Blobs)
                {
                    if(blob.Kind==PersistedStorage.SimpleBlob)
                    {
                        // Other simple blobs hold master data or blob locations.
                        if(blob.Type==PersistedStorage.IndexBlobType && blob.Id==IndexRoot) ReadIndex(blob);
                        continue;
                    }
                    InvalidDataException elementDamage;
                    foreach(PersistedStorage.DataElement element in PersistedStorage.ReadDataElements(blob,out elementDamage)) Add(element);
                    if(elementDamage!=null) { damage=elementDamage; break; }
                }
            }
            catch(InvalidDataException error) { damage=error; }
            if(damage==null) return;
            DamagedRows++;
            RecoveryLog.Warning("Stored chunk (stream "+row.StreamId+", BSN "+row.BSN+") is damaged: "+damage.Message+" The rest of this chunk was skipped.");
        }

        // Object records and root declarations are kept as they are parsed, so
        // those before damage in an element stay usable. The element itself is
        // registered only when it was read completely. Element types this
        // reader does not use are registered unparsed.
        private void Add(PersistedStorage.DataElement element)
        {
            if(element.Type==PersistedStorage.ObjectGroupFragmentType)
            {
                string group=FragmentOf(element);
                ReadFragment(element,record=>Keep(record,element.Sequence));
                if(group!=null) Append(fragments,group,element);
                return;
            }
            if(element.Type==PersistedStorage.ObjectDataBlobFragmentType) { ReadBlobPieces(element); return; }
            if(element.Type==PersistedStorage.StorageIndexType) ReadIndexElement(element);
            else if(element.Type==PersistedStorage.StorageManifestType) ReadStorage(element);
            else if(element.Type==PersistedStorage.CellManifestType) ReadCell(element);
            else if(element.Type==PersistedStorage.RevisionManifestType) ReadRevision(element,Declare);
            else if(element.Type==PersistedStorage.ObjectGroupType) ReadGroup(element,record=>Keep(record,element.Sequence));
            if(element.Id!=null) Append(elements,element.Id,element);
        }
        private static void Append(Dictionary<string,List<PersistedStorage.DataElement>> map,string key,PersistedStorage.DataElement element)
        {
            List<PersistedStorage.DataElement> list;
            if(!map.TryGetValue(key,out list)) { list=new List<PersistedStorage.DataElement>(); map.Add(key,list); }
            list.Add(element);
        }
        private void Declare(string root,string target) { roots.Add(new[] { root,target }); }
        // Generic states are keyed by object identity without the partition.
        // Partial records are leaves even if their header contains references.
        private void Keep(Record record,ulong sequence)
        {
            if(record.Kind==DeclarationRecord) return;
            string key=ObjectKey(record.Header.Id);
            if(record.Kind==InlineRecord)
                GenericObjects.AddInline(sequence,key,record.Header.References,
                    new ArraySegment<byte>(record.Header.Buffer,record.Header.Offset,record.Header.Count));
            else if(record.Kind==PartialRecord)
                GenericObjects.AddFragment(sequence,key,record.Piece.Start,
                    new ArraySegment<byte>(record.Piece.Buffer,record.Piece.Offset,record.Piece.Count));
            else if(record.Kind==BlobRecord)
                GenericObjects.AddBlobFallback(sequence,key,()=>ReadGenericBlob(record.BlobId));
        }
        internal static string ObjectKey(string id) { return id??String.Empty; }

        // Index entries: a cell (5) or another mapping (4) whose pointer names the
        // target. The pointer's kind is used when it is a known kind; otherwise
        // the manifest key marks the storage manifest and other keys revisions.
        private void ReadIndex(PersistedStorage.ContainedBlob blob)
        {
            var cursor=new PersistedStorage.Cursor(blob.Buffer,blob.DataStart,blob.DataEnd);
            cursor.Element(1);
            if(cursor.PeekItem()) cursor.SkipItem();
            cursor.Element(2);
            if(cursor.PeekItem()) cursor.SkipItem();
            int child;
            while((child=cursor.PeekElement())>=0)
            {
                if(child<3 || child>5) throw new InvalidDataException("The storage index contains an unknown entry.");
                cursor.Element(child);
                int end,start=cursor.Item(out end);
                if(child!=3)
                {
                    var fields=new PersistedStorage.Cursor(blob.Buffer,start,end);
                    string first=fields.Id(blob.Table), second=child==5 ? fields.Id(blob.Table) : null;
                    int length,pointer=fields.Atom(out length);
                    ulong kind;
                    string target=ReadPointer(blob,pointer,pointer+length,out kind);
                    if(child==5) AddMapping(new Mapping { Kind=CellMapping,First=first,Second=second,Target=target });
                    else if(kind==ManifestMapping || (kind!=RevisionMapping && first==ManifestKey)) AddMapping(new Mapping { Kind=ManifestMapping,First=first,Target=target });
                    else AddMapping(new Mapping { Kind=RevisionMapping,First=first,Target=target });
                }
                cursor.Position=end;
                cursor.Footer();
            }
            cursor.Footer();
            if(cursor.PeekFooter()) cursor.Footer();
        }
        // Pointer: an entry element whose item holds the target, a reference
        // count and the mapping kind; missing trailing fields leave kind 0.
        private static string ReadPointer(PersistedStorage.ContainedBlob blob,int start,int end,out ulong kind)
        {
            kind=0;
            var pointer=new PersistedStorage.Cursor(blob.Buffer,start,end);
            pointer.Element(1);
            int itemEnd,itemStart=pointer.Item(out itemEnd);
            var fields=new PersistedStorage.Cursor(blob.Buffer,itemStart,itemEnd);
            string target=fields.Id(blob.Table);
            if(fields.Position<fields.Limit) fields.UInt();
            if(fields.Position<fields.Limit) kind=fields.UInt();
            return target;
        }
        // MS-FSSHTTPB storage index data element: manifest, cell and revision entries.
        private void ReadIndexElement(PersistedStorage.DataElement element)
        {
            var cursor=Content(element);
            cursor.Element(1);
            int child;
            while((child=cursor.PeekElement())>=0)
            {
                if(child<1 || child>3) throw new InvalidDataException("A storage index data element contains an unknown entry.");
                cursor.Element(child);
                int end,start=cursor.Item(out end);
                var fields=new PersistedStorage.Cursor(element.Buffer,start,end);
                if(child==2) AddMapping(new Mapping { Kind=ManifestMapping,Target=fields.Id(element.Table) });
                else if(child==3)
                {
                    string context=fields.Id(element.Table),value=fields.Id(element.Table);
                    AddMapping(new Mapping { Kind=CellMapping,First=context,Second=value,Target=fields.Id(element.Table) });
                }
                else
                {
                    string revision=fields.Id(element.Table);
                    AddMapping(new Mapping { Kind=RevisionMapping,First=revision,Target=fields.Id(element.Table) });
                }
                cursor.Position=end;
                cursor.Footer();
            }
            cursor.Footer();
        }
        // An entry without a target is ignored; for a repeated key the first entry is kept.
        private void AddMapping(Mapping mapping)
        {
            if(mapping.Target==null) return;
            foreach(Mapping existing in mappings)
                if(existing.Kind==mapping.Kind && (mapping.Kind==ManifestMapping || (existing.First==mapping.First && existing.Second==mapping.Second))) return;
            mappings.Add(mapping);
        }

        // The object that stored revisions name for a root; the last one read wins.
        internal string LastRoot(string rootId)
        {
            string target=null;
            foreach(string[] root in roots) if(root[0]==rootId) target=root[1];
            return target;
        }

        // Resident object-data BLOB recovery is a fallback for the generic
        // reader. It is used only when no inline or partial leaf is present.
        private IList<ArraySegment<byte>> ReadGenericBlob(string id)
        {
            List<Piece> pieces;
            if(id==null || !blobs.TryGetValue(id,out pieces) || pieces.Count==0)
                throw new ContentUnavailableException("A stored object's data BLOB is absent from this content source.");
            return GenericDocumentStore.JoinFragments(pieces.Select(piece=>new GenericDocumentStore.Fragment {
                Start=piece.Start,Data=new ArraySegment<byte>(piece.Buffer,piece.Offset,piece.Count) }));
        }

        internal IEnumerable<string> ElementIds(int type)
        {
            foreach(var pair in elements) if(pair.Value[0].Type==type) yield return pair.Key;
        }
        // Group identities stored only as fragments.
        internal IEnumerable<string> FragmentedGroupIds()
        {
            foreach(string group in fragments.Keys) if(!elements.ContainsKey(group)) yield return group;
        }
        internal int? TypeOf(string id)
        {
            List<PersistedStorage.DataElement> list;
            if(id!=null && elements.TryGetValue(id,out list)) return list[0].Type;
            if(id!=null && fragments.ContainsKey(id)) return PersistedStorage.ObjectGroupType;
            return null;
        }
        private List<PersistedStorage.DataElement> Copies(string id,int type)
        {
            List<PersistedStorage.DataElement> list;
            if(id==null || !elements.TryGetValue(id,out list)) return null;
            foreach(PersistedStorage.DataElement element in list)
                if(element.Type!=type) throw new InvalidDataException("A storage manifest or data element has an unexpected type.");
            return list;
        }

        internal StorageManifest Storage(string id)
        {
            List<PersistedStorage.DataElement> copies=Copies(id,PersistedStorage.StorageManifestType);
            if(copies==null) return null;
            StorageManifest result=null;
            foreach(PersistedStorage.DataElement element in copies)
            {
                StorageManifest manifest=ReadStorage(element);
                if(result==null) result=manifest;
                else if(result.Schema!=manifest.Schema || !SameRows(result.Roots,manifest.Roots)) throw new InvalidDataException("An immutable storage manifest has conflicting stored values.");
            }
            return result;
        }
        // A repeated schema replaces the earlier one.
        private static StorageManifest ReadStorage(PersistedStorage.DataElement element)
        {
            var manifest=new StorageManifest();
            var cursor=Content(element);
            int child;
            while((child=cursor.PeekElement())>=0)
            {
                if(child==1)
                {
                    cursor.Element(1);
                    int end,start=cursor.Item(out end);
                    manifest.Schema=new PersistedStorage.Cursor(element.Buffer,start,end).Guid();
                    cursor.Position=end; cursor.Footer();
                }
                else if(child==2)
                {
                    cursor.Element(2);
                    int root;
                    while((root=cursor.PeekElement())>=0)
                    {
                        if(root!=3) throw new InvalidDataException("A storage manifest contains an unknown root.");
                        cursor.Element(3);
                        int end,start=cursor.Item(out end);
                        var fields=new PersistedStorage.Cursor(element.Buffer,start,end);
                        manifest.Roots.Add(new[] { fields.Id(element.Table),fields.Id(element.Table),fields.Id(element.Table) });
                        cursor.Position=end; cursor.Footer();
                    }
                    cursor.Footer();
                }
                else throw new InvalidDataException("A storage manifest contains an unknown field.");
            }
            return manifest;
        }

        internal string CellRevision(string id)
        {
            List<PersistedStorage.DataElement> copies=Copies(id,PersistedStorage.CellManifestType);
            if(copies==null) return null;
            string result=null;
            bool first=true;
            foreach(PersistedStorage.DataElement element in copies)
            {
                string revision=ReadCell(element);
                if(!first && result!=revision) throw new InvalidDataException("An immutable cell manifest has conflicting stored values.");
                result=revision;
                first=false;
            }
            return result;
        }
        private static string ReadCell(PersistedStorage.DataElement element)
        {
            var cursor=Content(element);
            cursor.Element(1);
            int end,start=cursor.Item(out end);
            string revision=new PersistedStorage.Cursor(element.Buffer,start,end).Id(element.Table);
            cursor.Position=end; cursor.Footer();
            return revision;
        }

        internal RevisionManifest Revision(string id)
        {
            List<PersistedStorage.DataElement> copies=Copies(id,PersistedStorage.RevisionManifestType);
            if(copies==null) return null;
            RevisionManifest result=null;
            foreach(PersistedStorage.DataElement element in copies)
            {
                RevisionManifest revision=ReadRevision(element,null);
                if(result==null) result=revision;
                else if(result.Id!=revision.Id || result.Basis!=revision.Basis || !SameRows(result.Roots,revision.Roots) || !result.Groups.SequenceEqual(revision.Groups,StringComparer.Ordinal))
                    throw new InvalidDataException("An immutable revision manifest has conflicting stored values.");
            }
            return result;
        }
        // Each root is passed to declare as soon as it is parsed. Repeated
        // identities replace earlier ones; absent group identities are skipped.
        private static RevisionManifest ReadRevision(PersistedStorage.DataElement element,Action<string,string> declare)
        {
            var revision=new RevisionManifest();
            var cursor=Content(element);
            int child;
            while((child=cursor.PeekElement())>=0)
            {
                if(child==1)
                {
                    cursor.Element(1);
                    int end,start=cursor.Item(out end);
                    var fields=new PersistedStorage.Cursor(element.Buffer,start,end);
                    revision.Id=fields.Id(element.Table);
                    revision.Basis=fields.Id(element.Table);
                    cursor.Position=end; cursor.Footer();
                }
                else if(child==2)
                {
                    cursor.Element(2);
                    int root;
                    while((root=cursor.PeekElement())>=0)
                    {
                        if(root!=3) throw new InvalidDataException("A revision manifest contains an unknown root.");
                        cursor.Element(3);
                        int end,start=cursor.Item(out end);
                        var fields=new PersistedStorage.Cursor(element.Buffer,start,end);
                        string rootId=fields.Id(element.Table),objectId=fields.Id(element.Table);
                        revision.Roots.Add(new[] { rootId,objectId });
                        if(declare!=null) declare(rootId,objectId);
                        cursor.Position=end; cursor.Footer();
                    }
                    cursor.Footer();
                }
                else if(child==4)
                {
                    cursor.Element(4);
                    if(cursor.PeekItem())
                    {
                        int end,start=cursor.Item(out end);
                        var fields=new PersistedStorage.Cursor(element.Buffer,start,end);
                        if(fields.Position<fields.Limit)
                        {
                            ulong count=fields.UInt();
                            if(count>(ulong)(end-fields.Position)) throw new InvalidDataException("A revision group list is truncated.");
                            for(ulong i=0;i<count;i++) AddGroup(revision,fields.Id(element.Table));
                        }
                        cursor.Position=end;
                    }
                    int group;
                    while((group=cursor.PeekElement())>=0)
                    {
                        if(group!=5) throw new InvalidDataException("A revision manifest contains an unknown group reference.");
                        cursor.Element(5);
                        int end,start=cursor.Item(out end);
                        AddGroup(revision,new PersistedStorage.Cursor(element.Buffer,start,end).Id(element.Table));
                        cursor.Position=end; cursor.Footer();
                    }
                    cursor.Footer();
                }
                else throw new InvalidDataException("A revision manifest contains an unknown field.");
            }
            return revision;
        }
        private static void AddGroup(RevisionManifest revision,string group)
        {
            if(group!=null) revision.Groups.Add(group);
        }

        // One resolved object, or the deferred reason it cannot be used.
        internal sealed class ObjectEntry
        {
            internal StoredObject Value;
            internal Exception Error;
        }
        // Resolved objects of one group, in stored order.
        internal sealed class ObjectSet
        {
            internal readonly List<string> Keys=new List<string>();
            internal readonly Dictionary<string,ObjectEntry> Entries=new Dictionary<string,ObjectEntry>(StringComparer.Ordinal);
        }
        private const int InlineRecord=0, BlobRecord=1, PartialRecord=2, DeclarationRecord=3;
        private sealed class Record
        {
            internal int Kind;
            internal StoredObject Header;
            internal string BlobId;
            internal Piece Piece;
        }

        // Resolves one object group from its whole copies and fragments, as a
        // OneNote package needs it. Copies of one data element must agree, and
        // one copy cannot hold an object twice. An inline state takes precedence;
        // declarations never hide data. Returns null when the group is absent.
        internal ObjectSet GroupObjects(string group)
        {
            List<PersistedStorage.DataElement> whole=Copies(group,PersistedStorage.ObjectGroupType),parts;
            fragments.TryGetValue(group??String.Empty,out parts);
            if(whole==null && parts==null) return null;
            var sources=new List<List<Record>>();
            if(whole!=null) sources.Add(Agreed(whole,false));
            if(parts!=null) foreach(var copies in parts.GroupBy(part=>part.Id,StringComparer.Ordinal)) sources.Add(Agreed(copies.ToList(),true));
            var order=new List<string>();
            var known=new HashSet<string>(StringComparer.Ordinal);
            var inline=new Dictionary<string,StoredObject>(StringComparer.Ordinal);
            var blobObjects=new Dictionary<string,BlobObject>(StringComparer.Ordinal);
            var partial=new Dictionary<string,PartialObject>(StringComparer.Ordinal);
            var conflicts=new HashSet<string>(StringComparer.Ordinal);
            foreach(List<Record> source in sources)
            {
                var states=new HashSet<string>(StringComparer.Ordinal);
                foreach(Record record in source)
                {
                    string key=Key(record.Header.Id,record.Header.Partition);
                    if((record.Kind==InlineRecord || record.Kind==BlobRecord) && !states.Add(key))
                        throw new InvalidDataException("A stored object group declares one object partition twice.");
                    if(record.Kind==DeclarationRecord) continue;
                    if(known.Add(key)) order.Add(key);
                    if(record.Kind==InlineRecord) AddInline(inline,conflicts,key,record.Header);
                    else if(record.Kind==BlobRecord) AddBlob(blobObjects,key,record.Header,record.BlobId);
                    else AddPiece(partial,key,record.Header,record.Piece);
                }
            }
            var result=new ObjectSet();
            foreach(string key in order)
            {
                ObjectEntry entry;
                StoredObject value;
                BlobObject blob;
                PartialObject pieces;
                if(conflicts.Contains(key)) entry=new ObjectEntry { Error=new InvalidDataException("An immutable stored object has conflicting stored values.") };
                else if(inline.TryGetValue(key,out value)) entry=new ObjectEntry { Value=value };
                else if(blobObjects.TryGetValue(key,out blob)) entry=Capture(()=>Resolve(blob));
                else if(partial.TryGetValue(key,out pieces)) entry=Capture(()=>Assemble(pieces.Header,pieces.Size,pieces.Pieces,false));
                else continue;
                result.Keys.Add(key);
                result.Entries.Add(key,entry);
            }
            return result;
        }
        private static List<Record> Agreed(List<PersistedStorage.DataElement> copies,bool fragment)
        {
            List<Record> first=null;
            foreach(PersistedStorage.DataElement element in copies)
            {
                List<Record> records=fragment ? ReadFragment(element,null) : ReadGroup(element,null);
                if(first==null) first=records;
                else if(!SameRecords(first,records)) throw new InvalidDataException("An immutable object group has conflicting stored values.");
            }
            return first;
        }
        private static bool SameRecords(List<Record> a,List<Record> b)
        {
            if(a.Count!=b.Count) return false;
            for(int i=0;i<a.Count;i++)
            {
                Record x=a[i],y=b[i];
                if(x.Kind!=y.Kind || x.Header.Id!=y.Header.Id || x.Header.Partition!=y.Header.Partition || x.BlobId!=y.BlobId || !SameLinks(x.Header,y.Header)) return false;
                if(x.Kind==InlineRecord && !SameObject(x.Header,y.Header)) return false;
                if(x.Kind==PartialRecord && !SamePiece(x.Piece,y.Piece)) return false;
            }
            return true;
        }
        private static bool SamePiece(Piece a,Piece b)
        {
            if(a.Size!=b.Size || a.Start!=b.Start || a.Count!=b.Count) return false;
            for(int i=0;i<a.Count;i++) if(a.Buffer[a.Offset+i]!=b.Buffer[b.Offset+i]) return false;
            return true;
        }
        private static ObjectEntry Capture(Func<StoredObject> resolve)
        {
            try { return new ObjectEntry { Value=resolve() }; }
            catch(InvalidDataException error) { return new ObjectEntry { Error=error }; }
            catch(NotSupportedException error) { return new ObjectEntry { Error=error }; }
        }
        internal static string Key(string id,ulong partition) { return id+"|"+partition.ToString(System.Globalization.CultureInfo.InvariantCulture); }

        // Object group: inline states (2), data BLOB references (3) and excluded
        // objects (4). Each record is passed to keep as soon as it is parsed.
        private static List<Record> ReadGroup(PersistedStorage.DataElement element,Action<Record> keep)
        {
            var records=new List<Record>();
            var cursor=Content(element);
            cursor.Element(1);
            int child;
            while((child=cursor.PeekElement())>=0)
            {
                cursor.Element(child);
                int end,start=cursor.Item(out end);
                var fields=new PersistedStorage.Cursor(element.Buffer,start,end);
                Record record=null;
                if(child==2) record=new Record { Kind=InlineRecord,Header=Inline(fields,element,Header(fields,element,false)) };
                else if(child==3) record=BlobReference(fields,element);
                else if(child!=4) throw new InvalidDataException("A stored object group contains an unknown object.");
                cursor.Position=end;
                cursor.Footer();
                if(record==null) continue;
                records.Add(record);
                if(keep!=null) keep(record);
            }
            cursor.Footer();
            return records;
        }
        // A whole group's BLOB reference: fields missing at the end of its item
        // stay empty.
        private static Record BlobReference(PersistedStorage.Cursor fields,PersistedStorage.DataElement element)
        {
            var header=new StoredObject { Id=fields.Id(element.Table),References=new List<string>(),Cells=new List<string[]>() };
            if(fields.Position<fields.Limit) header.Partition=fields.UInt();
            if(fields.Position<fields.Limit) header.References=References(fields,element);
            if(fields.Position<fields.Limit) header.Cells=Cells(fields,element);
            string blob=fields.Position<fields.Limit ? fields.Id(element.Table) : null;
            return new Record { Kind=BlobRecord,Header=header,BlobId=blob };
        }
        private static StoredObject Header(PersistedStorage.Cursor fields,PersistedStorage.DataElement element,bool fragment)
        {
            var header=new StoredObject { Id=fields.Id(element.Table),Partition=fields.UInt() };
            if(fragment) fields.UInt();
            header.References=References(fields,element);
            header.Cells=Cells(fields,element);
            return header;
        }
        private static StoredObject Inline(PersistedStorage.Cursor fields,PersistedStorage.DataElement element,StoredObject header)
        {
            int length;
            header.Buffer=element.Buffer;
            header.Offset=fields.Atom(out length);
            header.Count=length;
            return header;
        }
        private static string FragmentOf(PersistedStorage.DataElement element)
        {
            var cursor=Content(element);
            cursor.Element(1);
            int end,start=cursor.Item(out end);
            return new PersistedStorage.Cursor(element.Buffer,start,end).Id(element.Table);
        }
        // Object group fragment: inline states (3), data BLOB references (4),
        // descriptions (5), partial data (6) and excluded objects (7). Each record
        // is passed to keep as soon as it is parsed.
        private static List<Record> ReadFragment(PersistedStorage.DataElement element,Action<Record> keep)
        {
            var records=new List<Record>();
            var cursor=Content(element);
            cursor.Element(1);
            cursor.SkipItem();
            int child;
            while((child=cursor.PeekElement())>=0)
            {
                if(child==8) { cursor.SkipElement(); continue; }
                if(child!=2) throw new InvalidDataException("A stored object group fragment contains an unknown field.");
                cursor.Element(2);
                int kind;
                while((kind=cursor.PeekElement())>=0)
                {
                    if(kind<3 || kind>7) throw new InvalidDataException("A stored object group fragment contains an unknown object.");
                    cursor.Element(kind);
                    int end,start=cursor.Item(out end);
                    var fields=new PersistedStorage.Cursor(element.Buffer,start,end);
                    StoredObject header=Header(fields,element,true);
                    Record record;
                    if(kind==3) record=new Record { Kind=InlineRecord,Header=Inline(fields,element,header) };
                    else if(kind==4) record=new Record { Kind=BlobRecord,Header=header,BlobId=fields.Position<fields.Limit ? fields.Id(element.Table) : null };
                    else if(kind==6)
                    {
                        // The declared length is not compared with the stored bytes.
                        ulong size=fields.UInt(),offset=fields.UInt();
                        fields.UInt();
                        int length,data=fields.Atom(out length);
                        record=new Record { Kind=PartialRecord,Header=header,Piece=new Piece { Size=size,Start=offset,Buffer=element.Buffer,Offset=data,Count=length } };
                    }
                    else
                    {
                        if(fields.Position<fields.Limit) fields.UInt();
                        record=new Record { Kind=DeclarationRecord,Header=header };
                    }
                    cursor.Position=end;
                    cursor.Footer();
                    records.Add(record);
                    if(keep!=null) keep(record);
                }
                cursor.Footer();
            }
            cursor.Footer();
            return records;
        }
        // A piece without a BLOB identity is ignored; ranges are not checked.
        private void ReadBlobPieces(PersistedStorage.DataElement element)
        {
            var cursor=Content(element);
            int child;
            while((child=cursor.PeekElement())>=0)
            {
                if(child!=1) { cursor.SkipElement(); continue; }
                cursor.Element(1);
                int end,start=cursor.Item(out end);
                var fields=new PersistedStorage.Cursor(element.Buffer,start,end);
                string blob=fields.Id(element.Table);
                ulong sequence;
                if(fields.Serial(element.Table,out sequence)==null)
                    throw new InvalidDataException("A stored object data BLOB fragment has no serial number.");
                ulong size=fields.UInt(),offset=fields.UInt();
                fields.UInt();
                int length,data=fields.Atom(out length);
                cursor.Position=end;
                cursor.Footer();
                // ODB pieces are generic leaves at their FragmentOf identity.
                // Their threshold uses the serial's sequence, not host sequence.
                if(blob!=null)
                {
                    List<Piece> pieces;
                    if(!blobs.TryGetValue(blob,out pieces)) { pieces=new List<Piece>(); blobs.Add(blob,pieces); }
                    pieces.Add(new Piece { Size=size,Start=offset,Buffer=element.Buffer,Offset=data,Count=length });
                    GenericObjects.AddFragment(sequence,ObjectKey(blob),offset,
                        new ArraySegment<byte>(element.Buffer,data,length));
                }
            }
        }

        // Conflicting inline copies always conflict. Conflicting BLOB references or
        // partial declarations only matter when no inline state exists.
        private static void AddInline(Dictionary<string,StoredObject> inline,HashSet<string> conflicts,string key,StoredObject value)
        {
            StoredObject existing;
            if(!inline.TryGetValue(key,out existing)) { inline.Add(key,value); return; }
            if(!SameObject(existing,value)) conflicts.Add(key);
        }
        private static void AddBlob(Dictionary<string,BlobObject> blobObjects,string key,StoredObject header,string blob)
        {
            if(blob==null) throw new InvalidDataException("A data BLOB object reference is incomplete.");
            BlobObject existing;
            if(!blobObjects.TryGetValue(key,out existing)) { blobObjects.Add(key,new BlobObject { Header=header,BlobId=blob }); return; }
            if(existing.BlobId!=blob || !SameLinks(existing.Header,header)) existing.BlobId=null;
        }
        private static void AddPiece(Dictionary<string,PartialObject> partial,string key,StoredObject header,Piece piece)
        {
            PartialObject existing;
            if(!partial.TryGetValue(key,out existing)) { existing=new PartialObject { Header=header,Size=piece.Size }; partial.Add(key,existing); }
            else if(existing.Size!=piece.Size || !SameLinks(existing.Header,header)) existing.Size=UInt64.MaxValue;
            existing.Pieces.Add(piece);
        }
        private StoredObject Resolve(BlobObject value)
        {
            if(value.BlobId==null) throw new InvalidDataException("A stored object declares conflicting data BLOB references.");
            List<Piece> pieces;
            if(!blobs.TryGetValue(value.BlobId,out pieces) || pieces.Count==0) throw new ContentUnavailableException("A stored object's data BLOB is absent from this content source.");
            ulong size=pieces[0].Size;
            foreach(Piece piece in pieces) if(piece.Size!=size) throw new InvalidDataException("A data BLOB has conflicting declared sizes.");
            return Assemble(value.Header,size,pieces,true);
        }
        // OneNote packages: pieces must cover the whole object; overlapping bytes must agree.
        private static StoredObject Assemble(StoredObject header,ulong size,List<Piece> pieces,bool blob)
        {
            if(size==UInt64.MaxValue) throw new InvalidDataException("A stored object declares conflicting partial fragments.");
            if(size>Int32.MaxValue) throw new NotSupportedException("A fragmented stored object exceeds the supported size.");
            var data=new byte[(int)size];
            long covered=0;
            foreach(Piece piece in pieces.OrderBy(p=>p.Start))
            {
                if(piece.Start>(ulong)covered) throw new ContentUnavailableException("A fragmented stored object is missing a byte range.");
                ulong end=piece.Start+(ulong)piece.Count;
                if(end<piece.Start || end>size) throw new InvalidDataException("A stored object fragment exceeds its declared size.");
                for(long position=(long)piece.Start;position<(long)end;position++)
                {
                    byte value=piece.Buffer[piece.Offset+(int)(position-(long)piece.Start)];
                    if(position<covered) { if(data[position]!=value) throw new InvalidDataException("Overlapping stored object fragments disagree."); }
                    else data[position]=value;
                }
                if((long)end>covered) covered=(long)end;
            }
            if(covered!=(long)size) throw new ContentUnavailableException("A fragmented stored object is incomplete.");
            return new StoredObject { Id=header.Id,Partition=header.Partition,References=header.References,Cells=header.Cells,Buffer=data,Offset=0,Count=data.Length,FromBlob=blob };
        }

        // A count larger than the remaining bytes is truncated structure.
        private static List<string> References(PersistedStorage.Cursor fields,PersistedStorage.DataElement element)
        {
            ulong count=fields.UInt();
            if(count>(ulong)(fields.Limit-fields.Position)) throw new InvalidDataException("A stored object's reference list is truncated.");
            var references=new List<string>((int)count);
            for(ulong i=0;i<count;i++) references.Add(fields.Id(element.Table));
            return references;
        }
        private static List<string[]> Cells(PersistedStorage.Cursor fields,PersistedStorage.DataElement element)
        {
            ulong count=fields.UInt();
            if(count>(ulong)(fields.Limit-fields.Position)) throw new InvalidDataException("A stored object's cell reference list is truncated.");
            var cells=new List<string[]>((int)count);
            for(ulong i=0;i<count;i++) cells.Add(new[] { fields.Id(element.Table),fields.Id(element.Table) });
            return cells;
        }
        private static bool SameObject(StoredObject a,StoredObject b)
        {
            if(a.Count!=b.Count || !SameLinks(a,b)) return false;
            for(int i=0;i<a.Count;i++) if(a.Buffer[a.Offset+i]!=b.Buffer[b.Offset+i]) return false;
            return true;
        }
        private static bool SameLinks(StoredObject a,StoredObject b)
        {
            return a.References.SequenceEqual(b.References,StringComparer.Ordinal) && SameRows(a.Cells,b.Cells);
        }
        private static bool SameRows(List<string[]> a,List<string[]> b)
        {
            if(a.Count!=b.Count) return false;
            for(int i=0;i<a.Count;i++) if(!a[i].SequenceEqual(b[i],StringComparer.Ordinal)) return false;
            return true;
        }
        private static PersistedStorage.Cursor Content(PersistedStorage.DataElement element)
        {
            return new PersistedStorage.Cursor(element.Buffer,element.Start,element.End);
        }
    }
}
