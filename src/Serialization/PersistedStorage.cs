using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace SharePointExplorer
{
    // Structured reader for SharePoint's persisted shredded-storage rows. Each SQL
    // stream row holds one host blob: a tree of elements, each with an identifier,
    // an optional length-prefixed item, child elements and a footer. Contained
    // blobs carry the storage index and the MS-FSSHTTPB data elements of a cell.
    // Identifiers can be compressed against the GUID table of the host blob.
    // Nothing is located by scanning for bytes. Only broken
    // structure is damage: an element, item or footer where another is expected,
    // a bad signature or truncated fields. Identities, trailing bytes and value
    // ranges are not checked. Damage ends a read: what was read before it is
    // kept and the damage is recorded.
    internal static class PersistedStorage
    {
        internal const int StorageIndexType=1, StorageManifestType=2, CellManifestType=3, RevisionManifestType=4,
            ObjectGroupType=5, ObjectGroupFragmentType=11, ObjectDataBlobFragmentType=12;
        internal const int SimpleBlob=3, MoveableBlob=5, NonAggregatableBlob=6;
        internal const ulong IndexBlobType=1;
        private const uint DataElementSignature=0x2E4E32BA;

        internal sealed class HostBlob
        {
            // The first GUID table the host declares, and the one in effect at its end.
            internal Dictionary<ulong,Guid> FirstTable, Table;
            internal readonly List<ContainedBlob> Blobs=new List<ContainedBlob>();
            internal InvalidDataException Damage;
        }
        internal sealed class ContainedBlob
        {
            internal int Kind;
            internal string Id;
            internal ulong Sequence, Type;
            internal byte[] Buffer;
            internal int DataStart, DataEnd;
            internal Dictionary<ulong,Guid> Table;
        }
        internal sealed class DataElement
        {
            internal string Id, BlobId;
            internal int Type;
            internal ulong Sequence;
            internal byte[] Buffer;
            internal int Start, End;
            internal Dictionary<ulong,Guid> Table;
        }

        // A host blob without its own GUID table uses the inherited table. A later
        // table replaces an earlier one. Blobs before structural damage are kept;
        // nothing after it is read.
        internal static HostBlob ReadHostBlob(byte[] buffer,Dictionary<ulong,Guid> inherited)
        {
            var host=new HostBlob();
            if(buffer==null) { host.Damage=new InvalidDataException("A persisted stream row has no content."); return host; }
            try { ReadHost(new Cursor(buffer,0,buffer.Length),host,inherited); }
            catch(InvalidDataException error) { host.Damage=error; }
            return host;
        }
        private static void ReadHost(Cursor cursor,HostBlob host,Dictionary<ulong,Guid> inherited)
        {
            cursor.Element(1);
            if(cursor.PeekItem()) cursor.SkipItem();
            Dictionary<ulong,Guid> active=inherited;
            int child;
            while((child=cursor.PeekElement())>=0)
            {
                if(child==4)
                {
                    cursor.Element(4);
                    if(cursor.PeekItem()) cursor.SkipItem();
                    active=ReadGuidTable(cursor,1);
                    cursor.Footer();
                    if(host.FirstTable==null) host.FirstTable=active;
                    host.Table=active;
                }
                else if(child==2) ReadContainedBlobs(cursor,host,active);
                else cursor.SkipElement();
            }
            cursor.Footer();
        }

        // A repeated index takes the later GUID.
        private static Dictionary<ulong,Guid> ReadGuidTable(Cursor cursor,int id)
        {
            cursor.Element(id);
            var table=new Dictionary<ulong,Guid>();
            int child;
            while((child=cursor.PeekElement())>=0)
            {
                if(child!=1) throw new InvalidDataException("A persisted GUID table contains an unexpected entry.");
                cursor.Element(1);
                int end,start=cursor.Item(out end);
                var fields=new Cursor(cursor.Buffer,start,end);
                ulong index=fields.UInt();
                table[index]=fields.Guid();
                cursor.Position=end;
                cursor.Footer();
            }
            cursor.Footer();
            return table;
        }

        // Fields missing at the end of a contained blob's item keep their defaults.
        private static void ReadContainedBlobs(Cursor cursor,HostBlob host,Dictionary<ulong,Guid> table)
        {
            cursor.Element(2);
            int child;
            while((child=cursor.PeekElement())>=0)
            {
                if(child!=SimpleBlob && child!=MoveableBlob && child!=NonAggregatableBlob) throw new InvalidDataException("A host blob contains an unknown blob kind.");
                cursor.Element(child);
                int end,start=cursor.Item(out end);
                var fields=new Cursor(cursor.Buffer,start,end);
                var blob=new ContainedBlob { Kind=child,Buffer=cursor.Buffer,Table=table,DataStart=end,DataEnd=end };
                if(fields.Position<fields.Limit) blob.Id=fields.Id(table);
                if(fields.Position<fields.Limit) blob.Sequence=fields.UInt();
                if(fields.Position<fields.Limit) blob.Type=fields.UInt();
                if(fields.Position<fields.Limit) fields.UInt();
                if(fields.Position<fields.Limit) fields.UInt();
                if(fields.Position<fields.Limit)
                {
                    int length;
                    blob.DataStart=fields.Atom(out length);
                    blob.DataEnd=blob.DataStart+length;
                }
                cursor.Position=end;
                cursor.Footer();
                host.Blobs.Add(blob);
            }
            cursor.Footer();
        }

        // A moveable or non-aggregatable blob holds data element identity items
        // (1) and contents (2); each contents is read with the identity before
        // it. The element's own GUID table is read but not used: identifiers
        // resolve against the host blob's table. Elements found before damage are
        // returned with the damage recorded; contents cut by damage end at the
        // blob's end and stay usable up to it.
        internal static List<DataElement> ReadDataElements(ContainedBlob blob,out InvalidDataException damage)
        {
            var elements=new List<DataElement>();
            damage=null;
            try
            {
                var cursor=new Cursor(blob.Buffer,blob.DataStart,blob.DataEnd);
                if(cursor.UInt32()!=DataElementSignature) throw new InvalidDataException("A data element blob has an invalid signature.");
                cursor.Element(1);
                if(cursor.PeekItem()) cursor.SkipItem();
                if(cursor.PeekElement()==2) throw new NotSupportedException("This store uses a deprecated data element blob layout.");
                cursor.Element(3);
                string id=null;
                int type=0;
                bool identified=false;
                int child;
                while((child=cursor.PeekElement())>=0)
                {
                    if(child==1)
                    {
                        cursor.Element(1);
                        int end,start=cursor.Item(out end);
                        var info=new Cursor(blob.Buffer,start,end);
                        id=info.Id(blob.Table);
                        ulong value=info.UInt();
                        info.Serial(blob.Table);
                        type=value>Int32.MaxValue ? -1 : (int)value;
                        identified=true;
                        cursor.Position=end;
                        cursor.Footer();
                    }
                    else if(child==2)
                    {
                        if(!identified) throw new InvalidDataException("A data element's contents precede its identity.");
                        cursor.Element(2);
                        var element=new DataElement { Id=id,BlobId=blob.Id,Type=type,Sequence=blob.Sequence,Buffer=blob.Buffer,Start=cursor.Position,End=blob.DataEnd,Table=blob.Table };
                        elements.Add(element);
                        cursor.SkipContents();
                        element.End=cursor.Position;
                        cursor.Footer();
                    }
                    else if(child==3) ReadGuidTable(cursor,3);
                    else if(child==4 || child==5) cursor.SkipElement();
                    else throw new InvalidDataException("A data element blob contains an unknown child.");
                }
                cursor.Footer();
            }
            catch(InvalidDataException error) { damage=error; }
            return elements;
        }

        internal sealed class Cursor
        {
            internal readonly byte[] Buffer;
            internal int Position;
            internal readonly int Limit;
            internal Cursor(byte[] buffer,int start,int limit)
            {
                if(buffer==null) throw new ArgumentNullException("buffer");
                if(start<0 || limit>buffer.Length || start>limit) throw new InvalidDataException("A persisted range is outside its buffer.");
                Buffer=buffer; Position=start; Limit=limit;
            }
            internal Cursor(byte[] buffer,int start,int limit,int position) : this(buffer,start,limit) { Position=position; }
            private void Need(int count) { if(count<0 || Position>Limit-count) throw new InvalidDataException("A persisted structure is truncated."); }

            // Header byte forms: element (xx01, or 16-bit xx11), item (x100 16-bit,
            // x010 32-bit with an optional 64-bit length) and footer (x000).
            private int Header(out int kind,out long length,bool consume)
            {
                Need(1);
                int start=Position;
                byte first=Buffer[Position];
                int id;
                length=0;
                if((first&3)==1) { kind=0; id=first>>2; Position++; }
                else if((first&3)==3) { Need(2); kind=0; id=BitConverter.ToUInt16(Buffer,Position)>>2; Position+=2; }
                else if((first&7)==0) { kind=2; id=-1; Position++; }
                else if((first&7)==4)
                {
                    Need(2); int value=BitConverter.ToUInt16(Buffer,Position); Position+=2;
                    kind=1; id=(value>>3)&3; length=value>>5;
                }
                else if((first&7)==2)
                {
                    Need(4); uint value=BitConverter.ToUInt32(Buffer,Position); Position+=4;
                    kind=1; id=(int)((value>>3)&0xFF); length=value>>11;
                    if(length==0x1FFFFF)
                    {
                        Need(8); ulong full=BitConverter.ToUInt64(Buffer,Position); Position+=8;
                        if(full>Int32.MaxValue) throw new InvalidDataException("A persisted item is too large.");
                        length=(long)full;
                    }
                }
                else throw new InvalidDataException("A persisted structure has an unknown header.");
                if(!consume) Position=start;
                return id;
            }
            // Returns the next child element identity, or -1 at a footer or range end.
            internal int PeekElement()
            {
                if(Position>=Limit) return -1;
                int kind; long length;
                int id=Header(out kind,out length,false);
                if(kind==1) throw new InvalidDataException("A persisted item appears where an element was expected.");
                return kind==0 ? id : -1;
            }
            internal bool PeekItem()
            {
                if(Position>=Limit) return false;
                int kind; long length;
                Header(out kind,out length,false);
                return kind==1;
            }
            internal bool PeekFooter()
            {
                if(Position>=Limit) return false;
                int kind; long length;
                Header(out kind,out length,false);
                return kind==2;
            }
            internal void Element(int id)
            {
                int kind; long length;
                int actual=Header(out kind,out length,true);
                if(kind!=0 || actual!=id) throw new InvalidDataException("A persisted structure has an unexpected element.");
            }
            internal void Footer()
            {
                int kind; long length;
                Header(out kind,out length,true);
                if(kind!=2) throw new InvalidDataException("A persisted element lacks its footer.");
            }
            internal int Item(out int end)
            {
                int kind; long length;
                int id=Header(out kind,out length,true);
                if(kind!=1 || id!=0) throw new InvalidDataException("A persisted element has an unexpected item.");
                if(length>Limit-Position) throw new InvalidDataException("A persisted item is truncated.");
                end=Position+(int)length;
                return Position;
            }
            internal void SkipItem() { int end; Item(out end); Position=end; }
            // Consumes a complete child element, including its footer.
            internal void SkipElement()
            {
                int kind; long length;
                Header(out kind,out length,true);
                if(kind!=0) throw new InvalidDataException("A persisted structure has an unexpected header.");
                SkipContents();
                Footer();
            }
            // Advances to the footer closing the current element, or to the range end.
            internal void SkipContents()
            {
                int depth=0,visits=0;
                while(true)
                {
                    if(depth==0 && Position>=Limit) return;
                    if(++visits>10000000) throw new InvalidDataException("A persisted element exceeds its traversal budget.");
                    int start=Position;
                    int kind; long length;
                    Header(out kind,out length,true);
                    if(kind==0) { if(++depth>256) throw new InvalidDataException("A persisted element is nested too deeply."); continue; }
                    if(kind==1) { if(length>Limit-Position) throw new InvalidDataException("A persisted item is truncated."); Position+=(int)length; continue; }
                    if(depth==0) { Position=start; return; }
                    depth--;
                }
            }
            // MS-FSSHTTPB compact unsigned 64-bit integer.
            internal ulong UInt()
            {
                Need(1);
                byte first=Buffer[Position];
                if(first==0) { Position++; return 0; }
                int zeros=0;
                while((first&(1<<zeros))==0) zeros++;
                if(zeros==7) { Need(9); ulong full=BitConverter.ToUInt64(Buffer,Position+1); Position+=9; return full; }
                return Little(zeros+1)>>(zeros+1);
            }
            internal uint UInt32() { Need(4); uint value=BitConverter.ToUInt32(Buffer,Position); Position+=4; return value; }
            internal Guid Guid() { Need(16); var bytes=new byte[16]; System.Buffer.BlockCopy(Buffer,Position,bytes,0,16); Position+=16; return new Guid(bytes); }
            internal int Atom(out int length)
            {
                ulong value=UInt();
                if(value>(ulong)(Limit-Position)) throw new InvalidDataException("A persisted value is truncated.");
                length=(int)value;
                int start=Position;
                Position+=length;
                return start;
            }
            private ulong Little(int width)
            {
                Need(width);
                ulong value=0;
                for(int i=0;i<width;i++) value|=(ulong)Buffer[Position+i]<<(8*i);
                Position+=width;
                return value;
            }
            // Extended GUIDs use the public MS-FSSHTTPB forms, plus persisted forms
            // whose GUID is an index into the declaring GUID table.
            internal string Id(Dictionary<ulong,Guid> table)
            {
                Need(1);
                byte first=Buffer[Position];
                if(first==0) { Position++; return null; }
                ulong value,raw;
                if((first&1)==1) { raw=Little(2); return Indexed(table,raw>>11,(raw>>1)&0x1FF); }
                if((first&3)==2) { raw=Little(4); return Indexed(table,raw>>28,(raw>>2)&0x3FFFFFF); }
                if((first&15)==8) { raw=Little(6); return Indexed(table,raw>>37,(raw>>4)&0xFFFFFFFF); }
                if((first&7)==4) { Position++; value=(ulong)(first>>3); }
                else if((first&63)==32) value=Little(2)>>6;
                else if((first&127)==64) value=Little(3)>>7;
                else if(first==128) { Position++; value=Little(4); }
                else throw new InvalidDataException("A persisted extended GUID has an unknown form.");
                return Format(Guid(),value);
            }
            private static string Indexed(Dictionary<ulong,Guid> table,ulong index,ulong value)
            {
                Guid guid;
                if(table==null || !table.TryGetValue(index,out guid)) throw new InvalidDataException("A compressed extended GUID references an undeclared GUID table entry.");
                return Format(guid,value);
            }
            internal static string Format(Guid guid,ulong value)
            {
                return "G:"+guid.ToString("N")+":"+value.ToString(CultureInfo.InvariantCulture);
            }
            // Serial numbers: null, two persisted table-indexed forms and the public form.
            internal string Serial(Dictionary<ulong,Guid> table)
            {
                ulong sequence;
                return Serial(table,out sequence);
            }
            internal string Serial(Dictionary<ulong,Guid> table,out ulong sequence)
            {
                Need(1);
                byte first=Buffer[Position];
                sequence=0;
                if(first==0) { Position++; return null; }
                ulong raw;
                Guid guid;
                if((first&1)==1) { raw=Little(4); sequence=(raw>>6)&0x3FFFFFF; guid=Lookup(table,(raw>>2)&0xF); }
                else if((first&3)==2) { raw=Little(6); sequence=(raw>>13)&0x7FFFFFFFF; guid=Lookup(table,(raw>>3)&0x3FF); }
                else if(first==128) { Position++; guid=Guid(); sequence=Little(8); }
                else throw new InvalidDataException("A persisted serial number has an unknown form.");
                return "S:"+guid.ToString("N")+":"+sequence.ToString(CultureInfo.InvariantCulture);
            }
            private static Guid Lookup(Dictionary<ulong,Guid> table,ulong index)
            {
                Guid guid;
                if(table==null || !table.TryGetValue(index,out guid)) throw new InvalidDataException("A compressed serial number references an undeclared GUID table entry.");
                return guid;
            }
        }
    }
}
