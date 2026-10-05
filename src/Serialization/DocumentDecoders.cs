using System;
using System.Collections.Generic;
using System.IO;

namespace SharePointExplorer
{
    // Schema dispatch is independent of SQL retrieval and file publication.
    // Additional codecs can be registered without changing either caller.
    public sealed class DocumentDecoderRegistry
    {
        private readonly List<IDocumentDecoder> decoders;
        public DocumentDecoderRegistry(IEnumerable<IDocumentDecoder> decoders)
        {
            if(decoders==null) throw new ArgumentNullException("decoders");
            this.decoders=new List<IDocumentDecoder>();
            foreach(IDocumentDecoder decoder in decoders)
            {
                if(decoder==null || String.IsNullOrWhiteSpace(decoder.Name)) throw new ArgumentException("Each document decoder requires a name.","decoders");
                this.decoders.Add(decoder);
            }
        }
        public static DocumentDecoderRegistry CreateDefault()
        {
            return new DocumentDecoderRegistry(new IDocumentDecoder[] { new PlainDocumentDecoder(),new GenericDocumentDecoder(),new OneNoteDocumentDecoder(false),new OneNoteDocumentDecoder(true) });
        }
        public bool Supports(byte streamSchema)
        {
            return Find(streamSchema)!=null;
        }
        private IDocumentDecoder Find(byte streamSchema)
        {
            IDocumentDecoder result=null;
            foreach(IDocumentDecoder decoder in decoders)
            {
                if(!decoder.Supports(streamSchema)) continue;
                if(result!=null) throw new InvalidOperationException("Multiple decoders are registered for document storage schema "+streamSchema+".");
                result=decoder;
            }
            return result;
        }
        public IDocumentDecoder Resolve(byte streamSchema)
        {
            IDocumentDecoder result=Find(streamSchema);
            if(result!=null) return result;
            throw new NotSupportedException("Document storage schema "+streamSchema+" is not supported by this version. No file was exported.");
        }
    }
    // AllDocs.StreamSchema is a bit field: file layout (none 0, native 1,
    // shredded 2), cell layout (generic 0, non-generic 16, OneNote table of
    // contents 32) and host-blob tagging (pre-release 0, RTM 64). Readers are
    // chosen from these flags; reserved bits and undefined combinations are rejected.
    internal static class StreamSchemaFlags
    {
        internal const int FileMask=3, Native=1, Shredded=2, ReservedMask=12, CellMask=48, NonGenericCell=16, TocCell=32, HostMask=192;
        internal static bool IsDefined(byte schema)
        {
            return (schema&ReservedMask)==0 && (schema&FileMask)!=FileMask && (schema&CellMask)!=CellMask && (schema&HostMask)<=64;
        }
        internal static bool IsPlain(byte schema) { return IsDefined(schema) && (schema&FileMask)!=Shredded && (schema&CellMask)==0; }
        internal static bool IsShredded(byte schema,int cell) { return IsDefined(schema) && (schema&FileMask)==Shredded && (schema&CellMask)==cell; }
    }
    internal sealed class PlainDocumentDecoder : IDocumentDecoder
    {
        public string Name { get { return "Resident plain stream"; } }
        // Native resident streams (1, or 65 with RTM host tagging) and the legacy
        // stream-less convention (0, 64); all need exact inline bytes.
        public bool Supports(byte streamSchema) { return StreamSchemaFlags.IsPlain(streamSchema); }
        public void Write(IList<StoredChunk> chunks,long expectedSize,Stream output)
        {
            if(chunks==null) throw new ArgumentNullException("chunks");
            if(expectedSize<0) throw new InvalidDataException("The document has an invalid declared size.");
            if(expectedSize==0 && chunks.Count==0) return;
            if(chunks.Count==0) throw new ContentUnavailableException("The document's bytes are absent from this content source.");
            if(chunks.Count!=1 || chunks[0]==null || chunks[0].Partition!=0 || chunks[0].Content==null)
                throw new NotSupportedException("This file has no supported resident plain stream. Its content may require an unavailable template or another storage decoder.");
            if(chunks[0].Content.LongLength!=expectedSize)
                throw new InvalidDataException("The resident plain stream length disagrees with the document's declared size.");
            output.Write(chunks[0].Content,0,chunks[0].Content.Length);
        }
    }
    // The output is whatever the stored tree holds; its length is
    // not compared with the document's declared size.
    internal sealed class GenericDocumentDecoder : IDocumentDecoder,IReconstructedDocumentDecoder
    {
        public string Name { get { return "SharePoint generic document tree"; } }
        public bool Supports(byte streamSchema) { return StreamSchemaFlags.IsShredded(streamSchema,0); }
        public ReconstructedContent Reconstruct(IList<StoredChunk> chunks,long sourceSize)
        {
            if(chunks==null) throw new ArgumentNullException("chunks");
            bool main=false;
            foreach(StoredChunk chunk in chunks)
            {
                if(chunk==null) throw new InvalidDataException("A storage chunk descriptor is missing.");
                if(chunk.Partition==0 && chunk.StreamId==1 && chunk.Content!=null) main=true;
            }
            if(sourceSize>0 && !main) throw new ContentUnavailableException("The document's primary storage stream is absent from this content source.");
            return new ReconstructedContent(StoredDocumentReader.Reconstruct(chunks));
        }
        public void Write(IList<StoredChunk> chunks,long expectedSize,Stream output)
        {
            if(output==null || !output.CanWrite) throw new ArgumentException("A writable export stream is required.","output");
            Reconstruct(chunks,expectedSize).WriteTo(output);
        }
    }
}