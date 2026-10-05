using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SharePointExplorer
{
    // Reconstructed formats can have a serialized length different from SQL's
    // logical file size. Freeze the reconstructed output before any publication.
    internal interface IReconstructedDocumentDecoder
    {
        ReconstructedContent Reconstruct(IList<StoredChunk> chunks,long sourceSize);
    }
    // Reconstructed output: ordered segments of stored or synthesized buffers,
    // with the length they add up to.
    internal sealed class ReconstructedContent
    {
        private readonly ArraySegment<byte>[] parts;
        internal ReconstructedContent(IEnumerable<ArraySegment<byte>> parts)
        {
            this.parts=parts.ToArray();
            foreach(ArraySegment<byte> part in this.parts) Length=checked(Length+part.Count);
        }
        internal ReconstructedContent(byte[] bytes):this(new[]{new ArraySegment<byte>(bytes)}){}
        internal long Length { get; private set; }
        internal void WriteTo(Stream output)
        {
            foreach(ArraySegment<byte> part in parts) output.Write(part.Array,part.Offset,part.Count);
        }
    }
    internal sealed class OneNoteDocumentDecoder : IDocumentDecoder,IReconstructedDocumentDecoder
    {
        private readonly bool toc;
        internal OneNoteDocumentDecoder(bool toc){this.toc=toc;}
        public string Name {get{return toc?"OneNote server notebook index":"OneNote server section";}}
        public bool Supports(byte schema){return StreamSchemaFlags.IsShredded(schema,toc?StreamSchemaFlags.TocCell:StreamSchemaFlags.NonGenericCell);}
        public ReconstructedContent Reconstruct(IList<StoredChunk> chunks,long sourceSize)
        {
            if(sourceSize<0)throw new InvalidDataException("The OneNote document has an invalid declared size.");
            if(chunks==null)throw new ArgumentNullException("chunks");
            bool main=false;
            foreach(StoredChunk chunk in chunks)if(chunk!=null && chunk.Partition==0 && chunk.StreamId==1 && chunk.Content!=null)main=true;
            if(!main)throw new ContentUnavailableException("The OneNote primary storage stream is absent from this content source.");
            return new ReconstructedContent(StoredDocumentReader.ReconstructOneNote(chunks,toc));
        }
        public void Write(IList<StoredChunk> chunks,long expectedSize,Stream output)
        {
            if(output==null || !output.CanWrite)throw new ArgumentException("A writable export stream is required.","output");
            Reconstruct(chunks,expectedSize).WriteTo(output);
        }
    }
}
