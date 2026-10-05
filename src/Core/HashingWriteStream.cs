using System;
using System.IO;
using System.Security.Cryptography;

namespace SharePointExplorer
{
    // Recovery writes have one byte limit and one checksum, including when the
    // destination is Stream.Null for verification without local file creation.
    internal sealed class HashingWriteStream : Stream
    {
        private readonly Stream destination;
        private readonly long expectedSize;
        private readonly SHA256 hash;
        private readonly bool leaveOpen;
        private bool disposed, completed;
        private string digest;
        internal long BytesWritten { get; private set; }

        internal HashingWriteStream(Stream destination, long expectedSize, bool leaveOpen)
        {
            if(destination==null) throw new ArgumentNullException("destination");
            if(!destination.CanWrite) throw new ArgumentException("A writable recovery stream is required.","destination");
            if(expectedSize<0) throw new InvalidDataException("The document has an invalid declared size.");
            this.destination=destination;
            this.expectedSize=expectedSize;
            this.leaveOpen=leaveOpen;
            hash=SHA256.Create();
        }
        public override bool CanRead { get { return false; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return !disposed && !completed && destination.CanWrite; } }
        public override long Length { get { EnsureOpen(); return BytesWritten; } }
        public override long Position
        {
            get { EnsureOpen(); return BytesWritten; }
            set { throw new NotSupportedException("The recovery stream is forward-only."); }
        }
        public override void Write(byte[] buffer,int offset,int count)
        {
            EnsureOpen();
            if(completed) throw new InvalidOperationException("The recovery checksum is already finalized.");
            if(buffer==null) throw new ArgumentNullException("buffer");
            if(offset<0 || count<0 || offset>buffer.Length-count) throw new ArgumentOutOfRangeException("offset");
            if(count>expectedSize-BytesWritten) throw new InvalidDataException("The decoder exceeds the document's declared size.");
            if(count==0) return;
            destination.Write(buffer,offset,count);
            hash.TransformBlock(buffer,offset,count,null,0);
            BytesWritten+=count;
        }
        internal string Complete()
        {
            EnsureOpen();
            if(completed) return digest;
            if(BytesWritten!=expectedSize) throw new InvalidDataException("Reconstructed length does not match the stored file size.");
            hash.TransformFinalBlock(new byte[0],0,0);
            digest=BitConverter.ToString(hash.Hash).Replace("-","").ToLowerInvariant();
            completed=true;
            return digest;
        }
        public override void Flush() { EnsureOpen(); destination.Flush(); }
        public override int Read(byte[] buffer,int offset,int count) { throw new NotSupportedException("The recovery stream does not support reading."); }
        public override long Seek(long offset,SeekOrigin origin) { throw new NotSupportedException("The recovery stream is forward-only."); }
        public override void SetLength(long value) { throw new NotSupportedException("The decoder cannot change the document length."); }
        protected override void Dispose(bool disposing)
        {
            if(!disposed)
            {
                disposed=true;
                if(disposing)
                {
                    hash.Dispose();
                    if(!leaveOpen) destination.Dispose();
                }
            }
            base.Dispose(disposing);
        }
        private void EnsureOpen() { if(disposed) throw new ObjectDisposedException("HashingWriteStream"); }
    }
}