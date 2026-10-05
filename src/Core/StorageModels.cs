using System;

namespace SharePointExplorer
{
    public sealed class StoredChunk
    {
        public long BSN;
        public long StreamId;
        public byte Partition;
        public byte Type;
        public byte[] Content;
    }

    public sealed class ExportResult
    {
        public string Path;
        public long Bytes;
        public string Sha256;
        public string Decoder;
    }
}
