using System;
using System.Collections.Generic;
using System.IO;

namespace SharePointExplorer
{
    // Catalog queries and document content retrieval have separate contracts.
    // A session can therefore host a different restored-content source without
    // changing the desktop, storage decoder or file publication service.
    public interface ISharePointCatalog
    {
        string SourceName { get; }
        void ValidateSchema();
        List<string> CheckDatabase();
        List<Node> GetRootSites();
        List<Node> GetChildren(Node parent);
        Node GetFile(Guid siteId, Guid fileId);
        IEnumerable<Node> EnumerateCurrentFiles(Guid? siteId);
    }


    // Optional catalog capability for a complete library, independent of which
    // folders the desktop has expanded. Implementations must validate its scope.
    public interface ISharePointLibraryCatalog
    {
        IEnumerable<Node> EnumerateCurrentLibraryFiles(Node library);
        Node GetCurrentLibraryFile(Node library, Guid fileId);
    }
    // Version recovery is explicit. A catalog must reload the complete selected
    // identity rather than substituting the document's current state.
    public interface ISharePointVersionCatalog
    {
        List<Node> GetFileVersions(Guid siteId, Guid fileId);
        Node GetFileVersion(Node selectedVersion);
    }
    public interface IDocumentChunkStore
    {
        IList<StoredChunk> ReadChunks(Node document);
    }

    public interface IDocumentDecoder
    {
        string Name { get; }
        bool Supports(byte streamSchema);
        void Write(IList<StoredChunk> chunks, long expectedSize, Stream output);
    }

    public sealed class RecoveryResult
    {
        public Node Document { get; set; }
        public long Bytes { get; set; }
        public string Sha256 { get; set; }
        public string Decoder { get; set; }
    }

    public sealed class ContentUnavailableException : NotSupportedException
    {
        public ContentUnavailableException(string message) : base(message) { }
    }

    public enum RecoveryStatus
    {
        Success,
        Unsupported,
        Unavailable,
        Corrupt,
        SqlError,
        Failed
    }

    public static class RecoveryErrors
    {
        public static RecoveryStatus Classify(Exception error)
        {
            if(error is ContentUnavailableException) return RecoveryStatus.Unavailable;
            if(error is NotSupportedException) return RecoveryStatus.Unsupported;
            if(error is InvalidDataException || error is EndOfStreamException || error is OverflowException) return RecoveryStatus.Corrupt;
            if(error is Microsoft.Data.SqlClient.SqlException) return RecoveryStatus.SqlError;
            return RecoveryStatus.Failed;
        }
    }
}
