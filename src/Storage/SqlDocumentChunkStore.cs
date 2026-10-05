using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.IO;

namespace SharePointExplorer
{
    // Only retrieves physical SQL bytes. It does not interpret the shredded storage graph
    // or publish files; those responsibilities belong to other recovery layers.
    public sealed class SqlDocumentChunkStore : IDocumentChunkStore
    {
        private readonly SqlRepository repository;
        public SqlDocumentChunkStore(SqlRepository repository)
        {
            if(repository == null) throw new ArgumentNullException("repository");
            this.repository = repository;
        }
        public IList<StoredChunk> ReadChunks(Node file)
        {
            if(file == null) throw new ArgumentNullException("file");
            if(file.HasStream == false)
                throw new ContentUnavailableException("This is a metadata-only SharePoint template/ghosted file; its bytes are absent from the content database.");
            if(file.HistoryVersion<0 || (file.HistoryVersion>0 && file.HistoryVersion!=file.UiVersion))
                throw new ArgumentException("The document has an invalid selected version identity.","file");
            // Generations without the external ABS columns report them as NULL.
            // Preserve source delivery order.
            // The parser handles the main stream first; SQL BSN does not
            // select roots or resolve equal embedded sequence numbers.
            string sql = @"
SELECT m.BSN, m.StreamId, m.Partition, s.Type, s.Size, s.RbsId,
       " + repository.StreamCompressionColumns("s") + @", s.Content
FROM dbo.DocsToStreams AS m
LEFT JOIN dbo.DocStreams AS s ON s.SiteId=m.SiteId AND s.DocId=m.DocId
    AND s.Partition=m.Partition AND s.BSN=m.BSN
WHERE m.SiteId=@site AND m.DocId=@doc AND m.HistVersion=@version AND m.Level=@level AND m.Partition=0;";
            var chunks = new List<StoredChunk>();
            using (SqlConnection connection = repository.OpenConnection())
            using (SqlCommand command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 120;
                command.Parameters.Add("@site", SqlDbType.UniqueIdentifier).Value = file.SiteId;
                command.Parameters.Add("@doc", SqlDbType.UniqueIdentifier).Value = file.Id;
                command.Parameters.Add("@level", SqlDbType.TinyInt).Value = file.Level;
                command.Parameters.Add("@version", SqlDbType.Int).Value = file.HistoryVersion;
                using (SqlDataReader reader = command.ExecuteReader(CommandBehavior.SequentialAccess))
                {
                    while (reader.Read()) chunks.Add(ReadChunk(reader));
                }
            }
            if(chunks.Count == 0 && file.Size > 0)
                throw new ContentUnavailableException("No primary content streams are present for the selected document version in the restored database.");
            return chunks;
        }

        // Parse the actual SQL projection through IDataRecord so resident and
        // external-reference rules can be checked without a SQL service.
        internal static StoredChunk ReadChunk(IDataRecord reader)
        {
            if(reader==null) throw new ArgumentNullException("reader");
            long bsn = reader.GetInt64(0);
            long streamId = reader.GetInt64(1);
            byte partition = reader.GetByte(2);
            if (reader.IsDBNull(3)) throw new InvalidDataException("A referenced storage chunk is missing (BSN " + bsn + ").");
            byte type = reader.GetByte(3);
            int storedSize = reader.GetInt32(4);
            // An empty varbinary reference is absent, just as SQL NULL is.
            // Nonempty references still require a configured external provider.
            bool remoteBlob = !reader.IsDBNull(5) && reader.GetBytes(5, 0, null, 0, 0) > 0;
            int? compressedSize = reader.IsDBNull(6) ? (int?)null : reader.GetInt32(6);
            bool abs = !reader.IsDBNull(7);
            if (remoteBlob || (abs && reader.IsDBNull(8)))
                throw new ContentUnavailableException("This file requires an external BLOB store; its bytes are not all in this SQL database. No external content provider is configured.");
            if(storedSize<0 || (compressedSize.HasValue && compressedSize.Value<0))
                throw new InvalidDataException("A storage chunk has an invalid declared length (BSN "+bsn+").");
            // CompressedSize records a compressed external ABS replacement. The
            // source compression-update procedure changes the ABS reference and
            // this field without changing resident Content or its original Size.
            // An available exact-size resident payload remains original bytes;
            // decompressing it based on this metadata would corrupt recovery.
            if(compressedSize.HasValue && !abs)
                throw new NotSupportedException("This storage chunk declares compression without its external ABS context; an inline compression format cannot be identified safely.");
            if (reader.IsDBNull(8)) throw new InvalidDataException("Storage chunk content is missing (BSN " + bsn + ").");
            long length = reader.GetBytes(8, 0, null, 0, 0);
            if (length != storedSize || length > Int32.MaxValue)
                throw new InvalidDataException("Storage chunk length is inconsistent (BSN " + bsn + ").");
            byte[] bytes = new byte[(int)length];
            long read = 0;
            while (read < length)
            {
                int count = (int)Math.Min(65536, length - read);
                long actual = reader.GetBytes(8, read, bytes, (int)read, count);
                if (actual <= 0) throw new EndOfStreamException("Storage chunk was truncated.");
                read += actual;
            }
            return new StoredChunk { BSN=bsn, StreamId=streamId, Partition=partition, Type=type, Content=bytes };
        }

    }
}
