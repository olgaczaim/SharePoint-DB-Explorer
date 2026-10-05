using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Data.SqlClient;

namespace SharePointExplorer
{
    // Optional content-database columns differ between SharePoint generations.
    // SizeRead/SizeWrite and the external ABS stream columns are absent from
    // older schemas, so queries include only the columns that actually exist.
    internal sealed class ContentColumns
    {
        internal bool DocsSizeRead, DocsSizeWrite, VersionsSizeRead, VersionsSizeWrite, StreamCompressedSize, StreamAbsId;
    }

    public sealed partial class SqlRepository
    {
        private readonly object columnGate = new object();
        private ContentColumns columns;

        // The content database build recorded in dbo.Versions, and its generation.
        public string SourceBuild { get; private set; }
        public SharePointGeneration Generation { get; private set; }

        internal ContentColumns Columns
        {
            get { lock (columnGate) { if (columns == null) columns = ReadColumns(); return columns; } }
        }
        // Lets checks emulate an older generation's optional-column profile.
        internal void UseColumns(ContentColumns profile)
        {
            if (profile == null) throw new ArgumentNullException("profile");
            lock (columnGate) columns = profile;
        }

        private ContentColumns ReadColumns()
        {
            var found = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (SqlConnection connection = OpenConnection())
            using (SqlCommand command = CreateCommand(connection, @"
SELECT t.name + N'.' + c.name
FROM sys.columns AS c
INNER JOIN sys.objects AS t ON t.object_id = c.object_id
INNER JOIN sys.schemas AS s ON s.schema_id = t.schema_id
WHERE s.name = N'dbo' AND t.type IN (N'U', N'V') AND t.name IN (N'AllDocs', N'AllDocVersions', N'DocStreams')
  AND c.name IN (N'SizeRead', N'SizeWrite', N'CompressedSize', N'ABSId');"))
            using (SqlDataReader reader = command.ExecuteReader())
                while (reader.Read()) found.Add(reader.GetString(0));
            return new ContentColumns
            {
                DocsSizeRead = found.Contains("AllDocs.SizeRead"), DocsSizeWrite = found.Contains("AllDocs.SizeWrite"),
                VersionsSizeRead = found.Contains("AllDocVersions.SizeRead"), VersionsSizeWrite = found.Contains("AllDocVersions.SizeWrite"),
                StreamCompressedSize = found.Contains("DocStreams.CompressedSize"), StreamAbsId = found.Contains("DocStreams.ABSId")
            };
        }

        // Logical document size: SizeRead when present, then Size, then SizeWrite.
        // Size is a 32-bit column; the 64-bit columns hold large-file sizes.
        internal string DocumentSize(string alias)
        {
            ContentColumns available = Columns;
            return SizeExpression(alias, available.DocsSizeRead, available.DocsSizeWrite, true);
        }
        internal string VersionSize(string alias)
        {
            ContentColumns available = Columns;
            return SizeExpression(alias, available.VersionsSizeRead, available.VersionsSizeWrite, false);
        }
        internal static string SizeExpression(string alias, bool read, bool write, bool zero)
        {
            string prefix = String.IsNullOrEmpty(alias) ? String.Empty : alias + ".";
            var parts = new List<string>();
            if (read) parts.Add("CONVERT(bigint," + prefix + "SizeRead)");
            parts.Add("CONVERT(bigint," + prefix + "Size)");
            if (write) parts.Add("CONVERT(bigint," + prefix + "SizeWrite)");
            if (zero) parts.Add("0");
            return parts.Count == 1 ? parts[0] : "COALESCE(" + String.Join(",", parts) + ")";
        }
        // Stream columns in the order SqlDocumentChunkStore.ReadChunk expects.
        internal string StreamCompressionColumns(string alias)
        {
            ContentColumns available = Columns;
            return (available.StreamCompressedSize ? alias + ".CompressedSize" : "CAST(NULL AS int) AS CompressedSize") + ", " +
                (available.StreamAbsId ? alias + ".ABSId" : "CAST(NULL AS uniqueidentifier) AS ABSId");
        }

        // dbo.Versions records the content database build under the empty VersionId.
        private void ReadBuild(SqlConnection connection)
        {
            Version best = null;
            string text = null;
            using (SqlCommand command = CreateCommand(connection, "SELECT Version FROM dbo.Versions WHERE VersionId = '00000000-0000-0000-0000-000000000000';"))
            using (SqlDataReader reader = command.ExecuteReader())
                while (reader.Read())
                {
                    if (reader.IsDBNull(0)) continue;
                    string value = Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture);
                    Version parsed;
                    if (Version.TryParse(value, out parsed) && (best == null || parsed > best)) { best = parsed; text = value; }
                }
            if (best == null)
                throw new InvalidOperationException("This database does not record a SharePoint content database build in dbo.Versions. Use a restored SharePoint content database.");
            SourceBuild = text;
            Generation = SharePointBuilds.Classify(best);
            if (!SharePointBuilds.IsSupported(Generation))
                throw new InvalidOperationException("This content database was created by " + SharePointBuilds.DisplayName(Generation) + " (build " + text +
                    "). This explorer supports SharePoint Server 2016, SharePoint Server 2019 and Subscription Edition content databases.");
        }
    }
}
