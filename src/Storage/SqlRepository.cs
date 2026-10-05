using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;

namespace SharePointExplorer
{
    public sealed partial class SqlRepository : ISharePointCatalog, ISharePointLibraryCatalog, ISharePointVersionCatalog, ISharePointAttachmentCatalog, ISharePointMigrationCatalog, ISharePointMigrationHistoryCatalog
    {
        public List<Node> GetItemAttachments(Node item) { return new SqlAttachmentCatalog(this).GetItemAttachments(item); }
        public IEnumerable<Node> EnumerateCurrentListAttachments(Node list) { return new SqlAttachmentCatalog(this).EnumerateCurrentListAttachments(list); }
        public Node GetCurrentItemAttachment(Node item,Guid attachmentId) { return new SqlAttachmentCatalog(this).GetCurrentItemAttachment(item,attachmentId); }
        public MigrationListSnapshot ReadMigrationList(Node list) { return new SqlMigrationCatalog(this).ReadMigrationList(list); }
        public MigrationHistorySnapshot ReadMigrationHistory(MigrationListSnapshot snapshot) { return new SqlMigrationCatalog(this).ReadMigrationHistory(snapshot); }
        private const int QueryTimeout = 120;
        public string ConnectionString { get; private set; }
        public string SourceName { get; private set; }

        public SqlRepository(string server, string database)
            : this(new SqlConnectionOptions { Server = server, Database = database })
        {
        }

        public SqlRepository(SqlConnectionOptions options)
        {
            if (options == null) throw new ArgumentNullException("options");
            SqlConnectionOptions snapshot = options.Clone();
            SqlConnectionStringBuilder builder = snapshot.CreateConnectionStringBuilder();
            ConnectionString = builder.ConnectionString;
            SourceName = snapshot.SourceName;
        }

        public SqlConnection OpenConnection()
        {
            SqlConnection connection = new SqlConnection(ConnectionString);
            try
            {
                connection.Open();
                return connection;
            }
            catch
            {
                connection.Dispose();
                throw;
            }
        }

        public void ValidateSchema()
        {
            Dictionary<string, string[]> required = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            required.Add("Sites", new string[] { "Id", "RootWebId", "FullUrl", "Deleted" });
            required.Add("Webs", new string[] { "Id", "SiteId", "ParentWebId", "FullUrl", "Title", "DeleteTransactionId" });
            required.Add("AllLists", new string[] { "tp_ID", "tp_SiteId", "tp_WebId", "tp_Title", "tp_BaseType", "tp_RootFolder", "tp_DeleteTransactionId" });
            // SizeRead/SizeWrite and the ABS stream columns are optional; they are
            // absent from older generations and are only queried when present.
            required.Add("AllDocs", new string[] { "Id", "SiteId", "WebId", "ListId", "ParentId", "DirName", "LeafName", "Type", "Level", "IsCurrentVersion", "DeleteTransactionId", "InternalVersion", "UIVersion", "HasStream", "TimeLastModified", "StreamSchema", "Size" });
            required.Add("DocsToStreams", new string[] { "SiteId", "DocId", "HistVersion", "Level", "Partition", "BSN", "StreamId" });
            required.Add("DocStreams", new string[] { "SiteId", "DocId", "Partition", "BSN", "Type", "Size", "RbsId", "Content" });
            required.Add("Versions", new string[] { "VersionId", "Version" });
            Dictionary<string, HashSet<string>> actual = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            using (SqlConnection connection = OpenConnection())
            using (SqlCommand command = CreateCommand(connection, @"
SELECT t.name AS TableName, c.name AS ColumnName
FROM sys.objects AS t
INNER JOIN sys.schemas AS s ON s.schema_id = t.schema_id
INNER JOIN sys.columns AS c ON c.object_id = t.object_id
WHERE s.name = N'dbo' AND t.type IN (N'U', N'V') AND t.name IN (N'Sites', N'Webs', N'AllLists', N'AllDocs', N'DocsToStreams', N'DocStreams', N'Versions');"))
            using (SqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    string table = reader.GetString(0);
                    HashSet<string> columns;
                    if (!actual.TryGetValue(table, out columns))
                    {
                        columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        actual.Add(table, columns);
                    }
                    columns.Add(reader.GetString(1));
                }
            }
            List<string> missing = new List<string>();
            foreach (KeyValuePair<string, string[]> table in required)
            {
                HashSet<string> columns;
                if (!actual.TryGetValue(table.Key, out columns))
                {
                    missing.Add("dbo." + table.Key + " (table or view missing or not visible to this account)");
                    continue;
                }
                foreach (string column in table.Value)
                    if (!columns.Contains(column)) missing.Add("dbo." + table.Key + "." + column);
            }
            if (missing.Count != 0)
                throw new InvalidOperationException("This database does not expose the required SharePoint content schema: " + String.Join(", ", missing.ToArray()) + ". Use a restored SharePoint content database and grant this account SELECT access.");
            using (SqlConnection connection = OpenConnection()) ReadBuild(connection);
        }

        public List<string> CheckDatabase()
        {
            List<string> summary = new List<string>();
            using (SqlConnection connection = OpenConnection())
            using (SqlCommand command = CreateCommand(connection, @"
SELECT DB_NAME() AS DatabaseName,
       (SELECT COUNT_BIG(*) FROM dbo.Sites WHERE Deleted = 0) AS Sites,
       (SELECT COUNT_BIG(*) FROM dbo.Webs WHERE DeleteTransactionId = 0x) AS Webs,
       (SELECT COUNT_BIG(*) FROM dbo.AllLists WHERE tp_BaseType = 1 AND tp_DeleteTransactionId = 0x) AS Libraries,
       (SELECT COUNT_BIG(*) FROM (SELECT SiteId, Id FROM dbo.AllDocs WHERE Type=0 AND IsCurrentVersion=1 AND DeleteTransactionId=0x GROUP BY SiteId, Id) AS f) AS CurrentFiles;"))
            using (SqlDataReader reader = command.ExecuteReader())
            {
                reader.Read();
                summary.Add("Database: " + reader.GetString(0));
                summary.Add("Site collections: " + reader.GetInt64(1));
                summary.Add("Sites (webs): " + reader.GetInt64(2));
                summary.Add("Document libraries: " + reader.GetInt64(3));
                summary.Add("Distinct current files (all lists): " + reader.GetInt64(4));
            }
            return summary;
        }

        public List<Node> GetRootSites()
        {
            List<Node> nodes = new List<Node>();
            using (SqlConnection connection = OpenConnection())
            using (SqlCommand command = CreateCommand(connection, @"
SELECT s.Id AS SiteId, s.RootWebId AS WebId,
       COALESCE(NULLIF(w.Title, N''), NULLIF(s.FullUrl, N''), NULLIF(w.FullUrl, N''), N'Site ' + CONVERT(nvarchar(36), s.Id)) AS Name,
       COALESCE(NULLIF(s.FullUrl, N''), w.FullUrl, N'') AS FullUrl
FROM dbo.Sites AS s
LEFT JOIN dbo.Webs AS w ON w.Id = s.RootWebId AND w.SiteId = s.Id
WHERE s.Deleted = 0 AND (w.Id IS NULL OR w.DeleteTransactionId = 0x)
ORDER BY Name, s.Id;"))
            using (SqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                    nodes.Add(new Node { Kind = NodeKind.Site, SiteId = reader.GetGuid(0), WebId = reader.GetGuid(1), Id = reader.GetGuid(1), Name = reader.GetString(2), Path = reader.GetString(3) });
            }
            return nodes;
        }

        public List<Node> GetChildren(Node parent)
        {
            if (parent == null) return GetRootSites();
            if (!parent.IsContainer) return new List<Node>();
            if (parent.Kind == NodeKind.List || (parent.Kind == NodeKind.Folder && parent.ListBaseType.HasValue && parent.ListBaseType.Value != 1))
                return new ListCatalog(this).GetChildren(parent);
            return parent.Kind == NodeKind.Site ? GetWebChildren(parent) : GetFolderChildren(parent);
        }

        private List<Node> GetWebChildren(Node parent)
        {
            List<Node> nodes = new List<Node>();
            using (SqlConnection connection = OpenConnection())
            {
                using (SqlCommand command = CreateCommand(connection, @"
SELECT Id, COALESCE(NULLIF(Title, N''), NULLIF(FullUrl, N''), CONVERT(nvarchar(36), Id)) AS Name, COALESCE(FullUrl, N'') AS FullUrl
FROM dbo.Webs
WHERE SiteId = @SiteId AND ParentWebId = @WebId AND DeleteTransactionId = 0x
ORDER BY Name, Id;"))
                {
                    AddGuid(command, "@SiteId", parent.SiteId);
                    AddGuid(command, "@WebId", parent.WebId);
                    using (SqlDataReader reader = command.ExecuteReader())
                        while (reader.Read())
                            nodes.Add(new Node { Kind = NodeKind.Site, SiteId = parent.SiteId, WebId = reader.GetGuid(0), Id = reader.GetGuid(0), Name = reader.GetString(1), Path = reader.GetString(2) });
                }
                using (SqlCommand command = CreateCommand(connection, @"
SELECT l.tp_ID, l.tp_RootFolder,
       COALESCE(NULLIF(l.tp_Title, N''), NULLIF(r.LeafName, N''), CONVERT(nvarchar(36), l.tp_ID)) AS Name,
       r.DirName, r.LeafName, r.TimeLastModified
FROM dbo.AllLists AS l
OUTER APPLY
(
    SELECT TOP (1) d.DirName, d.LeafName, d.TimeLastModified
    FROM dbo.AllDocs AS d
    WHERE d.SiteId = l.tp_SiteId AND d.Id = l.tp_RootFolder AND d.Type = 1
      AND d.IsCurrentVersion = 1 AND d.DeleteTransactionId = 0x
    ORDER BY d.Level DESC, d.InternalVersion DESC
) AS r
WHERE l.tp_SiteId = @SiteId AND l.tp_WebId = @WebId
  AND l.tp_BaseType = 1 AND l.tp_DeleteTransactionId = 0x
ORDER BY Name, l.tp_ID;"))
                {
                    AddGuid(command, "@SiteId", parent.SiteId);
                    AddGuid(command, "@WebId", parent.WebId);
                    using (SqlDataReader reader = command.ExecuteReader())
                        while (reader.Read())
                            nodes.Add(new Node { Kind = NodeKind.Library, SiteId = parent.SiteId, WebId = parent.WebId, ListId = reader.GetGuid(0), Id = reader.GetGuid(1), Name = reader.GetString(2), Path = CombinePath(ReadString(reader, 3), ReadString(reader, 4)), Modified = reader.IsDBNull(5) ? DateTime.MinValue : reader.GetDateTime(5) });
                }
            }
            nodes.AddRange(new ListCatalog(this).GetWebLists(parent));
            return nodes;
        }

        private List<Node> GetFolderChildren(Node parent)
        {
            List<Node> nodes = new List<Node>();
            using (SqlConnection connection = OpenConnection())
            using (SqlCommand command = CreateCommand(connection, @"
WITH CurrentChildren AS
(
    SELECT Id, SiteId, WebId, ListId, DirName, LeafName, Type,
           " + DocumentSize("") + @" AS FileSize,
           TimeLastModified, StreamSchema, Level, InternalVersion, UIVersion, ParentId, HasStream,
           ROW_NUMBER() OVER (PARTITION BY Id ORDER BY Level DESC, InternalVersion DESC) AS rn
    FROM dbo.AllDocs
    WHERE SiteId = @SiteId AND WebId = @WebId AND ListId = @ListId AND ParentId = @ParentId
      AND Type IN (0, 1) AND IsCurrentVersion = 1 AND DeleteTransactionId = 0x
)
SELECT Id, SiteId, WebId, ListId, DirName, LeafName, Type, FileSize, TimeLastModified, StreamSchema, Level, InternalVersion, UIVersion, ParentId, HasStream
FROM CurrentChildren WHERE rn = 1
ORDER BY Type DESC, LeafName, Id;"))
            {
                AddGuid(command, "@SiteId", parent.SiteId);
                AddGuid(command, "@WebId", parent.WebId);
                AddGuid(command, "@ListId", parent.ListId);
                AddGuid(command, "@ParentId", parent.Id);
                using (SqlDataReader reader = command.ExecuteReader())
                    while (reader.Read()) nodes.Add(ReadDocument(reader));
            }
            return nodes;
        }

        public Node GetFile(Guid siteId, Guid fileId)
        {
            using (SqlConnection connection = OpenConnection())
            using (SqlCommand command = CreateCommand(connection, @"
SELECT TOP (1) Id, SiteId, WebId, ListId, DirName, LeafName, Type,
       " + DocumentSize("") + @" AS FileSize,
       TimeLastModified, StreamSchema, Level, InternalVersion, UIVersion, ParentId, HasStream
FROM dbo.AllDocs
WHERE SiteId = @SiteId AND Id = @FileId AND Type = 0
  AND IsCurrentVersion = 1 AND DeleteTransactionId = 0x
ORDER BY Level DESC, InternalVersion DESC;"))
            {
                AddGuid(command, "@SiteId", siteId);
                AddGuid(command, "@FileId", fileId);
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (!reader.Read()) throw new InvalidOperationException("No active current file was found with file ID " + fileId + " in site collection " + siteId + ". Refresh the location to inspect available files.");
                    return ReadDocument(reader);
                }
            }
        }

        public List<Node> GetFileVersions(Guid siteId,Guid fileId)
        {
            return new SqlVersionCatalog(this).GetFileVersions(siteId,fileId);
        }

        public Node GetFileVersion(Node selectedVersion)
        {
            return new SqlVersionCatalog(this).GetFileVersion(selectedVersion);
        }

        // Keyset pagination releases each SQL reader before document recovery starts.
        // Scope is active current files in active document libraries, independent of
        // whether the console tree can reach their folder hierarchy.
        public IEnumerable<Node> EnumerateCurrentFiles(Guid? siteId)
        {
            const int pageSize = 250;
            bool hasAfter = false;
            Guid afterSite = Guid.Empty;
            Guid afterFile = Guid.Empty;
            while (true)
            {
                List<Node> page = new List<Node>();
                using (SqlConnection connection = OpenConnection())
                using (SqlCommand command = CreateCommand(connection, @"
WITH CurrentFiles AS
(
    SELECT d.Id, d.SiteId, d.WebId, d.ListId, d.DirName, d.LeafName, d.Type,
           " + DocumentSize("d") + @" AS FileSize,
           d.TimeLastModified, d.StreamSchema, d.Level, d.InternalVersion, d.UIVersion, d.ParentId, d.HasStream,
           ROW_NUMBER() OVER (PARTITION BY d.SiteId, d.Id ORDER BY d.Level DESC, d.InternalVersion DESC) AS rn
    FROM dbo.AllDocs AS d
    WHERE d.Type=0 AND d.IsCurrentVersion=1 AND d.DeleteTransactionId=0x
      AND (@SiteFilter IS NULL OR d.SiteId=@SiteFilter)
      AND (@HasAfter=0 OR d.SiteId>@AfterSite OR (d.SiteId=@AfterSite AND d.Id>@AfterFile))
      AND EXISTS (SELECT 1 FROM dbo.AllLists AS l WHERE l.tp_SiteId=d.SiteId AND l.tp_ID=d.ListId
                  AND l.tp_BaseType=1 AND l.tp_DeleteTransactionId=0x)
      AND EXISTS (SELECT 1 FROM dbo.Sites AS s WHERE s.Id=d.SiteId AND s.Deleted=0)
      AND EXISTS (SELECT 1 FROM dbo.Webs AS w WHERE w.SiteId=d.SiteId AND w.Id=d.WebId AND w.DeleteTransactionId=0x)
)
SELECT TOP (@PageSize) Id, SiteId, WebId, ListId, DirName, LeafName, Type, FileSize,
       TimeLastModified, StreamSchema, Level, InternalVersion, UIVersion, ParentId, HasStream
FROM CurrentFiles WHERE rn=1 ORDER BY SiteId, Id;"))
                {
                    command.Parameters.Add("@SiteFilter", SqlDbType.UniqueIdentifier).Value = siteId.HasValue ? (object)siteId.Value : DBNull.Value;
                    command.Parameters.Add("@HasAfter", SqlDbType.Bit).Value = hasAfter;
                    command.Parameters.Add("@PageSize", SqlDbType.Int).Value = pageSize;
                    AddGuid(command, "@AfterSite", afterSite);
                    AddGuid(command, "@AfterFile", afterFile);
                    using (SqlDataReader reader = command.ExecuteReader())
                        while (reader.Read()) page.Add(ReadDocument(reader));
                }
                if (page.Count == 0) yield break;
                foreach (Node file in page) yield return file;
                Node last = page[page.Count - 1];
                afterSite = last.SiteId;
                afterFile = last.Id;
                hasAfter = true;
                if (page.Count < pageSize) yield break;
            }
        }

        public IEnumerable<Node> EnumerateCurrentLibraryFiles(Node library)
        {
            if(library == null) throw new ArgumentNullException("library");
            if(library.Kind != NodeKind.Library || library.HistoryVersion != 0 || library.SiteId == Guid.Empty ||
                library.WebId == Guid.Empty || library.ListId == Guid.Empty || library.Id == Guid.Empty)
                throw new ArgumentException("Select a current document library with its complete source scope.","library");
            // Copy before returning an iterator; callers cannot change scope during discovery.
            return EnumerateLibraryFiles(library.SiteId,library.WebId,library.ListId,library.Id);
        }


        public Node GetCurrentLibraryFile(Node library,Guid fileId)
        {
            if(library == null) throw new ArgumentNullException("library");
            if(library.Kind != NodeKind.Library || library.HistoryVersion != 0 || library.SiteId == Guid.Empty ||
                library.WebId == Guid.Empty || library.ListId == Guid.Empty || library.Id == Guid.Empty || fileId == Guid.Empty)
                throw new ArgumentException("Select a current document library and file with complete source scope.","library");
            using(SqlConnection connection=OpenConnection())
            using(SqlCommand command=CreateCommand(connection,@"
WITH CurrentDocument AS
(
    SELECT d.Id,d.SiteId,d.WebId,d.ListId,d.DirName,d.LeafName,d.Type,
           " + DocumentSize("d") + @" AS FileSize,
           d.TimeLastModified,d.StreamSchema,d.Level,d.InternalVersion,d.UIVersion,d.ParentId,d.HasStream,
           ROW_NUMBER() OVER(PARTITION BY d.Id ORDER BY d.Level DESC,d.InternalVersion DESC) AS rn
    FROM dbo.AllDocs AS d
    WHERE d.SiteId=@SiteId AND d.Id=@FileId AND d.Type=0
      AND d.IsCurrentVersion=1 AND d.DeleteTransactionId=0x
)
SELECT Id,SiteId,WebId,ListId,DirName,LeafName,Type,FileSize,
       TimeLastModified,StreamSchema,Level,InternalVersion,UIVersion,ParentId,HasStream
FROM CurrentDocument AS d
WHERE rn=1 AND WebId=@WebId AND ListId=@ListId
  AND EXISTS(SELECT 1 FROM dbo.AllLists AS l WHERE l.tp_SiteId=d.SiteId AND l.tp_WebId=d.WebId
              AND l.tp_ID=d.ListId AND l.tp_RootFolder=@RootId AND l.tp_BaseType=1 AND l.tp_DeleteTransactionId=0x)
  AND EXISTS(SELECT 1 FROM dbo.Sites AS s WHERE s.Id=d.SiteId AND s.Deleted=0)
  AND EXISTS(SELECT 1 FROM dbo.Webs AS w WHERE w.SiteId=d.SiteId AND w.Id=d.WebId AND w.DeleteTransactionId=0x);"))
            {
                AddGuid(command,"@SiteId",library.SiteId); AddGuid(command,"@FileId",fileId);
                AddGuid(command,"@WebId",library.WebId); AddGuid(command,"@ListId",library.ListId); AddGuid(command,"@RootId",library.Id);
                using(SqlDataReader reader=command.ExecuteReader()) return reader.Read() ? ReadDocument(reader) : null;
            }
        }

        private IEnumerable<Node> EnumerateLibraryFiles(Guid siteId,Guid webId,Guid listId,Guid rootId)
        {
            using(SqlConnection connection=OpenConnection())
            using(SqlCommand command=CreateCommand(connection,@"
SELECT COUNT_BIG(*) FROM dbo.AllLists AS l
WHERE l.tp_SiteId=@SiteId AND l.tp_WebId=@WebId AND l.tp_ID=@ListId AND l.tp_RootFolder=@RootId
  AND l.tp_BaseType=1 AND l.tp_DeleteTransactionId=0x
  AND EXISTS(SELECT 1 FROM dbo.Sites AS s WHERE s.Id=l.tp_SiteId AND s.Deleted=0)
  AND EXISTS(SELECT 1 FROM dbo.Webs AS w WHERE w.SiteId=l.tp_SiteId AND w.Id=l.tp_WebId AND w.DeleteTransactionId=0x);"))
            {
                AddGuid(command,"@SiteId",siteId); AddGuid(command,"@WebId",webId);
                AddGuid(command,"@ListId",listId); AddGuid(command,"@RootId",rootId);
                if(Convert.ToInt64(command.ExecuteScalar()) != 1)
                    throw new ContentUnavailableException("The selected document library is no longer active in this source.");
            }
            const int pageSize=250;
            bool hasAfter=false;
            Guid afterFile=Guid.Empty;
            while(true)
            {
                var page=new List<Node>();
                using(SqlConnection connection=OpenConnection())
                using(SqlCommand command=CreateCommand(connection,@"
WITH CurrentFiles AS
(
    SELECT d.Id,d.SiteId,d.WebId,d.ListId,d.DirName,d.LeafName,d.Type,
           " + DocumentSize("d") + @" AS FileSize,
           d.TimeLastModified,d.StreamSchema,d.Level,d.InternalVersion,d.UIVersion,d.ParentId,d.HasStream,
           ROW_NUMBER() OVER(PARTITION BY d.Id ORDER BY d.Level DESC,d.InternalVersion DESC) AS rn
    FROM dbo.AllDocs AS d
    WHERE d.SiteId=@SiteId AND d.Type=0 AND d.IsCurrentVersion=1 AND d.DeleteTransactionId=0x
      AND (@HasAfter=0 OR d.Id>@AfterFile)
)
SELECT TOP (@PageSize) Id,SiteId,WebId,ListId,DirName,LeafName,Type,FileSize,
       TimeLastModified,StreamSchema,Level,InternalVersion,UIVersion,ParentId,HasStream
FROM CurrentFiles AS d
WHERE rn=1 AND WebId=@WebId AND ListId=@ListId
  AND EXISTS(SELECT 1 FROM dbo.AllLists AS l WHERE l.tp_SiteId=d.SiteId AND l.tp_WebId=d.WebId
              AND l.tp_ID=d.ListId AND l.tp_RootFolder=@RootId AND l.tp_BaseType=1 AND l.tp_DeleteTransactionId=0x)
  AND EXISTS(SELECT 1 FROM dbo.Sites AS s WHERE s.Id=d.SiteId AND s.Deleted=0)
  AND EXISTS(SELECT 1 FROM dbo.Webs AS w WHERE w.SiteId=d.SiteId AND w.Id=d.WebId AND w.DeleteTransactionId=0x)
ORDER BY Id;"))
                {
                    AddGuid(command,"@SiteId",siteId); AddGuid(command,"@WebId",webId);
                    AddGuid(command,"@ListId",listId); AddGuid(command,"@RootId",rootId);
                    AddGuid(command,"@AfterFile",afterFile);
                    command.Parameters.Add("@HasAfter",SqlDbType.Bit).Value=hasAfter;
                    command.Parameters.Add("@PageSize",SqlDbType.Int).Value=pageSize;
                    using(SqlDataReader reader=command.ExecuteReader())
                        while(reader.Read()) page.Add(ReadDocument(reader));
                }
                foreach(Node file in page) yield return file;
                if(page.Count<pageSize) yield break;
                afterFile=page[page.Count-1].Id; hasAfter=true;
            }
        }

        private static Node ReadDocument(SqlDataReader reader)
        {
            return new Node
            {
                Id = reader.GetGuid(0), SiteId = reader.GetGuid(1), WebId = reader.GetGuid(2), ListId = reader.IsDBNull(3) ? Guid.Empty : reader.GetGuid(3),
                Path = CombinePath(ReadString(reader, 4), ReadString(reader, 5)), Name = ReadString(reader, 5),
                Kind = Convert.ToInt32(reader.GetValue(6)) == 1 ? NodeKind.Folder : NodeKind.File,
                Size = reader.GetInt64(7), Modified = reader.IsDBNull(8) ? DateTime.MinValue : reader.GetDateTime(8),
                StreamSchema = reader.IsDBNull(9) ? (byte)0 : Convert.ToByte(reader.GetValue(9)), Level = reader.IsDBNull(10) ? (byte)0 : Convert.ToByte(reader.GetValue(10)),
                InternalVersion = reader.IsDBNull(11) ? 0 : Convert.ToInt32(reader.GetValue(11)),
                UiVersion = reader.IsDBNull(12) ? 0 : Convert.ToInt32(reader.GetValue(12)),
                ParentId = reader.IsDBNull(13) ? Guid.Empty : reader.GetGuid(13), HistoryVersion = 0,
                HasStream = reader.IsDBNull(14) ? (bool?)null : Convert.ToInt32(reader.GetValue(14)) != 0
            };
        }

        private static string ReadString(SqlDataReader reader, int index)
        {
            return reader.IsDBNull(index) ? String.Empty : reader.GetString(index);
        }

        private static string CombinePath(string directory, string leaf)
        {
            if (String.IsNullOrEmpty(directory)) return leaf;
            if (String.IsNullOrEmpty(leaf)) return directory;
            return directory.TrimEnd('/') + "/" + leaf;
        }

        private static SqlCommand CreateCommand(SqlConnection connection, string sql)
        {
            SqlCommand command = connection.CreateCommand();
            command.CommandText = sql;
            command.CommandTimeout = QueryTimeout;
            return command;
        }

        private static void AddGuid(SqlCommand command, string name, Guid value)
        {
            command.Parameters.Add(name, SqlDbType.UniqueIdentifier).Value = value;
        }
    }
}
