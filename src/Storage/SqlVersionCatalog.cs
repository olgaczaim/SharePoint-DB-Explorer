using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;

namespace SharePointExplorer
{
    // AllDocVersions.UIVersion is DocsToStreams.HistVersion. InternalVersion
    // remains a separate state counter and is never used as the history key.
    internal sealed class SqlVersionCatalog : ISharePointVersionCatalog
    {
        private readonly SqlRepository repository;
        internal SqlVersionCatalog(SqlRepository repository)
        {
            if(repository==null) throw new ArgumentNullException("repository");
            this.repository=repository;
        }
        public List<Node> GetFileVersions(Guid siteId,Guid fileId)
        {
            if(siteId==Guid.Empty || fileId==Guid.Empty) throw new ArgumentException("Select a document with its source identity.");
            using(SqlConnection connection=repository.OpenConnection())
            {
                ValidateVersionSchema(connection);
                Node current=GetActiveCurrent(connection,siteId,fileId);
                if(current==null) throw new ContentUnavailableException("The selected document is no longer active in this source.");
                var versions=new List<Node> {current};
                using(SqlCommand command=Command(connection,@"
SELECT UIVersion,InternalVersion,Level,TimeCreated,
       " + repository.VersionSize("") + @" AS FileSize,StreamSchema,HasStream
FROM dbo.AllDocVersions
WHERE SiteId=@SiteId AND Id=@FileId AND DeleteTransactionId=0x AND UIVersion>0
ORDER BY UIVersion DESC;"))
                {
                    AddGuid(command,"@SiteId",siteId);AddGuid(command,"@FileId",fileId);
                    using(SqlDataReader reader=command.ExecuteReader())
                        while(reader.Read()) versions.Add(ReadHistorical(reader,current));
                }
                return versions;
            }
        }
        public Node GetFileVersion(Node selectedVersion)
        {
            ValidateIdentity(selectedVersion);
            // Caller mutation cannot alter the identity between the current scope
            // lookup and the historical metadata query.
            var selected=new Node {Kind=selectedVersion.Kind,SiteId=selectedVersion.SiteId,Id=selectedVersion.Id,
                WebId=selectedVersion.WebId,ListId=selectedVersion.ListId,HistoryVersion=selectedVersion.HistoryVersion,
                UiVersion=selectedVersion.UiVersion,Level=selectedVersion.Level,InternalVersion=selectedVersion.InternalVersion};
            using(SqlConnection connection=repository.OpenConnection())
            {
                ValidateVersionSchema(connection);
                Node current=GetActiveCurrent(connection,selected.SiteId,selected.Id);
                if(current==null || current.WebId!=selected.WebId || current.ListId!=selected.ListId) return null;
                if(selected.HistoryVersion==0)
                    return current.UiVersion==selected.UiVersion && current.Level==selected.Level &&
                        current.InternalVersion==selected.InternalVersion ? current : null;
                using(SqlCommand command=Command(connection,@"
SELECT UIVersion,InternalVersion,Level,TimeCreated,
       " + repository.VersionSize("") + @" AS FileSize,StreamSchema,HasStream
FROM dbo.AllDocVersions
WHERE SiteId=@SiteId AND Id=@FileId AND DeleteTransactionId=0x
  AND UIVersion=@Version AND Level=@Level AND COALESCE(InternalVersion,0)=@InternalVersion;"))
                {
                    AddGuid(command,"@SiteId",selected.SiteId);AddGuid(command,"@FileId",selected.Id);
                    command.Parameters.Add("@Version",SqlDbType.Int).Value=selected.HistoryVersion;
                    command.Parameters.Add("@Level",SqlDbType.TinyInt).Value=selected.Level;
                    command.Parameters.Add("@InternalVersion",SqlDbType.Int).Value=selected.InternalVersion;
                    using(SqlDataReader reader=command.ExecuteReader()) return reader.Read() ? ReadHistorical(reader,current) : null;
                }
            }
        }
        private static void ValidateIdentity(Node selected)
        {
            if(selected==null) throw new ArgumentNullException("selectedVersion");
            if(selected.IsDeleted || selected.Kind!=NodeKind.File || selected.SiteId==Guid.Empty || selected.Id==Guid.Empty ||
                selected.HistoryVersion<0 || selected.UiVersion<0 || selected.InternalVersion<0 ||
                (selected.HistoryVersion>0 && selected.HistoryVersion!=selected.UiVersion))
                throw new ArgumentException("Select a complete document version identity.","selectedVersion");
        }
        private static void ValidateVersionSchema(SqlConnection connection)
        {
            using(SqlCommand command=Command(connection,@"
SELECT COUNT_BIG(*) FROM sys.columns
WHERE object_id=OBJECT_ID(N'dbo.AllDocVersions') AND name IN
 (N'SiteId',N'Id',N'UIVersion',N'InternalVersion',N'Level',N'TimeCreated',N'DeleteTransactionId',N'Size',N'StreamSchema',N'HasStream');"))
                if(Convert.ToInt64(command.ExecuteScalar())!=10)
                    throw new NotSupportedException("Historical document metadata is not available in this database or is not visible to this account.");
        }
        private Node GetActiveCurrent(SqlConnection connection,Guid siteId,Guid fileId)
        {
            using(SqlCommand command=Command(connection,@"
WITH CurrentFile AS
(
    SELECT d.*,ROW_NUMBER() OVER(ORDER BY d.Level DESC,d.InternalVersion DESC) AS rn
    FROM dbo.AllDocs AS d
    WHERE d.SiteId=@SiteId AND d.Id=@FileId AND d.Type=0 AND d.IsCurrentVersion=1 AND d.DeleteTransactionId=0x
)
SELECT d.Id,d.SiteId,d.WebId,d.ListId,d.DirName,d.LeafName,
       " + repository.DocumentSize("d") + @" AS FileSize,
       d.TimeLastModified,d.StreamSchema,d.Level,d.InternalVersion,d.UIVersion,d.ParentId,d.HasStream
FROM CurrentFile AS d
WHERE rn=1
  AND EXISTS(SELECT 1 FROM dbo.Sites AS s WHERE s.Id=d.SiteId AND s.Deleted=0)
  AND EXISTS(SELECT 1 FROM dbo.Webs AS w WHERE w.SiteId=d.SiteId AND w.Id=d.WebId AND w.DeleteTransactionId=0x)
  AND (d.ListId IS NULL OR EXISTS(SELECT 1 FROM dbo.AllLists AS l WHERE l.tp_SiteId=d.SiteId AND l.tp_WebId=d.WebId
                                  AND l.tp_ID=d.ListId AND l.tp_DeleteTransactionId=0x));"))
            {
                AddGuid(command,"@SiteId",siteId);AddGuid(command,"@FileId",fileId);
                using(SqlDataReader reader=command.ExecuteReader())
                {
                    if(!reader.Read()) return null;
                    string directory=reader.GetString(4),name=reader.GetString(5);
                    return new Node {Kind=NodeKind.File,Id=reader.GetGuid(0),SiteId=reader.GetGuid(1),WebId=reader.GetGuid(2),
                        ListId=reader.IsDBNull(3) ? Guid.Empty : reader.GetGuid(3),Name=name,
                        Path=String.IsNullOrEmpty(directory) ? name : directory.TrimEnd('/')+"/"+name,
                        Size=reader.GetInt64(6),Modified=reader.IsDBNull(7) ? DateTime.MinValue : reader.GetDateTime(7),
                        StreamSchema=reader.IsDBNull(8) ? (byte)0 : Convert.ToByte(reader.GetValue(8)),
                        Level=Convert.ToByte(reader.GetValue(9)),InternalVersion=reader.IsDBNull(10) ? 0 : Convert.ToInt32(reader.GetValue(10)),
                        UiVersion=Convert.ToInt32(reader.GetValue(11)),ParentId=reader.GetGuid(12),HistoryVersion=0,
                        HasStream=reader.IsDBNull(13) ? (bool?)null : Convert.ToInt32(reader.GetValue(13))!=0};
                }
            }
        }
        internal static Node ReadHistorical(IDataRecord reader,Node current)
        {
            // History does not retain path/name columns. Its location is the active
            // document's location; all content metadata below comes from the old row.
            if(reader.IsDBNull(4)) throw new ContentUnavailableException("A historical version has no recorded content length.");
            if(reader.IsDBNull(5)) throw new NotSupportedException("A historical version has no recorded storage schema; its content cannot be decoded safely.");
            int version=Convert.ToInt32(reader.GetValue(0));
            return new Node {Kind=NodeKind.File,SiteId=current.SiteId,Id=current.Id,WebId=current.WebId,ListId=current.ListId,
                ParentId=current.ParentId,Name=current.Name,Path=current.Path,HistoryVersion=version,UiVersion=version,
                InternalVersion=reader.IsDBNull(1) ? 0 : Convert.ToInt32(reader.GetValue(1)),Level=Convert.ToByte(reader.GetValue(2)),
                Modified=reader.IsDBNull(3) ? DateTime.MinValue : reader.GetDateTime(3),Size=Convert.ToInt64(reader.GetValue(4)),
                StreamSchema=Convert.ToByte(reader.GetValue(5)),HasStream=reader.IsDBNull(6) ? (bool?)null : reader.GetBoolean(6)};
        }
        private static SqlCommand Command(SqlConnection connection,string text)
        {
            SqlCommand command=connection.CreateCommand();command.CommandText=text;command.CommandTimeout=120;return command;
        }
        private static void AddGuid(SqlCommand command,string name,Guid value)
        {
            command.Parameters.Add(name,SqlDbType.UniqueIdentifier).Value=value;
        }
    }
}