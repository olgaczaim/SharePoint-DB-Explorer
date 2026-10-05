using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using Microsoft.Data.SqlClient;

namespace SharePointExplorer
{
    public sealed partial class SqlRepository : ISharePointDeletedCatalog
    {
        public List<Node> GetDeletedItems(Node site) { return new SqlDeletedCatalog(this).GetDeletedItems(site); }
        public List<Node> GetDeletedChildren(Node container) { return new SqlDeletedCatalog(this).GetDeletedChildren(container); }
        public Node GetDeletedFile(Node selected) { return new SqlDeletedCatalog(this).GetDeletedFile(selected); }
        public List<Node> GetDeletedFileVersions(Node selected) { return new SqlDeletedCatalog(this).GetDeletedFileVersions(selected); }
        public DeletedListItemSnapshot GetDeletedListItem(Node selected) { return new SqlDeletedCatalog(this).GetDeletedListItem(selected); }
    }

    // Retained recycle rows are a separate namespace. Every read uses the
    // stored binary deletion transaction; active rows never supply file bytes.
    internal sealed class SqlDeletedCatalog : ISharePointDeletedCatalog
    {
        private readonly SqlRepository repository;
        internal SqlDeletedCatalog(SqlRepository repository)
        { this.repository=repository ?? throw new ArgumentNullException("repository"); }

        private const string SizeToken="/*DocumentSize*/";
        private string RowsSql { get { return Rows.Replace(SizeToken,repository.DocumentSize("d")); } }
        private const string Rows=@"
WITH DeletedRows AS
(
 SELECT d.*,ROW_NUMBER() OVER(PARTITION BY d.Id,d.DeleteTransactionId ORDER BY d.Level DESC,d.InternalVersion DESC) rn
 FROM dbo.AllDocs d WHERE d.SiteId=@SiteId AND d.WebId=@WebId AND d.DeleteTransactionId<>0x
 AND d.IsCurrentVersion=1 AND d.Type IN(0,1)
)
SELECT d.Id,d.SiteId,d.WebId,d.ListId,d.ParentId,d.DirName,d.LeafName,d.Type,
 /*DocumentSize*/,d.StreamSchema,d.HasStream,
 d.Level,d.InternalVersion,d.UIVersion,d.TimeLastModified,d.TimeCreated,d.DeleteTransactionId,
 l.tp_BaseType,l.tp_RootFolder,l.tp_Title,u.tp_ID,u.tp_GUID,u.tp_Version,u.tp_HasAttachment,r.Title,r.DeleteDate
FROM DeletedRows d
OUTER APPLY(SELECT TOP(1) z.tp_BaseType,z.tp_RootFolder,z.tp_Title FROM dbo.AllLists z
 WHERE z.tp_SiteId=d.SiteId AND z.tp_WebId=d.WebId AND z.tp_ID=d.ListId
 AND z.tp_DeleteTransactionId IN(0x,d.DeleteTransactionId)
 ORDER BY CASE WHEN z.tp_DeleteTransactionId=d.DeleteTransactionId THEN 0 ELSE 1 END) l
OUTER APPLY(SELECT TOP(1) z.tp_ID,z.tp_GUID,z.tp_Version,z.tp_HasAttachment FROM dbo.AllUserData z
 WHERE z.tp_SiteId=d.SiteId AND z.tp_ListId=d.ListId AND z.tp_DocId=d.Id AND z.tp_Level=d.Level
 AND z.tp_DeleteTransactionId=d.DeleteTransactionId AND z.tp_RowOrdinal=0
 AND z.tp_IsCurrent=1 AND z.tp_IsCurrentVersion=1 AND z.tp_CalculatedVersion=0 ORDER BY z.tp_Version DESC) u
OUTER APPLY(SELECT TOP(1) z.Title,z.DeleteDate FROM dbo.RecycleBin z
 WHERE z.SiteId=d.SiteId AND z.WebId=d.WebId AND z.EffectiveDeleteTransactionId=d.DeleteTransactionId
 ORDER BY CASE WHEN z.DocId=d.Id THEN 0 ELSE 1 END,z.DeleteDate DESC) r
WHERE d.rn=1 ";

        public List<Node> GetDeletedItems(Node site)
        {
            if(site==null || (site.Kind!=NodeKind.Site && site.Kind!=NodeKind.DeletedItems) || site.SiteId==Guid.Empty || site.WebId==Guid.Empty)
                throw new ArgumentException("Select a site to browse its retained deleted items.","site");
            var scope=ScopeCopy(site);
            using(SqlConnection connection=repository.OpenConnection())
            {
                ValidateSchema(connection);
                List<Node> result=ReadRows(connection,scope,@"
 AND NOT EXISTS(SELECT 1 FROM DeletedRows p WHERE p.rn=1 AND p.Id=d.ParentId
 AND p.DeleteTransactionId=d.DeleteTransactionId AND p.ListId=d.ListId AND p.Type=1)
 ORDER BY d.DirName,d.LeafName,d.Id;",null);
                foreach(Node version in ReadVersions(connection,scope,null))
                    if(!result.Exists(node=>node.Id==version.Id && String.Equals(node.DeletionTransactionId,version.DeletionTransactionId,StringComparison.OrdinalIgnoreCase)))
                        result.Add(version);
                return result;
            }
        }
        public List<Node> GetDeletedChildren(Node container)
        {
            DeletedIdentity.Validate(container,false);
            if(container.Kind==NodeKind.File)throw new ArgumentException("Select a deleted folder, list or list item.","container");
            Node selected=DeletedIdentity.Copy(container);
            using(SqlConnection connection=repository.OpenConnection())
            {
                ValidateSchema(connection);
                Node current=ResolveRow(connection,selected);
                if(!DeletedIdentity.Same(selected,current))throw new ContentUnavailableException("The selected deleted container is no longer retained with this identity.");
                if(current.Kind==NodeKind.ListItem)return ReadItemAttachments(connection,current);
                return ReadRows(connection,current," AND d.ListId=@ListId AND d.ParentId=@ParentId AND d.DeleteTransactionId=@Deletion ORDER BY d.Type DESC,d.LeafName,d.Id;",
                    command=>{GuidParameter(command,"@ListId",current.ListId);GuidParameter(command,"@ParentId",current.Id);DeletionParameter(command,current.DeletionTransactionId);});
            }
        }
        public Node GetDeletedFile(Node selected)
        {
            DeletedIdentity.Validate(selected,true);Node identity=DeletedIdentity.Copy(selected);
            using(SqlConnection connection=repository.OpenConnection())
            {
                ValidateSchema(connection);
                Node document;
                if(identity.HistoryVersion>0)
                    document=ReadVersions(connection,identity,identity).Find(node=>DeletedIdentity.Same(node,identity));
                else if(identity.AttachmentOwnerId.HasValue)
                {
                    Node owner=ReadRows(connection,identity," AND d.ListId=@ListId AND d.Id=@OwnerId AND d.DeleteTransactionId=@Deletion;",
                        command=>{GuidParameter(command,"@ListId",identity.ListId);GuidParameter(command,"@OwnerId",identity.AttachmentOwnerId.Value);DeletionParameter(command,identity.DeletionTransactionId);}).Find(node=>node.Kind==NodeKind.ListItem);
                    document=owner==null ? null : ReadItemAttachments(connection,owner).Find(node=>DeletedIdentity.Same(node,identity));
                }
                else document=ResolveRow(connection,identity);
                return DeletedIdentity.Same(document,identity) ? document : null;
            }
        }
        public List<Node> GetDeletedFileVersions(Node selected)
        {
            DeletedIdentity.Validate(selected,true);Node identity=DeletedIdentity.Copy(selected);
            using(SqlConnection connection=repository.OpenConnection())
            {
                ValidateSchema(connection);
                var result=new List<Node>();
                Node current=ReadRows(connection,identity," AND d.Id=@DocId AND d.ListId=@ListId AND d.DeleteTransactionId=@Deletion;",
                    command=>{GuidParameter(command,"@DocId",identity.Id);GuidParameter(command,"@ListId",identity.ListId);DeletionParameter(command,identity.DeletionTransactionId);}).Find(node=>node.Kind==NodeKind.File);
                if(current!=null && !identity.AttachmentOwnerId.HasValue)result.Add(current);
                else if(current!=null && identity.HistoryVersion==0)
                {
                    Node exactAttachment=GetDeletedFile(identity);if(exactAttachment!=null)result.Add(exactAttachment);
                }
                result.AddRange(ReadVersions(connection,identity,null));
                if(!result.Exists(node=>DeletedIdentity.Same(node,identity)))throw new ContentUnavailableException("The selected deleted document or version is no longer retained.");
                return result;
            }
        }
        public DeletedListItemSnapshot GetDeletedListItem(Node selected)
        {
            DeletedIdentity.Validate(selected,false);
            if(selected.Kind!=NodeKind.ListItem || !selected.ListItemId.HasValue || !selected.ItemUniqueId.HasValue)
                throw new ArgumentException("Select a deleted ordinary list item with its retained item identity.","selected");
            Node identity=DeletedIdentity.Copy(selected);
            using(SqlConnection connection=repository.OpenConnection())
            {
                ValidateSchema(connection);Node current=ResolveRow(connection,identity);
                if(!DeletedIdentity.Same(current,identity))throw new ContentUnavailableException("The deleted list item is no longer retained with its selected identity.");
                string fields=String.Empty;
                using(SqlCommand command=Command(connection,@"SELECT TOP(1) tp_Fields FROM dbo.AllLists WHERE tp_SiteId=@SiteId AND tp_WebId=@WebId AND tp_ID=@ListId
 AND tp_DeleteTransactionId IN(0x,@Deletion) ORDER BY CASE WHEN tp_DeleteTransactionId=@Deletion THEN 0 ELSE 1 END;"))
                {
                    Scope(command,identity);GuidParameter(command,"@ListId",identity.ListId);DeletionParameter(command,identity.DeletionTransactionId);
                    object raw=command.ExecuteScalar();if(raw!=null && raw!=DBNull.Value)fields=MigrationFieldSchema.Decode((byte[])raw,"Fields");
                }
                // AllUserData uses a SQL sparse column set. SELECT * substitutes
                // that XML column for the individual user fields, losing their
                // typed values. An explicit trusted schema projection reads them.
                var columns=new List<string>();
                using(SqlCommand schema=Command(connection,"SELECT name FROM sys.columns WHERE object_id=OBJECT_ID(N'dbo.AllUserData') AND is_column_set=0 ORDER BY column_id;"))
                using(SqlDataReader reader=schema.ExecuteReader())while(reader.Read())columns.Add(reader.GetString(0));
                string projection=BuildRetainedProjection(columns);
                var values=new List<DeletedListItemValue>();
                using(SqlCommand command=Command(connection,@"SELECT "+projection+@" FROM dbo.AllUserData u WHERE u.tp_SiteId=@SiteId AND u.tp_ListId=@ListId
 AND u.tp_ID=@ItemId AND u.tp_GUID=@UniqueId AND u.tp_DocId=@DocId AND u.tp_DeleteTransactionId=@Deletion
 AND u.tp_Level=@Level AND u.tp_Version=@InternalVersion AND u.tp_IsCurrent=1 AND u.tp_IsCurrentVersion=1 AND u.tp_CalculatedVersion=0
 ORDER BY u.tp_RowOrdinal;"))
                {
                    Scope(command,identity);GuidParameter(command,"@ListId",identity.ListId);GuidParameter(command,"@UniqueId",identity.ItemUniqueId.Value);
                    GuidParameter(command,"@DocId",identity.Id);DeletionParameter(command,identity.DeletionTransactionId);
                    command.Parameters.Add("@ItemId",SqlDbType.Int).Value=identity.ListItemId.Value;
                    command.Parameters.Add("@Level",SqlDbType.TinyInt).Value=identity.Level;
                    command.Parameters.Add("@InternalVersion",SqlDbType.Int).Value=identity.InternalVersion;
                    using(SqlDataReader reader=command.ExecuteReader())
                    {
                        int ordinal=reader.GetOrdinal("tp_RowOrdinal");bool found=false;
                        while(reader.Read())
                        {
                            found=true;int row=Convert.ToInt32(reader.GetValue(ordinal),CultureInfo.InvariantCulture);
                            for(int column=0;column<reader.FieldCount;column++)
                            {
                                object raw=reader.GetValue(column);bool absent=raw==DBNull.Value;
                                string text=absent ? String.Empty : raw is byte[] ? Convert.ToBase64String((byte[])raw) : raw is DateTime ? ((DateTime)raw).ToString("O",CultureInfo.InvariantCulture) : Convert.ToString(raw,CultureInfo.InvariantCulture);
                                values.Add(new DeletedListItemValue(row,reader.GetName(column),reader.GetDataTypeName(column),text,absent));
                            }
                        }
                        if(!found)throw new ContentUnavailableException("The selected deleted item's retained field rows are absent.");
                    }
                }
                if(!DeletedIdentity.Same(current,ResolveRow(connection,identity)))throw new ContentUnavailableException("The deleted list item changed while its metadata was being read.");
                return new DeletedListItemSnapshot(current,fields,values);
            }
        }
        internal static string BuildRetainedProjection(IEnumerable<string> columns)
        {
            if(columns==null)throw new ArgumentNullException("columns");
            var quoted=new List<string>();var names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach(string column in columns)
            {
                if(String.IsNullOrEmpty(column) || !names.Add(column))throw new InvalidDataException("The retained item schema contains an empty or repeated column identity.");
                quoted.Add("u.["+column.Replace("]","]]")+"]");
            }
            if(quoted.Count==0 || quoted.Count>4096)throw new NotSupportedException("The retained item schema exceeds SQL's supported explicit projection limit.");
            return String.Join(",",quoted.ToArray());
        }
        private Node ResolveRow(SqlConnection connection,Node selected)
        {
            return ReadRows(connection,selected," AND d.Id=@DocId AND d.ListId=@ListId AND d.DeleteTransactionId=@Deletion;",
                command=>{GuidParameter(command,"@DocId",selected.Id);GuidParameter(command,"@ListId",selected.ListId);DeletionParameter(command,selected.DeletionTransactionId);}).Find(node=>node.Kind==selected.Kind && node.Level==selected.Level);
        }
        private List<Node> ReadRows(SqlConnection connection,Node scope,string condition,Action<SqlCommand> parameters)
        {
            var result=new List<Node>();
            using(SqlCommand command=Command(connection,RowsSql+condition))
            {
                Scope(command,scope);if(parameters!=null)parameters(command);
                using(SqlDataReader reader=command.ExecuteReader())while(reader.Read())result.Add(ReadNode(reader));
            }
            return result;
        }
        internal static Node ReadNode(IDataRecord row)
        {
            byte type=Convert.ToByte(row.GetValue(7),CultureInfo.InvariantCulture);
            if(type!=0 && type!=1)throw new InvalidDataException("The source returned an unsupported deleted document type.");
            Guid id=row.GetGuid(0);int? baseType=row.IsDBNull(17)?(int?)null:Convert.ToInt32(row.GetValue(17));
            bool listRoot=type==1 && !row.IsDBNull(18) && id==row.GetGuid(18);
            bool item=type==0 && baseType.HasValue && baseType.Value!=1 && !row.IsDBNull(20);
            string name=Text(row,6),title=Text(row,24);
            NodeKind kind=listRoot ? baseType==1 ? NodeKind.Library : NodeKind.List : type==1 ? NodeKind.Folder : item ? NodeKind.ListItem : NodeKind.File;
            var node=new Node {Kind=kind,Id=id,SiteId=row.GetGuid(1),WebId=row.GetGuid(2),ListId=row.IsDBNull(3)?Guid.Empty:row.GetGuid(3),
                ParentId=row.GetGuid(4),Path=Combine(Text(row,5),name),Name=listRoot && !String.IsNullOrWhiteSpace(Text(row,19)) ? Text(row,19) : item ? String.IsNullOrWhiteSpace(title) ? "Item "+Convert.ToInt32(row.GetValue(20)).ToString(CultureInfo.InvariantCulture) : title : name,
                Size=Convert.ToInt64(row.GetValue(8),CultureInfo.InvariantCulture),StreamSchema=row.IsDBNull(9)?(byte)0:Convert.ToByte(row.GetValue(9)),
                HasStream=row.IsDBNull(10)?(bool?)null:Convert.ToInt32(row.GetValue(10))!=0,Level=Convert.ToByte(row.GetValue(11)),
                InternalVersion=item ? Convert.ToInt32(row.GetValue(22)) : row.IsDBNull(12)?0:Convert.ToInt32(row.GetValue(12)),UiVersion=Convert.ToInt32(row.GetValue(13)),HistoryVersion=0,
                Modified=row.IsDBNull(14)?DateTime.MinValue:row.GetDateTime(14),Created=row.IsDBNull(15)?(DateTime?)null:row.GetDateTime(15),
                DeletionTransactionId=DeletedIdentity.Format((byte[])row.GetValue(16)),ListBaseType=baseType,
                ListItemId=item?(int?)Convert.ToInt32(row.GetValue(20)):null,ItemUniqueId=item && !row.IsDBNull(21)?(Guid?)row.GetGuid(21):null,
                HasAttachments=item && !row.IsDBNull(23)?(bool?)Convert.ToBoolean(row.GetValue(23)):null,Title=item?title:null,
                DeletedAt=row.IsDBNull(25)?(DateTime?)null:row.GetDateTime(25)};
            if(item)node.HasStream=false;
            return node;
        }
        private List<Node> ReadItemAttachments(SqlConnection connection,Node owner)
        {
            if(owner.Kind!=NodeKind.ListItem || !owner.ListItemId.HasValue)throw new ArgumentException("Select a retained deleted ordinary list item.","owner");
            string condition=@" AND d.ListId=@ListId AND d.Type=0 AND d.DeleteTransactionId=@Deletion
 AND EXISTS(SELECT 1 FROM dbo.AllDocs f JOIN dbo.AllDocs a ON a.SiteId=f.SiteId AND a.WebId=f.WebId AND a.ListId=f.ListId
 AND a.Id=f.ParentId AND a.Type=1 AND a.IsCurrentVersion=1 AND a.DeleteTransactionId IN(0x,f.DeleteTransactionId) AND a.LeafName=N'Attachments'
 WHERE f.SiteId=d.SiteId AND f.WebId=d.WebId AND f.ListId=d.ListId AND f.Id=d.ParentId AND f.Type=1 AND f.IsCurrentVersion=1
 AND f.DeleteTransactionId=d.DeleteTransactionId AND f.LeafName=@OwnerIdText AND EXISTS(SELECT 1 FROM dbo.AllLists l WHERE l.tp_SiteId=f.SiteId AND l.tp_WebId=f.WebId AND l.tp_ID=f.ListId AND l.tp_RootFolder=a.ParentId AND l.tp_DeleteTransactionId IN(0x,f.DeleteTransactionId)))
 ORDER BY d.LeafName,d.Id;";
            // The source list root and folder GUIDs establish ownership.
            var result=ReadRows(connection,owner,condition,command=>{
                GuidParameter(command,"@ListId",owner.ListId);DeletionParameter(command,owner.DeletionTransactionId);
                command.Parameters.Add("@OwnerIdText",SqlDbType.NVarChar,400).Value=owner.ListItemId.Value.ToString(CultureInfo.InvariantCulture);
            });
            foreach(Node file in result)
            {
                if(file.Kind!=NodeKind.File)throw new InvalidDataException("Deleted attachment metadata resolved to an item backing row.");
                file.AttachmentOwnerId=owner.Id;file.ListItemId=owner.ListItemId;file.ItemUniqueId=owner.ItemUniqueId;
            }
            return result;
        }
        private List<Node> ReadVersions(SqlConnection connection,Node scope,Node selected)
        {
            var result=new List<Node>();
            using(SqlCommand validation=Command(connection,@"SELECT COUNT_BIG(*) FROM sys.columns WHERE object_id=OBJECT_ID(N'dbo.AllDocVersions')
 AND name IN(N'SiteId',N'Id',N'DeleteTransactionId',N'UIVersion',N'InternalVersion',N'Level',N'TimeCreated',N'Size',N'StreamSchema',N'HasStream');"))
                if(Convert.ToInt64(validation.ExecuteScalar())!=10)return result;
            string filter=scope.Id!=Guid.Empty && scope.Kind!=NodeKind.Site && scope.Kind!=NodeKind.DeletedItems ? " AND v.Id=@DocId AND d.ListId=@ListId AND v.DeleteTransactionId=@Deletion" : String.Empty;
            if(selected!=null)filter+=" AND v.UIVersion=@Version AND v.Level=@Level AND COALESCE(v.InternalVersion,0)=@InternalVersion";
            using(SqlCommand command=Command(connection,@"
SELECT v.UIVersion,v.InternalVersion,v.Level,v.TimeCreated,"+repository.VersionSize("v")+@",v.StreamSchema,v.HasStream,
 v.DeleteTransactionId,d.Id,d.SiteId,d.WebId,d.ListId,d.ParentId,d.DirName,d.LeafName,r.DeleteDate
FROM dbo.AllDocVersions v
CROSS APPLY(SELECT TOP(1) a.* FROM dbo.AllDocs a WHERE a.SiteId=v.SiteId AND a.Id=v.Id AND a.WebId=@WebId
 AND a.Type=0 AND a.IsCurrentVersion=1 AND a.DeleteTransactionId IN(0x,v.DeleteTransactionId) ORDER BY a.Level DESC,a.InternalVersion DESC) d
OUTER APPLY(SELECT TOP(1) z.DeleteDate FROM dbo.RecycleBin z WHERE z.SiteId=v.SiteId AND z.EffectiveDeleteTransactionId=v.DeleteTransactionId
 ORDER BY z.DeleteDate DESC) r
WHERE v.SiteId=@SiteId AND v.DeleteTransactionId<>0x AND v.UIVersion>0"+filter+" ORDER BY d.DirName,d.LeafName,v.UIVersion DESC;"))
            {
                Scope(command,scope);
                if(filter.Length>0){GuidParameter(command,"@DocId",scope.Id);GuidParameter(command,"@ListId",scope.ListId);DeletionParameter(command,scope.DeletionTransactionId);}
                if(selected!=null)
                {
                    command.Parameters.Add("@Version",SqlDbType.Int).Value=selected.HistoryVersion;
                    command.Parameters.Add("@Level",SqlDbType.TinyInt).Value=selected.Level;
                    command.Parameters.Add("@InternalVersion",SqlDbType.Int).Value=selected.InternalVersion;
                }
                using(SqlDataReader reader=command.ExecuteReader())while(reader.Read())
                {
                    var context=new Node {Kind=NodeKind.File,Id=reader.GetGuid(8),SiteId=reader.GetGuid(9),WebId=reader.GetGuid(10),ListId=reader.IsDBNull(11)?Guid.Empty:reader.GetGuid(11),
                        ParentId=reader.GetGuid(12),Path=Combine(Text(reader,13),Text(reader,14)),Name=Text(reader,14)};
                    Node version=SqlVersionCatalog.ReadHistorical(reader,context);version.DeletionTransactionId=DeletedIdentity.Format((byte[])reader.GetValue(7));
                    version.DeletedAt=reader.IsDBNull(15)?(DateTime?)null:reader.GetDateTime(15);result.Add(version);
                }
            }
            return result;
        }
        private static void ValidateSchema(SqlConnection connection)
        {
            using(SqlCommand command=Command(connection,@"SELECT COUNT_BIG(*) FROM sys.columns WHERE object_id=OBJECT_ID(N'dbo.RecycleBin')
 AND name IN(N'SiteId',N'WebId',N'DocId',N'EffectiveDeleteTransactionId',N'Title',N'DeleteDate');"))
                if(Convert.ToInt64(command.ExecuteScalar())!=6)throw new NotSupportedException("This restored database does not expose the required retained recycle-bin metadata.");
        }
        private static Node ScopeCopy(Node node) {return new Node {Kind=node.Kind,SiteId=node.SiteId,WebId=node.WebId,Id=node.Id,ListId=node.ListId};}
        private static SqlCommand Command(SqlConnection connection,string sql){var command=connection.CreateCommand();command.CommandText=sql;command.CommandTimeout=120;return command;}
        private static void Scope(SqlCommand command,Node node){GuidParameter(command,"@SiteId",node.SiteId);GuidParameter(command,"@WebId",node.WebId);}
        private static void GuidParameter(SqlCommand command,string name,Guid id){command.Parameters.Add(name,SqlDbType.UniqueIdentifier).Value=id;}
        private static void DeletionParameter(SqlCommand command,string transaction){command.Parameters.Add("@Deletion",SqlDbType.VarBinary,16).Value=DeletedIdentity.Parse(transaction);}
        private static string Text(IDataRecord reader,int index){return reader.IsDBNull(index)?String.Empty:reader.GetString(index);}
        private static string Combine(string directory,string name){return String.IsNullOrEmpty(directory)?name:String.IsNullOrEmpty(name)?directory:directory.TrimEnd('/')+"/"+name;}
    }
}