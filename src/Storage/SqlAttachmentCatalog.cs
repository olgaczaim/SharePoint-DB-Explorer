using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using Microsoft.Data.SqlClient;

namespace SharePointExplorer
{
    public sealed class SqlAttachmentCatalog : ISharePointAttachmentCatalog
    {
        private readonly SqlRepository repository;
        private const string CurrentDocs=@"
WITH CurrentDocs AS
(
    SELECT d.*,ROW_NUMBER() OVER(PARTITION BY d.Id ORDER BY d.Level DESC,d.InternalVersion DESC) rn
    FROM dbo.AllDocs d WHERE d.SiteId=@SiteId AND d.IsCurrentVersion=1 AND d.DeleteTransactionId=0x
)";
        private const string CurrentItems=@",
CurrentItems AS
(
    SELECT u.*,ROW_NUMBER() OVER(PARTITION BY u.tp_ID ORDER BY u.tp_Level DESC,u.tp_Version DESC) rn
    FROM dbo.AllUserData u
    WHERE u.tp_SiteId=@SiteId AND u.tp_ListId=@ListId AND u.tp_RowOrdinal=0
      AND u.tp_IsCurrent=1 AND u.tp_IsCurrentVersion=1 AND u.tp_CalculatedVersion=0 AND u.tp_DeleteTransactionId=0x
)";
        private const string OwnerColumns=@"
SELECT u.tp_DocId,u.tp_ID,u.tp_GUID,u.tp_ParentId,u.tp_Modified,u.tp_Created,u.tp_HasAttachment,
       u.tp_Level,u.tp_Version,u.tp_UIVersion,d.DirName,d.LeafName
FROM CurrentItems u JOIN CurrentDocs d ON d.Id=u.tp_DocId AND d.rn=1 AND d.Type=0
    AND d.WebId=@WebId AND d.ListId=@ListId AND d.DoclibRowId=u.tp_ID AND d.ParentId=u.tp_ParentId
    AND d.Level=u.tp_Level AND d.UIVersion=u.tp_UIVersion
WHERE u.rn=1";
        public SqlAttachmentCatalog(SqlRepository repository)
        {
            if(repository==null) throw new ArgumentNullException("repository");
            this.repository=repository;
        }
        public List<Node> GetItemAttachments(Node listItem)
        {
            ValidateItem(listItem);
            Node selection=CopyItem(listItem);
            using(SqlConnection connection=repository.OpenConnection())
            {
                ValidateSchema(connection);
                Node list=GetList(connection,selection,false);
                Node owner=GetOwner(connection,selection,list.ListBaseType.Value);
                if(owner==null) throw new ContentUnavailableException("The selected attachment owner is no longer a current item in this list.");
                if(owner.HasAttachments!=true) return new List<Node>();
                Node folder=GetItemFolder(connection,list,owner);
                return ReadFiles(connection,owner,folder,null);
            }
        }
        public Node GetCurrentItemAttachment(Node listItem,Guid attachmentDocId)
        {
            ValidateItem(listItem);
            if(attachmentDocId==Guid.Empty) throw new ArgumentException("Select an attachment with its document identity.","attachmentDocId");
            Node selection=CopyItem(listItem);
            using(SqlConnection connection=repository.OpenConnection())
            {
                ValidateSchema(connection);
                Node list=GetList(connection,selection,false);
                Node owner=GetOwner(connection,selection,list.ListBaseType.Value);
                if(owner==null) throw new ContentUnavailableException("The selected attachment owner is no longer a current item in this list.");
                if(owner.HasAttachments!=true) return null;
                Node folder=GetItemFolder(connection,list,owner);
                List<Node> files=ReadFiles(connection,owner,folder,attachmentDocId);
                return files.Count==0 ? null : files[0];
            }
        }
        public IEnumerable<Node> EnumerateCurrentListAttachments(Node list)
        {
            ValidateList(list);
            // Freeze caller scope before creating the lazy iterator.
            Node selection=new Node {Kind=list.Kind,SiteId=list.SiteId,WebId=list.WebId,ListId=list.ListId,Id=list.Id,ListBaseType=list.ListBaseType};
            return Enumerate(selection);
        }
        private IEnumerable<Node> Enumerate(Node selection)
        {
            Node list;
            using(SqlConnection connection=repository.OpenConnection())
            {
                ValidateSchema(connection);list=GetList(connection,selection,true);
            }
            const int pageSize=250;
            int after=0;
            while(true)
            {
                var owners=new List<Node>();
                using(SqlConnection connection=repository.OpenConnection())
                {
                    // Revalidate activity/root on each page, not merely at discovery start.
                    GetList(connection,selection,true);
                    using(SqlCommand command=Command(connection,CurrentDocs+CurrentItems+OwnerColumns.Replace("SELECT u.tp_DocId","SELECT TOP(@PageSize) u.tp_DocId")+@"
  AND u.tp_HasAttachment=1 AND u.tp_ID>@After ORDER BY u.tp_ID;"))
                    {
                        Scope(command,list);command.Parameters.Add("@PageSize",SqlDbType.Int).Value=pageSize;
                        command.Parameters.Add("@After",SqlDbType.Int).Value=after;
                        using(SqlDataReader reader=command.ExecuteReader())
                            while(reader.Read()) owners.Add(ReadOwner(reader,list,list.ListBaseType.Value));
                    }
                }
                foreach(Node owner in owners)
                    foreach(Node file in GetItemAttachments(owner)) yield return file;
                if(owners.Count<pageSize) yield break;
                after=owners[owners.Count-1].ListItemId.Value;
            }
        }
        internal static void ValidateItem(Node item)
        {
            if(item==null) throw new ArgumentNullException("listItem");
            if(item.Kind!=NodeKind.ListItem || item.HistoryVersion!=0 || item.SiteId==Guid.Empty || item.WebId==Guid.Empty ||
                item.ListId==Guid.Empty || item.Id==Guid.Empty || !item.ListItemId.HasValue || item.ListItemId.Value<=0 ||
                !item.ItemUniqueId.HasValue || item.ItemUniqueId.Value==Guid.Empty || (item.ListBaseType.HasValue && item.ListBaseType.Value==1))
                throw new ArgumentException("Select a current ordinary-list item with its complete owner identity.","listItem");
        }
        internal static void ValidateList(Node list)
        {
            if(list==null) throw new ArgumentNullException("list");
            if(list.Kind!=NodeKind.List || list.HistoryVersion!=0 || list.SiteId==Guid.Empty || list.WebId==Guid.Empty ||
                list.ListId==Guid.Empty || list.Id==Guid.Empty || (list.ListBaseType.HasValue && list.ListBaseType.Value==1))
                throw new ArgumentException("Select a current ordinary list with its complete source scope.","list");
        }
        private static void ValidateSchema(SqlConnection connection)
        {
            using(SqlCommand command=Command(connection,@"
SELECT COUNT_BIG(*) FROM sys.columns WHERE object_id=OBJECT_ID(N'dbo.AllUserData') AND name IN
(N'tp_ID',N'tp_ListId',N'tp_SiteId',N'tp_RowOrdinal',N'tp_Version',N'tp_Modified',N'tp_Created',N'tp_HasAttachment',N'tp_IsCurrent',
 N'tp_GUID',N'tp_ParentId',N'tp_DocId',N'tp_DeleteTransactionId',N'tp_Level',N'tp_IsCurrentVersion',N'tp_UIVersion',N'tp_CalculatedVersion');"))
                if(Convert.ToInt64(command.ExecuteScalar())!=17)
                    throw new NotSupportedException("Current ordinary-list attachment metadata is not available in this database or is not visible to this account.");
        }
        private static Node GetList(SqlConnection connection,Node selection,bool requireRoot)
        {
            using(SqlCommand command=Command(connection,CurrentDocs+@"
SELECT l.tp_RootFolder,l.tp_BaseType,d.DirName,d.LeafName
FROM dbo.AllLists l JOIN CurrentDocs d ON d.Id=l.tp_RootFolder AND d.rn=1 AND d.Type=1
    AND d.WebId=l.tp_WebId AND d.ListId=l.tp_ID
WHERE l.tp_SiteId=@SiteId AND l.tp_WebId=@WebId AND l.tp_ID=@ListId AND l.tp_BaseType<>1 AND l.tp_DeleteTransactionId=0x
  AND EXISTS(SELECT 1 FROM dbo.Sites s WHERE s.Id=l.tp_SiteId AND s.Deleted=0)
  AND EXISTS(SELECT 1 FROM dbo.Webs w WHERE w.SiteId=l.tp_SiteId AND w.Id=l.tp_WebId AND w.DeleteTransactionId=0x);"))
            {
                Scope(command,selection);
                using(SqlDataReader reader=command.ExecuteReader())
                {
                    if(!reader.Read()) throw new ContentUnavailableException("The selected ordinary list is no longer active in this source.");
                    Guid root=reader.GetGuid(0);
                    if(requireRoot && root!=selection.Id) throw new InvalidDataException("The selected ordinary-list root differs from the current source.");
                    return new Node {Kind=NodeKind.List,SiteId=selection.SiteId,WebId=selection.WebId,ListId=selection.ListId,Id=root,
                        ListBaseType=reader.GetInt32(1),Path=Combine(Text(reader,2),Text(reader,3)),Name=Text(reader,3)};
                }
            }
        }
        private static Node GetOwner(SqlConnection connection,Node selection,int baseType)
        {
            using(SqlCommand command=Command(connection,CurrentDocs+CurrentItems+OwnerColumns+@"
  AND u.tp_ID=@ItemId AND u.tp_GUID=@ItemGuid AND u.tp_DocId=@OwnerDocId;"))
            {
                Scope(command,selection);command.Parameters.Add("@ItemId",SqlDbType.Int).Value=selection.ListItemId.Value;
                GuidParameter(command,"@ItemGuid",selection.ItemUniqueId.Value);GuidParameter(command,"@OwnerDocId",selection.Id);
                using(SqlDataReader reader=command.ExecuteReader()) return reader.Read() ? ReadOwner(reader,selection,baseType) : null;
            }
        }
        private static Node ReadOwner(IDataRecord reader,Node scope,int baseType)
        {
            int id=reader.GetInt32(1);
            return new Node {Kind=NodeKind.ListItem,SiteId=scope.SiteId,WebId=scope.WebId,ListId=scope.ListId,ListBaseType=baseType,
                Id=reader.GetGuid(0),ListItemId=id,ItemUniqueId=reader.GetGuid(2),ParentId=reader.GetGuid(3),
                Modified=reader.IsDBNull(4) ? DateTime.MinValue : reader.GetDateTime(4),Created=reader.IsDBNull(5) ? (DateTime?)null : reader.GetDateTime(5),
                HasAttachments=reader.GetBoolean(6),Level=reader.GetByte(7),InternalVersion=reader.GetInt32(8),UiVersion=reader.GetInt32(9),
                Name="Item "+id.ToString(CultureInfo.InvariantCulture),Path=Combine(Text(reader,10),Text(reader,11)),HasStream=false};
        }
        private static Node GetItemFolder(SqlConnection connection,Node list,Node owner)
        {
            Node attachments=GetFolder(connection,list,list.Id,list.Path,"Attachments");
            if(attachments==null) throw new ContentUnavailableException("The attachment folder for this ordinary list is absent from the source.");
            Node item=GetFolder(connection,list,attachments.Id,attachments.Path,owner.ListItemId.Value.ToString(CultureInfo.InvariantCulture));
            if(item==null) throw new ContentUnavailableException("The attachment folder for the selected item is absent from the source.");
            return item;
        }
        private static Node GetFolder(SqlConnection connection,Node scope,Guid parent,string directory,string name)
        {
            using(SqlCommand command=Command(connection,CurrentDocs+@"
SELECT Id FROM CurrentDocs
WHERE rn=1 AND WebId=@WebId AND ListId=@ListId AND Type=1 AND ParentId=@ParentId AND DirName=@Directory AND LeafName=@Name;"))
            {
                Scope(command,scope);GuidParameter(command,"@ParentId",parent);
                command.Parameters.Add("@Directory",SqlDbType.NVarChar,400).Value=directory;
                command.Parameters.Add("@Name",SqlDbType.NVarChar,400).Value=name;
                using(SqlDataReader reader=command.ExecuteReader())
                {
                    if(!reader.Read()) return null;
                    Guid id=reader.GetGuid(0);
                    if(reader.Read()) throw new InvalidDataException("Multiple current folders claim the same attachment owner path.");
                    return new Node {Kind=NodeKind.Folder,SiteId=scope.SiteId,WebId=scope.WebId,ListId=scope.ListId,Id=id,ParentId=parent,Name=name,Path=Combine(directory,name)};
                }
            }
        }
        private List<Node> ReadFiles(SqlConnection connection,Node owner,Node folder,Guid? fileId)
        {
            var files=new List<Node>();
            using(SqlCommand command=Command(connection,CurrentDocs+@"
SELECT Id,SiteId,WebId,ListId,DirName,LeafName,Type,"+repository.DocumentSize("")+@",
       TimeLastModified,StreamSchema,Level,InternalVersion,UIVersion,ParentId,HasStream
FROM CurrentDocs
WHERE rn=1 AND WebId=@WebId AND ListId=@ListId AND Type=0 AND ParentId=@ParentId AND DirName=@Directory
  AND Id<>@OwnerDocId AND (@FileId IS NULL OR Id=@FileId) ORDER BY LeafName,Id;"))
            {
                Scope(command,owner);GuidParameter(command,"@ParentId",folder.Id);GuidParameter(command,"@OwnerDocId",owner.Id);
                command.Parameters.Add("@Directory",SqlDbType.NVarChar,400).Value=folder.Path;
                command.Parameters.Add("@FileId",SqlDbType.UniqueIdentifier).Value=fileId.HasValue ? (object)fileId.Value : DBNull.Value;
                using(SqlDataReader reader=command.ExecuteReader())
                    while(reader.Read()) files.Add(ReadAttachment(reader,owner,folder.Id,folder.Path));
            }
            return files;
        }
        internal static Node ReadAttachment(IDataRecord reader,Node owner,Guid folderId,string folderPath)
        {
            Guid id=reader.GetGuid(0);
            if(id==Guid.Empty || id==owner.Id || reader.GetGuid(1)!=owner.SiteId || reader.GetGuid(2)!=owner.WebId || reader.IsDBNull(3) ||
                reader.GetGuid(3)!=owner.ListId || reader.GetByte(6)!=0 || reader.GetGuid(13)!=folderId ||
                !String.Equals(Text(reader,4),folderPath,StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("The source returned a file outside the selected attachment owner.");
            return new Node {Kind=NodeKind.File,Id=id,SiteId=owner.SiteId,WebId=owner.WebId,ListId=owner.ListId,
                ParentId=folderId,Name=Text(reader,5),Path=Combine(folderPath,Text(reader,5)),Size=reader.GetInt64(7),
                Modified=reader.IsDBNull(8) ? DateTime.MinValue : reader.GetDateTime(8),StreamSchema=reader.IsDBNull(9) ? (byte)0 : Convert.ToByte(reader.GetValue(9)),
                Level=reader.GetByte(10),InternalVersion=reader.IsDBNull(11) ? 0 : reader.GetInt32(11),UiVersion=reader.GetInt32(12),HistoryVersion=0,
                HasStream=reader.IsDBNull(14) ? (bool?)null : Convert.ToInt32(reader.GetValue(14))!=0,ListBaseType=owner.ListBaseType,
                AttachmentOwnerId=owner.Id,ListItemId=owner.ListItemId,ItemUniqueId=owner.ItemUniqueId};
        }
        private static Node CopyItem(Node item)
        {
            return new Node {Kind=item.Kind,SiteId=item.SiteId,WebId=item.WebId,ListId=item.ListId,Id=item.Id,HistoryVersion=item.HistoryVersion,
                ListItemId=item.ListItemId,ItemUniqueId=item.ItemUniqueId,ListBaseType=item.ListBaseType};
        }
        private static SqlCommand Command(SqlConnection connection,string text){SqlCommand command=connection.CreateCommand();command.CommandText=text;command.CommandTimeout=120;return command;}
        private static void Scope(SqlCommand command,Node node){GuidParameter(command,"@SiteId",node.SiteId);GuidParameter(command,"@WebId",node.WebId);GuidParameter(command,"@ListId",node.ListId);}
        private static void GuidParameter(SqlCommand command,string name,Guid value){command.Parameters.Add(name,SqlDbType.UniqueIdentifier).Value=value;}
        private static string Text(IDataRecord reader,int ordinal){return reader.IsDBNull(ordinal) ? String.Empty : reader.GetString(ordinal);}
        private static string Combine(string directory,string leaf){return String.IsNullOrEmpty(directory) ? leaf : directory.TrimEnd('/')+"/"+leaf;}
    }
}