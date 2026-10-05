using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Microsoft.Data.SqlClient;

namespace SharePointExplorer
{
    // Metadata-only capability. File and attachment bytes use the recovery engine.
    public sealed partial class SqlMigrationCatalog : ISharePointMigrationCatalog
    {
        private readonly SqlRepository repository;
        public SqlMigrationCatalog(SqlRepository repository)
        {if(repository==null)throw new ArgumentNullException("repository");this.repository=repository;}
        public MigrationListSnapshot ReadMigrationList(Node listOrLibrary)
        {
            if(listOrLibrary==null)throw new ArgumentNullException("listOrLibrary");
            if((listOrLibrary.Kind!=NodeKind.List && listOrLibrary.Kind!=NodeKind.Library) || listOrLibrary.HistoryVersion!=0 ||
                listOrLibrary.SiteId==Guid.Empty || listOrLibrary.WebId==Guid.Empty || listOrLibrary.ListId==Guid.Empty || listOrLibrary.Id==Guid.Empty)
                throw new ArgumentException("Select a current list or document library with its complete source scope.","listOrLibrary");
            Node selected=MigrationSnapshotCopy.Node(listOrLibrary);
            using(SqlConnection connection=repository.OpenConnection())
            {
                Dictionary<string,Dictionary<string,string>> schema=ReadSchema(connection);ValidateSchema(schema);
                MigrationListMetadata metadata=ReadMetadata(connection,selected,schema);
                List<MigrationFieldDefinition> fields=MigrationFieldSchema.Read(metadata.FieldSchemaXml,schema["AllUserData"],metadata.ListId,metadata.WebId,metadata.ContentTypesXml);
                Dictionary<Guid,int?> rowIds;Dictionary<Guid,MigrationTemplateFile> templateCandidates;
                Dictionary<Guid,Node> documents=ReadDocuments(connection,metadata,out rowIds,out templateCandidates);
                List<MigrationItemSnapshot> items=ReadItems(connection,metadata,fields,documents,rowIds).Select(row=>row.Snapshot).ToList();
                var views=ReadViews(connection,metadata,schema);
                var templates=ReadTemplateFiles(metadata,documents,templateCandidates,items,views);
                var folders=ReadParentFolders(metadata.RootFolderId,documents,items.Select(item=>item.Document).Concat(templates.Select(template=>template.Document)));
                ValidateFolderGraph(metadata.RootFolderId,folders,items);
                var files=items.Where(item=>item.Document.Kind==NodeKind.File).Select(item=>item.Document).ToList();
                if(metadata.BaseType==1)ValidateLibraryFiles(metadata,files);
                return new MigrationListSnapshot(metadata,fields,items,folders,files,
                    ReadUsers(connection,metadata,fields,items,schema,templates),views,templates,SqlDeploymentContext.Read(connection,metadata));
            }
        }
        private static Dictionary<string,Dictionary<string,string>> ReadSchema(SqlConnection connection)
        {
            var result=new Dictionary<string,Dictionary<string,string>>(StringComparer.OrdinalIgnoreCase);
            using(SqlCommand command=Command(connection,@"
SELECT o.name,c.name,TYPE_NAME(c.system_type_id) FROM sys.objects o JOIN sys.columns c ON c.object_id=o.object_id
WHERE o.schema_id=SCHEMA_ID(N'dbo') AND o.type IN(N'U',N'V')
 AND o.name IN(N'AllLists',N'AllUserData',N'AllDocs',N'Sites',N'Webs',N'UserInfo',N'Versions',N'AllWebParts');"))
            using(SqlDataReader reader=command.ExecuteReader())while(reader.Read())
            {
                string table=reader.GetString(0);Dictionary<string,string> columns;
                if(!result.TryGetValue(table,out columns))result.Add(table,columns=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase));
                columns.Add(reader.GetString(1),reader.GetString(2));
            }
            return result;
        }
        private static void Require(Dictionary<string,Dictionary<string,string>> schema,string table,params string[] names)
        {
            Dictionary<string,string> columns;
            if(!schema.TryGetValue(table,out columns))throw new NotSupportedException("Migration metadata requires dbo."+table+" in this content source.");
            foreach(string name in names)if(!columns.ContainsKey(name))throw new NotSupportedException("Migration metadata requires dbo."+table+"."+name+" in this content source.");
        }
        private static void ValidateSchema(Dictionary<string,Dictionary<string,string>> schema)
        {
            Require(schema,"AllLists","tp_SiteId","tp_WebId","tp_ID","tp_Title","tp_Description","tp_BaseType","tp_ServerTemplate","tp_FeatureId","tp_Template",
                "tp_RootFolder","tp_Flags","tp_Flags2","tp_ReadSecurity","tp_WriteSecurity","tp_Direction","tp_Created","tp_Author",
                "tp_MaxMajorVersionCount","tp_MaxMajorwithMinorVersionCount","tp_Fields","tp_ContentTypes","tp_DeleteTransactionId","tp_ImageUrl","tp_Version");
            Require(schema,"AllDocs","SiteId","WebId","ListId","Id","ParentId","DirName","LeafName","Type","Level","InternalVersion","UIVersion",
                "IsCurrentVersion","DeleteTransactionId","Size","StreamSchema","HasStream","TimeCreated","TimeLastModified","DoclibRowId","SetupPath","SetupPathVersion","SetupPathUser");
            Require(schema,"AllUserData","tp_SiteId","tp_ListId","tp_ID","tp_RowOrdinal","tp_Level","tp_Version","tp_IsCurrent","tp_IsCurrentVersion",
                "tp_CalculatedVersion","tp_DeleteTransactionId","tp_DocId","tp_GUID","tp_ParentId","tp_UIVersion","tp_Author","tp_Editor","tp_Created",
                "tp_Modified","tp_HasAttachment","tp_ContentTypeId","tp_ModerationStatus","tp_UIVersionString");
            Require(schema,"Sites","Id","Deleted","FullUrl","PlatformVersion","RootWebId","UserInfoListId");
            Require(schema,"Webs","SiteId","Id","DeleteTransactionId","FullUrl","Title","ProductVersion","Language");
        }
        private static MigrationListMetadata ReadMetadata(SqlConnection connection,Node selected,Dictionary<string,Dictionary<string,string>> schema)
        {
            string version=ReadSourceVersion(connection,schema);
            using(SqlCommand command=Command(connection,@"
SELECT l.tp_Title,l.tp_Description,l.tp_BaseType,l.tp_ServerTemplate,l.tp_FeatureId,l.tp_Template,
 l.tp_RootFolder,l.tp_Flags,l.tp_Flags2,l.tp_ReadSecurity,l.tp_WriteSecurity,l.tp_Direction,
 l.tp_Created,l.tp_Author,l.tp_MaxMajorVersionCount,l.tp_MaxMajorwithMinorVersionCount,l.tp_Fields,l.tp_ContentTypes,
 s.FullUrl,w.FullUrl,w.Title,d.DirName,d.LeafName,d.ParentId,d.Level,d.InternalVersion,d.UIVersion,d.TimeCreated,d.TimeLastModified,
 l.tp_ImageUrl,l.tp_Version,s.PlatformVersion,w.ProductVersion,w.Language
FROM dbo.AllLists l JOIN dbo.Sites s ON s.Id=l.tp_SiteId AND s.Deleted=0
JOIN dbo.Webs w ON w.SiteId=l.tp_SiteId AND w.Id=l.tp_WebId AND w.DeleteTransactionId=0x
OUTER APPLY
(
 SELECT TOP(1) a.DirName,a.LeafName,a.ParentId,a.Level,a.InternalVersion,a.UIVersion,a.TimeCreated,a.TimeLastModified
 FROM dbo.AllDocs a WHERE a.SiteId=l.tp_SiteId AND a.WebId=l.tp_WebId AND a.ListId=l.tp_ID AND a.Id=l.tp_RootFolder
 AND a.Type=1 AND a.IsCurrentVersion=1 AND a.DeleteTransactionId=0x ORDER BY a.Level DESC,a.InternalVersion DESC
) d
WHERE l.tp_SiteId=@SiteId AND l.tp_WebId=@WebId AND l.tp_ID=@ListId AND l.tp_RootFolder=@RootId AND l.tp_DeleteTransactionId=0x;"))
            {
                Scope(command,selected);GuidParameter(command,"@RootId",selected.Id);
                using(SqlDataReader reader=command.ExecuteReader())
                {
                    if(!reader.Read())throw new ContentUnavailableException("The selected list root is no longer active in this content source.");
                    int baseType=Int(reader,2);
                    if((selected.Kind==NodeKind.Library)!=(baseType==1) || (selected.ListBaseType.HasValue && selected.ListBaseType.Value!=baseType))
                        throw new InvalidDataException("The selected list kind does not match its active source metadata.");
                    if(reader.IsDBNull(21) || reader.IsDBNull(22))throw new NotSupportedException("The active list has no current root-folder path metadata.");
                    Node root=new Node {Kind=baseType==1?NodeKind.Library:NodeKind.List,SiteId=selected.SiteId,WebId=selected.WebId,ListId=selected.ListId,
                        Id=reader.GetGuid(6),Name=Text(reader,0),Path=PathOf(Text(reader,21),Text(reader,22)),ParentId=reader.GetGuid(23),ListBaseType=baseType,
                        Level=reader.GetByte(24),InternalVersion=DocumentInternalVersion(reader.GetValue(25),1),UiVersion=Int(reader,26),Created=Date(reader,27),Modified=Date(reader,28),HasStream=false};
                    if(String.IsNullOrWhiteSpace(root.Path))throw new NotSupportedException("The active list root has no recorded source URL.");
                    if(reader.IsDBNull(16))throw new NotSupportedException("The active list has no stored field definitions.");
                    string fields=MigrationFieldSchema.Decode((byte[])reader.GetValue(16),"Fields");
                    string contentTypes=reader.IsDBNull(17)?String.Empty:MigrationFieldSchema.Decode((byte[])reader.GetValue(17),"ContentTypes");
                    var metadata=new MigrationListMetadata(root,Text(reader,0),Text(reader,1),baseType,Int(reader,3),reader.GetGuid(4),
                        reader.IsDBNull(5)?(Guid?)null:reader.GetGuid(5),Long(reader,7),Long(reader,8),Int(reader,9),Int(reader,10),Int(reader,11),
                        Date(reader,12),Int(reader,13),NullableInt(reader,14),NullableInt(reader,15),Text(reader,18),Text(reader,19),Text(reader,20),fields,contentTypes,version,
                        Text(reader,29),Int(reader,30),Text(reader,31),Int(reader,32),Int(reader,33),NullableInt(reader,25));
                    if(reader.Read())throw new InvalidDataException("The source contains duplicate active list metadata.");return metadata;
                }
            }
        }
        private static string ReadSourceVersion(SqlConnection connection,Dictionary<string,Dictionary<string,string>> schema)
        {
            if(!schema.ContainsKey("Versions"))return String.Empty;Require(schema,"Versions","VersionId","Version");
            using(SqlCommand command=Command(connection,"SELECT Version FROM dbo.Versions WHERE VersionId=@VersionId;"))
            {
                GuidParameter(command,"@VersionId",Guid.Empty);
                using(SqlDataReader reader=command.ExecuteReader())
                {
                    if(!reader.Read())return String.Empty;string version=Text(reader,0);Version parsed;
                    if(!Version.TryParse(version,out parsed))throw new NotSupportedException("The recorded source build is not a supported version number.");
                    if(reader.Read())throw new InvalidDataException("The source records multiple product build versions.");return version;
                }
            }
        }
        private Dictionary<Guid,Node> ReadDocuments(SqlConnection connection,MigrationListMetadata metadata,out Dictionary<Guid,int?> rowIds,out Dictionary<Guid,MigrationTemplateFile> templateCandidates)
        {
            var result=new Dictionary<Guid,Node>();rowIds=new Dictionary<Guid,int?>();templateCandidates=new Dictionary<Guid,MigrationTemplateFile>();
            using(SqlCommand command=Command(connection,@"
WITH D AS
(
 SELECT d.Id,d.ParentId,d.DirName,d.LeafName,d.Type,d.Level,d.InternalVersion,d.UIVersion,d.TimeCreated,d.TimeLastModified,
 "+repository.DocumentSize("d")+@" Size,d.StreamSchema,d.HasStream,d.DoclibRowId,d.SetupPath,d.SetupPathVersion,d.SetupPathUser,
 ROW_NUMBER() OVER(PARTITION BY d.Id ORDER BY d.Level DESC,d.InternalVersion DESC) rn
 FROM dbo.AllDocs d WHERE d.SiteId=@SiteId AND d.WebId=@WebId AND d.ListId=@ListId
 AND d.IsCurrentVersion=1 AND d.DeleteTransactionId=0x
)
SELECT Id,ParentId,DirName,LeafName,Type,Level,InternalVersion,UIVersion,TimeCreated,TimeLastModified,Size,StreamSchema,HasStream,DoclibRowId,SetupPath,SetupPathVersion,SetupPathUser
FROM D WHERE rn=1 ORDER BY Type DESC,Id;"))
            {
                Scope(command,metadata.List);
                using(SqlDataReader reader=command.ExecuteReader())while(reader.Read())
                {
                    int documentType=Int(reader,4);
                    if(documentType!=0 && documentType!=1)throw new NotSupportedException("The selected list contains an unsupported document record kind.");
                    var node=new Node {Kind=documentType==1?NodeKind.Folder:NodeKind.File,SiteId=metadata.SiteId,WebId=metadata.WebId,ListId=metadata.ListId,
                        Id=reader.GetGuid(0),ParentId=reader.GetGuid(1),Name=Text(reader,3),Path=PathOf(Text(reader,2),Text(reader,3)),Level=reader.GetByte(5),
                        InternalVersion=DocumentInternalVersion(reader.GetValue(6),documentType),UiVersion=Int(reader,7),Created=Date(reader,8),Modified=Date(reader,9),Size=Long(reader,10),
                        StreamSchema=reader.IsDBNull(11)?(byte)0:reader.GetByte(11),HasStream=reader.IsDBNull(12)?(bool?)null:Int(reader,12)!=0,ListBaseType=metadata.BaseType};
                    if(node.Size<0)throw new InvalidDataException("A current source document has a negative size.");result.Add(node.Id,node);rowIds.Add(node.Id,reader.IsDBNull(13)?(int?)null:Int(reader,13));
                    if(node.Kind==NodeKind.File && node.HasStream==false && !String.IsNullOrWhiteSpace(Text(reader,14)))
                        templateCandidates.Add(node.Id,new MigrationTemplateFile(node,Text(reader,14),reader.IsDBNull(15)?(byte?)null:reader.GetByte(15),reader.IsDBNull(16)?null:reader.GetString(16)));
                }
            }
            return result;
        }
        private sealed class ItemRow
        {
            internal int Id,Version;internal byte Level;internal Node Document;internal Guid UniqueId;internal int Author,Editor;
            internal DateTime Created,Modified;internal bool HasAttachments;internal string ContentType;
            internal Dictionary<byte,Dictionary<string,object>> Ordinals=new Dictionary<byte,Dictionary<string,object>>();
            internal MigrationItemSnapshot Snapshot;
        }
        private static List<ItemRow> ReadItems(SqlConnection connection,MigrationListMetadata metadata,List<MigrationFieldDefinition> fields,Dictionary<Guid,Node> documents,Dictionary<Guid,int?> rowIds)
        {
            var result=new List<ItemRow>();var docIds=new HashSet<Guid>();
            using(SqlCommand command=Command(connection,@"
WITH CurrentItems AS
(
 SELECT u.tp_ID,u.tp_DocId,u.tp_GUID,u.tp_ParentId,u.tp_Level,u.tp_Version,u.tp_UIVersion,u.tp_Author,u.tp_Editor,
 u.tp_Created,u.tp_Modified,u.tp_HasAttachment,u.tp_ContentTypeId,u.tp_ModerationStatus,u.tp_UIVersionString,
 ROW_NUMBER() OVER(PARTITION BY u.tp_ID ORDER BY u.tp_Level DESC,u.tp_Version DESC) rn
 FROM dbo.AllUserData u WHERE u.tp_SiteId=@SiteId AND u.tp_ListId=@ListId AND u.tp_RowOrdinal=0
 AND u.tp_IsCurrent=1 AND u.tp_IsCurrentVersion=1 AND u.tp_CalculatedVersion=0 AND u.tp_DeleteTransactionId=0x
)
SELECT tp_ID,tp_DocId,tp_GUID,tp_ParentId,tp_Level,tp_Version,tp_UIVersion,tp_Author,tp_Editor,tp_Created,tp_Modified,tp_HasAttachment,
 tp_ContentTypeId,tp_ModerationStatus,tp_UIVersionString FROM CurrentItems WHERE rn=1 ORDER BY tp_ID;"))
            {
                Scope(command,metadata.List);
                using(SqlDataReader reader=command.ExecuteReader())while(reader.Read())
                {
                    Guid documentId=reader.GetGuid(1);Node document;
                    if(!documents.TryGetValue(documentId,out document))throw new NotSupportedException("A current item has no active AllDocs path metadata: item "+Int(reader,0)+".");
                    if(!docIds.Add(documentId))throw new InvalidDataException("Multiple current list items reference the same source document.");
                    if(documentId==metadata.RootFolderId)continue;
                    if(document.ParentId!=reader.GetGuid(3))throw new InvalidDataException("Current list-item and document parent-folder identities disagree.");
                    if(document.Level!=reader.GetByte(4) || document.UiVersion!=Int(reader,6))throw new InvalidDataException("Current item and document version metadata disagree.");
                    int? documentRowId=rowIds[documentId];
                    if(!documentRowId.HasValue || documentRowId.Value!=Int(reader,0))throw new InvalidDataException("Current item and document row identities disagree.");
                    Node node=MigrationSnapshotCopy.Node(document);
                    if(metadata.BaseType!=1 && node.Kind!=NodeKind.Folder) {node.Kind=NodeKind.ListItem;node.HasStream=false;}
                    node.ListItemId=Int(reader,0);node.ItemUniqueId=reader.GetGuid(2);node.Title=String.Empty;node.HasAttachments=Bool(reader,11);
                    var row=new ItemRow {Id=Int(reader,0),Document=node,UniqueId=reader.GetGuid(2),Level=reader.GetByte(4),Version=Int(reader,5),
                        Author=Int(reader,7),Editor=Int(reader,8),Created=Date(reader,9),Modified=Date(reader,10),HasAttachments=Bool(reader,11),ContentType=ContentType(reader.GetValue(12))};
                    var system=new Dictionary<string,object>(StringComparer.OrdinalIgnoreCase) {
                        {"tp_ID",row.Id},{"tp_DocId",documentId},{"tp_GUID",row.UniqueId},{"tp_ParentId",node.ParentId},{"tp_UIVersion",Int(reader,6)},
                        {"tp_Author",row.Author},{"tp_Editor",row.Editor},{"tp_Created",row.Created},{"tp_Modified",row.Modified},
                        {"tp_HasAttachment",row.HasAttachments},{"tp_ContentTypeId",row.ContentType},{"tp_ModerationStatus",reader.GetValue(13)},
                        {"tp_UIVersionString",reader.GetValue(14)} };
                    row.Ordinals.Add(0,system);result.Add(row);
                }
            }
            string[] columns=fields.SelectMany(field=>field.StorageMappings).Select(mapping=>mapping.Column).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(column=>column,StringComparer.Ordinal).ToArray();
            if(columns.Length>2000)throw new NotSupportedException("This list exceeds the supported per-query field-storage column limit.");
            for(int offset=0;offset<result.Count;offset+=200)ReadOrdinalValues(connection,metadata,result.Skip(offset).Take(200).ToList(),columns);
            foreach(ItemRow row in result)
            {
                var values=new List<MigrationFieldValue>();
                foreach(MigrationFieldDefinition field in fields)
                {
                    if(field.StorageMappings.Count==0) {AddSystemValue(values,field,row);continue;}
                    foreach(MigrationStorageMapping mapping in field.StorageMappings)
                    {
                        Dictionary<string,object> ordinal;object value;
                        if(!row.Ordinals.TryGetValue(mapping.RowOrdinal,out ordinal))throw new NotSupportedException("Current item "+row.Id+" has no recorded field row ordinal "+mapping.RowOrdinal+".");
                        if(!ordinal.TryGetValue(mapping.Column,out value))throw new InvalidDataException("The current field projection omitted its declared storage column.");
                        values.Add(new MigrationFieldValue(field.Name,field.Id,mapping.Component,mapping.SqlType,value));
                    }
                }
                MigrationFieldValue title=values.FirstOrDefault(value=>value.Name=="Title" && value.Component==1);
                if(title!=null && title.Value!=null)row.Document.Title=Convert.ToString(title.Value,CultureInfo.InvariantCulture);
                row.Document.Created=row.Created;row.Document.Modified=row.Modified;
                row.Snapshot=new MigrationItemSnapshot(row.Document,row.Id,row.UniqueId,row.Author,row.Editor,row.Created,row.Modified,row.HasAttachments,row.ContentType,values);
            }
            return result;
        }
        private static void ReadOrdinalValues(SqlConnection connection,MigrationListMetadata metadata,List<ItemRow> rows,string[] columns)
        {
            if(columns.Length==0)return;
            var byId=rows.ToDictionary(row=>row.Id);
            var parameters=rows.Select((row,index)=>"@Item"+index.ToString(CultureInfo.InvariantCulture)).ToArray();
            string projection=String.Join(",",columns.Select(column=>"u.["+column+"]"));
            using(SqlCommand command=Command(connection,@"
SELECT u.tp_ID,u.tp_RowOrdinal,u.tp_Level,u.tp_Version,"+projection+@"
FROM dbo.AllUserData u WHERE u.tp_SiteId=@SiteId AND u.tp_ListId=@ListId AND u.tp_ID IN("+String.Join(",",parameters)+@")
 AND u.tp_IsCurrent=1 AND u.tp_IsCurrentVersion=1 AND u.tp_CalculatedVersion=0 AND u.tp_DeleteTransactionId=0x
ORDER BY u.tp_ID,u.tp_RowOrdinal,u.tp_Level DESC,u.tp_Version DESC;"))
            {
                Scope(command,metadata.List);for(int index=0;index<rows.Count;index++)command.Parameters.Add(parameters[index],SqlDbType.Int).Value=rows[index].Id;
                var found=new HashSet<string>(StringComparer.Ordinal);
                using(SqlDataReader reader=command.ExecuteReader())while(reader.Read())
                {
                    ItemRow row=byId[Int(reader,0)];byte ordinal=reader.GetByte(1);
                    if(reader.GetByte(2)!=row.Level || Int(reader,3)!=row.Version)continue;
                    if(!found.Add(row.Id+":"+ordinal))throw new InvalidDataException("The source contains duplicate current item storage rows.");
                    Dictionary<string,object> values;if(!row.Ordinals.TryGetValue(ordinal,out values))row.Ordinals.Add(ordinal,values=new Dictionary<string,object>(StringComparer.OrdinalIgnoreCase));
                    for(int index=0;index<columns.Length;index++)values[columns[index]]=reader.IsDBNull(index+4)?null:reader.GetValue(index+4);
                }
            }
        }
        private static void AddSystemValue(List<MigrationFieldValue> values,MigrationFieldDefinition field,ItemRow row)
        {
            string column=null,type=null;object direct=null;bool hasDirect=false;
            switch(field.Name)
            {
                case "ID":column="tp_ID";type="int";break;case "GUID":column="tp_GUID";type="uniqueidentifier";break;
                case "Author":column="tp_Author";type="int";break;case "Editor":column="tp_Editor";type="int";break;
                case "Created":column="tp_Created";type="datetime";break;case "Modified":column="tp_Modified";type="datetime";break;
                case "Attachments":column="tp_HasAttachment";type="bit";break;case "ContentTypeId":column="tp_ContentTypeId";type="varchar";break;
                case "_ModerationStatus":column="tp_ModerationStatus";type="int";break;case "_UIVersion":column="tp_UIVersion";type="int";break;
                case "_UIVersionString":column="tp_UIVersionString";type="nvarchar";break;
                case "FileLeafRef":direct=row.Document.Name;type="nvarchar";hasDirect=true;break;
                case "FileRef":direct=row.Document.Path;type="nvarchar";hasDirect=true;break;
                case "FSObjType":direct=row.Document.Kind==NodeKind.Folder?1:0;type="int";hasDirect=true;break;
                case "UniqueId":direct=row.Document.Id;type="uniqueidentifier";hasDirect=true;break;
            }
            if(hasDirect)values.Add(new MigrationFieldValue(field.Name,field.Id,1,type,direct));
            else if(column!=null)values.Add(new MigrationFieldValue(field.Name,field.Id,1,type,row.Ordinals[0][column]));
            // A virtual/computed FieldRef without storage has no invented value.
        }
        private List<MigrationTemplateFile> ReadTemplateFiles(MigrationListMetadata metadata,Dictionary<Guid,Node> documents,
            Dictionary<Guid,MigrationTemplateFile> candidates,List<MigrationItemSnapshot> items,List<MigrationViewSnapshot> views)
        {
            var ownedAttachments=new HashSet<Guid>();
            if(metadata.BaseType!=1)
                foreach(MigrationItemSnapshot item in items.Where(item=>item.HasAttachments && !item.IsFolder))
                    foreach(Node attachment in repository.GetItemAttachments(item.Document))ownedAttachments.Add(attachment.Id);
            var primary=new HashSet<Guid>(items.Select(item=>item.DocumentId));var result=new List<MigrationTemplateFile>();
            foreach(Node document in documents.Values.Where(document=>document.Kind==NodeKind.File).OrderBy(document=>document.Path,StringComparer.Ordinal).ThenBy(document=>document.Id))
            {
                if(ownedAttachments.Contains(document.Id))continue;
                MigrationTemplateFile template;
                if(candidates.TryGetValue(document.Id,out template)) {result.Add(template);continue;}
                if(!primary.Contains(document.Id))throw new NotSupportedException("A non-item source file has no verified template metadata or attachment owner: "+document.Path+".");
                if(metadata.BaseType==1 && document.HasStream==false)throw new NotSupportedException("A library document has no payload or verified setup-path metadata: "+document.Path+".");
            }
            var templateIds=new HashSet<Guid>(result.Select(template=>template.Document.Id));
            foreach(MigrationViewSnapshot view in views)
                if(!view.FileId.HasValue || !templateIds.Contains(view.FileId.Value))
                    throw new NotSupportedException("A public view page has no verified metadata-only setup file. Custom stored view pages are not supported by this package profile.");
            return result;
        }
        private static List<Node> ReadParentFolders(Guid rootId,Dictionary<Guid,Node> documents,IEnumerable<Node> entries)
        {
            var result=new Dictionary<Guid,Node>();
            foreach(Node entry in entries)
            {
                if(entry.Kind==NodeKind.Folder && entry.Id!=rootId)result[entry.Id]=entry;
                Guid parent=entry.ParentId;var visited=new HashSet<Guid>();
                while(parent!=rootId)
                {
                    Node folder;if(!visited.Add(parent))throw new InvalidDataException("A source item/template has cyclic folder ancestry.");
                    if(!documents.TryGetValue(parent,out folder) || folder.Kind!=NodeKind.Folder)
                        throw new NotSupportedException("A source item/template has no recorded parent folder within the selected list.");
                    result[parent]=folder;parent=folder.ParentId;
                }
            }
            return result.Values.OrderBy(folder=>folder.Path.Count(character=>character=='/')).ThenBy(folder=>folder.Path,StringComparer.Ordinal).ThenBy(folder=>folder.Id).ToList();
        }
        internal static void ValidateFolderGraph(Guid rootId,IEnumerable<Node> folders,IEnumerable<MigrationItemSnapshot> items)
        {
            var byId=folders.ToDictionary(folder=>folder.Id);
            foreach(MigrationItemSnapshot item in items)
            {
                Guid parent=item.ParentFolderId;var visited=new HashSet<Guid>();
                while(parent!=rootId)
                {
                    Node folder;if(!visited.Add(parent))throw new InvalidDataException("The migration list contains a cyclic folder ancestry.");
                    if(!byId.TryGetValue(parent,out folder))throw new NotSupportedException("A current item has no recorded parent-folder metadata within the selected list.");
                    parent=folder.ParentId;
                }
            }
        }
        private void ValidateLibraryFiles(MigrationListMetadata metadata,List<Node> files)
        {
            var primary=new HashSet<Guid>(files.Select(file=>file.Id));
            foreach(Node file in repository.EnumerateCurrentLibraryFiles(metadata.List))
                if(file.HasStream!=false && !primary.Contains(file.Id))throw new NotSupportedException("A current library file has no authoritative item metadata: "+file.Path+".");
        }
        private static List<MigrationUserSnapshot> ReadUsers(SqlConnection connection,MigrationListMetadata metadata,List<MigrationFieldDefinition> fields,
            List<MigrationItemSnapshot> items,Dictionary<string,Dictionary<string,string>> schema,List<MigrationTemplateFile> templates)
        {
            var ids=new HashSet<int>();if(metadata.AuthorId>0)ids.Add(metadata.AuthorId);
            foreach(string setupUser in templates.Select(template=>template.SetupPathUser).Where(value=>!String.IsNullOrWhiteSpace(value)).Distinct(StringComparer.Ordinal))
            {
                int userId;
                if(Int32.TryParse(setupUser,NumberStyles.None,CultureInfo.InvariantCulture,out userId)) {if(userId>0)ids.Add(userId);continue;}
                Require(schema,"UserInfo","tp_SiteID","tp_ID","tp_Login");
                using(SqlCommand command=Command(connection,"SELECT tp_ID FROM dbo.UserInfo WHERE tp_SiteID=@SiteId AND tp_Login=@Login;"))
                {
                    GuidParameter(command,"@SiteId",metadata.SiteId);command.Parameters.Add("@Login",SqlDbType.NVarChar,setupUser.Length).Value=setupUser;
                    using(SqlDataReader reader=command.ExecuteReader())
                    {
                        if(!reader.Read())throw new NotSupportedException("A template setup user has no recorded source login identity.");
                        ids.Add(Int(reader,0));if(reader.Read())throw new InvalidDataException("A template setup user matches ambiguous source identities.");
                    }
                }
            }
            var userNames=new HashSet<string>(fields.Where(field=>field.SharePointType=="User" || field.SharePointType=="UserMulti").Select(field=>field.Name),StringComparer.Ordinal);
            foreach(MigrationItemSnapshot item in items)
            {
                if(item.AuthorId>0)ids.Add(item.AuthorId);if(item.EditorId>0)ids.Add(item.EditorId);
                foreach(MigrationFieldValue value in item.Values.Where(value=>userNames.Contains(value.Name) && !value.IsNull))
                {
                    if(!(value.Value is int))throw new NotSupportedException("A user field uses an unsupported multi-user storage representation.");
                    int id=(int)value.Value;if(id>0)ids.Add(id);
                }
            }
            var result=new List<MigrationUserSnapshot>();if(ids.Count==0)return result;
            Require(schema,"UserInfo","tp_SiteID","tp_ID","tp_Login","tp_Title","tp_Email","tp_DomainGroup","tp_SiteAdmin","tp_Deleted","tp_SystemID");
            foreach(int[] page in ids.OrderBy(id=>id).Select((id,index)=>new{id,index}).GroupBy(pair=>pair.index/200).Select(group=>group.Select(pair=>pair.id).ToArray()))
            {
                string[] parameters=page.Select((id,index)=>"@User"+index).ToArray();
                using(SqlCommand command=Command(connection,@"SELECT tp_ID,tp_Login,tp_Title,tp_Email,tp_DomainGroup,tp_SiteAdmin,tp_Deleted,tp_SystemID FROM dbo.UserInfo
WHERE tp_SiteID=@SiteId AND tp_ID IN("+String.Join(",",parameters)+") ORDER BY tp_ID;"))
                {
                    GuidParameter(command,"@SiteId",metadata.SiteId);for(int index=0;index<page.Length;index++)command.Parameters.Add(parameters[index],SqlDbType.Int).Value=page[index];
                    using(SqlDataReader reader=command.ExecuteReader())while(reader.Read())
                    {
                        string login=Text(reader,1);if(String.IsNullOrWhiteSpace(login))throw new NotSupportedException("A referenced source user has no recorded login identity.");
                        result.Add(new MigrationUserSnapshot(Int(reader,0),login,Text(reader,2),Text(reader,3),Bool(reader,4),Bool(reader,5),Int(reader,6),reader.IsDBNull(7)?null:(byte[])reader.GetValue(7)));
                    }
                }
            }
            if(result.Count!=ids.Count)throw new NotSupportedException("A referenced author, editor or user field has no source user metadata.");return result;
        }
        private static List<MigrationViewSnapshot> ReadViews(SqlConnection connection,MigrationListMetadata metadata,Dictionary<string,Dictionary<string,string>> schema)
        {
            var result=new List<MigrationViewSnapshot>();if(!schema.ContainsKey("AllWebParts"))return result;
            Require(schema,"AllWebParts","tp_SiteId","tp_ListId","tp_ID","tp_View","tp_UserID","tp_Level","tp_Version","tp_Deleted","tp_IsCurrentVersion",
                "tp_Type","tp_Flags","tp_PageUrlID","tp_DisplayName","tp_BaseViewID","tp_IsIncluded","tp_PartOrder","tp_ZoneID","tp_FrameState","tp_WebPartTypeId","tp_ContentTypeId");
            using(SqlCommand command=Command(connection,@"
WITH V AS
(
 SELECT p.*,ROW_NUMBER() OVER(PARTITION BY p.tp_ID ORDER BY p.tp_Level DESC,p.tp_Version DESC) rn
 FROM dbo.AllWebParts p WHERE p.tp_SiteId=@SiteId AND p.tp_ListId=@ListId AND p.tp_UserID IS NULL
 AND p.tp_Deleted=0 AND p.tp_IsCurrentVersion=1 AND p.tp_Type IN(0,1)
)
SELECT p.tp_ID,p.tp_View,p.tp_Flags,p.tp_PageUrlID,p.tp_DisplayName,p.tp_BaseViewID,p.tp_Type,
 p.tp_IsIncluded,p.tp_PartOrder,p.tp_ZoneID,p.tp_FrameState,p.tp_WebPartTypeId,p.tp_ContentTypeId,d.DirName,d.LeafName,p.tp_Level,p.tp_Version
FROM V p OUTER APPLY
(
 SELECT TOP(1) a.DirName,a.LeafName FROM dbo.AllDocs a WHERE a.SiteId=p.tp_SiteId AND a.WebId=@WebId AND a.ListId=@ListId
 AND a.Id=p.tp_PageUrlID AND a.Type=0 AND a.IsCurrentVersion=1 AND a.DeleteTransactionId=0x ORDER BY a.Level DESC,a.InternalVersion DESC
) d WHERE p.rn=1 ORDER BY p.tp_ID;"))
            {
                Scope(command,metadata.List);
                using(SqlDataReader reader=command.ExecuteReader())while(reader.Read())
                {
                    if(reader.IsDBNull(3) || reader.IsDBNull(13) || reader.IsDBNull(14))throw new NotSupportedException("An active public view has no recorded current page identity or URL.");
                    string fragment=reader.IsDBNull(1)?String.Empty:MigrationFieldSchema.DecodeView((byte[])reader.GetValue(1));
                    string contentType=reader.IsDBNull(12)?null:ContentType(reader.GetValue(12));if(contentType=="0x")contentType=null;
                    result.Add(new MigrationViewSnapshot(reader.GetGuid(0),fragment,unchecked((uint)Int(reader,2)),reader.GetGuid(3),PathOf(Text(reader,13),Text(reader,14)),
                        Text(reader,4),reader.IsDBNull(5)?(int?)null:Int(reader,5),reader.GetByte(6),Bool(reader,7),reader.IsDBNull(8)?(int?)null:Int(reader,8),
                        Text(reader,9),reader.IsDBNull(10)?(int?)null:Int(reader,10),reader.IsDBNull(11)?(Guid?)null:reader.GetGuid(11),contentType,reader.IsDBNull(15)?(byte?)null:reader.GetByte(15),reader.IsDBNull(16)?(int?)null:Int(reader,16)));
                }
            }
            return result;
        }
        private static string ContentType(object value)
        {
            if(value==null || value==DBNull.Value)throw new NotSupportedException("A current item has no recorded content-type identity.");
            byte[] bytes=value as byte[];if(bytes!=null)return "0x"+BitConverter.ToString(bytes).Replace("-","");
            string text=value as string;if(text==null || !Regex.IsMatch(text,@"\A0x[0-9a-fA-F]+\z"))throw new NotSupportedException("The content-type storage representation is not supported.");return text;
        }
        private static SqlCommand Command(SqlConnection connection,string sql) {var command=connection.CreateCommand();command.CommandText=sql;command.CommandTimeout=120;return command;}
        private static void Scope(SqlCommand command,Node node) {GuidParameter(command,"@SiteId",node.SiteId);GuidParameter(command,"@WebId",node.WebId);GuidParameter(command,"@ListId",node.ListId);}
        private static void GuidParameter(SqlCommand command,string name,Guid value) {command.Parameters.Add(name,SqlDbType.UniqueIdentifier).Value=value;}
        private static string Text(SqlDataReader reader,int ordinal) {return reader.IsDBNull(ordinal)?String.Empty:reader.GetString(ordinal);}
        internal static int DocumentInternalVersion(object stored,int documentType)
        {
            if(stored==null || stored==DBNull.Value)
            {
                // Metadata-only folders have no recoverable document stream version.
                // Node uses the existing zero sentinel; the root's raw nullable value is kept in metadata.
                if(documentType==1)return 0;
                throw new NotSupportedException("A file has no recorded internal document version.");
            }
            return Convert.ToInt32(stored,CultureInfo.InvariantCulture);
        }
        private static int? NullableInt(SqlDataReader reader,int ordinal) {return reader.IsDBNull(ordinal)?(int?)null:Int(reader,ordinal);}
        private static int Int(SqlDataReader reader,int ordinal) {if(reader.IsDBNull(ordinal))throw new NotSupportedException("Required integer migration metadata is absent: "+reader.GetName(ordinal)+".");return Convert.ToInt32(reader.GetValue(ordinal),CultureInfo.InvariantCulture);}
        private static long Long(SqlDataReader reader,int ordinal) {if(reader.IsDBNull(ordinal))throw new NotSupportedException("Required numeric migration metadata is absent: "+reader.GetName(ordinal)+".");return Convert.ToInt64(reader.GetValue(ordinal),CultureInfo.InvariantCulture);}
        private static bool Bool(SqlDataReader reader,int ordinal) {if(reader.IsDBNull(ordinal))throw new NotSupportedException("Required boolean migration metadata is absent: "+reader.GetName(ordinal)+".");return Convert.ToBoolean(reader.GetValue(ordinal),CultureInfo.InvariantCulture);}
        private static DateTime Date(SqlDataReader reader,int ordinal) {if(reader.IsDBNull(ordinal))throw new NotSupportedException("Required timestamp migration metadata is absent: "+reader.GetName(ordinal)+".");return reader.GetDateTime(ordinal);}
        private static string PathOf(string directory,string leaf) {return String.IsNullOrEmpty(directory)?leaf:String.IsNullOrEmpty(leaf)?directory:directory.TrimEnd('/')+"/"+leaf;}
    }
    internal static class MigrationFieldSchema
    {
        private const int MaximumXmlCharacters=16*1024*1024;
        internal static string Decode(byte[] stored,string rootName)
        {
            string text=new UTF8Encoding(false,true).GetString(ListFieldSchema.Decompress(stored));
            int start=text.IndexOf('<');if(start<0)throw new InvalidDataException("Stored migration metadata has no XML definitions.");
            if(start!=0 && !Regex.IsMatch(text.Substring(0,start),@"\A[0-9]+(?:\.[0-9]+)*\z"))throw new NotSupportedException("The stored migration metadata version prefix is unsupported.");
            string xml=text.Substring(start);LoadFragments(xml,rootName);return xml;
        }
        internal static List<MigrationFieldDefinition> Read(string xml,IDictionary<string,string> columns,Guid listId,Guid webId,string contentTypesXml=null)
        {
            Dictionary<string,Guid> contentTypeIds=ContentTypeFieldIds(contentTypesXml);
            XmlDocument document=LoadFragments(xml,"Fields");var result=new List<MigrationFieldDefinition>();var names=new Dictionary<string,string>(StringComparer.Ordinal);
            IEnumerable<XmlElement> elements=document.DocumentElement.ChildNodes.OfType<XmlElement>();
            if(elements.Count()==1 && elements.First().LocalName=="Fields")elements=elements.First().ChildNodes.OfType<XmlElement>();
            foreach(XmlElement element in elements)
            {
                if(element.LocalName!="Field" && element.LocalName!="FieldRef")throw new NotSupportedException("The stored list schema contains an unsupported field definition element.");
                string name=element.GetAttribute("Name");if(String.IsNullOrWhiteSpace(name))throw new InvalidDataException("A field definition has no internal name.");
                string existing;if(names.TryGetValue(name,out existing)) {if(existing!=element.OuterXml)throw new InvalidDataException("The list contains conflicting field definitions: "+name+".");continue;}
                names.Add(name,element.OuterXml);Guid? id=OptionalGuid(element.GetAttribute("ID"),"field ID");Guid referencedId;
                if(contentTypeIds.TryGetValue(name,out referencedId)) {if(id.HasValue && id.Value!=referencedId)throw new InvalidDataException("Field definitions disagree with content-type field identities.");id=referencedId;}
                var mappings=new List<MigrationStorageMapping>();
                foreach(XmlAttribute attribute in element.Attributes)
                {
                    Match match=Regex.Match(attribute.Name,@"\AColName(?<component>[2-9]|[1-3][0-9])?\z");
                    if(!match.Success) {if(attribute.Name.StartsWith("ColName",StringComparison.Ordinal))throw new NotSupportedException("A field declares an unsupported storage component.");continue;}
                    int component=match.Groups["component"].Success?Int32.Parse(match.Groups["component"].Value,CultureInfo.InvariantCulture):1;
                    string column=attribute.Value;
                    if(!Regex.IsMatch(column,@"\A(?:nvarchar|ntext|varchar|text|bit|int|bigint|smallint|tinyint|float|real|decimal|numeric|money|datetime|uniqueidentifier|varbinary|binary)[1-9][0-9]{0,3}\z"))
                        throw new NotSupportedException("Field "+name+" declares an unsupported storage column.");
                    string sqlType;if(!columns.TryGetValue(column,out sqlType))throw new InvalidDataException("Field "+name+" references a storage column absent from AllUserData.");
                    byte ordinal=0;string ordinalName=component==1?"RowOrdinal":"RowOrdinal"+component;string rawOrdinal=element.GetAttribute(ordinalName);
                    if(rawOrdinal.Length!=0 && !Byte.TryParse(rawOrdinal,NumberStyles.None,CultureInfo.InvariantCulture,out ordinal))throw new InvalidDataException("Field "+name+" declares an invalid row ordinal.");
                    mappings.Add(new MigrationStorageMapping(column,ordinal,component,sqlType));
                }
                mappings.Sort((first,second)=>first.Component.CompareTo(second.Component));
                string type=element.GetAttribute("Type");Guid? lookupList=null,lookupWeb=null;
                if(mappings.Count==0 && type!="Computed" && !KnownUnmappedField(name))
                    throw new NotSupportedException("Field "+name+" has no verified storage mapping or supported virtual-field semantics.");
                if(type=="Lookup" || type=="LookupMulti")
                {
                    string target=element.GetAttribute("List");if(target=="Self")lookupList=listId;else lookupList=OptionalGuid(target,"lookup-list ID");
                    lookupWeb=OptionalGuid(element.GetAttribute("WebId"),"lookup-web ID");
                    if(!lookupList.HasValue || lookupList.Value!=listId || (lookupWeb.HasValue && lookupWeb.Value!=webId))
                        throw new NotSupportedException("Field "+name+" depends on an external lookup list or web that this single-list package cannot include.");
                }
                if(type=="LookupMulti" || type=="UserMulti" || type=="TaxonomyFieldType" || type=="TaxonomyFieldTypeMulti" || type=="BusinessData")
                    throw new NotSupportedException("Field "+name+" requires an unsupported relational or external field-value representation.");
                result.Add(new MigrationFieldDefinition(id,name,type,element.LocalName=="FieldRef",element.OuterXml,mappings,lookupList,lookupWeb));
            }
            return result;
        }
        private static bool KnownUnmappedField(string name)
        {
            switch(name)
            {
                case "ID":case "GUID":case "Author":case "Editor":case "Created":case "Modified":case "Attachments":
                case "ContentTypeId":case "_ModerationStatus":case "_UIVersion":case "_UIVersionString":case "FileLeafRef":
                case "FileRef":case "FSObjType":case "UniqueId":return true;
                default:return false;
            }
        }
        private static Dictionary<string,Guid> ContentTypeFieldIds(string xml)
        {
            var result=new Dictionary<string,Guid>(StringComparer.Ordinal);if(String.IsNullOrWhiteSpace(xml))return result;
            XmlDocument document=LoadFragments(xml,"ContentTypes");
            foreach(XmlElement field in document.SelectNodes("//*[local-name()='FieldRef']").OfType<XmlElement>())
            {
                string name=field.GetAttribute("Name"),raw=field.GetAttribute("ID");if(name.Length==0 || raw.Length==0)continue;
                Guid? id=OptionalGuid(raw,"content-type field ID");Guid previous;
                if(result.TryGetValue(name,out previous) && previous!=id.Value)throw new InvalidDataException("Content types reference conflicting field identities.");
                result[name]=id.Value;
            }
            return result;
        }
        internal static string DecodeView(byte[] bytes)
        {
            if(bytes==null || bytes.Length==0)throw new InvalidDataException("An active source view has no definition bytes.");
            string text;
            if(bytes.Length>=2 && bytes[0]==0xA8 && bytes[1]==0xA9)text=Decode(bytes,"Views");
            else
            {
                try {text=bytes.Length>=2 && bytes[1]==0?new UnicodeEncoding(false,false,true).GetString(bytes):new UTF8Encoding(false,true).GetString(bytes);}
                catch(DecoderFallbackException error) {throw new NotSupportedException("An active view uses an unsupported binary definition format.",error);}
                text=text.TrimStart('\uFEFF');
            }
            XmlDocument document=LoadFragments(text,"Views");
            var children=new HashSet<string>(StringComparer.Ordinal) {"View","PagedRowset","Toolbar","Query","ViewFields","GroupByHeader","GroupByFooter","ViewHeader","ViewBody","ViewFooter",
                "RowLimitExceeded","ViewEmpty","PagedRecurrenceRowset","PagedClientCallbackRowset","Aggregations","OpenApplicationExtension","RowLimit","Mobile","ViewStyle",
                "CalendarSettings","CalendarViewStyles","ViewBidiHeader","Script","ViewData","Formats","InlineEdit","ProjectedFields","Joins","Method","ParameterBindings","Xsl","XslLink","JS","JSLink"};
            if(document.DocumentElement.ChildNodes.OfType<XmlElement>().Any(element=>!children.Contains(element.LocalName)))throw new NotSupportedException("An active view uses an unsupported definition envelope.");
            return text;
        }
        internal static XmlDocument LoadFragments(string xml,string rootName)
        {
            if(xml==null)throw new ArgumentNullException("xml");
            if(xml.Length>MaximumXmlCharacters)throw new NotSupportedException("Migration metadata exceeds the supported 16 MB XML limit.");
            var settings=new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=MaximumXmlCharacters,IgnoreComments=true};
            var document=new XmlDocument {XmlResolver=null};
            using(XmlReader reader=XmlReader.Create(new StringReader("<"+rootName+">"+xml+"</"+rootName+">"),settings))document.Load(reader);
            return document;
        }
        private static Guid? OptionalGuid(string text,string label)
        {if(String.IsNullOrEmpty(text))return null;Guid value;if(!Guid.TryParse(text,out value))throw new NotSupportedException("The stored "+label+" is not a supported GUID reference.");return value;}
    }
}
