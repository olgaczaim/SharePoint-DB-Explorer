using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace SharePointExplorer
{
    public interface ISharePointMigrationCatalog
    {
        MigrationListSnapshot ReadMigrationList(Node listOrLibrary);
    }

    public sealed class MigrationListMetadata
    {
        private readonly Node list;
        public Node List { get { return MigrationSnapshotCopy.Node(list); } }
        public Guid SiteId { get { return list.SiteId; } }
        public Guid WebId { get { return list.WebId; } }
        public Guid ListId { get { return list.ListId; } }
        public Guid RootFolderId { get { return list.Id; } }
        public string RootFolderUrl { get { return list.Path; } }
        public string Title { get; private set; }
        public string Description { get; private set; }
        public int BaseType { get; private set; }
        public int ServerTemplate { get; private set; }
        public Guid TemplateFeatureId { get; private set; }
        public Guid? TemplateId { get; private set; }
        public long Flags { get; private set; }
        public long Flags2 { get; private set; }
        public int ReadSecurity { get; private set; }
        public int WriteSecurity { get; private set; }
        public int Direction { get; private set; }
        public DateTime Created { get; private set; }
        public int AuthorId { get; private set; }
        public int? MaxMajorVersions { get; private set; }
        public int? MaxMajorWithMinorVersions { get; private set; }
        // Stored values can be relative paths. They are not inferred absolute URLs.
        public string StoredSiteUrl { get; private set; }
        public string StoredWebUrl { get; private set; }
        public string WebTitle { get; private set; }
        public string FieldSchemaXml { get; private set; }
        public string ContentTypesXml { get; private set; }
        public string SourceVersion { get; private set; }
        public string ImageUrl { get; private set; }
        public int? Version { get; private set; }
        public string PlatformVersion { get; private set; }
        public int? ProductVersion { get; private set; }
        public int? Language { get; private set; }
        public int? RootInternalVersion { get; private set; }
        public MigrationListMetadata(Node list,string title,string description,int baseType,int serverTemplate,
            Guid featureId,Guid? templateId,long flags,long flags2,int readSecurity,int writeSecurity,int direction,
            DateTime created,int authorId,int? maxMajorVersions,int? maxMajorWithMinorVersions,string storedSiteUrl,
            string storedWebUrl,string webTitle,string fieldSchemaXml,string contentTypesXml,string sourceVersion,
            string imageUrl=null,int? version=null,string platformVersion=null,int? productVersion=null,int? language=null,int? rootInternalVersion=null)
        {
            if(list==null) throw new ArgumentNullException("list");
            this.list=MigrationSnapshotCopy.Node(list);Title=title;Description=description;BaseType=baseType;ServerTemplate=serverTemplate;
            TemplateFeatureId=featureId;TemplateId=templateId;Flags=flags;Flags2=flags2;ReadSecurity=readSecurity;WriteSecurity=writeSecurity;
            Direction=direction;Created=created;AuthorId=authorId;MaxMajorVersions=maxMajorVersions;MaxMajorWithMinorVersions=maxMajorWithMinorVersions;
            StoredSiteUrl=storedSiteUrl;StoredWebUrl=storedWebUrl;WebTitle=webTitle;FieldSchemaXml=fieldSchemaXml;ContentTypesXml=contentTypesXml;SourceVersion=sourceVersion;
            ImageUrl=imageUrl;Version=version;PlatformVersion=platformVersion;ProductVersion=productVersion;Language=language;RootInternalVersion=rootInternalVersion;
        }
    }

    public sealed class MigrationStorageMapping
    {
        public string Column { get; private set; }
        public byte RowOrdinal { get; private set; }
        public int Component { get; private set; }
        public string SqlType { get; private set; }
        public MigrationStorageMapping(string column,byte rowOrdinal,int component,string sqlType)
        {Column=column;RowOrdinal=rowOrdinal;Component=component;SqlType=sqlType;}
    }
    public sealed class MigrationFieldDefinition
    {
        public Guid? Id { get; private set; }
        public string Name { get; private set; }
        public string SharePointType { get; private set; }
        public bool IsReference { get; private set; }
        public string SchemaXml { get; private set; }
        public Guid? LookupListId { get; private set; }
        public Guid? LookupWebId { get; private set; }
        public ReadOnlyCollection<MigrationStorageMapping> StorageMappings { get; private set; }
        public MigrationFieldDefinition(Guid? id,string name,string sharePointType,bool isReference,string schemaXml,
            IEnumerable<MigrationStorageMapping> storageMappings,Guid? lookupListId,Guid? lookupWebId)
        {
            Id=id;Name=name;SharePointType=sharePointType;IsReference=isReference;SchemaXml=schemaXml;
            StorageMappings=new List<MigrationStorageMapping>(storageMappings??Enumerable.Empty<MigrationStorageMapping>()).AsReadOnly();
            LookupListId=lookupListId;LookupWebId=lookupWebId;
        }
    }
    public sealed class MigrationFieldValue
    {
        private readonly object value;
        public string Name { get; private set; }
        public Guid? FieldId { get; private set; }
        public int Component { get; private set; }
        public string SqlType { get; private set; }
        public object Value { get { return MigrationSnapshotCopy.Value(value); } }
        public bool IsNull { get { return value==null; } }
        public MigrationFieldValue(string name,Guid? fieldId,int component,string sqlType,object value)
        {Name=name;FieldId=fieldId;Component=component;SqlType=sqlType;this.value=MigrationSnapshotCopy.Value(value);}
    }
    public sealed class MigrationItemSnapshot
    {
        private readonly Node document;
        public Node Document { get { return MigrationSnapshotCopy.Node(document); } }
        public int ItemId { get; private set; }
        public Guid ItemUniqueId { get; private set; }
        public Guid DocumentId { get { return document.Id; } }
        public Guid ParentFolderId { get { return document.ParentId; } }
        public bool IsFolder { get { return document.Kind==NodeKind.Folder; } }
        public int AuthorId { get; private set; }
        public int EditorId { get; private set; }
        public DateTime Created { get; private set; }
        public DateTime Modified { get; private set; }
        public bool HasAttachments { get; private set; }
        public string ContentTypeId { get; private set; }
        public ReadOnlyCollection<MigrationFieldValue> Values { get; private set; }
        public MigrationItemSnapshot(Node document,int itemId,Guid itemUniqueId,int authorId,int editorId,
            DateTime created,DateTime modified,bool hasAttachments,string contentTypeId,IEnumerable<MigrationFieldValue> values)
        {
            if(document==null) throw new ArgumentNullException("document");
            this.document=MigrationSnapshotCopy.Node(document);ItemId=itemId;ItemUniqueId=itemUniqueId;AuthorId=authorId;EditorId=editorId;
            Created=created;Modified=modified;HasAttachments=hasAttachments;ContentTypeId=contentTypeId;
            Values=new List<MigrationFieldValue>(values??Enumerable.Empty<MigrationFieldValue>()).AsReadOnly();
        }
    }
    public sealed class MigrationUserSnapshot
    {
        private readonly byte[] systemId;
        public int Id { get; private set; }
        public string Login { get; private set; }
        public string Name { get; private set; }
        public string Email { get; private set; }
        public bool IsDomainGroup { get; private set; }
        public bool IsSiteAdmin { get; private set; }
        public int Deleted { get; private set; }
        public byte[] SystemId { get { return systemId==null?null:(byte[])systemId.Clone(); } }
        public MigrationUserSnapshot(int id,string login,string name,string email,bool isDomainGroup,bool isSiteAdmin,int deleted,byte[] systemId)
        {Id=id;Login=login;Name=name;Email=email;IsDomainGroup=isDomainGroup;IsSiteAdmin=isSiteAdmin;Deleted=deleted;this.systemId=systemId==null?null:(byte[])systemId.Clone();}
    }
    public sealed class MigrationViewSnapshot
    {
        public Guid Id { get; private set; }
        // Stored view XML is often a child-element fragment, not a full View.
        public string SchemaXml { get; private set; }
        public uint? Flags { get; private set; }
        public Guid? FileId { get; private set; }
        public string Url { get; private set; }
        public string DisplayName { get; private set; }
        public int? BaseViewId { get; private set; }
        public byte? PageType { get; private set; }
        public bool? IsIncluded { get; private set; }
        public int? PartOrder { get; private set; }
        public string ZoneId { get; private set; }
        public int? FrameState { get; private set; }
        public Guid? WebPartTypeId { get; private set; }
        public string ContentTypeId { get; private set; }
        public byte? Level { get; private set; }
        public int? Version { get; private set; }
        public MigrationViewSnapshot(Guid id,string schemaXml,uint? flags=null,Guid? fileId=null,string url=null,
            string displayName=null,int? baseViewId=null,byte? pageType=null,bool? isIncluded=null,int? partOrder=null,
            string zoneId=null,int? frameState=null,Guid? webPartTypeId=null,string contentTypeId=null,byte? level=null,int? version=null)
        {
            Id=id;SchemaXml=schemaXml;Flags=flags;FileId=fileId;Url=url;DisplayName=displayName;BaseViewId=baseViewId;
            PageType=pageType;IsIncluded=isIncluded;PartOrder=partOrder;ZoneId=zoneId;FrameState=frameState;
            WebPartTypeId=webPartTypeId;ContentTypeId=contentTypeId;Level=level;Version=version;
        }
    }
    public sealed class MigrationTemplateFile
    {
        private readonly Node document;
        public Node Document { get { return MigrationSnapshotCopy.Node(document); } }
        public string SetupPath { get; private set; }
        public byte? SetupPathVersion { get; private set; }
        public string SetupPathUser { get; private set; }
        public MigrationTemplateFile(Node document,string setupPath,byte? setupPathVersion,string setupPathUser)
        {if(document==null)throw new ArgumentNullException("document");this.document=MigrationSnapshotCopy.Node(document);SetupPath=setupPath;SetupPathVersion=setupPathVersion;SetupPathUser=setupPathUser;}
    }
    public sealed class MigrationSystemObject
    {
        public Guid Id { get; private set; }
        public string Type { get; private set; }
        public string Url { get; private set; }
        public string Role { get; private set; }
        public MigrationSystemObject(Guid id,string type,string url,string role)
        {Id=id;Type=type;Url=url;Role=role;}
    }
    public sealed class MigrationListSnapshot
    {
        private readonly List<Node> folders,files;
        public MigrationListMetadata Metadata { get; private set; }
        public Node List { get { return Metadata.List; } }
        public ReadOnlyCollection<MigrationFieldDefinition> Fields { get; private set; }
        public ReadOnlyCollection<MigrationItemSnapshot> Items { get; private set; }
        public ReadOnlyCollection<Node> Folders { get { return folders.Select(MigrationSnapshotCopy.Node).ToList().AsReadOnly(); } }
        public ReadOnlyCollection<Node> Files { get { return files.Select(MigrationSnapshotCopy.Node).ToList().AsReadOnly(); } }
        public ReadOnlyCollection<MigrationUserSnapshot> Users { get; private set; }
        public ReadOnlyCollection<MigrationViewSnapshot> Views { get; private set; }
        public ReadOnlyCollection<MigrationTemplateFile> TemplateFiles { get; private set; }
        public ReadOnlyCollection<MigrationSystemObject> SystemObjects { get; private set; }
        public MigrationListSnapshot(MigrationListMetadata metadata,IEnumerable<MigrationFieldDefinition> fields,
            IEnumerable<MigrationItemSnapshot> items,IEnumerable<Node> folders,IEnumerable<Node> files,
            IEnumerable<MigrationUserSnapshot> users,IEnumerable<MigrationViewSnapshot> views,IEnumerable<MigrationTemplateFile> templateFiles=null,IEnumerable<MigrationSystemObject> systemObjects=null)
        {
            if(metadata==null) throw new ArgumentNullException("metadata");
            Metadata=metadata;Fields=new List<MigrationFieldDefinition>(fields??Enumerable.Empty<MigrationFieldDefinition>()).AsReadOnly();
            Items=new List<MigrationItemSnapshot>(items??Enumerable.Empty<MigrationItemSnapshot>()).AsReadOnly();
            this.folders=(folders??Enumerable.Empty<Node>()).Select(MigrationSnapshotCopy.Node).ToList();
            this.files=(files??Enumerable.Empty<Node>()).Select(MigrationSnapshotCopy.Node).ToList();
            Users=new List<MigrationUserSnapshot>(users??Enumerable.Empty<MigrationUserSnapshot>()).AsReadOnly();
            Views=new List<MigrationViewSnapshot>(views??Enumerable.Empty<MigrationViewSnapshot>()).AsReadOnly();
            TemplateFiles=new List<MigrationTemplateFile>(templateFiles??Enumerable.Empty<MigrationTemplateFile>()).AsReadOnly();
            SystemObjects=new List<MigrationSystemObject>(systemObjects??Enumerable.Empty<MigrationSystemObject>()).AsReadOnly();
        }
    }
    internal static class MigrationSnapshotCopy
    {
        internal static object Value(object value)
        {if(value==null || value==DBNull.Value)return null;byte[] bytes=value as byte[];return bytes==null?value:(byte[])bytes.Clone();}
        internal static Node Node(Node source)
        {
            return new Node {Kind=source.Kind,SiteId=source.SiteId,WebId=source.WebId,ListId=source.ListId,Id=source.Id,Name=source.Name,Path=source.Path,
                Size=source.Size,StreamSchema=source.StreamSchema,Level=source.Level,InternalVersion=source.InternalVersion,HistoryVersion=source.HistoryVersion,
                UiVersion=source.UiVersion,ParentId=source.ParentId,HasStream=source.HasStream,Modified=source.Modified,ListBaseType=source.ListBaseType,
                ListItemId=source.ListItemId,ItemUniqueId=source.ItemUniqueId,Title=source.Title,Created=source.Created,HasAttachments=source.HasAttachments,AttachmentOwnerId=source.AttachmentOwnerId,DeletionTransactionId=source.DeletionTransactionId,DeletedAt=source.DeletedAt};
        }
    }
}