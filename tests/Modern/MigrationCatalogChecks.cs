using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

namespace SharePointExplorer.Tests
{
    internal static class MigrationCatalogChecks
    {
        internal static void Run()
        {
            Guid listId=Guid.NewGuid(),webId=Guid.NewGuid(),titleId=Guid.NewGuid(),urlId=Guid.NewGuid();
            var columns=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase) {
                {"nvarchar8","nvarchar"},{"nvarchar9","nvarchar"},{"nvarchar10","nvarchar"},{"int1","int"},{"bit1","bit"} };
            string fields="<FieldRef Name='Title' ColName='nvarchar8' RowOrdinal='2'/><Field ID='"+urlId+"' Name='SourceLink' Type='URL' ColName='nvarchar9' RowOrdinal='1' ColName2='nvarchar10' RowOrdinal2='3'/>";
            string contentTypes="<ContentType ID='0x01'><FieldRefs><FieldRef Name='Title' ID='"+titleId+"'/></FieldRefs></ContentType>";
            List<MigrationFieldDefinition> parsed=MigrationFieldSchema.Read(fields,columns,listId,webId,contentTypes);
            Check(parsed.Count==2 && parsed[0].Name=="Title" && parsed[0].IsReference && parsed[0].Id==titleId && parsed[0].SharePointType==String.Empty &&
                parsed[0].StorageMappings.Single().Column=="nvarchar8" && parsed[0].StorageMappings.Single().RowOrdinal==2 && parsed[0].StorageMappings.Single().SqlType=="nvarchar",
                "Migration fields guessed a title column/type or failed to bind the source content-type field ID.");
            Check(parsed[1].Id==urlId && parsed[1].SharePointType=="URL" && parsed[1].StorageMappings.Select(mapping=>mapping.Component).SequenceEqual(new[]{1,2}) &&
                parsed[1].StorageMappings.Select(mapping=>(int)mapping.RowOrdinal).SequenceEqual(new[]{1,3}),
                "Migration URL metadata lost its independently declared second column/row ordinal.");
            Check(parsed[0].SchemaXml.Contains("FieldRef") && !parsed[0].SchemaXml.Contains("ID="),"Resolving a source field identity rewrote the raw FieldRef definition.");
            string wrapped="<Fields>"+fields+"</Fields>";
            Check(MigrationFieldSchema.Read(wrapped,columns,listId,webId,contentTypes).Count==2,"A stored Fields wrapper changed field enumeration.");
            Reject(()=>MigrationFieldSchema.Read("<FieldRef Name='Title' ColName='nvarchar8];SELECT 1--'/>",columns,listId,webId),"SQL identifier injection");
            Reject(()=>MigrationFieldSchema.Read("<FieldRef Name='Title' ColName='nvarchar7'/>",columns,listId,webId),"missing source storage column");
            Reject(()=>MigrationFieldSchema.Read("<FieldRef Name='Title' ColName='nvarchar8' RowOrdinal='256'/>",columns,listId,webId),"out-of-range ordinal");
            Reject(()=>MigrationFieldSchema.Read("<FieldRef Name='Title' ColName='nvarchar8'/><FieldRef Name='Title' ColName='nvarchar9'/>",columns,listId,webId),"conflicting field mappings");
            Reject(()=>MigrationFieldSchema.Read("<FieldRef Name='Title' ColName40='nvarchar8'/>",columns,listId,webId),"unknown storage component");
            Reject(()=>MigrationFieldSchema.Read("<Field Name='CustomText' Type='Text'/>",columns,listId,webId),"unresolved concrete stored field");
            Reject(()=>MigrationFieldSchema.Read("<FieldRef Name='UnknownInherited' ID='"+Guid.NewGuid()+"'/>",columns,listId,webId),"unresolved inherited custom field");
            Check(MigrationFieldSchema.Read("<FieldRef Name='ContentTypeId'/><Field Name='LinkTitle' Type='Computed'><FieldRefs><FieldRef Name='Title'/></FieldRefs></Field>",columns,listId,webId).Count==2,
                "Explicit system storage or computed virtual definitions were rejected.");
            string conflictingTypes="<ContentType><FieldRefs><FieldRef Name='Title' ID='"+titleId+"'/><FieldRef Name='Title' ID='"+Guid.NewGuid()+"'/></FieldRefs></ContentType>";
            Reject(()=>MigrationFieldSchema.Read(fields,columns,listId,webId,conflictingTypes),"conflicting source field IDs");
            Reject(()=>MigrationFieldSchema.Read("<!DOCTYPE Field [<!ENTITY entity SYSTEM 'file:///not-read'>]><Field Name='Title'/>",columns,listId,webId),"DTD/entity resolution");
            Reject(()=>MigrationFieldSchema.Read("<Field Name='Related' Type='Lookup' List='"+Guid.NewGuid()+"' ColName='int1'/>",columns,listId,webId),"unpackaged external lookup list");
            Check(MigrationFieldSchema.Read("<Field Name='Related' Type='Lookup' List='Self' ColName='int1'/>",columns,listId,webId).Single().LookupListId==listId,
                "An evidenced self lookup was treated as an external dependency.");
            Reject(()=>MigrationFieldSchema.Read("<Field Name='Users' Type='UserMulti' ColName='int1'/>",columns,listId,webId),"unimplemented user junction storage");
            byte[] stored=Compress("16.0.0.14326.0.0"+fields);
            Check(MigrationFieldSchema.Decode(stored,"Fields")==fields,"Migration field decompression modified the source XML.");
            Check(MigrationFieldSchema.Decode(Compress(contentTypes),"ContentTypes")==contentTypes,"Unprefixed source content-type XML was rejected or rewritten.");
            stored[stored.Length-1]^=1;Reject(()=>MigrationFieldSchema.Decode(stored,"Fields"),"compressed metadata checksum mismatch");
            string view="<View Name='"+Guid.NewGuid()+"' Flags='0' />";
            Check(MigrationFieldSchema.DecodeView(Encoding.UTF8.GetBytes(view))==view,"A plain UTF-8 view was not preserved.");
            Reject(()=>MigrationFieldSchema.DecodeView(Encoding.UTF8.GetBytes("<Unknown />")),"unknown view envelope");
            string fragment="<Toolbar Type='Standard'/><ViewFields><FieldRef Name='Title'/></ViewFields>";
            Check(MigrationFieldSchema.DecodeView(Compress(fragment))==fragment,"Actual compressed public-view child fragments were lost or rewritten.");
            Check(SqlMigrationCatalog.DocumentInternalVersion(DBNull.Value,1)==0 && SqlMigrationCatalog.DocumentInternalVersion(1025,0)==1025,
                "Nullable metadata-only folder counters changed an authoritative file version.");
            Reject(()=>SqlMigrationCatalog.DocumentInternalVersion(DBNull.Value,0),"absent file recovery identity");
            CheckSnapshotsAndFolders(listId,webId,parsed,fields,contentTypes);
            Console.WriteLine("PASS migration field mappings, source IDs, typed immutable values, bounded XML/decompression and folder ancestry");
        }
        private static void CheckSnapshotsAndFolders(Guid listId,Guid webId,List<MigrationFieldDefinition> fields,string raw,string contentTypes)
        {
            Guid root=Guid.NewGuid(),folderId=Guid.NewGuid();
            var list=new Node {Kind=NodeKind.List,SiteId=Guid.NewGuid(),WebId=webId,ListId=listId,Id=root,Path="Lists/Items",ParentId=webId};
            var metadata=new MigrationListMetadata(list,"Items","Source description",0,100,Guid.NewGuid(),null,0,0,1,1,0,
                new DateTime(2025,1,2),7,10,0,"","","Source web",raw,contentTypes,"16.0.14326.20450","/_layouts/images/itgen.png",3,"15.0.36.0",15,1033);
            var unversionedMetadata=new MigrationListMetadata(list,"Items","",0,100,Guid.NewGuid(),null,0,0,1,1,0,
                new DateTime(2025,1,2),7,null,null,"","","Source web",raw,contentTypes,"16.0.14326.20450",rootInternalVersion:null);
            Check(!unversionedMetadata.MaxMajorVersions.HasValue && !unversionedMetadata.MaxMajorWithMinorVersions.HasValue && !unversionedMetadata.RootInternalVersion.HasValue,
                "Absent list retention/folder counters were replaced with invented numeric values.");
            var folder=new Node {Kind=NodeKind.Folder,SiteId=list.SiteId,WebId=webId,ListId=listId,Id=folderId,ParentId=root,Name="Nested",Path="Lists/Items/Nested"};
            var document=new Node {Kind=NodeKind.ListItem,SiteId=list.SiteId,WebId=webId,ListId=listId,Id=Guid.NewGuid(),ParentId=folderId,Name="1_.000",Path="Lists/Items/Nested/1_.000"};
            byte[] bytes={1,2,3};var value=new MigrationFieldValue("Binary",null,1,"varbinary",bytes);bytes[0]=9;
            var values=new[]{value,new MigrationFieldValue("Enabled",null,1,"bit",true),new MigrationFieldValue("Missing",null,1,"nvarchar",null)};
            var item=new MigrationItemSnapshot(document,1,Guid.NewGuid(),7,8,new DateTime(2025,1,2),new DateTime(2025,2,3),false,"0x01",values);
            var folderItem=new MigrationItemSnapshot(folder,2,Guid.NewGuid(),7,8,new DateTime(2025,1,2),new DateTime(2025,2,3),false,"0x0120",Array.Empty<MigrationFieldValue>());
            byte[] sid={1,2,3};var user=new MigrationUserSnapshot(7,"example\\author","Source author","author@example.invalid",false,false,0,sid);sid[0]=9;
            var templateNode=new Node {Kind=NodeKind.File,SiteId=list.SiteId,WebId=webId,ListId=listId,Id=Guid.NewGuid(),ParentId=folderId,
                Name="AllItems.aspx",Path="Lists/Items/Nested/AllItems.aspx",HasStream=false,Level=1,InternalVersion=3,UiVersion=512};
            var template=new MigrationTemplateFile(templateNode,"pages\\viewpage.aspx",4,"example\\author");
            var sourceView=new MigrationViewSnapshot(Guid.NewGuid(),"<ViewFields><FieldRef Name='Title'/></ViewFields>",0x00800009,templateNode.Id,
                templateNode.Path,"All Items",1,0,true,2,"Main",0,Guid.NewGuid(),null,1,3);
            var snapshot=new MigrationListSnapshot(metadata,fields,new[]{folderItem,item},new[]{folder},Array.Empty<Node>(),new[]{user},new[]{sourceView},new[]{template});
            templateNode.Path="caller mutation";Node exposedTemplate=template.Document;exposedTemplate.ParentId=Guid.Empty;
            Check(snapshot.TemplateFiles.Count==1 && template.Document.Path=="Lists/Items/Nested/AllItems.aspx" && template.Document.ParentId==folderId &&
                template.Document.HasStream==false && template.SetupPath=="pages\\viewpage.aspx" && template.SetupPathVersion==4 && template.SetupPathUser=="example\\author" &&
                snapshot.Views[0].FileId==template.Document.Id && snapshot.Views[0].Flags==0x00800009 && snapshot.Views[0].Level==1 && snapshot.Views[0].Version==3,
                "Migration template/view snapshots lost exact setup metadata or exposed a mutable file identity.");
            list.Path="caller mutation";document.ParentId=Guid.Empty;folder.Name="caller mutation";
            Node exposed=snapshot.Folders[0];exposed.ParentId=Guid.Empty;byte[] exposedValue=(byte[])value.Value;exposedValue[0]=8;byte[] exposedSid=user.SystemId;exposedSid[0]=8;
            Check(snapshot.List.Path=="Lists/Items" && snapshot.List.ParentId==webId && snapshot.Folders[0].Name=="Nested" && snapshot.Folders[0].ParentId==root &&
                item.ParentFolderId==folderId && (byte[])value.Value is byte[] saved && saved[0]==1 && user.SystemId[0]==1 && values[1].Value is bool && values[2].IsNull,
                "Migration snapshot exposed mutable source identities or lost typed/null values.");
            Check(metadata.ImageUrl=="/_layouts/images/itgen.png" && metadata.Version==3 && metadata.PlatformVersion=="15.0.36.0" && metadata.ProductVersion==15 && metadata.Language==1033,
                "Migration metadata discarded evidenced source list/build properties.");
            SqlMigrationCatalog.ValidateFolderGraph(root,snapshot.Folders,snapshot.Items);
            Reject(()=>SqlMigrationCatalog.ValidateFolderGraph(root,Array.Empty<Node>(),new[]{item}),"missing parent-folder source metadata");
            var cycle=new Node {Kind=NodeKind.Folder,Id=folderId,ParentId=folderId};
            Reject(()=>SqlMigrationCatalog.ValidateFolderGraph(root,new[]{cycle},new[]{item}),"cyclic folder ancestry");
        }
        private static void Reject(Action action,string cause)
        {
            try{action();}catch(NotSupportedException){return;}catch(InvalidDataException){return;}catch(XmlException){return;}
            throw new Exception("Migration metadata accepted "+cause+".");
        }
        private static void Check(bool condition,string message){if(!condition)throw new Exception(message);}
        private static byte[] Compress(string text)
        {
            byte[] data=Encoding.UTF8.GetBytes(text);
            using(var output=new MemoryStream())
            {
                output.Write(new byte[]{0xA8,0xA9,0x30,0x31,12,0,0,0});output.Write(BitConverter.GetBytes(data.Length));output.WriteByte(0x78);output.WriteByte(0x9C);
                using(var deflate=new DeflateStream(output,CompressionMode.Compress,true))deflate.Write(data);
                uint a=1,b=0;foreach(byte value in data){a=(a+value)%65521;b=(b+a)%65521;}uint checksum=(b<<16)|a;
                output.WriteByte((byte)(checksum>>24));output.WriteByte((byte)(checksum>>16));output.WriteByte((byte)(checksum>>8));output.WriteByte((byte)checksum);return output.ToArray();
            }
        }
    }
}