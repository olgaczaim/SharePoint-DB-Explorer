using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.Data.SqlClient;
namespace SharePointExplorer
{
    public sealed partial class SqlMigrationCatalog : ISharePointMigrationHistoryCatalog
    {
        public MigrationHistorySnapshot ReadMigrationHistory(MigrationListSnapshot snapshot)
        {
            if(snapshot==null)throw new ArgumentNullException("snapshot");
            using(SqlConnection connection=repository.OpenConnection())
            {
                var schema=ReadSchema(connection);ValidateSchema(schema);
                var historical=new List<MigrationItemSnapshot>();var files=new List<MigrationHistoryFile>();
                foreach(MigrationItemSnapshot current in snapshot.Items)
                {
                    List<Node> retained=current.Document.Kind==NodeKind.File ? repository.GetFileVersions(snapshot.Metadata.SiteId,current.DocumentId).Where(file=>file.HistoryVersion>0).ToList() : new List<Node>();
                    var rows=new List<ItemRow>();
                    using(SqlCommand command=Command(connection,@"
SELECT u.tp_ID,u.tp_DocId,u.tp_GUID,u.tp_ParentId,u.tp_Level,u.tp_Version,u.tp_UIVersion,u.tp_Author,u.tp_Editor,u.tp_Created,u.tp_Modified,
 u.tp_HasAttachment,u.tp_ContentTypeId,u.tp_ModerationStatus,u.tp_UIVersionString
FROM dbo.AllUserData u WHERE u.tp_SiteId=@SiteId AND u.tp_ListId=@ListId AND u.tp_ID=@ItemId AND u.tp_DocId=@DocId AND u.tp_GUID=@UniqueId
 AND u.tp_RowOrdinal=0 AND u.tp_DeleteTransactionId=0x AND u.tp_UIVersion>0 AND u.tp_UIVersion<@CurrentUiVersion
ORDER BY u.tp_UIVersion,u.tp_Level,u.tp_Version;"))
                    {
                        Scope(command,snapshot.List);command.Parameters.Add("@ItemId",SqlDbType.Int).Value=current.ItemId;
                        GuidParameter(command,"@DocId",current.DocumentId);GuidParameter(command,"@UniqueId",current.ItemUniqueId);
                        command.Parameters.Add("@CurrentUiVersion",SqlDbType.Int).Value=current.Document.UiVersion;
                        using(SqlDataReader reader=command.ExecuteReader())while(reader.Read())
                        {
                            Node document=MigrationSnapshotCopy.Node(current.Document);int ui=Int(reader,6);byte level=reader.GetByte(4);
                            if(document.Kind==NodeKind.File)
                            {
                                Node version=retained.SingleOrDefault(file=>file.UiVersion==ui && file.Level==level);
                                if(version==null)throw new ContentUnavailableException("Historical item metadata has no matching retained file version: "+document.Path+" version "+ui+".");
                                document=MigrationSnapshotCopy.Node(version);document.ListItemId=current.ItemId;document.ItemUniqueId=current.ItemUniqueId;
                            }
                            else {document.HistoryVersion=ui;document.UiVersion=ui;document.Level=level;document.InternalVersion=Int(reader,5);}
                            if(document.ParentId!=reader.GetGuid(3))throw new NotSupportedException("Historical item folder movement requires source path-history metadata which is not retained in this package profile.");
                            var row=new ItemRow {Id=Int(reader,0),Document=document,UniqueId=reader.GetGuid(2),Level=level,Version=Int(reader,5),Author=Int(reader,7),Editor=Int(reader,8),Created=Date(reader,9),Modified=Date(reader,10),HasAttachments=Bool(reader,11),ContentType=ContentType(reader.GetValue(12))};
                            row.Ordinals.Add(0,new Dictionary<string,object>(StringComparer.OrdinalIgnoreCase) {{"tp_ID",row.Id},{"tp_DocId",document.Id},{"tp_GUID",row.UniqueId},{"tp_ParentId",document.ParentId},{"tp_UIVersion",ui},{"tp_Author",row.Author},{"tp_Editor",row.Editor},{"tp_Created",row.Created},{"tp_Modified",row.Modified},{"tp_HasAttachment",row.HasAttachments},{"tp_ContentTypeId",row.ContentType},{"tp_ModerationStatus",reader.GetValue(13)},{"tp_UIVersionString",reader.GetValue(14)}});
                            if(rows.Any(previous=>previous.Document.UiVersion==ui))throw new NotSupportedException("Multiple stored publishing states represent one historical UI version; a verified state selection is required.");
                            rows.Add(row);
                        }
                    }
                    foreach(ItemRow row in rows)
                    {
                        string[] columns=snapshot.Fields.SelectMany(field=>field.StorageMappings).Select(mapping=>mapping.Column).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(column=>column,StringComparer.Ordinal).ToArray();
                        if(columns.Length>2000)throw new NotSupportedException("Historical field storage exceeds the supported SQL projection limit.");
                        if(columns.Length>0)
                        using(SqlCommand command=Command(connection,"SELECT u.tp_RowOrdinal,"+String.Join(",",columns.Select(column=>"u.["+column+"]"))+@"
FROM dbo.AllUserData u WHERE u.tp_SiteId=@SiteId AND u.tp_ListId=@ListId AND u.tp_ID=@ItemId AND u.tp_DocId=@DocId
 AND u.tp_Level=@Level AND u.tp_Version=@Version AND u.tp_UIVersion=@UiVersion AND u.tp_DeleteTransactionId=0x ORDER BY u.tp_RowOrdinal;"))
                        {
                            Scope(command,snapshot.List);command.Parameters.Add("@ItemId",SqlDbType.Int).Value=row.Id;GuidParameter(command,"@DocId",row.Document.Id);
                            command.Parameters.Add("@Level",SqlDbType.TinyInt).Value=row.Level;command.Parameters.Add("@Version",SqlDbType.Int).Value=row.Version;command.Parameters.Add("@UiVersion",SqlDbType.Int).Value=row.Document.UiVersion;
                            var ordinals=new HashSet<byte>();
                            using(SqlDataReader reader=command.ExecuteReader())while(reader.Read())
                            {
                                byte ordinal=reader.GetByte(0);if(!ordinals.Add(ordinal))throw new InvalidDataException("Duplicate historical storage row ordinal.");
                                Dictionary<string,object> values;if(!row.Ordinals.TryGetValue(ordinal,out values))row.Ordinals.Add(ordinal,values=new Dictionary<string,object>(StringComparer.OrdinalIgnoreCase));
                                for(int index=0;index<columns.Length;index++)values[columns[index]]=reader.IsDBNull(index+1)?null:reader.GetValue(index+1);
                            }
                        }
                        var fieldValues=new List<MigrationFieldValue>();
                        foreach(MigrationFieldDefinition field in snapshot.Fields)
                        {
                            if(field.StorageMappings.Count==0){AddSystemValue(fieldValues,field,row);continue;}
                            foreach(MigrationStorageMapping mapping in field.StorageMappings)
                            {
                                Dictionary<string,object> values;object value;
                                if(!row.Ordinals.TryGetValue(mapping.RowOrdinal,out values) || !values.TryGetValue(mapping.Column,out value))throw new ContentUnavailableException("A historical field's declared storage row is absent: "+field.Name+".");
                                fieldValues.Add(new MigrationFieldValue(field.Name,field.Id,mapping.Component,mapping.SqlType,value));
                            }
                        }
                        row.Document.Created=row.Created;row.Document.Modified=row.Modified;
                        historical.Add(new MigrationItemSnapshot(row.Document,row.Id,row.UniqueId,row.Author,row.Editor,row.Created,row.Modified,row.HasAttachments,row.ContentType,fieldValues));
                    }
                    foreach(Node file in retained)
                    {
                        if(!historical.Any(item=>item.DocumentId==file.Id && item.Document.UiVersion==file.UiVersion && item.Document.Level==file.Level))throw new ContentUnavailableException("A retained file version has no matching historical item metadata: "+file.Path+".");
                        using(SqlCommand command=Command(connection,"SELECT CheckinComment FROM dbo.AllDocVersions WHERE SiteId=@SiteId AND Id=@DocId AND UIVersion=@UiVersion AND Level=@Level AND COALESCE(InternalVersion,0)=@InternalVersion AND DeleteTransactionId=0x;"))
                        {
                            GuidParameter(command,"@SiteId",file.SiteId);GuidParameter(command,"@DocId",file.Id);command.Parameters.Add("@UiVersion",SqlDbType.Int).Value=file.UiVersion;command.Parameters.Add("@Level",SqlDbType.TinyInt).Value=file.Level;command.Parameters.Add("@InternalVersion",SqlDbType.Int).Value=file.InternalVersion;
                            using(SqlDataReader reader=command.ExecuteReader()){if(!reader.Read())throw new ContentUnavailableException("A retained file version changed during history collection.");string comment=reader.IsDBNull(0)?null:reader.GetString(0);if(reader.Read())throw new InvalidDataException("Duplicate retained file identity.");files.Add(new MigrationHistoryFile(file,comment));}
                        }
                    }
                }
                var users=ReadUsers(connection,snapshot.Metadata,snapshot.Fields.ToList(),historical,schema,snapshot.TemplateFiles.ToList());
                return new MigrationHistorySnapshot(historical.OrderBy(item=>item.ItemId).ThenBy(item=>item.Document.UiVersion),users,files.OrderBy(file=>file.Document.Id).ThenBy(file=>file.Document.UiVersion));
            }
        }
    }
}
