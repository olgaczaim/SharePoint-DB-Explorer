using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Microsoft.Data.SqlClient;

namespace SharePointExplorer.Tests
{
    internal static class MigrationSqlChecks
    {
        internal static void Run(SqlConnectionOptions options)
        {
            var repository=new SqlRepository(options);var catalog=new SqlMigrationCatalog(repository);var scopes=new List<Node>();
            using(SqlConnection connection=repository.OpenConnection())
            using(SqlCommand command=connection.CreateCommand())
            {
                command.CommandText=@"SELECT l.tp_SiteId,l.tp_WebId,l.tp_ID,l.tp_RootFolder,l.tp_BaseType,l.tp_Title FROM dbo.AllLists l
JOIN dbo.Sites s ON s.Id=l.tp_SiteId AND s.Deleted=0 JOIN dbo.Webs w ON w.SiteId=l.tp_SiteId AND w.Id=l.tp_WebId AND w.DeleteTransactionId=0x
WHERE l.tp_DeleteTransactionId=0x AND l.tp_Title IN(N'List1_WF',N'doclib1',N'test') ORDER BY l.tp_SiteId,l.tp_ID;";
                using(SqlDataReader reader=command.ExecuteReader())while(reader.Read())scopes.Add(new Node {Kind=reader.GetInt32(4)==1?NodeKind.Library:NodeKind.List,
                    SiteId=reader.GetGuid(0),WebId=reader.GetGuid(1),ListId=reader.GetGuid(2),Id=reader.GetGuid(3),ListBaseType=reader.GetInt32(4),Name=reader.GetString(5)});
            }
            if(scopes.Count==0){Console.WriteLine("SKIP migration SQL fixtures: recorded list/library scopes are absent.");return;}
            foreach(Node scope in scopes)
            {
                MigrationListSnapshot snapshot=catalog.ReadMigrationList(scope);
                Check(snapshot.List.SiteId==scope.SiteId && snapshot.List.WebId==scope.WebId && snapshot.List.ListId==scope.ListId && snapshot.List.Id==scope.Id &&
                    snapshot.Items.Select(item=>item.ItemId).SequenceEqual(snapshot.Items.Select(item=>item.ItemId).OrderBy(id=>id)) &&
                    snapshot.Items.All(item=>item.Document.HistoryVersion==0 && item.Document.SiteId==scope.SiteId && item.Document.WebId==scope.WebId && item.Document.ListId==scope.ListId),
                    "SQL migration snapshot lost exact active scope/current identity or stable item ordering.");
                using(SqlConnection connection=repository.OpenConnection())
                {
                    using(SqlCommand command=connection.CreateCommand())
                    {
                        command.CommandText=@"SELECT l.tp_Title,l.tp_ImageUrl,l.tp_Version,l.tp_BaseType,l.tp_ServerTemplate,l.tp_Flags,l.tp_Description,s.FullUrl,w.FullUrl,
 (SELECT COUNT_BIG(DISTINCT u.tp_ID) FROM dbo.AllUserData u WHERE u.tp_SiteId=l.tp_SiteId AND u.tp_ListId=l.tp_ID AND u.tp_RowOrdinal=0 AND u.tp_DocId<>l.tp_RootFolder
 AND u.tp_IsCurrent=1 AND u.tp_IsCurrentVersion=1 AND u.tp_CalculatedVersion=0 AND u.tp_DeleteTransactionId=0x),
 l.tp_MaxMajorVersionCount,l.tp_MaxMajorwithMinorVersionCount,
 (SELECT TOP(1) root.InternalVersion FROM dbo.AllDocs root WHERE root.SiteId=l.tp_SiteId AND root.WebId=l.tp_WebId AND root.ListId=l.tp_ID
 AND root.Id=l.tp_RootFolder AND root.Type=1 AND root.IsCurrentVersion=1 AND root.DeleteTransactionId=0x ORDER BY root.Level DESC,root.InternalVersion DESC)
FROM dbo.AllLists l JOIN dbo.Sites s ON s.Id=l.tp_SiteId JOIN dbo.Webs w ON w.SiteId=l.tp_SiteId AND w.Id=l.tp_WebId
WHERE l.tp_SiteId=@SiteId AND l.tp_WebId=@WebId AND l.tp_ID=@ListId AND l.tp_RootFolder=@RootId AND l.tp_DeleteTransactionId=0x;";
                        Scope(command,scope);command.Parameters.Add("@RootId",SqlDbType.UniqueIdentifier).Value=scope.Id;
                        using(SqlDataReader reader=command.ExecuteReader())
                        {
                            Check(reader.Read(),"The SQL metadata fixture disappeared during migration validation.");
                            Check(snapshot.Metadata.Title==reader.GetString(0) && snapshot.Metadata.ImageUrl==reader.GetString(1) && snapshot.Metadata.Version==reader.GetInt32(2) &&
                                snapshot.Metadata.BaseType==reader.GetInt32(3) && snapshot.Metadata.ServerTemplate==reader.GetInt32(4) && snapshot.Metadata.Flags==reader.GetInt64(5) &&
                                snapshot.Metadata.Description==(reader.IsDBNull(6)?String.Empty:reader.GetString(6)) && snapshot.Metadata.StoredSiteUrl==reader.GetString(7) &&
                                snapshot.Metadata.StoredWebUrl==reader.GetString(8) && snapshot.Items.Count==reader.GetInt64(9) &&
                                snapshot.Metadata.MaxMajorVersions==(reader.IsDBNull(10)?(int?)null:reader.GetInt32(10)) &&
                                snapshot.Metadata.MaxMajorWithMinorVersions==(reader.IsDBNull(11)?(int?)null:reader.GetInt32(11)) &&
                                snapshot.Metadata.RootInternalVersion==(reader.IsDBNull(12)?(int?)null:reader.GetInt32(12)),
                                "SQL migration metadata or whole-list item counts differ from the restored source.");
                        }
                    }
                    DeploymentContextChecks.Verify(connection,snapshot);
                    CheckStoredTemplatesAndViews(connection,scope,snapshot);
                    foreach(MigrationItemSnapshot item in snapshot.Items)
                    using(SqlCommand command=connection.CreateCommand())
                    {
                        command.CommandText=@"SELECT TOP(1) u.tp_DocId,u.tp_GUID,u.tp_ParentId,u.tp_Level,u.tp_UIVersion,u.tp_Author,u.tp_Editor,u.tp_Created,u.tp_Modified,u.tp_HasAttachment,
 d.DoclibRowId,d.DirName,d.LeafName,d.Level,d.UIVersion FROM dbo.AllUserData u JOIN dbo.AllDocs d ON d.SiteId=u.tp_SiteId AND d.WebId=@WebId AND d.ListId=u.tp_ListId AND d.Id=u.tp_DocId
WHERE u.tp_SiteId=@SiteId AND u.tp_ListId=@ListId AND u.tp_ID=@ItemId AND u.tp_RowOrdinal=0 AND u.tp_IsCurrent=1 AND u.tp_IsCurrentVersion=1
 AND u.tp_CalculatedVersion=0 AND u.tp_DeleteTransactionId=0x AND d.IsCurrentVersion=1 AND d.DeleteTransactionId=0x
ORDER BY u.tp_Level DESC,u.tp_Version DESC,d.Level DESC,d.InternalVersion DESC;";
                        Scope(command,scope);command.Parameters.Add("@ItemId",SqlDbType.Int).Value=item.ItemId;
                        using(SqlDataReader reader=command.ExecuteReader())
                        {
                            Check(reader.Read(),"The source item disappeared during migration validation.");Node document=item.Document;
                            Check(item.DocumentId==reader.GetGuid(0) && item.ItemUniqueId==reader.GetGuid(1) && item.ParentFolderId==reader.GetGuid(2) &&
                                document.Level==reader.GetByte(3) && document.UiVersion==reader.GetInt32(4) && item.AuthorId==reader.GetInt32(5) && item.EditorId==reader.GetInt32(6) &&
                                item.Created==reader.GetDateTime(7) && item.Modified==reader.GetDateTime(8) && item.HasAttachments==reader.GetBoolean(9) && item.ItemId==reader.GetInt32(10) &&
                                document.Name==reader.GetString(12) && document.Level==reader.GetByte(13) && document.UiVersion==reader.GetInt32(14),
                                "SQL migration item metadata combines different item/document source identities.");
                        }
                        Check(snapshot.Users.Any(user=>user.Id==item.AuthorId) && snapshot.Users.Any(user=>user.Id==item.EditorId),
                            "A migration item omitted its recorded author/editor user identity.");
                    }
                }
                if(scope.Kind==NodeKind.List)
                    Check(snapshot.Files.Count==0 && snapshot.Items.Where(item=>!item.IsFolder).All(item=>item.Document.Kind==NodeKind.ListItem && item.Document.HasStream==false),
                        "Ordinary list metadata items or attachment files became primary library content.");
                else
                    Check(snapshot.Files.Count==snapshot.Items.Count(item=>item.Document.Kind==NodeKind.File),"Library document metadata lost its owning list-item relation.");
                Node modified=MigrationSnapshotCopy.Node(scope);modified.Id=Guid.NewGuid();bool rejected=false;
                try{catalog.ReadMigrationList(modified);}catch(ContentUnavailableException){rejected=true;}
                Check(rejected,"Migration snapshot accepted an unrelated source root folder.");
                Console.WriteLine("PASS migration SQL "+scope.Name+": "+snapshot.Items.Count+" current items, "+snapshot.Files.Count+" documents, "+snapshot.Users.Count+" users, "+snapshot.Views.Count+" views, "+snapshot.TemplateFiles.Count+" setup files");
            }
        }
        private static void CheckStoredTemplatesAndViews(SqlConnection connection,Node scope,MigrationListSnapshot snapshot)
        {
            using(SqlCommand command=connection.CreateCommand())
            {
                command.CommandText=@"SELECT s.PlatformVersion,w.ProductVersion,w.Language,
 (SELECT Version FROM dbo.Versions WHERE VersionId='00000000-0000-0000-0000-000000000000'),
 (SELECT COUNT(DISTINCT p.tp_ID) FROM dbo.AllWebParts p WHERE p.tp_SiteId=s.Id AND p.tp_ListId=@ListId AND p.tp_UserID IS NULL
 AND p.tp_Deleted=0 AND p.tp_IsCurrentVersion=1 AND p.tp_Type IN(0,1))
FROM dbo.Sites s JOIN dbo.Webs w ON w.SiteId=s.Id AND w.Id=@WebId WHERE s.Id=@SiteId AND s.Deleted=0 AND w.DeleteTransactionId=0x;";
                Scope(command,scope);
                using(SqlDataReader reader=command.ExecuteReader())
                {
                    Check(reader.Read(),"The source site/web metadata disappeared during migration validation.");
                    Check(snapshot.Metadata.PlatformVersion==reader.GetString(0) && snapshot.Metadata.ProductVersion==Convert.ToInt32(reader.GetValue(1)) &&
                        snapshot.Metadata.Language==reader.GetInt32(2) && snapshot.Metadata.SourceVersion==reader.GetString(3) && snapshot.Views.Count==reader.GetInt32(4),
                        "Migration source generation metadata or active shared-view count differs from SQL.");
                }
            }
            foreach(MigrationTemplateFile template in snapshot.TemplateFiles)
            using(SqlCommand command=connection.CreateCommand())
            {
                command.CommandText=@"SELECT TOP(1) ParentId,DirName,LeafName,HasStream,SetupPath,SetupPathVersion,SetupPathUser,Level,InternalVersion,UIVersion
FROM dbo.AllDocs WHERE SiteId=@SiteId AND WebId=@WebId AND ListId=@ListId AND Id=@DocumentId AND Type=0 AND IsCurrentVersion=1 AND DeleteTransactionId=0x
ORDER BY Level DESC,InternalVersion DESC;";
                Scope(command,scope);command.Parameters.Add("@DocumentId",SqlDbType.UniqueIdentifier).Value=template.Document.Id;
                using(SqlDataReader reader=command.ExecuteReader())
                {
                    Check(reader.Read(),"A template page disappeared during migration validation.");Node document=template.Document;
                    Check(document.ParentId==reader.GetGuid(0) && document.Name==reader.GetString(2) && document.HasStream==false &&
                        Convert.ToInt32(reader.GetValue(3))==0 && !String.IsNullOrWhiteSpace(template.SetupPath) && template.SetupPath==reader.GetString(4) &&
                        template.SetupPathVersion==(reader.IsDBNull(5)?(byte?)null:reader.GetByte(5)) &&
                        template.SetupPathUser==(reader.IsDBNull(6)?null:reader.GetString(6)) && document.Level==reader.GetByte(7) &&
                        document.InternalVersion==reader.GetInt32(8) && document.UiVersion==reader.GetInt32(9),
                        "Migration ghost file lacks exact recorded setup metadata or current source identity.");
                    Guid parent=document.ParentId;var visited=new HashSet<Guid>();
                    while(parent!=snapshot.Metadata.RootFolderId)
                    {
                        Check(visited.Add(parent),"A template folder chain is cyclic.");Node folder=snapshot.Folders.SingleOrDefault(candidate=>candidate.Id==parent);
                        Check(folder!=null,"A template file omitted a recorded folder ancestor.");parent=folder.ParentId;
                    }
                }
            }
            foreach(MigrationViewSnapshot view in snapshot.Views)
            using(SqlCommand command=connection.CreateCommand())
            {
                command.CommandText=@"SELECT TOP(1) tp_PageUrlID,tp_Flags,tp_Type,tp_Level,tp_Version,tp_DisplayName,tp_View
FROM dbo.AllWebParts WHERE tp_SiteId=@SiteId AND tp_ListId=@ListId AND tp_ID=@ViewId AND tp_UserID IS NULL
AND tp_Deleted=0 AND tp_IsCurrentVersion=1 AND tp_Type IN(0,1) ORDER BY tp_Level DESC,tp_Version DESC;";
                Scope(command,scope);command.Parameters.Add("@ViewId",SqlDbType.UniqueIdentifier).Value=view.Id;
                using(SqlDataReader reader=command.ExecuteReader())
                {
                    Check(reader.Read(),"A public view disappeared during migration validation.");
                    Check(view.FileId==reader.GetGuid(0) && view.Flags==unchecked((uint)reader.GetInt32(1)) && view.PageType==reader.GetByte(2) &&
                        view.Level==reader.GetByte(3) && view.Version==reader.GetInt32(4) && view.DisplayName==(reader.IsDBNull(5)?String.Empty:reader.GetString(5)) &&
                        snapshot.TemplateFiles.Any(template=>template.Document.Id==view.FileId) && (!reader.IsDBNull(6) || view.SchemaXml==String.Empty),
                        "Migration public view lost its source flags, level, page identity or NULL-fragment distinction.");
                }
            }
        }
        private static void Scope(SqlCommand command,Node scope)
        {command.Parameters.Add("@SiteId",SqlDbType.UniqueIdentifier).Value=scope.SiteId;command.Parameters.Add("@WebId",SqlDbType.UniqueIdentifier).Value=scope.WebId;command.Parameters.Add("@ListId",SqlDbType.UniqueIdentifier).Value=scope.ListId;}
        private static void Check(bool condition,string message){if(!condition)throw new Exception(message);}
    }
}