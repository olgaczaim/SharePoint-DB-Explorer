using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Microsoft.Data.SqlClient;

namespace SharePointExplorer.Tests
{
    internal static class DeploymentContextChecks
    {
        // Compare emitted deployment identity context with independently scoped
        // source records. List payload IDs cannot substitute for system IDs.
        internal static void Verify(SqlConnection connection,MigrationListSnapshot snapshot)
        {
            var metadata=snapshot.Metadata;
            Guid rootWeb,userList;
            using(var command=connection.CreateCommand())
            {
                command.CommandText="SELECT RootWebId,UserInfoListId FROM dbo.Sites WHERE Id=@SiteId AND Deleted=0;";
                command.Parameters.Add("@SiteId",SqlDbType.UniqueIdentifier).Value=metadata.SiteId;
                using(var reader=command.ExecuteReader())
                {
                    Check(reader.Read(),"The deployment source site is absent.");
                    rootWeb=reader.GetGuid(0);userList=reader.GetGuid(1);
                    Check(!reader.Read(),"The deployment source site is ambiguous.");
                }
            }
            Check(snapshot.SystemObjects.Count==(rootWeb==metadata.WebId?3:4) &&
                snapshot.SystemObjects.Select(value=>value.Id).Distinct().Count()==snapshot.SystemObjects.Count &&
                snapshot.SystemObjects.Select(value=>value.Role).Distinct().Count()==snapshot.SystemObjects.Count,
                "Deployment context has missing, duplicate or unrelated system identities.");
            MigrationSystemObject root=Find(snapshot,"RootWeb","Web",rootWeb);
            using(var command=connection.CreateCommand())
            {
                command.CommandText="SELECT FullUrl FROM dbo.Webs WHERE SiteId=@SiteId AND Id=@RootWeb AND DeleteTransactionId=0x;";
                command.Parameters.Add("@SiteId",SqlDbType.UniqueIdentifier).Value=metadata.SiteId;
                command.Parameters.Add("@RootWeb",SqlDbType.UniqueIdentifier).Value=rootWeb;
                Check(root.Url==(string)command.ExecuteScalar(),"Root-web context used the selected subsite URL instead of the source root URL.");
            }
            if(rootWeb!=metadata.WebId)
                Check(Find(snapshot,"SelectedWeb","Web",metadata.WebId).Url==metadata.StoredWebUrl,
                    "Selected-subsite context is missing or uses another source URL.");
            MigrationSystemObject users=Find(snapshot,"UserInfoList","List",userList);
            using(var command=connection.CreateCommand())
            {
                command.CommandText=@"SELECT d.DirName,d.LeafName FROM dbo.AllLists l JOIN dbo.AllDocs d
 ON d.SiteId=l.tp_SiteId AND d.WebId=l.tp_WebId AND d.ListId=l.tp_ID AND d.Id=l.tp_RootFolder
WHERE l.tp_SiteId=@SiteId AND l.tp_WebId=@RootWeb AND l.tp_ID=@UserList AND l.tp_DeleteTransactionId=0x
 AND d.Type=1 AND d.IsCurrentVersion=1 AND d.DeleteTransactionId=0x;";
                command.Parameters.Add("@SiteId",SqlDbType.UniqueIdentifier).Value=metadata.SiteId;
                command.Parameters.Add("@RootWeb",SqlDbType.UniqueIdentifier).Value=rootWeb;
                command.Parameters.Add("@UserList",SqlDbType.UniqueIdentifier).Value=userList;
                using(var reader=command.ExecuteReader())
                {
                    Check(reader.Read() && users.Url==PathOf(reader.GetString(0),reader.GetString(1)),
                        "User-information-list context lost its source list GUID or root URL.");
                }
            }
            MigrationSystemObject parent=Find(snapshot,"RootParentFolder","Folder",snapshot.List.ParentId);
            using(var command=connection.CreateCommand())
            {
                command.CommandText=@"SELECT DirName,LeafName,Type FROM dbo.AllDocs
WHERE SiteId=@SiteId AND WebId=@WebId AND Id=@ParentId AND Type IN(1,2)
 AND IsCurrentVersion=1 AND DeleteTransactionId=0x;";
                command.Parameters.Add("@SiteId",SqlDbType.UniqueIdentifier).Value=metadata.SiteId;
                command.Parameters.Add("@WebId",SqlDbType.UniqueIdentifier).Value=metadata.WebId;
                command.Parameters.Add("@ParentId",SqlDbType.UniqueIdentifier).Value=parent.Id;
                using(var reader=command.ExecuteReader())
                    Check(reader.Read() && parent.Url==PathOf(reader.GetString(0),reader.GetString(1)),
                        "List root-parent context does not identify an active source folder.");
            }
            Console.WriteLine("PASS deployment SQL "+metadata.Title+": exact source root web, user-information list and parent folder identities/URLs");
        }
        private static MigrationSystemObject Find(MigrationListSnapshot snapshot,string role,string type,Guid id)
        {
            MigrationSystemObject value=snapshot.SystemObjects.SingleOrDefault(item=>item.Role==role);
            Check(value!=null && value.Type==type && value.Id==id && value.Url!=null,
                "Deployment context is missing or misidentifies "+role+".");
            return value;
        }
        private static string PathOf(string directory,string leaf)
        {return String.IsNullOrEmpty(directory)?leaf:String.IsNullOrEmpty(leaf)?directory:directory.TrimEnd('/')+"/"+leaf;}
        private static void Check(bool success,string message)
        {if(!success)throw new Exception(message);}
    }
}