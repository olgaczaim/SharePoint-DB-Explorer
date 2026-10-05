using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Microsoft.Data.SqlClient;

namespace SharePointExplorer
{
    // Read deployment identity context without invoking the farm deployment
    // procedure, creating temporary SQL tables, or changing the source.
    internal static class SqlDeploymentContext
    {
        internal static List<MigrationSystemObject> Read(SqlConnection connection,MigrationListMetadata metadata)
        {
            if(connection==null || metadata==null)throw new ArgumentNullException();
            var objects=new Dictionary<Guid,MigrationSystemObject>();
            Guid rootWeb;
            using(SqlCommand command=connection.CreateCommand())
            {
                command.CommandTimeout=120;
                command.CommandText=@"SELECT s.RootWebId,w.FullUrl,s.UserInfoListId,d.DirName,d.LeafName
FROM dbo.Sites s JOIN dbo.Webs w ON w.SiteId=s.Id AND w.Id=s.RootWebId AND w.DeleteTransactionId=0x
LEFT JOIN dbo.AllLists l ON l.tp_SiteId=s.Id AND l.tp_WebId=s.RootWebId AND l.tp_ID=s.UserInfoListId AND l.tp_DeleteTransactionId=0x
OUTER APPLY(SELECT TOP(1) a.DirName,a.LeafName FROM dbo.AllDocs a
 WHERE a.SiteId=s.Id AND a.WebId=s.RootWebId AND a.ListId=l.tp_ID AND a.Id=l.tp_RootFolder AND a.Type IN(1,2)
 AND a.IsCurrentVersion=1 AND a.DeleteTransactionId=0x ORDER BY a.Level DESC,a.InternalVersion DESC) d
WHERE s.Id=@SiteId AND s.Deleted=0;";
                command.Parameters.Add("@SiteId",SqlDbType.UniqueIdentifier).Value=metadata.SiteId;
                using(SqlDataReader reader=command.ExecuteReader())
                {
                    if(!reader.Read() || reader.IsDBNull(0) || reader.IsDBNull(1) || reader.IsDBNull(2) || reader.IsDBNull(3) || reader.IsDBNull(4))
                        throw new ContentUnavailableException("The source site has no complete active root-web/user-information-list identity context.");
                    rootWeb=reader.GetGuid(0);
                    Add(objects,new MigrationSystemObject(rootWeb,"Web",reader.GetString(1),"RootWeb"));
                    Add(objects,new MigrationSystemObject(reader.GetGuid(2),"List",PathOf(reader.GetString(3),reader.GetString(4)),"UserInfoList"));
                    if(reader.Read())throw new InvalidDataException("The source site has duplicate deployment context.");
                }
            }
            if(metadata.WebId!=rootWeb)
                Add(objects,new MigrationSystemObject(metadata.WebId,"Web",metadata.StoredWebUrl,"SelectedWeb"));
            using(SqlCommand command=connection.CreateCommand())
            {
                command.CommandTimeout=120;
                command.CommandText=@"SELECT TOP(1) d.DirName,d.LeafName FROM dbo.AllDocs d
WHERE d.SiteId=@SiteId AND d.WebId=@WebId AND d.Id=@ParentId AND d.Type IN(1,2)
 AND d.IsCurrentVersion=1 AND d.DeleteTransactionId=0x ORDER BY d.Level DESC,d.InternalVersion DESC;";
                command.Parameters.Add("@SiteId",SqlDbType.UniqueIdentifier).Value=metadata.SiteId;
                command.Parameters.Add("@WebId",SqlDbType.UniqueIdentifier).Value=metadata.WebId;
                command.Parameters.Add("@ParentId",SqlDbType.UniqueIdentifier).Value=metadata.List.ParentId;
                using(SqlDataReader reader=command.ExecuteReader())
                {
                    if(!reader.Read() || reader.IsDBNull(0) || reader.IsDBNull(1))
                        throw new ContentUnavailableException("The selected list root has no active parent-folder identity context.");
                    Add(objects,new MigrationSystemObject(metadata.List.ParentId,"Folder",PathOf(reader.GetString(0),reader.GetString(1)),"RootParentFolder"));
                }
            }
            return objects.Values.OrderBy(value=>value.Role,StringComparer.Ordinal).ThenBy(value=>value.Id).ToList();
        }
        private static void Add(Dictionary<Guid,MigrationSystemObject> objects,MigrationSystemObject value)
        {
            if(value.Id==Guid.Empty || value.Url==null)throw new InvalidDataException("A deployment context identity or URL is absent.");
            MigrationSystemObject previous;
            if(objects.TryGetValue(value.Id,out previous))
            {
                if(previous.Type!=value.Type || previous.Url!=value.Url || previous.Role!=value.Role)
                    throw new InvalidDataException("A deployment context identity has conflicting roles or URLs.");
                return;
            }
            objects.Add(value.Id,value);
        }
        private static string PathOf(string directory,string leaf)
        {return String.IsNullOrEmpty(directory)?leaf:String.IsNullOrEmpty(leaf)?directory:directory.TrimEnd('/')+"/"+leaf;}
    }
}
