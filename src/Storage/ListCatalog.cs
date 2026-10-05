using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace SharePointExplorer
{
    // Ordinary list items have AllDocs identity/path records but are metadata,
    // not exportable documents. Read their current ordinal-zero item row and
    // only the title column explicitly declared by the stored field schema.
    public sealed class ListCatalog
    {
        private readonly SqlRepository repository;
        public ListCatalog(SqlRepository repository)
        {
            if(repository==null) throw new ArgumentNullException("repository");
            this.repository=repository;
        }

        public List<Node> GetWebLists(Node web)
        {
            if(web==null || web.Kind!=NodeKind.Site) throw new ArgumentException("Select a site to browse lists.","web");
            var result=new List<Node>();
            using(SqlConnection connection=repository.OpenConnection())
            using(SqlCommand command=Command(connection,@"
SELECT l.tp_ID,l.tp_RootFolder,l.tp_BaseType,
       COALESCE(NULLIF(l.tp_Title,N''),NULLIF(d.LeafName,N''),CONVERT(nvarchar(36),l.tp_ID)),
       d.DirName,d.LeafName,d.TimeLastModified
FROM dbo.AllLists l
OUTER APPLY
(
    SELECT TOP(1) r.DirName,r.LeafName,r.TimeLastModified
    FROM dbo.AllDocs r
    WHERE r.SiteId=l.tp_SiteId AND r.WebId=l.tp_WebId AND r.ListId=l.tp_ID
      AND r.Id=l.tp_RootFolder AND r.Type=1 AND r.IsCurrentVersion=1 AND r.DeleteTransactionId=0x
    ORDER BY r.Level DESC,r.InternalVersion DESC
) d
WHERE l.tp_SiteId=@SiteId AND l.tp_WebId=@WebId AND l.tp_BaseType<>1 AND l.tp_DeleteTransactionId=0x
  AND EXISTS(SELECT 1 FROM dbo.Sites s WHERE s.Id=l.tp_SiteId AND s.Deleted=0)
  AND EXISTS(SELECT 1 FROM dbo.Webs w WHERE w.SiteId=l.tp_SiteId AND w.Id=l.tp_WebId AND w.DeleteTransactionId=0x)
ORDER BY l.tp_Title,l.tp_ID;"))
            {
                GuidParameter(command,"@SiteId",web.SiteId); GuidParameter(command,"@WebId",web.WebId);
                using(SqlDataReader reader=command.ExecuteReader())
                    while(reader.Read()) result.Add(new Node {
                        Kind=NodeKind.List,SiteId=web.SiteId,WebId=web.WebId,ListId=reader.GetGuid(0),
                        Id=reader.GetGuid(1),ListBaseType=reader.GetInt32(2),Name=reader.GetString(3),
                        Path=CombinePath(Text(reader,4),Text(reader,5)),Modified=Date(reader,6)
                    });
            }
            return result;
        }

        public List<Node> GetChildren(Node parent)
        {
            if(parent==null || (parent.Kind!=NodeKind.List && parent.Kind!=NodeKind.Folder))
                throw new ArgumentException("Select a list or a list folder.","parent");
            using(SqlConnection connection=repository.OpenConnection())
            {
                HashSet<string> columns=ValidateItemSchema(connection);
                int baseType;
                TitleField title=GetTitleField(connection,parent,columns,out baseType);
                var result=ReadFolders(connection,parent,baseType);
                ReadItems(connection,parent,baseType,title,result);
                return result;
            }
        }

        private static HashSet<string> ValidateItemSchema(SqlConnection connection)
        {
            var actual=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using(SqlCommand command=Command(connection,"SELECT c.name FROM sys.columns c WHERE c.object_id=OBJECT_ID(N'dbo.AllUserData');"))
            using(SqlDataReader reader=command.ExecuteReader()) while(reader.Read()) actual.Add(reader.GetString(0));
            string[] required={"tp_ID","tp_ListId","tp_SiteId","tp_RowOrdinal","tp_Version","tp_Modified","tp_Created","tp_HasAttachment","tp_IsCurrent","tp_GUID","tp_ParentId","tp_DocId","tp_DeleteTransactionId","tp_Level","tp_IsCurrentVersion","tp_UIVersion","tp_CalculatedVersion"};
            foreach(string name in required)
                if(!actual.Contains(name)) throw new NotSupportedException("Ordinary list browsing requires dbo.AllUserData."+name+" in this content database.");
            return actual;
        }

        private static TitleField GetTitleField(SqlConnection connection,Node parent,HashSet<string> columns,out int baseType)
        {
            using(SqlCommand command=Command(connection,@"
SELECT l.tp_BaseType,l.tp_RootFolder,l.tp_Fields
FROM dbo.AllLists l
WHERE l.tp_SiteId=@SiteId AND l.tp_WebId=@WebId AND l.tp_ID=@ListId AND l.tp_BaseType<>1 AND l.tp_DeleteTransactionId=0x
  AND EXISTS(SELECT 1 FROM dbo.Sites s WHERE s.Id=l.tp_SiteId AND s.Deleted=0)
  AND EXISTS(SELECT 1 FROM dbo.Webs w WHERE w.SiteId=l.tp_SiteId AND w.Id=l.tp_WebId AND w.DeleteTransactionId=0x);"))
            {
                Scope(command,parent);
                using(SqlDataReader reader=command.ExecuteReader())
                {
                    if(!reader.Read()) throw new ContentUnavailableException("The selected ordinary list is no longer present in this content source.");
                    baseType=reader.GetInt32(0);
                    if(parent.Kind==NodeKind.List && parent.Id!=reader.GetGuid(1)) throw new InvalidDataException("The selected list root does not match the current catalog.");
                    TitleField field=reader.IsDBNull(2) ? null : ListFieldSchema.ReadTitle((byte[])reader.GetValue(2));
                    if(field!=null && !columns.Contains(field.Column)) throw new InvalidDataException("The list title schema references a storage column absent from dbo.AllUserData.");
                    return field;
                }
            }
        }

        private static List<Node> ReadFolders(SqlConnection connection,Node parent,int baseType)
        {
            var result=new List<Node>();
            using(SqlCommand command=Command(connection,@"
WITH Folders AS
(
    SELECT d.Id,d.ParentId,d.DirName,d.LeafName,d.TimeLastModified,d.Level,d.InternalVersion,d.UIVersion,
           ROW_NUMBER() OVER(PARTITION BY d.Id ORDER BY d.Level DESC,d.InternalVersion DESC) rn
    FROM dbo.AllDocs d
    WHERE d.SiteId=@SiteId AND d.WebId=@WebId AND d.ListId=@ListId AND d.ParentId=@ParentId
      AND d.Type=1 AND d.IsCurrentVersion=1 AND d.DeleteTransactionId=0x
)
SELECT Id,ParentId,DirName,LeafName,TimeLastModified,Level,InternalVersion,UIVersion FROM Folders WHERE rn=1 ORDER BY LeafName,Id;"))
            {
                Scope(command,parent); GuidParameter(command,"@ParentId",parent.Id);
                using(SqlDataReader reader=command.ExecuteReader())
                    while(reader.Read()) result.Add(new Node {
                        Kind=NodeKind.Folder,SiteId=parent.SiteId,WebId=parent.WebId,ListId=parent.ListId,
                        ListBaseType=baseType,Id=reader.GetGuid(0),ParentId=reader.GetGuid(1),
                        Path=CombinePath(Text(reader,2),Text(reader,3)),Name=Text(reader,3),Modified=Date(reader,4),
                        Level=reader.IsDBNull(5)?(byte)0:reader.GetByte(5),InternalVersion=reader.IsDBNull(6)?0:reader.GetInt32(6),
                        UiVersion=reader.IsDBNull(7)?0:reader.GetInt32(7)
                    });
            }
            return result;
        }

        private static void ReadItems(SqlConnection connection,Node parent,int baseType,TitleField title,List<Node> result)
        {
            // Column names come only from validated Field/FieldRef attributes,
            // an allowlist and sys.columns, never from a user-supplied SQL name.
            string titleExpression=title==null ? "CAST(NULL AS nvarchar(max))" : "CONVERT(nvarchar(max),t.["+title.Column+"])";
            string titleJoin=title==null ? "" : @"
OUTER APPLY
(
    SELECT TOP(1) v.["+title.Column+@"]
    FROM dbo.AllUserData v
    WHERE v.tp_SiteId=u.tp_SiteId AND v.tp_ListId=u.tp_ListId AND v.tp_ID=u.tp_ID
      AND v.tp_Level=u.tp_Level AND v.tp_Version=u.tp_Version AND v.tp_RowOrdinal=@TitleOrdinal
      AND v.tp_IsCurrent=1 AND v.tp_IsCurrentVersion=1 AND v.tp_CalculatedVersion=0 AND v.tp_DeleteTransactionId=0x
) t";
            string sql=@"
WITH CurrentItems AS
(
    SELECT u.tp_DocId,u.tp_ID,u.tp_GUID,u.tp_ParentId,u.tp_Modified,u.tp_Created,u.tp_HasAttachment,
           u.tp_Level,u.tp_Version,u.tp_UIVersion,u.tp_SiteId,u.tp_ListId,
           ROW_NUMBER() OVER(PARTITION BY u.tp_ID ORDER BY u.tp_Level DESC,u.tp_Version DESC) rn
    FROM dbo.AllUserData u
    WHERE u.tp_SiteId=@SiteId AND u.tp_ListId=@ListId AND u.tp_ParentId=@ParentId AND u.tp_RowOrdinal=0
      AND u.tp_IsCurrent=1 AND u.tp_IsCurrentVersion=1 AND u.tp_CalculatedVersion=0 AND u.tp_DeleteTransactionId=0x
      AND NOT EXISTS(SELECT 1 FROM dbo.AllDocs f WHERE f.SiteId=u.tp_SiteId AND f.ListId=u.tp_ListId
                     AND f.Id=u.tp_DocId AND f.Type=1 AND f.IsCurrentVersion=1 AND f.DeleteTransactionId=0x)
)
SELECT u.tp_DocId,u.tp_ID,u.tp_GUID,u.tp_ParentId,u.tp_Modified,u.tp_Created,u.tp_HasAttachment,
       u.tp_Level,u.tp_Version,u.tp_UIVersion,"+titleExpression+@" AS Title,d.DirName,d.LeafName
FROM CurrentItems u"+titleJoin+@"
OUTER APPLY
(
    SELECT TOP(1) a.DirName,a.LeafName
    FROM dbo.AllDocs a
    WHERE a.SiteId=u.tp_SiteId AND a.WebId=@WebId AND a.ListId=u.tp_ListId AND a.Id=u.tp_DocId
      AND a.Level=u.tp_Level AND a.Type=0 AND a.IsCurrentVersion=1 AND a.DeleteTransactionId=0x
    ORDER BY a.InternalVersion DESC
) d
WHERE u.rn=1 ORDER BY u.tp_ID;";
            using(SqlCommand command=Command(connection,sql))
            {
                Scope(command,parent); GuidParameter(command,"@ParentId",parent.Id);
                if(title!=null) command.Parameters.Add("@TitleOrdinal",SqlDbType.TinyInt).Value=title.RowOrdinal;
                using(SqlDataReader reader=command.ExecuteReader())
                    while(reader.Read())
                    {
                        int itemId=reader.GetInt32(1);
                        string itemTitle=Text(reader,10);
                        result.Add(new Node {
                            Kind=NodeKind.ListItem,SiteId=parent.SiteId,WebId=parent.WebId,ListId=parent.ListId,ListBaseType=baseType,
                            Id=reader.GetGuid(0),ListItemId=itemId,ItemUniqueId=reader.GetGuid(2),ParentId=reader.GetGuid(3),
                            Modified=Date(reader,4),Created=reader.IsDBNull(5)?(DateTime?)null:reader.GetDateTime(5),HasAttachments=reader.GetBoolean(6),
                            Level=reader.GetByte(7),InternalVersion=reader.GetInt32(8),UiVersion=reader.GetInt32(9),Title=itemTitle,
                            Name=String.IsNullOrWhiteSpace(itemTitle)?"Item "+itemId.ToString(CultureInfo.InvariantCulture):itemTitle,
                            Path=CombinePath(Text(reader,11),Text(reader,12)),HasStream=false
                        });
                    }
            }
        }

        private static SqlCommand Command(SqlConnection connection,string sql) { var command=connection.CreateCommand();command.CommandText=sql;command.CommandTimeout=120;return command; }
        private static void Scope(SqlCommand command,Node node) { GuidParameter(command,"@SiteId",node.SiteId);GuidParameter(command,"@WebId",node.WebId);GuidParameter(command,"@ListId",node.ListId); }
        private static void GuidParameter(SqlCommand command,string name,Guid value) { command.Parameters.Add(name,SqlDbType.UniqueIdentifier).Value=value; }
        private static string Text(SqlDataReader reader,int ordinal) { return reader.IsDBNull(ordinal)?String.Empty:reader.GetString(ordinal); }
        private static DateTime Date(SqlDataReader reader,int ordinal) { return reader.IsDBNull(ordinal)?DateTime.MinValue:reader.GetDateTime(ordinal); }
        private static string CombinePath(string directory,string leaf) { return String.IsNullOrEmpty(directory)?leaf:String.IsNullOrEmpty(leaf)?directory:directory.TrimEnd('/')+"/"+leaf; }
    }

    internal sealed class TitleField
    {
        internal string Column;
        internal byte RowOrdinal;
    }

    internal static class ListFieldSchema
    {
        // WSS compressed metadata is distinct from document chunk compression.
        // MS-WSSFO2 2.2.4.8; MS-WSSFO3 3.1.5.32.1 and 2.2.7.3.5.
        private const int MaxSchemaBytes=16*1024*1024;
        internal static TitleField ReadTitle(byte[] stored)
        {
            string text=new UTF8Encoding(false,true).GetString(Decompress(stored));
            int xmlStart=text.IndexOf('<');
            if(xmlStart<0) throw new InvalidDataException("The list field metadata contains no XML definitions.");
            string prefix=text.Substring(0,xmlStart);
            if(!Regex.IsMatch(prefix,@"\A[0-9]+(?:\.[0-9]+)*\z")) throw new NotSupportedException("The list field metadata version prefix is not supported.");
            var settings=new XmlReaderSettings { DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=MaxSchemaBytes,IgnoreComments=true };
            var document=new XmlDocument { XmlResolver=null };
            using(var reader=XmlReader.Create(new StringReader("<Fields>"+text.Substring(xmlStart)+"</Fields>"),settings)) document.Load(reader);
            TitleField result=null;
            foreach(XmlNode node in document.DocumentElement.ChildNodes)
            {
                var element=node as XmlElement;
                if(element==null || (element.LocalName!="Field" && element.LocalName!="FieldRef") || element.GetAttribute("Name")!="Title") continue;
                string column=element.GetAttribute("ColName");
                if(String.IsNullOrEmpty(column)) continue;
                if(!Regex.IsMatch(column,@"\A(?:nvarchar|ntext)[1-9][0-9]{0,2}\z")) throw new InvalidDataException("The list title references an unsupported storage column.");
                byte ordinal=0;
                string declaredOrdinal=element.GetAttribute("RowOrdinal");
                if(declaredOrdinal.Length!=0 && !Byte.TryParse(declaredOrdinal,NumberStyles.None,CultureInfo.InvariantCulture,out ordinal)) throw new InvalidDataException("The list title has an invalid row ordinal.");
                if(result!=null && (result.Column!=column || result.RowOrdinal!=ordinal)) throw new InvalidDataException("The list schema declares conflicting title storage mappings.");
                result=new TitleField { Column=column,RowOrdinal=ordinal };
            }
            return result;
        }

        internal static byte[] Decompress(byte[] stored)
        {
            if(stored==null) throw new ArgumentNullException("stored");
            if(stored.Length<18 || stored[0]!=0xA8 || stored[1]!=0xA9 || stored[2]!=0x30 || stored[3]!=0x31 || BitConverter.ToUInt32(stored,4)!=12)
                throw new NotSupportedException("The stored list field metadata compression header is not supported.");
            uint expected=BitConverter.ToUInt32(stored,8);
            if(expected>MaxSchemaBytes) throw new NotSupportedException("The list field metadata exceeds the supported 16 MB limit.");
            int cmf=stored[12],flg=stored[13];
            if((cmf&15)!=8 || (cmf>>4)>7 || ((cmf<<8)+flg)%31!=0 || (flg&32)!=0) throw new InvalidDataException("The list field metadata has an invalid ZLIB header.");
            using(var input=new MemoryStream(stored,14,stored.Length-18,false))
            using(var inflater=new DeflateStream(input,CompressionMode.Decompress))
            using(var output=new MemoryStream((int)expected))
            {
                byte[] buffer=new byte[8192];int count;
                uint sum1=1,sum2=0;
                while((count=inflater.Read(buffer,0,buffer.Length))!=0)
                {
                    if(output.Length+count>expected) throw new InvalidDataException("The list field metadata exceeds its declared uncompressed length.");
                    output.Write(buffer,0,count);
                    for(int index=0;index<count;index++) { sum1=(sum1+buffer[index])%65521;sum2=(sum2+sum1)%65521; }
                }
                if(output.Length!=expected) throw new InvalidDataException("The list field metadata has an incorrect uncompressed length.");
                int end=stored.Length-4;
                uint checksum=((uint)stored[end]<<24)|((uint)stored[end+1]<<16)|((uint)stored[end+2]<<8)|stored[end+3];
                if(checksum!=((sum2<<16)|sum1)) throw new InvalidDataException("The list field metadata checksum is invalid.");
                return output.ToArray();
            }
        }
    }
}
