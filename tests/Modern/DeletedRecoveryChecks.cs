#nullable enable
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Microsoft.Data.SqlClient;

namespace SharePointExplorer.Tests
{
    internal static class DeletedRecoveryChecks
    {
        public static void Run()
        {
            Mapping();IdentityAndRecovery();MetadataPublication();
            Console.WriteLine("PASS deleted transaction/version identity, no current fallback, source revalidation, immutable snapshots and retained metadata XML");
        }
        private static void Mapping()
        {
            Check(SqlDeletedCatalog.BuildRetainedProjection(new[] {"tp_ID","nvarchar1","name]with[brackets"})=="u.[tp_ID],u.[nvarchar1],u.[name]]with[brackets]",
                "Retained sparse field projection failed to preserve typed values or quote source column identities.");
            Reject<InvalidDataException>(()=>SqlDeletedCatalog.BuildRetainedProjection(new[] {"nvarchar1","NVARCHAR1"}));
            Reject<NotSupportedException>(()=>SqlDeletedCatalog.BuildRetainedProjection(Enumerable.Range(1,4097).Select(index=>"c"+index)));
            var catalog=new Catalog();Node file=catalog.Document;
            using var table=new DataTable();
            Type[] types={typeof(Guid),typeof(Guid),typeof(Guid),typeof(Guid),typeof(Guid),typeof(string),typeof(string),typeof(byte),typeof(long),typeof(byte),typeof(int),typeof(byte),typeof(int),typeof(int),typeof(DateTime),typeof(DateTime),typeof(byte[]),typeof(int),typeof(Guid),typeof(string),typeof(int),typeof(Guid),typeof(int),typeof(bool),typeof(string),typeof(DateTime)};
            for(int index=0;index<types.Length;index++)table.Columns.Add("c"+index,types[index]);
            object[] row={file.Id,file.SiteId,file.WebId,file.ListId,file.ParentId,"Library","doc.txt",(byte)0,4L,(byte)1,1,(byte)1,513,512,DateTime.UtcNow,DateTime.UtcNow,DeletedIdentity.Parse(file.DeletionTransactionId),1,Guid.NewGuid(),"Library",DBNull.Value,DBNull.Value,DBNull.Value,DBNull.Value,"doc.txt",DateTime.UtcNow};
            table.Rows.Add(row);object[] item=(object[])row.Clone();item[17]=0;item[20]=2;item[21]=Guid.NewGuid();item[22]=1;item[23]=false;item[24]="bbb";table.Rows.Add(item);
            object[] list=(object[])row.Clone();list[7]=(byte)1;list[17]=0;list[18]=file.Id;list[19]="Notes";table.Rows.Add(list);
            object[] incomplete=(object[])row.Clone();incomplete[16]=Array.Empty<byte>();table.Rows.Add(incomplete);
            using var reader=table.CreateDataReader();Check(reader.Read(),"Deleted file fixture absent.");Node mapped=SqlDeletedCatalog.ReadNode(reader);
            Check(mapped.Kind==NodeKind.File && mapped.IsDeleted && mapped.Size==4 && mapped.HistoryVersion==0 && mapped.DeletionTransactionId==file.DeletionTransactionId && mapped.Path=="Library/doc.txt","Deleted file mapper lost transaction or document metadata.");
            Check(reader.Read(),"Deleted item fixture absent.");mapped=SqlDeletedCatalog.ReadNode(reader);
            Check(mapped.Kind==NodeKind.ListItem && mapped.Name=="bbb" && mapped.ListItemId==2 && mapped.InternalVersion==1 && mapped.HasStream==false,"Deleted ordinary item was exposed as backing-file bytes or wrong version counter.");
            Check(reader.Read(),"Deleted list fixture absent.");mapped=SqlDeletedCatalog.ReadNode(reader);
            Check(mapped.Kind==NodeKind.List && mapped.Name=="Notes","Deleted list root was exposed as a plain folder.");
            Check(reader.Read(),"Incomplete deleted transaction fixture absent.");Reject<InvalidDataException>(()=>SqlDeletedCatalog.ReadNode(reader));
        }
        private static void IdentityAndRecovery()
        {
            for(int variant=0;variant<10;variant++)
            {
                var catalog=new Catalog();var store=new Store();Node changed=DeletedIdentity.Copy(catalog.Document);
                if(variant==0)changed.DeletionTransactionId="11111111111111111111111111111111";
                if(variant==1)changed.DeletionTransactionId=null!;
                if(variant==2)changed.SiteId=Guid.NewGuid();if(variant==3)changed.Id=Guid.NewGuid();if(variant==4)changed.WebId=Guid.NewGuid();
                if(variant==5)changed.ListId=Guid.NewGuid();if(variant==6)changed.UiVersion=1024;if(variant==7)changed.InternalVersion++;
                if(variant==8){changed.HistoryVersion=512;changed.UiVersion=512;}if(variant==9)changed.Level++;
                catalog.Lookup=_=>changed;
                using var session=Session(catalog,store);
                Reject<InvalidDataException>(()=>session.Engine.PrepareDeleted(catalog.Document));
                Check(store.Reads==0 && catalog.CurrentReads==0 && catalog.DeletedReads==1,"Wrong deleted transaction or version reached content retrieval.");
            }
            foreach(string invalid in new[] {"","00","XYZ",new string('G',32)})
            {
                var catalog=new Catalog();Node selected=DeletedIdentity.Copy(catalog.Document);selected.DeletionTransactionId=invalid;
                using var session=Session(catalog,new Store());Reject<ArgumentException>(()=>session.Engine.PrepareDeleted(selected));
                Check(catalog.DeletedReads==0,"Invalid transaction identity read the source.");
            }
            {
                var catalog=new Catalog();var store=new Store();using var session=Session(catalog,store);
                Reject<NotSupportedException>(()=>session.Engine.Prepare(catalog.Document));
                Reject<NotSupportedException>(()=>session.Engine.PrepareVersion(catalog.Document));
                Check(catalog.CurrentReads==0 && catalog.DeletedReads==0 && store.Reads==0,"Ordinary export substituted current bytes for a deleted selection.");
                Node selected=DeletedIdentity.Copy(catalog.Document);PreparedDocument prepared=session.Engine.PrepareDeleted(selected);
                selected.DeletionTransactionId="22222222222222222222222222222222";selected.Name="mutated.txt";
                Node exposed=prepared.Document;exposed.DeletionTransactionId=null!;
                using var output=new MemoryStream();RecoveryResult result=prepared.Recover(output);
                Check(Encoding.UTF8.GetString(output.ToArray())=="past" && result.Document.IsDeleted && result.Document.DeletionTransactionId==catalog.Document.DeletionTransactionId &&
                    catalog.DeletedReads==2 && catalog.CurrentReads==0 && result.Sha256==Digest(output.ToArray()),"Prepared deleted content lost frozen identity or returned current bytes.");
            }
            {
                var catalog=new Catalog();var store=new Store();using var session=Session(catalog,store);
                catalog.Lookup=_=>catalog.DeletedReads==1 ? DeletedIdentity.Copy(catalog.Document) : null;
                Reject<ContentUnavailableException>(()=>session.Engine.PrepareDeleted(catalog.Document));
                Check(store.Reads==1 && catalog.DeletedReads==2,"Deleted source purge was not detected after stream reads.");
            }
            {
                var catalog=new Catalog();var store=new Store();catalog.Document.HistoryVersion=512;catalog.Document.UiVersion=512;
                using var session=Session(catalog,store);RecoveryResult recovered=session.Engine.PrepareDeleted(catalog.Document).Recover(Stream.Null);
                Check(recovered.Document.HistoryVersion==512 && store.LastHistory==512,"Deleted historical recovery requested current stream mapping.");
            }
            {
                var catalog=new Catalog();var store=new Store();catalog.Lookup=request=>{request.DeletionTransactionId=null!;return DeletedIdentity.Copy(request);};
                using var session=Session(catalog,store);Reject<InvalidDataException>(()=>session.Engine.PrepareDeleted(catalog.Document));
                Check(store.Reads==0,"Mutated provider query changed the validated deletion identity.");
            }
        }
        private static void MetadataPublication()
        {
            var catalog=new Catalog();Node item=DeletedIdentity.Copy(catalog.Document);item.Kind=NodeKind.ListItem;item.ListBaseType=0;item.ListItemId=2;item.ItemUniqueId=Guid.NewGuid();item.Name="bbb";
            var snapshot=new DeletedListItemSnapshot(item,"<Fields><Field Name=\"Title\" ColName=\"nvarchar1\" /></Fields>",new[] {
                new DeletedListItemValue(0,"nvarchar1","nvarchar","bbb <&> \"quoted\"",false),new DeletedListItemValue(0,"int1","int","",true),new DeletedListItemValue(1,"varbinary1","varbinary",Convert.ToBase64String(new byte[] {0,255}),false)});
            string root=Path.GetFullPath(Path.Combine(".scratch","deleted-recovery-checks",Guid.NewGuid().ToString("N")));
            var exporter=new DeletedItemExportService();ExportResult first=exporter.Export(snapshot,root),second=exporter.Export(snapshot,root);
            item.Name="mutated";Node copy=snapshot.Item;copy.DeletionTransactionId=null!;
            XDocument xml=XDocument.Load(first.Path);XElement element=xml.Root!;
            Check(element.Name.LocalName=="DeletedListItemRecovery" && (string?)element.Attribute("DeletionTransactionId")==catalog.Document.DeletionTransactionId &&
                element.Element("StoredFieldSchemaXml")!.Value.Contains("ColName=\"nvarchar1\"") && element.Descendants("Value").First().Value=="bbb <&> \"quoted\"" &&
                element.Descendants("Value").ElementAt(1).Attribute("IsNull")!.Value=="true" && element.Descendants("Value").ElementAt(2).Value=="AP8=",
                "Recovered metadata XML changed the transaction, field schema, escaping, null or binary encoding.");
            Check(File.ReadAllBytes(first.Path).Length==first.Bytes && Digest(File.ReadAllBytes(first.Path))==first.Sha256 && first.Path!=second.Path && !Directory.GetFiles(root,"*.partial").Any(),"Metadata publication overwrote output, reported another checksum or left incomplete artifacts.");
        }
        public static void RunSql(SqlConnectionOptions options)
        {
            using var session=RecoverySession.OpenSql(options);
            var deleted=session.Catalog as ISharePointDeletedCatalog ?? throw new Exception("SQL retained deleted capability absent.");
            var repository=(SqlRepository)session.Catalog;int recoveredFiles=0,recoveredItems=0;
            foreach(Node site in session.Catalog.GetRootSites())
            {
                List<Node> nodes=deleted.GetDeletedItems(site);
                using(SqlConnection connection=repository.OpenConnection())using(SqlCommand command=connection.CreateCommand())
                {
                    command.CommandText=@"SELECT COUNT_BIG(*) FROM dbo.AllDocs d WHERE d.SiteId=@Site AND d.WebId=@Web AND d.Type=0 AND d.IsCurrentVersion=1 AND d.DeleteTransactionId<>0x
 AND NOT EXISTS(SELECT 1 FROM dbo.AllDocs p WHERE p.SiteId=d.SiteId AND p.WebId=d.WebId AND p.ListId=d.ListId AND p.Id=d.ParentId AND p.Type=1 AND p.IsCurrentVersion=1 AND p.DeleteTransactionId=d.DeleteTransactionId);";
                    command.Parameters.Add("@Site",SqlDbType.UniqueIdentifier).Value=site.SiteId;command.Parameters.Add("@Web",SqlDbType.UniqueIdentifier).Value=site.WebId;
                    long expected=Convert.ToInt64(command.ExecuteScalar());
                    Check(nodes.Count(node=>node.Kind==NodeKind.File && node.HistoryVersion==0 || node.Kind==NodeKind.ListItem)==expected,"SQL deleted root enumeration omitted retained document or ordinary-item rows.");
                }
                foreach(Node node in nodes)
                {
                    Check(node.IsDeleted && node.DeletionTransactionId.Length==32,"SQL deleted metadata exposed an active or incomplete transaction identity.");
                    if(node.Kind==NodeKind.File && session.Engine.CanExport(node))
                    {
                        Node exact=deleted.GetDeletedFile(node) ?? throw new Exception("SQL selected deleted file disappeared.");
                        Node foreign=DeletedIdentity.Copy(node);foreign.DeletionTransactionId="FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF";
                        Check(deleted.GetDeletedFile(foreign)==null,"SQL deleted revalidation accepted another deletion transaction.");
                        Reject<NotSupportedException>(()=>session.Engine.Prepare(node));
                        using var bytes=new MemoryStream();RecoveryResult recovered=session.Engine.PrepareDeleted(node).Recover(bytes);
                        Check(recovered.Bytes==exact.Size && recovered.Sha256==Digest(bytes.ToArray()),"SQL deleted byte length/checksum differs from frozen selected metadata.");
                        if(node.Id==new Guid("f0675d2e-4225-4ea9-8358-da3f7da48220"))
                            Check(recovered.Bytes==212541 && recovered.Sha256=="48c41a496332b7953b71429e76ad4a662210b50754626042e0c1d7b6affe64ef","Deleted JPEG differs from prior independently captured fixture bytes.");
                        if(node.Id==new Guid("2b7b249f-4f3c-412e-a715-f9dbf77cb062"))
                            Check(recovered.Bytes==49622 && bytes.ToArray().Take(8).SequenceEqual(new byte[] {137,80,78,71,13,10,26,10}),"Deleted PNG lost its native PNG header or declared length.");
                        Check(deleted.GetDeletedFileVersions(node).Exists(version=>DeletedIdentity.Same(version,node)),"SQL deleted version browser omitted the selected retained document.");
                        recoveredFiles++;
                    }
                    else if(node.Kind==NodeKind.ListItem)
                    {
                        DeletedListItemSnapshot snapshot=deleted.GetDeletedListItem(node);
                        Check(snapshot.Values.Count>0 && snapshot.Item.ItemUniqueId==node.ItemUniqueId && snapshot.StoredFieldsXml.Contains("<Field"),"SQL deleted ordinary-item recovery lost field rows or source field schema.");
                        if(node.ListItemId==2 && node.ListId==new Guid("cb9a0845-257b-40e1-8379-7468dcce5500"))
                            Check(node.Name=="bbb" && snapshot.Values.Any(value=>value.Column=="nvarchar1" && value.Value=="bbb" && !value.IsNull),"SQL second-stage ordinary item title/content was not retained.");
                        string root=Path.GetFullPath(Path.Combine(".scratch","deleted-sql-checks",Guid.NewGuid().ToString("N")));
                        ExportResult result=new DeletedItemExportService().Export(snapshot,root);
                        Check(XDocument.Load(result.Path).Root!.Attribute("DeletionTransactionId")!.Value==node.DeletionTransactionId,"SQL metadata export lost its deleted transaction.");
                        foreach(Node attachment in deleted.GetDeletedChildren(node))Check(attachment.AttachmentOwnerId==node.Id && attachment.DeletionTransactionId==node.DeletionTransactionId,"SQL retained attachment lost deleted owner identity.");
                        recoveredItems++;
                    }
                }
            }
            Console.WriteLine("PASS SQL retained deleted recovery: "+recoveredFiles+" documents and "+recoveredItems+" ordinary list items; exact deletion identity and checksums");
        }
        private sealed class Catalog : ISharePointCatalog,ISharePointDeletedCatalog
        {
            internal Node Document=new() {Kind=NodeKind.File,SiteId=Guid.NewGuid(),WebId=Guid.NewGuid(),ListId=Guid.NewGuid(),Id=Guid.NewGuid(),ParentId=Guid.NewGuid(),
                Name="doc.txt",Path="Library/doc.txt",Size=4,StreamSchema=1,HasStream=true,Level=1,InternalVersion=513,UiVersion=512,DeletionTransactionId="1234567890ABCDEF1234567890ABCDEF",DeletedAt=DateTime.UtcNow};
            internal int CurrentReads,DeletedReads;internal Func<Node,Node?>? Lookup;
            public string SourceName=>"retained deleted fixture";public void ValidateSchema(){}public List<string> CheckDatabase()=>new();
            public List<Node> GetRootSites()=>new();public List<Node> GetChildren(Node parent)=>new();
            public Node GetFile(Guid site,Guid id){CurrentReads++;throw new Exception("Deleted recovery called active catalog lookup.");}
            public IEnumerable<Node> EnumerateCurrentFiles(Guid? site)=>Array.Empty<Node>();
            public List<Node> GetDeletedItems(Node site)=>new() {DeletedIdentity.Copy(Document)};
            public List<Node> GetDeletedChildren(Node container)=>new();
            public Node GetDeletedFile(Node selected){DeletedReads++;return Lookup!=null ? Lookup(selected)! : DeletedIdentity.Copy(Document);}
            public List<Node> GetDeletedFileVersions(Node selected)=>new() {DeletedIdentity.Copy(Document)};
            public DeletedListItemSnapshot GetDeletedListItem(Node selected)=>throw new NotSupportedException();
        }
        private sealed class Store : IDocumentChunkStore
        {
            internal int Reads,LastHistory;
            public IList<StoredChunk> ReadChunks(Node document){Reads++;LastHistory=document.HistoryVersion;return new[] {new StoredChunk {BSN=1,StreamId=1,Partition=0,Type=0,Content=Encoding.UTF8.GetBytes("past")}};}
        }
        private static RecoverySession Session(Catalog catalog,Store store)=>new(catalog,store,DocumentDecoderRegistry.CreateDefault());
        private static string Digest(byte[] bytes)=>Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        private static void Check(bool condition,string message){if(!condition)throw new Exception(message);}
        private static void Reject<T>(Action action) where T:Exception {try{action();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name+".");}
    }
}