#nullable enable
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using SharePointExplorer.Desktop;

namespace SharePointExplorer.Tests
{
    internal static class AttachmentCatalogChecks
    {
        public static void Run()
        {
            string root=Path.GetFullPath(Path.Combine(".scratch","attachment-checks",Guid.NewGuid().ToString("N")));
            InputAndMapper();ScopedRecovery();
            BatchExports(Path.Combine(root,"batch"));
            OwnershipAndRuntime(Path.Combine(root,"scope"));
            CancellationAndEmpty(Path.Combine(root,"cancel"));
            Console.WriteLine("PASS current attachment ownership, strict folder mapping, mutation guards, grouped/flat batches, reports and cancellation");
        }
        public static void RunSql()
        {
            using var session=RecoverySession.OpenSql(new SqlConnectionOptions { Server = "SQL", Database = "WSS_Content" });
            var catalog=session.Catalog as ISharePointAttachmentCatalog ?? throw new Exception("SQL attachment capability is absent.");
            Guid site=new("cbd6e0be-6ee8-4b9d-9b04-d9031327831b"),web=new("f9c7aeec-64be-4f81-87e9-2a65625e7bc9"),listId=new("cb9a0845-257b-40e1-8379-7468dcce5500");
            var list=new Node {Kind=NodeKind.List,SiteId=site,WebId=web,ListId=listId,Id=new Guid("86337905-4800-49f3-9b90-2531f797552b"),ListBaseType=0,Name="test",Path="Lists/test"};
            var owner=new Node {Kind=NodeKind.ListItem,SiteId=site,WebId=web,ListId=listId,Id=new Guid("701fd014-9894-426f-8f2e-bcf01130b2a8"),
                ListItemId=1,ItemUniqueId=new Guid("1b46809a-cd66-43ef-81df-3984879fefd2"),ListBaseType=0,HasAttachments=false};
            List<Node> files=catalog.GetItemAttachments(owner);
            Check(files.Count==1,"The actual item attachment fixture changed.");
            Node pdf=files[0];
            Check(pdf.Id==new Guid("046b196d-ec40-48eb-90d6-4cba4ee06810") && pdf.Size==454216 && pdf.StreamSchema==66 &&
                pdf.Path=="Lists/test/Attachments/1/dummy-1000-chars-lorem (2).pdf" && pdf.AttachmentOwnerId==owner.Id &&
                pdf.ListItemId==1 && pdf.ItemUniqueId==owner.ItemUniqueId && pdf.ParentId==new Guid("7f4b02e5-17d6-42a5-b439-588a52fccd35"),
                "Actual attachment metadata lost its exact file/owner/folder identity.");
            Node[] all=catalog.EnumerateCurrentListAttachments(list).ToArray();
            Check(all.Length==1 && all[0].Id==pdf.Id,"Whole-list attachment enumeration selected another item or a backing file.");
            Check(catalog.GetCurrentItemAttachment(owner,owner.Id)==null && catalog.GetCurrentItemAttachment(owner,Guid.NewGuid())==null &&
                catalog.GetCurrentItemAttachment(owner,new Guid("f0675d2e-4225-4ea9-8358-da3f7da48220"))==null,
                "The attachment lookup returned a metadata backing file, nonexistent file or unrelated library document.");
            Node foreign=Copy(owner);foreign.ListItemId=2;
            Reject<ContentUnavailableException>(()=>catalog.GetItemAttachments(foreign));
            foreign=Copy(owner);foreign.ItemUniqueId=Guid.NewGuid();
            Reject<ContentUnavailableException>(()=>catalog.GetCurrentItemAttachment(foreign,pdf.Id));
            PreparedDocument prepared=session.Engine.PrepareItemAttachment(pdf,owner);
            using var bytes=new MemoryStream();RecoveryResult recovered=prepared.Recover(bytes);
            Check(recovered.Bytes==454216 && recovered.Sha256=="18132453d1f58ee4f8a73a85f346a26956bbbac40287cc039c602b5ef63970f5" &&
                Encoding.ASCII.GetString(bytes.ToArray(),0,5)=="%PDF-","Actual SQL attachment bytes disagree with independently captured/decoded PDF.");
            string root=Path.GetFullPath(Path.Combine(".scratch","attachment-sql-checks",Guid.NewGuid().ToString("N")));
            using var controller=new ExplorerController(RecoverySession.OpenSql(new SqlConnectionOptions { Server = "SQL", Database = "WSS_Content" }));
            DesktopExportSummary item=controller.ExportAttachments(owner,Path.Combine(root,"item"),CancellationToken.None,null!);
            DesktopExportSummary grouped=controller.ExportAttachments(list,Path.Combine(root,"list"),CancellationToken.None,null!);
            Check(item.Total==1 && item.Success==1 && item.Failed==0 && item.Skipped==0 && Path.GetDirectoryName(item.Entries[0].Path)==item.Directory &&
                grouped.Total==1 && grouped.Success==1 && Path.GetDirectoryName(grouped.Entries[0].Path)==Path.Combine(grouped.Directory,"1") &&
                item.Entries[0].Sha256==recovered.Sha256 && grouped.Entries[0].Sha256==recovered.Sha256 &&
                Digest(File.ReadAllBytes(item.Entries[0].Path))==recovered.Sha256,"SQL attachment item/list export paths or checksum differ.");
            Console.WriteLine("PASS SQL actual ordinary-list PDF attachment: owner/folder identity, 454216 bytes and independently verified SHA-256; item flat/list grouped: "+root);
        }
        private static void InputAndMapper()
        {
            var catalog=new Catalog();Node owner=catalog.Owner(1);
            for(int variant=0;variant<8;variant++)
            {
                Node bad=Copy(owner);
                if(variant==0)bad.Id=Guid.Empty;if(variant==1)bad.ItemUniqueId=null;if(variant==2)bad.ListItemId=0;
                if(variant==3)bad.Kind=NodeKind.File;if(variant==4)bad.HistoryVersion=512;if(variant==5)bad.WebId=Guid.Empty;
                if(variant==6)bad.ListBaseType=1;if(variant==7)bad.SiteId=Guid.Empty;
                Reject<ArgumentException>(()=>SqlAttachmentCatalog.ValidateItem(bad));
            }
            SqlAttachmentCatalog.ValidateItem(owner);SqlAttachmentCatalog.ValidateList(catalog.List);
            Node invalidList=Copy(catalog.List);invalidList.Kind=NodeKind.Library;
            Reject<ArgumentException>(()=>SqlAttachmentCatalog.ValidateList(invalidList));
            // Pure row mapping verifies exact directory and GUID identity, not a prefix match.
            Guid fileId=Guid.NewGuid(),folder=Guid.NewGuid();string directory="Lists/Notes/Attachments/1";
            using var table=new DataTable();Type[] types={typeof(Guid),typeof(Guid),typeof(Guid),typeof(Guid),typeof(string),typeof(string),typeof(byte),typeof(long),typeof(DateTime),typeof(byte),typeof(byte),typeof(int),typeof(int),typeof(Guid),typeof(int)};
            for(int index=0;index<types.Length;index++)table.Columns.Add("c"+index,types[index]);
            object[] valid={fileId,owner.SiteId,owner.WebId,owner.ListId,directory,"real.aspx",(byte)0,3L,DateTime.UtcNow,(byte)0,(byte)1,257,512,folder,1};
            table.Rows.Add(valid);
            for(int variant=0;variant<8;variant++)
            {
                object[] bad=(object[])valid.Clone();bad[14]=0;
                if(variant==0)bad[0]=owner.Id;if(variant==1)bad[1]=Guid.NewGuid();if(variant==2)bad[2]=Guid.NewGuid();
                if(variant==3)bad[3]=Guid.NewGuid();if(variant==4)bad[4]="Lists/Notes/Attachments/10";
                if(variant==5)bad[13]=Guid.NewGuid();if(variant==6)bad[6]=(byte)1;if(variant==7)bad[0]=Guid.Empty;
                table.Rows.Add(bad);
            }
            using DataTableReader rows=table.CreateDataReader();rows.Read();
            Node file=SqlAttachmentCatalog.ReadAttachment(rows,owner,folder,directory);
            Check(file.Id==fileId && file.ParentId==folder && file.AttachmentOwnerId==owner.Id && file.ItemUniqueId==owner.ItemUniqueId &&
                file.ListItemId==1 && file.Name=="real.aspx" && file.HasStream==true,"Attachment mapper changed real parent/owner metadata or filtered an extension.");
            while(rows.Read())Reject<InvalidDataException>(()=>SqlAttachmentCatalog.ReadAttachment(rows,owner,folder,directory));
        }
        private static void ScopedRecovery()
        {
            for(int variant=0;variant<11;variant++)
            {
                var catalog=new Catalog();Node owner=catalog.Owner(1);Node selected=catalog.Add(owner,"file.txt","bytes");Node changed=Copy(selected);
                if(variant==0)changed.Id=Guid.NewGuid();if(variant==1)changed.SiteId=Guid.NewGuid();if(variant==2)changed.WebId=Guid.NewGuid();
                if(variant==3)changed.ListId=Guid.NewGuid();if(variant==4)changed.AttachmentOwnerId=Guid.NewGuid();if(variant==5)changed.ListItemId=10;
                if(variant==6)changed.ItemUniqueId=Guid.NewGuid();if(variant==7)changed.Kind=NodeKind.ListItem;if(variant==8)changed.HistoryVersion=512;
                if(variant==9)changed.AttachmentOwnerId=null;
                if(variant==10)changed.DeletionTransactionId="1234567890ABCDEF1234567890ABCDEF";
                catalog.Lookup=(_,__)=>changed;var store=new Store(catalog);using var session=Session(catalog,store);
                Reject<InvalidDataException>(()=>session.Engine.PrepareItemAttachment(selected,owner));
                Check(store.Reads==0 && catalog.UnscopedReads==0,"A mismatched attachment identity reached content retrieval.");
            }
            {
                var catalog=new Catalog();Node owner=catalog.Owner(1);Node selected=catalog.Add(owner,"one.txt","one");Node other=catalog.Owner(10);
                catalog.Lookup=(request,id)=>{request.Id=other.Id;request.ListItemId=other.ListItemId;request.ItemUniqueId=other.ItemUniqueId;
                    Node changed=Copy(catalog.Current[id]);changed.AttachmentOwnerId=other.Id;changed.ListItemId=other.ListItemId;changed.ItemUniqueId=other.ItemUniqueId;return changed;};
                var store=new Store(catalog);using var session=Session(catalog,store);
                Reject<InvalidDataException>(()=>session.Engine.PrepareItemAttachment(selected,owner));
                Check(store.Reads==0,"Mutating the catalog owner request bypassed the frozen owner guard.");
            }
            {
                var catalog=new Catalog();Node owner=catalog.Owner(1);Node selected=catalog.Add(owner,"one.txt","one");Node other=catalog.Add(owner,"two.txt","two");
                catalog.Lookup=(_,__)=>{selected.Id=other.Id;return Copy(other);};
                var store=new Store(catalog);using var session=Session(catalog,store);
                Reject<InvalidDataException>(()=>session.Engine.PrepareItemAttachment(selected,owner));
                Check(store.Reads==0,"Mutating the caller selection bypassed the frozen file guard.");
            }
            {
                var catalog=new Catalog();Node owner=catalog.Owner(1);Node selected=catalog.Add(owner,"one.txt","one");var store=new Store(catalog);using var session=Session(catalog,store);
                Node bad=Copy(selected);bad.ItemUniqueId=Guid.NewGuid();
                Reject<InvalidDataException>(()=>session.Engine.PrepareItemAttachment(bad,owner));
                bad=Copy(selected);bad.HistoryVersion=512;
                Reject<NotSupportedException>(()=>session.Engine.PrepareItemAttachment(bad,owner));
                bad=Copy(selected);bad.Id=Guid.Empty;
                Reject<ArgumentException>(()=>session.Engine.PrepareItemAttachment(bad,owner));
                Check(catalog.ScopedReads==0 && store.Reads==0,"An invalid selected attachment made source reads.");
                Guid ownerId=owner.Id;PreparedDocument prepared=session.Engine.PrepareItemAttachment(selected,owner);
                selected.ListItemId=10;owner.Id=Guid.NewGuid();
                RecoveryResult recovered=prepared.Recover(Stream.Null);
                Check(recovered.Bytes==3 && recovered.Document.ListItemId==1 && prepared.Document.AttachmentOwnerId==ownerId && recovered.Document.AttachmentOwnerId==ownerId,
                    "Prepared attachment lost frozen owner metadata.");
            }
        }
        private static void BatchExports(string root)
        {
            var catalog=new Catalog();Node one=catalog.Owner(1),two=catalog.Owner(2);
            Node first=catalog.Add(one,"same.txt","first");Node collision=catalog.Add(one,"SAME.txt","second");Node another=catalog.Add(two,"same.txt","third");
            Node ghost=catalog.Add(one,"default.aspx","template");ghost.HasStream=false;
            catalog.Add(one,"unsupported.bin","native").StreamSchema=67;
            Node broken=catalog.Add(two,"broken.bin","bad");broken.StreamSchema=7;
            catalog.Current[broken.Id].StreamSchema=7;
            Node last=catalog.Add(two,"last.txt","last");catalog.Discovered.Add(Copy(first));
            var store=new Store(catalog);using var controller=new ExplorerController(Session(catalog,store));
            string grouped=Path.Combine(root,"grouped");Directory.CreateDirectory(Path.Combine(grouped,"1"));File.WriteAllText(Path.Combine(grouped,"1","same.txt"),"original");
            var progress=new List<DesktopExportProgress>();DesktopExportSummary result=controller.ExportAttachments(catalog.List,grouped,CancellationToken.None,progress.Add);
            Check(result.Total==5 && result.Success==4 && result.Failed==1 && result.Skipped==0 && result.Entries.Count==5 && store.Reads==5 && catalog.ScopedReads==5 &&
                catalog.UnscopedReads==0 && progress.Count==6 && progress[0].Document==null && progress[0].Total==5,"Attachment batching counted ignored files or failed to deduplicate/continue.");
            Check(Path.GetFileName(result.Entries[0].Path)=="same (2).txt" && Path.GetFileName(result.Entries[1].Path)=="SAME (3).txt" &&
                Path.GetDirectoryName(result.Entries[2].Path)==Path.Combine(grouped,"2") && File.ReadAllText(Path.Combine(grouped,"1","same.txt"))=="original",
                "Grouped attachment export mixed owners or overwrote a case-insensitive collision.");
            Check(result.Entries.Last().Document.Id==last.Id && File.ReadAllText(result.Entries.Last().Path)=="last" &&
                !Directory.GetFiles(grouped,"*.partial",SearchOption.AllDirectories).Any(),"Attachment failure stopped continuation or left partial bytes.");
            List<string[]> rows=Csv(File.ReadAllText(result.ReportPath));
            Check(rows.Count==6 && rows[0].Length==12 && rows.Single(row=>row[2]==last.Id.ToString("D"))[9]==two.Id.ToString("D") &&
                rows[4][8]=="bad, \"attachment\"\r\nretry","Attachment CSV lost owner fields or multiline quoted errors.");
            foreach(DesktopExportEntry entry in result.Entries.Where(entry=>entry.Status==RecoveryStatus.Success))
                Check(entry.Sha256==Digest(File.ReadAllBytes(entry.Path)) && entry.Bytes==catalog.Content[entry.Document.Id].Length,"Attachment checksum or byte report differs from output.");
            DesktopExportSummary item=controller.ExportAttachments(one,Path.Combine(root,"item"),CancellationToken.None,null!);
            Check(item.Total==2 && item.Success==2 && item.Entries.All(entry=>Path.GetDirectoryName(entry.Path)==item.Directory),"Individual-item attachments were not saved directly in the chosen folder.");
            // The same prepared current snapshot supplies filename, report and bytes.
            catalog.Current[first.Id].Name="new.txt";catalog.Current[first.Id].Path="Lists/Notes/Attachments/1/new.txt";
            item=controller.ExportAttachments(one,Path.Combine(root,"renamed"),CancellationToken.None,null!);
            Check(item.Entries[0].Document.Name=="new.txt" && Path.GetFileName(item.Entries[0].Path)=="new.txt" &&
                Csv(File.ReadAllText(item.ReportPath))[1][3].EndsWith("/new.txt"),"Attachment filename/report retained stale discovery metadata.");
        }
        private static void OwnershipAndRuntime(string root)
        {
            for(int variant=0;variant<5;variant++)
            {
                var catalog=new Catalog();Node owner=catalog.Owner(1);Node file=catalog.Add(owner,"bad.txt","bad");file.HasStream=false;
                if(variant==0)file.SiteId=Guid.NewGuid();if(variant==1)file.ListId=Guid.NewGuid();if(variant==2)file.WebId=Guid.NewGuid();
                if(variant==3)file.AttachmentOwnerId=Guid.NewGuid();if(variant==4)file.ItemUniqueId=Guid.NewGuid();
                var store=new Store(catalog);using var controller=new ExplorerController(Session(catalog,store));
                string destination=Path.Combine(root,"scope-"+variant);
                Reject<InvalidDataException>(()=>controller.ExportAttachments(owner,destination,CancellationToken.None,null!));
                Check(store.Reads==0 && !Directory.Exists(destination),"Quiet eligibility concealed an out-of-scope attachment.");
            }
            {
                var catalog=new Catalog();Node owner=catalog.Owner(1),other=catalog.Owner(10);Node file=catalog.Add(owner,"same.txt","same");
                Node ambiguous=Copy(file);ambiguous.AttachmentOwnerId=other.Id;ambiguous.ListItemId=10;ambiguous.ItemUniqueId=other.ItemUniqueId;ambiguous.HasStream=false;
                catalog.Discovered.Add(ambiguous);var store=new Store(catalog);using var controller=new ExplorerController(Session(catalog,store));
                string destination=Path.Combine(root,"ambiguous");Reject<InvalidDataException>(()=>controller.ExportAttachments(catalog.List,destination,CancellationToken.None,null!));
                Check(store.Reads==0 && !Directory.Exists(destination),"Duplicate file identities with different owners were silently merged.");
            }
            for(int variant=0;variant<2;variant++)
            {
                var catalog=new Catalog();Node owner=catalog.Owner(1);Node file=catalog.Add(owner,"changed.aspx","stored");
                if(variant==0)catalog.Current[file.Id].HasStream=false;else catalog.Current[file.Id].StreamSchema=67;
                var store=new Store(catalog);using var controller=new ExplorerController(Session(catalog,store));
                DesktopExportSummary result=controller.ExportAttachments(owner,Path.Combine(root,"runtime-"+variant),CancellationToken.None,null!);
                Check(result.Total==1 && result.Skipped==1 && result.Failed==0 && Csv(File.ReadAllText(result.ReportPath)).Count==2,
                    "Freshly unavailable/unsupported attachment lost its runtime audit outcome.");
            }
        }
        private static void CancellationAndEmpty(string root)
        {
            var catalog=new Catalog();Node owner=catalog.Owner(1);catalog.Add(owner,"one.txt","one");catalog.Add(owner,"two.txt","two");
            var store=new Store(catalog);using var controller=new ExplorerController(Session(catalog,store));
            using var pre=new CancellationTokenSource();pre.Cancel();string destination=Path.Combine(root,"pre");
            DesktopExportSummary result=controller.ExportAttachments(owner,destination,pre.Token,null!);
            Check(result.Cancelled && catalog.Discoveries==0 && store.Reads==0 && !Directory.Exists(destination),"Pre-cancelled attachments accessed source/output.");
            using var discovery=new CancellationTokenSource();catalog.OnDiscovery=()=>discovery.Cancel();destination=Path.Combine(root,"discovery");
            result=controller.ExportAttachments(catalog.List,destination,discovery.Token,null!);
            Check(result.Cancelled && store.Reads==0 && !Directory.Exists(destination),"Discovery cancellation exported a partial attachment selection.");
            catalog.OnDiscovery=null;using var between=new CancellationTokenSource();
            result=controller.ExportAttachments(owner,Path.Combine(root,"between"),between.Token,item=>{if(item.Completed==1)between.Cancel();});
            Check(result.Cancelled && result.Success==1 && result.Total==2 && store.Reads==1 && Csv(File.ReadAllText(result.ReportPath)).Count==2,
                "Between-file attachment cancellation read the next file or lost the completed report.");
            controller.Dispose();Reject<ObjectDisposedException>(()=>controller.ExportAttachments(owner,Path.Combine(root,"disposed"),CancellationToken.None,null!));
            var emptyCatalog=new Catalog();Node emptyOwner=emptyCatalog.Owner(1);Node ignored=emptyCatalog.Add(emptyOwner,"default.aspx","template");ignored.HasStream=false;
            var emptyStore=new Store(emptyCatalog);using var empty=new ExplorerController(Session(emptyCatalog,emptyStore));
            result=empty.ExportAttachments(emptyOwner,Path.Combine(root,"ignored"),CancellationToken.None,null!);
            Check(result.Total==0 && result.Success==0 && result.Skipped==0 && emptyStore.Reads==0 && emptyCatalog.ScopedReads==0 &&
                Csv(File.ReadAllText(result.ReportPath)).Count==1,"Known unavailable attachments created warnings, reads or audit outcomes.");
        }
        private static RecoverySession Session(Catalog catalog,Store store)=>new(catalog,store,new DocumentDecoderRegistry(new IDocumentDecoder[] {DocumentDecoderRegistry.CreateDefault().Resolve(0),new BrokenDecoder()}));
        private static Node Copy(Node source)=>new() {Kind=source.Kind,SiteId=source.SiteId,WebId=source.WebId,ListId=source.ListId,Id=source.Id,ParentId=source.ParentId,
            Name=source.Name,Path=source.Path,Size=source.Size,StreamSchema=source.StreamSchema,HistoryVersion=source.HistoryVersion,Level=source.Level,UiVersion=source.UiVersion,
            InternalVersion=source.InternalVersion,HasStream=source.HasStream,ListBaseType=source.ListBaseType,ListItemId=source.ListItemId,ItemUniqueId=source.ItemUniqueId,
            AttachmentOwnerId=source.AttachmentOwnerId,HasAttachments=source.HasAttachments};
        private static string Digest(byte[] bytes)=>Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        private static void Check(bool value,string message){if(!value)throw new Exception(message);}
        private static void Reject<T>(Action action) where T:Exception{try{action();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name);}
        private static List<string[]> Csv(string text)
        {
            var result=new List<string[]>();var row=new List<string>();var field=new StringBuilder();bool quoted=false;
            for(int index=0;index<text.Length;index++)
            {
                char c=text[index];
                if(c=='"'){if(quoted && index+1<text.Length && text[index+1]=='"'){field.Append('"');index++;}else quoted=!quoted;}
                else if(!quoted && c==','){row.Add(field.ToString());field.Clear();}
                else if(!quoted && (c=='\r'||c=='\n')){row.Add(field.ToString());field.Clear();result.Add(row.ToArray());row.Clear();if(c=='\r' && index+1<text.Length && text[index+1]=='\n')index++;}
                else field.Append(c);
            }
            Check(!quoted,"CSV contains unterminated quotes.");return result;
        }
        private sealed class Catalog : ISharePointCatalog,ISharePointAttachmentCatalog
        {
            internal readonly Node List=new() {Kind=NodeKind.List,SiteId=Guid.NewGuid(),WebId=Guid.NewGuid(),ListId=Guid.NewGuid(),Id=Guid.NewGuid(),ListBaseType=0,Name="Notes",Path="Lists/Notes"};
            internal readonly List<Node> Discovered=new();internal readonly Dictionary<Guid,Node> Current=new();internal readonly Dictionary<Guid,byte[]> Content=new();
            internal readonly Dictionary<int,Node> Owners=new();internal Func<Node,Guid,Node?>? Lookup;internal Action? OnDiscovery;internal int Discoveries,ScopedReads,UnscopedReads;
            public string SourceName=>"Synthetic, \"attachments\"";
            internal Node Owner(int id)
            {
                if(Owners.TryGetValue(id,out Node? owner))return owner;
                owner=new Node {Kind=NodeKind.ListItem,SiteId=List.SiteId,WebId=List.WebId,ListId=List.ListId,Id=Guid.NewGuid(),ParentId=List.Id,ListItemId=id,ItemUniqueId=Guid.NewGuid(),ListBaseType=0,HasAttachments=true};
                Owners.Add(id,owner);return owner;
            }
            internal Node Add(Node owner,string name,string text)
            {
                byte[] bytes=Encoding.UTF8.GetBytes(text);var file=new Node {Kind=NodeKind.File,SiteId=List.SiteId,WebId=List.WebId,ListId=List.ListId,Id=Guid.NewGuid(),ParentId=Guid.NewGuid(),
                    AttachmentOwnerId=owner.Id,ListItemId=owner.ListItemId,ItemUniqueId=owner.ItemUniqueId,ListBaseType=0,Name=name,Path=List.Path+"/Attachments/"+owner.ListItemId+"/"+name,Size=bytes.Length,HasStream=true};
                Discovered.Add(file);Current.Add(file.Id,Copy(file));Content.Add(file.Id,bytes);return file;
            }
            public List<Node> GetItemAttachments(Node item){Discoveries++;return Discovered.Where(file=>file.AttachmentOwnerId==item.Id || file.HasStream==false).Select(Copy).ToList();}
            public IEnumerable<Node> EnumerateCurrentListAttachments(Node list){Discoveries++;foreach(Node file in Discovered){OnDiscovery?.Invoke();yield return Copy(file);}}
            public Node GetCurrentItemAttachment(Node owner,Guid fileId)
            {
                ScopedReads++;if(Lookup!=null)return Lookup(owner,fileId)!;
                return Current.TryGetValue(fileId,out Node? file) && file.SiteId==owner.SiteId && file.WebId==owner.WebId && file.ListId==owner.ListId &&
                    file.AttachmentOwnerId==owner.Id && file.ListItemId==owner.ListItemId && file.ItemUniqueId==owner.ItemUniqueId ? Copy(file) : null!;
            }
            public Node GetFile(Guid siteId,Guid fileId){UnscopedReads++;throw new Exception("Attachment recovery used unscoped current-file lookup.");}
            public void ValidateSchema(){}public List<string> CheckDatabase()=>new();public List<Node> GetRootSites()=>new();public List<Node> GetChildren(Node node)=>new();public IEnumerable<Node> EnumerateCurrentFiles(Guid? site)=>Array.Empty<Node>();
        }
        private sealed class Store : IDocumentChunkStore
        {
            private readonly Catalog catalog;internal int Reads;internal Store(Catalog catalog){this.catalog=catalog;}
            public IList<StoredChunk> ReadChunks(Node file){Reads++;if(file.HasStream==false)throw new ContentUnavailableException("Template content unavailable.");return new List<StoredChunk> {new() {Partition=0,Content=catalog.Content[file.Id]}};}
        }
        private sealed class BrokenDecoder : IDocumentDecoder
        {
            public string Name=>"Synthetic corrupt attachment";public bool Supports(byte schema)=>schema==7;
            public void Write(IList<StoredChunk> chunks,long expected,Stream output){output.WriteByte(1);throw new InvalidDataException("bad, \"attachment\"\r\nretry");}
        }
    }
}