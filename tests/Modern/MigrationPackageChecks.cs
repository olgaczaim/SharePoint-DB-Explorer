using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Xml.Linq;
using SharePointExplorer;

namespace SharePointExplorer.Tests
{
    internal static class MigrationPackageChecks
    {
        private static readonly XNamespace M = "urn:deployment-manifest-schema";
        public static void Run()
        {
            string root = Path.GetFullPath(Path.Combine(".scratch", "migration-package-checks", Guid.NewGuid().ToString("N")));
            TestLibrary(Path.Combine(root, "library"));
            TestGenerations(Path.Combine(root, "generations"));
            TestHistory(Path.Combine(root, "history"));
            TestHistoricalAttachments(Path.Combine(root, "historical-attachments"));
            TestAttachments(Path.Combine(root, "attachments"));
            TestContentTypes(Path.Combine(root, "content-types"));
            TestListSettings(Path.Combine(root, "list-settings"));
            TestSystemContext(Path.Combine(root, "system-context"));
            TestFailures(Path.Combine(root, "failures"));
            TestDeletedPackageScopes(Path.Combine(root, "deleted-scopes"));
            TestMetadataMutation(Path.Combine(root, "metadata-race"));
            TestCancellation(Path.Combine(root, "cancel"));
            Console.WriteLine("PASS deployment XML schemas, scoped files/items/attachments, exact payload provenance, tamper rejection and atomic cancellation");
        }
        // Each supported source generation receives its content-deployment
        // compatibility profile; only 2016/2019 declare the extra export settings.
        private static void TestGenerations(string root)
        {
            var cases = new[] { ("16.0.4351.1000", "38455", true), ("16.0.10337.12109", "12710", true), ("16.0.17928.20238", "7123", false) };
            foreach (var (build, databaseVersion, serverSettings) in cases)
            {
                var catalog = new Catalog(true); catalog.AddFile("ok.txt", catalog.List.Id, "Documents/ok.txt", "ok"); catalog.Build = build;
                using (RecoverySession session = catalog.Session())
                {
                    MigrationPackageResult result = new MigrationPackageExporter().Export(session, catalog.List, Path.Combine(root, databaseVersion), CancellationToken.None, null);
                    MigrationPackageExporter.ValidatePackage(result.PackagePath);
                    XElement version = XDocument.Load(Path.Combine(result.PackagePath, "SystemData.xml")).Descendants().Single(e => e.Name.LocalName == "SchemaVersion");
                    Require((string)version.Attribute("Build") == build && (string)version.Attribute("DatabaseVersion") == databaseVersion, "Source generation " + build + " received the wrong compatibility profile.");
                    XElement settings = XDocument.Load(Path.Combine(result.PackagePath, "ExportSettings.xml")).Root;
                    Require((settings.Attribute("IncludeAlerts") != null) == serverSettings && (settings.Attribute("AzureContainerSourceUri") != null) == serverSettings &&
                        (settings.Attribute("AzureContainerManifestUri") != null) == serverSettings, "Generation-specific export settings are wrong for " + build + ".");
                }
            }
        }
        private static void TestLibrary(string root)
        {
            var catalog = new Catalog(true);
            Node first = catalog.AddFile("alpha.txt", catalog.List.Id, "Documents/alpha.txt", "alpha<&\n");
            Node folder = catalog.AddFolder("nested");
            Node second = catalog.AddFile("beta.txt", folder.Id, "Documents/nested/beta.txt", "distinct beta payload");
            catalog.Title = new string('L', 255);
            using (RecoverySession session = catalog.Session())
            {
                var progress = new List<MigrationPackageProgress>();
                MigrationPackageResult result = new MigrationPackageExporter().Export(session, catalog.List, root, CancellationToken.None, p => progress.Add(p));
                Require(result.FileCount == 2 && result.ItemCount == 3 && result.AttachmentCount == 0 && !result.Cancelled, "Library package counts changed.");
                Require(Path.GetFileName(result.PackagePath).Length < 180 && Directory.GetDirectories(root, "*.staging").Length == 0, "Long source title broke publication or leaked staging.");
                Require(progress.Count == 2 && progress[1].Total == 2 && progress[1].Completed == 2 && catalog.ScopedReads == 2 && catalog.UnscopedReads == 0, "Package recovery ignored its scoped snapshot.");
                MigrationPackageExporter.ValidatePackage(result.PackagePath);
                XDocument manifest = XDocument.Load(Path.Combine(result.PackagePath, "Manifest.xml"));
                Require((string)manifest.Descendants(M + "DocumentLibrary").Single().Attribute("Title") == catalog.Title, "Folder-name truncation modified source metadata.");
                XElement item = manifest.Descendants(M + "ListItem").Single(e => (string)e.Attribute("DocId") == first.Id.ToString("D"));
                Require((string)item.Element(M + "Fields").Element(M + "Field").Attribute("Value") == "Title <& alpha.txt", "Typed field text was not XML-preserved.");
                foreach (Node file in new[] { first, second }) CheckPayload(manifest, result.PackagePath, file, catalog.Content[file.Id]);
                Require(File.ReadAllText(result.ReportPath).Contains(Hash(catalog.Content[second.Id])) && File.ReadAllText(result.AuditPath).Contains("\"FarmImportVerified\": false"), "Package recovery audit overstates verification or omits hashes.");
                Require(!File.ReadAllText(Path.Combine(result.PackagePath, "IMPORT.txt")).Contains("IncludeUserSecurity"), "Package import instructions claimed absent ACLs.");
                MigrationPackageResult again = new MigrationPackageExporter().Export(session, catalog.List, root, CancellationToken.None, null);
                Require(again.PackagePath != result.PackagePath && File.Exists(Path.Combine(result.PackagePath, "00000001.dat")), "Second package export overwrote existing output.");
                // Two valid payloads cannot be assigned to the wrong documents.
                XElement[] files = manifest.Descendants(M + "File").ToArray();
                string original = (string)files[0].Attribute("FileValue");
                files[0].SetAttributeValue("FileValue", (string)files[1].Attribute("FileValue")); files[1].SetAttributeValue("FileValue", original);
                manifest.Save(Path.Combine(result.PackagePath, "Manifest.xml"));
                Reject<InvalidDataException>(() => MigrationPackageExporter.ValidatePackage(result.PackagePath));
                manifest = XDocument.Load(Path.Combine(again.PackagePath, "Manifest.xml"));
                manifest.Descendants(M + "File").First().SetAttributeValue("FileValue", "../evil.dat");
                manifest.Save(Path.Combine(again.PackagePath, "Manifest.xml"));
                Reject<InvalidDataException>(() => MigrationPackageExporter.ValidatePackage(again.PackagePath));
            }
        }
        private static void TestHistoricalAttachments(string root)
        {
            var catalog=new Catalog(false);Node current=catalog.AddItem("1_.000",true);current.UiVersion=1024;
            catalog.AddAttachment(current,"current.pdf",Encoding.ASCII.GetBytes("%PDF-current-only-attachment"));
            MigrationItemSnapshot currentItem=catalog.ReadMigrationList(catalog.List).Items.Single();Node old=MigrationSnapshotCopy.Node(current);
            old.UiVersion=old.HistoryVersion=512;old.InternalVersion=513;old.Modified=old.Modified.AddDays(-1);
            var olderItem=new MigrationItemSnapshot(old,currentItem.ItemId,currentItem.ItemUniqueId,1,1,currentItem.Created,old.Modified,true,"0x01",currentItem.Values);
            catalog.History=new MigrationHistorySnapshot(new[]{olderItem},Array.Empty<MigrationUserSnapshot>(),Array.Empty<MigrationHistoryFile>());
            using(RecoverySession session=catalog.Session())
            {
                MigrationPackageResult result=new MigrationPackageExporter{IncludeHistory=true}.Export(session,catalog.List,root,CancellationToken.None,null);
                MigrationPackageExporter.ValidatePackage(result.PackagePath);XDocument manifest=XDocument.Load(Path.Combine(result.PackagePath,"Manifest.xml"));
                XElement item=manifest.Descendants(M+"ListItem").Single(entry=>!entry.Ancestors(M+"Versions").Any());
                Require(item.Element(M+"Attachments").Elements().Count()==1 && item.Element(M+"Versions").Elements().All(version=>version.Element(M+"Attachments")==null),"All-version metadata assigned current attachments to an unverified historical attachment set.");
                Require(File.ReadAllText(result.AuditPath).Contains("historical attachment sets"),"Package audit omitted historical attachment coverage limits.");
            }
        }
        private static void TestHistory(string root)
        {
            var catalog=new Catalog(true);
            Node current=catalog.AddFile("versioned.txt",catalog.List.Id,"Documents/versioned.txt","current bytes");
            current.UiVersion=1024;current.InternalVersion=2;catalog.Current[current.Id]=MigrationSnapshotCopy.Node(current);
            MigrationListSnapshot snapshot=catalog.ReadMigrationList(catalog.List);
            MigrationItemSnapshot owner=snapshot.Items.Single();
            Node old=MigrationSnapshotCopy.Node(current);old.UiVersion=old.HistoryVersion=512;old.InternalVersion=513;old.Size=Encoding.UTF8.GetByteCount("independent old bytes");old.Modified=old.Modified.AddDays(-1);
            catalog.VersionContent[old.Id.ToString("N")+":"+old.HistoryVersion]=Encoding.UTF8.GetBytes("independent old bytes");
            var historical=new MigrationItemSnapshot(old,owner.ItemId,owner.ItemUniqueId,2,2,owner.Created,old.Modified,false,"0x0101",
                new[]{new MigrationFieldValue("Title",snapshot.Fields.First().Id,1,"nvarchar","Historical title"),new MigrationFieldValue("ContentTypeId",null,1,"varchar","0x0101")});
            catalog.History=new MigrationHistorySnapshot(new[]{historical},new[]{new MigrationUserSnapshot(2,@"CONTOSO\historical","Historical author","history@example.test",false,false,0,new byte[]{4,5,6})},new[]{new MigrationHistoryFile(old,"original approval")});
            using(RecoverySession session=catalog.Session())
            {
                MigrationPackageResult result=new MigrationPackageExporter {IncludeHistory=true}.Export(session,catalog.List,root,CancellationToken.None,null);
                MigrationPackageExporter.ValidatePackage(result.PackagePath);
                XDocument manifest=XDocument.Load(Path.Combine(result.PackagePath,"Manifest.xml"));XDocument settings=XDocument.Load(Path.Combine(result.PackagePath,"ExportSettings.xml"));
                Require((string)settings.Root.Attribute("IncludeVersions")=="All","Retained-history package still claims only current versions.");
                XElement file=manifest.Descendants(M+"File").Single(e=>!e.Ancestors(M+"Versions").Any() && (string)e.Attribute("Id")==old.Id.ToString("D"));
                XElement oldFile=file.Element(M+"Versions").Elements(M+"File").Single(version=>(string)version.Attribute("Version")=="1.0");
                XElement currentFile=file.Element(M+"Versions").Elements(M+"File").Single(version=>(string)version.Attribute("Version")=="2.0");
                Require(file.Attribute("FileValue")==null && File.ReadAllText(Path.Combine(result.PackagePath,(string)currentFile.Attribute("FileValue")))=="current bytes","All-version layout lacks its nested current payload or repeats it at the root.");
                Require((string)oldFile.Attribute("Version")=="1.0" && (string)oldFile.Attribute("CheckinComment")=="original approval" && (string)oldFile.Attribute("Author")=="2","Historical file attribution/version/comment were replaced by current metadata.");
                Require(File.ReadAllBytes(Path.Combine(result.PackagePath,(string)oldFile.Attribute("FileValue"))).SequenceEqual(catalog.VersionContent[old.Id.ToString("N")+":"+old.HistoryVersion]),"XML history substituted current document bytes.");
                XElement oldItem=manifest.Descendants(M+"ListItem").Single(e=>e.Ancestors(M+"Versions").Any() && (string)e.Attribute("Version")=="1.0");
                Require((string)oldItem.Element(M+"Fields").Elements(M+"Field").Single(e=>(string)e.Attribute("Name")=="Title").Attribute("Value")=="Historical title" && (string)oldItem.Attribute("ModifiedBy")=="2","Historical item fields/editors were not retained.");
                Require(XDocument.Load(Path.Combine(result.PackagePath,"UserGroup.xml")).Descendants().Any(e=>e.Name.LocalName=="User" && (string)e.Attribute("Id")=="2"),"Historical-only author is absent from the user map.");
                string first=(string)currentFile.Attribute("FileValue");currentFile.SetAttributeValue("FileValue",(string)oldFile.Attribute("FileValue"));oldFile.SetAttributeValue("FileValue",first);manifest.Save(Path.Combine(result.PackagePath,"Manifest.xml"));
                Reject<InvalidDataException>(()=>MigrationPackageExporter.ValidatePackage(result.PackagePath));
                MigrationPackageResult currentOnly=new MigrationPackageExporter().Export(session,catalog.List,Path.Combine(root,"current-only"),CancellationToken.None,null);
                Require(!XDocument.Load(Path.Combine(currentOnly.PackagePath,"Manifest.xml")).Descendants(M+"Versions").Any(),"Default export changed to historical coverage.");
                int reads=0;catalog.BeforeHistory=()=>{if(++reads==2)catalog.History=MigrationHistorySnapshot.Empty;};
                string changed=Path.Combine(root,"changed");Reject<InvalidDataException>(()=>new MigrationPackageExporter {IncludeHistory=true}.Export(session,catalog.List,changed,CancellationToken.None,null));NoOutputs(changed);catalog.BeforeHistory=null;
            }
        }        private static void TestAttachments(string root)
        {
            var catalog = new Catalog(false); Node owner = catalog.AddItem("1_.000", true);
            Node attachment = catalog.AddAttachment(owner, "report.pdf", Encoding.ASCII.GetBytes("%PDF-native-fixture-payload"));
            catalog.AddTemplateView();
            using (RecoverySession session = catalog.Session())
            {
                MigrationPackageResult result = new MigrationPackageExporter().Export(session, catalog.List, root, CancellationToken.None, null);
                Require(result.ItemCount == 1 && result.FileCount == 0 && result.AttachmentCount == 1 && result.Bytes == attachment.Size, "Ordinary list attachment package counts changed.");
                XDocument manifest = XDocument.Load(Path.Combine(result.PackagePath, "Manifest.xml"));
                XElement entry = manifest.Descendants(M + "Attachment").Single();
                Require((string)entry.Parent.Parent.Attribute("DocId") == owner.Id.ToString("D") && (string)entry.Attribute("Id") == attachment.Id.ToString("D"), "Attachment was not nested in its authoritative owner.");
                CheckPayload(manifest, result.PackagePath, attachment, catalog.Content[attachment.Id]);
                Require((string)entry.Attribute("TimeCreated") == "2024-01-02T03:04:05Z" && (string)entry.Attribute("Author") == "1" && (string)entry.Attribute("ModifiedBy") == "1", "Attachment creation or authoritative owner attribution was omitted.");
                Require(catalog.AttachmentReads == 1 && catalog.UnscopedReads == 0 && manifest.Descendants(M + "File").All(e => (string)e.Attribute("IsGhosted") == "true"), "List backing .000 records were treated as payload files.");
                XElement view = manifest.Descendants(M + "View").Single();
                Require((string)view.Attribute("DefaultView") == "true" && view.Element(M + "ViewFields").Elements().Count() == 1, "The stored view flags/page type/field fragment were lost.");
                Require(manifest.Descendants(M + "WebPart").Single().Attribute("Name").Value == view.Attribute("Name").Value, "Ghost view file lost its matching Web Part.");
                byte[] data = File.ReadAllBytes(Path.Combine(result.PackagePath, "00000001.dat")); data[0] ^= 1;
                File.WriteAllBytes(Path.Combine(result.PackagePath, "00000001.dat"), data);
                Reject<InvalidDataException>(() => MigrationPackageExporter.ValidatePackage(result.PackagePath));
            }
        }
        private static void TestContentTypes(string root)
        {
            foreach (bool library in new[] { false, true })
            {
                var catalog = new Catalog(library);
                if (library) catalog.AddFile("one.txt", catalog.List.Id, "Documents/one.txt", "known source bytes");
                else catalog.AddItem("1_.000", false);
                string id = library ? "0x0101" : "0x01";
                catalog.ContentTypesOverride = "<ContentTypes><ContentType ID='" + id + "' Name='" + (library ? "$Resources:core,Document;" : "$Resources:core,Item;") +
                    "' Version='7' DelayActivateTemplateBinding='GROUP,SPSPERS,SITEPAGEPUBLISHING' FeatureId='{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}'/></ContentTypes>";
                string destination = Path.Combine(root, library ? "library" : "list");
                using (RecoverySession session = catalog.Session())
                {
                    MigrationPackageResult result = new MigrationPackageExporter().Export(session, catalog.List, destination, CancellationToken.None, null);
                    XElement contentType = XDocument.Load(Path.Combine(result.PackagePath, "Manifest.xml")).Descendants(M + "ContentType").Single();
                    Require(contentType.Attribute("DelayActivateTemplateBinding") == null && (string)contentType.Attribute("Version") == "7", "Bounded CAML conversion removed source content-type behavior beyond the one incompatible marker.");
                    using (JsonDocument audit = JsonDocument.Parse(File.ReadAllText(result.AuditPath)))
                    {
                        Require(audit.RootElement.GetProperty("SourceContentTypesXml").GetString() == catalog.ContentTypesOverride &&
                            audit.RootElement.GetProperty("CamlConversions").GetArrayLength() == 1, "Original content-type XML or exact conversion evidence was lost.");
                    }
                    catalog.ContentTypesOverride = catalog.ContentTypesOverride.Replace("GROUP,SPSPERS,SITEPAGEPUBLISHING", "UNKNOWN_TEMPLATE");
                    string rejected = Path.Combine(root, "unknown-value-" + library);
                    Reject<NotSupportedException>(() => new MigrationPackageExporter().Export(session, catalog.List, rejected, CancellationToken.None, null)); NoOutputs(rejected);
                    catalog.ContentTypesOverride = catalog.ContentTypesOverride.Replace("UNKNOWN_TEMPLATE", "GROUP,SPSPERS,SITEPAGEPUBLISHING").Replace(id + "'", "0x010200" + Guid.NewGuid().ToString("N") + "'");
                    rejected = Path.Combine(root, "unknown-content-type-" + library);
                    Reject<NotSupportedException>(() => new MigrationPackageExporter().Export(session, catalog.List, rejected, CancellationToken.None, null)); NoOutputs(rejected);
                }
            }
        }
        private static void TestListSettings(string root)
        {
            var catalog = new Catalog(true);
            catalog.AddFile("one.txt", catalog.List.Id, "Documents/one.txt", "settings payload");
            catalog.Flags = unchecked((long)(0x8UL | 0x80UL | 0x4000UL | 0x40000UL | 0x80000UL | 0x200000UL | 0x400000UL | 0x20000000UL |
                0x800000000UL | 0x4000000000UL | 0x8000000000000UL | 0x20000000000000UL | 0x40000000000000UL | 0x80000000000000UL | 0x100000000000000UL));
            catalog.Flags2 = 4096; catalog.TemplateId = Guid.NewGuid();
            using (RecoverySession session = catalog.Session())
            {
                MigrationPackageResult result = new MigrationPackageExporter().Export(session, catalog.List, root, CancellationToken.None, null);
                XDocument manifest = XDocument.Load(Path.Combine(result.PackagePath, "Manifest.xml"));
                XElement list = manifest.Descendants(M + "DocumentLibrary").Single();
                Require((string)list.Attribute("EnableVersioning") == "true" && (string)list.Attribute("EnableMinorVersions") == "true" &&
                    (string)list.Attribute("ForceCheckout") == "true" && (string)list.Attribute("DraftVersionVisibility") == "Approver", "Versioning/checkout/draft settings fell back to target defaults.");
                Require((string)list.Attribute("EnableAttachments") == "false" && (string)list.Attribute("EnableFolderCreation") == "false" &&
                    (string)list.Attribute("EnableContentTypes") == "true" && (string)list.Attribute("NoCrawl") == "true" && (string)list.Attribute("EnableSyndication") == "false" &&
                    (string)list.Attribute("BrowserFileHandling") == "Strict" && (string)list.Attribute("DisableGridEditing") == "true" &&
                    (string)list.Attribute("NavigateForFormsPages") == "true" && (string)list.Attribute("StrictTypeCoercion") == "true", "Source list interaction/content settings were lost.");
                Require((string)list.Attribute("DocumentTemplateId") == catalog.TemplateId.Value.ToString("D"), "Stored document template identity was omitted.");
                using (JsonDocument audit = JsonDocument.Parse(File.ReadAllText(result.AuditPath)))
                    Require(audit.RootElement.GetProperty("SourceListFlags2").GetUInt64() == 4096, "Extended source flags were lost from the audit.");
                XDocument system = XDocument.Load(Path.Combine(result.PackagePath, "SystemData.xml"));
                XElement rootOnly = system.Root.Elements().Single(e => e.Name.LocalName == "RootWebOnlyLists");
                Require((string)rootOnly.Elements().Single().Attribute("Id") == catalog.List.ListId.ToString("D"), "Root-web-only restriction used a folder identity instead of the list identity.");
                rootOnly.Elements().Single().SetAttributeValue("Id", catalog.List.Id.ToString("D")); system.Save(Path.Combine(result.PackagePath, "SystemData.xml"));
                Reject<InvalidDataException>(() => MigrationPackageExporter.ValidatePackage(result.PackagePath));
                catalog.Flags = 8;
                result = new MigrationPackageExporter().Export(session, catalog.List, root, CancellationToken.None, null);
                list = XDocument.Load(Path.Combine(result.PackagePath, "Manifest.xml")).Descendants(M + "DocumentLibrary").Single();
                Require((string)list.Attribute("EnableVersioning") == "false" && (string)list.Attribute("EnableMinorVersions") == "false" &&
                    (string)list.Attribute("ForceCheckout") == "false" && (string)list.Attribute("DraftVersionVisibility") == "Reader", "Disabled source versioning settings were invented.");
                system = XDocument.Load(Path.Combine(result.PackagePath, "SystemData.xml"));
                Require(!system.Descendants().Any(e => e.Name.LocalName == "RootWebOnlyLists"), "A root-web restriction was invented for an ordinary list.");
                list.SetAttributeValue("EnableVersioning", "true"); list.Document.Save(Path.Combine(result.PackagePath, "Manifest.xml"));
                Reject<InvalidDataException>(() => MigrationPackageExporter.ValidatePackage(result.PackagePath));
                foreach (long unsupported in new[] { 0x400L, 0x100000000000L, 0x4000000000000L })
                {
                    catalog.Flags = unsupported;
                    string policy = Path.Combine(root, "policy-" + unsupported);
                    Reject<NotSupportedException>(() => new MigrationPackageExporter().Export(session, catalog.List, policy, CancellationToken.None, null)); NoOutputs(policy);
                }
                using (var recovered = new MemoryStream())
                {
                    Node current = catalog.Items.Single(item => item.Kind == NodeKind.File);
                    session.Engine.PrepareLibraryFile(current, catalog.List).Recover(recovered);
                    Require(recovered.ToArray().SequenceEqual(catalog.Content[current.Id]), "Unsupported package policy blocked normal raw file recovery.");
                }
                catalog.Flags = 0x100000000L;
                string blocked = Path.Combine(root, "blocked");
                Reject<NotSupportedException>(() => new MigrationPackageExporter().Export(session, catalog.List, blocked, CancellationToken.None, null)); NoOutputs(blocked);
            }
        }
        private static void TestSystemContext(string root)
        {
            var catalog = new Catalog(false);
            catalog.StoredSiteUrl = "sites/source"; catalog.StoredWebUrl = "sites/source/subsite"; catalog.RootWebId = Guid.NewGuid();
            catalog.List.Path = "sites/source/subsite/Lists/Items"; catalog.AddItem("1_.000", false);
            using (RecoverySession session = catalog.Session())
            {
                MigrationPackageResult result = new MigrationPackageExporter().Export(session, catalog.List, root, CancellationToken.None, null);
                XDocument system = XDocument.Load(Path.Combine(result.PackagePath, "SystemData.xml"));
                XElement[] contexts = system.Descendants().Where(element => element.Name.LocalName == "SystemObject").ToArray();
                Require(contexts.Length == 4 && contexts.Any(element => (string)element.Attribute("Id") == catalog.RootWebId.ToString("D") && (string)element.Attribute("Url") == "/sites/source") &&
                    contexts.Any(element => (string)element.Attribute("Id") == catalog.List.WebId.ToString("D") && (string)element.Attribute("Url") == "/sites/source/subsite") &&
                    contexts.Any(element => (string)element.Attribute("Id") == catalog.UserInfoListId.ToString("D") && (string)element.Attribute("Type") == "List" && (string)element.Attribute("Url") == "/sites/source/_catalogs/users"), "Subsite/root-web/user-info context was omitted or substituted.");
                XElement users = contexts.Single(element => (string)element.Attribute("Id") == catalog.UserInfoListId.ToString("D"));
                users.SetAttributeValue("Url", "/sites/source/Lists/unrelated"); system.Save(Path.Combine(result.PackagePath, "SystemData.xml"));
                Reject<InvalidDataException>(() => MigrationPackageExporter.ValidatePackage(result.PackagePath));
                MigrationSystemObject[] recorded = catalog.ReadMigrationList(catalog.List).SystemObjects.ToArray();
                catalog.SystemObjectsOverride = Array.Empty<MigrationSystemObject>();
                string empty = Path.Combine(root, "empty-context");
                Reject<ContentUnavailableException>(() => new MigrationPackageExporter().Export(session, catalog.List, empty, CancellationToken.None, null)); NoOutputs(empty);
                catalog.SystemObjectsOverride = recorded.Where(context => context.Role != "UserInfoList");
                string missing = Path.Combine(root, "missing-context");
                Reject<InvalidDataException>(() => new MigrationPackageExporter().Export(session, catalog.List, missing, CancellationToken.None, null)); NoOutputs(missing);
                catalog.SystemObjectsOverride = recorded.Select(context => context.Role == "RootParentFolder" ? new MigrationSystemObject(catalog.List.Id, context.Type, context.Url, context.Role) : context);
                string collision = Path.Combine(root, "colliding-context");
                Reject<InvalidDataException>(() => new MigrationPackageExporter().Export(session, catalog.List, collision, CancellationToken.None, null)); NoOutputs(collision);
                catalog.SystemObjectsOverride = recorded.Select(context => context.Role == "UserInfoList" ? new MigrationSystemObject(context.Id, "Web", context.Url, context.Role) : context);
                string wrongType = Path.Combine(root, "wrong-context-type");
                Reject<InvalidDataException>(() => new MigrationPackageExporter().Export(session, catalog.List, wrongType, CancellationToken.None, null)); NoOutputs(wrongType);
            }
        }
        private static void TestDeletedPackageScopes(string root)
        {
            foreach(bool library in new[]{false,true})
            {
                var catalog=new Catalog(library);Node retained=MigrationSnapshotCopy.Node(catalog.List);retained.DeletionTransactionId="01";
                using(RecoverySession session=catalog.Session())
                {
                    foreach(bool all in new[]{false,true})
                        Reject<ArgumentException>(()=>new MigrationPackageExporter{IncludeHistory=all}.Export(session,retained,root,CancellationToken.None,null));
                    Require(catalog.StoreReads==0,"Current migration package read bytes for a deleted list/library scope.");NoOutputs(root);
                }
            }
        }
        private static void TestFailures(string root)
        {
            var catalog = new Catalog(true); Node file = catalog.AddFile("one.txt", catalog.List.Id, "Documents/one.txt", "one");
            using (RecoverySession session = catalog.Session())
            {
                catalog.BeforeCurrent = () => catalog.Current[file.Id].UiVersion++;
                Reject<InvalidDataException>(() => new MigrationPackageExporter().Export(session, catalog.List, root, CancellationToken.None, null));
                Require(catalog.StoreReads == 1, "The fresh prepared snapshot was not checked before publication.");
                NoOutputs(root);
                catalog.BeforeCurrent = null; catalog.Current[file.Id] = MigrationSnapshotCopy.Node(file); catalog.Current[file.Id].StreamSchema = 67;
                Reject<NotSupportedException>(() => new MigrationPackageExporter().Export(session, catalog.List, root, CancellationToken.None, null)); NoOutputs(root);
                catalog.Current[file.Id] = MigrationSnapshotCopy.Node(file); catalog.Content[file.Id] = Encoding.ASCII.GetBytes("wrong length");
                Reject<InvalidDataException>(() => new MigrationPackageExporter().Export(session, catalog.List, root, CancellationToken.None, null)); NoOutputs(root);
            }
            var list = new Catalog(false); list.AddItem("1_.000", true);
            using (RecoverySession session = list.Session())
            {
                Reject<ContentUnavailableException>(() => new MigrationPackageExporter().Export(session, list.List, root, CancellationToken.None, null)); NoOutputs(root);
                list.HasMissingUser = true;
                Reject<ContentUnavailableException>(() => new MigrationPackageExporter().Export(session, list.List, root, CancellationToken.None, null)); NoOutputs(root);
            }
            var attachmentRace = new Catalog(false); Node raceOwner = attachmentRace.AddItem("1_.000", true);
            attachmentRace.AddAttachment(raceOwner, "old.pdf", Encoding.ASCII.GetBytes("old payload"));
            using (RecoverySession session = attachmentRace.Session())
            {
                Reject<InvalidDataException>(() => new MigrationPackageExporter().Export(session, attachmentRace.List, root, CancellationToken.None, p =>
                    attachmentRace.AddAttachment(raceOwner, "new.pdf", Encoding.ASCII.GetBytes("new independently added payload"))));
                NoOutputs(root);
            }
            var metadata = new Catalog(true); metadata.AddFile("ok.txt", metadata.List.Id, "Documents/ok.txt", "ok"); metadata.Build = "15.0.4569.1000";
            using (RecoverySession session = metadata.Session())
            {
                Reject<NotSupportedException>(() => new MigrationPackageExporter().Export(session, metadata.List, root, CancellationToken.None, null)); NoOutputs(root);
            }
            var dtd = new Catalog(false); dtd.AddItem("1_.000", false); dtd.FieldsOverride = "<!DOCTYPE Fields [<!ENTITY injected SYSTEM 'file:///C:/Windows/win.ini'>]><Fields>&injected;</Fields>";
            using (RecoverySession session = dtd.Session())
            { Reject<System.Xml.XmlException>(() => new MigrationPackageExporter().Export(session, dtd.List, root, CancellationToken.None, null)); NoOutputs(root); }
        }
        private static void TestMetadataMutation(string root)
        {
            var catalog = new Catalog(true); catalog.AddFile("one.txt", catalog.List.Id, "Documents/one.txt", "known bytes");
            using (RecoverySession session = catalog.Session())
            {
                Reject<InvalidDataException>(() => new MigrationPackageExporter().Export(session, catalog.List, root, CancellationToken.None, p => catalog.Title = "Changed source metadata"));
                NoOutputs(root);
                Reject<IOException>(() => new MigrationPackageExporter().Export(session, catalog.List, root, CancellationToken.None, p => throw new IOException("Progress consumer failed")));
                NoOutputs(root);
            }
        }
        private static void TestCancellation(string root)
        {
            var catalog = new Catalog(true); catalog.AddFile("alpha.txt", catalog.List.Id, "Documents/alpha.txt", "alpha");
            catalog.AddFile("beta.txt", catalog.List.Id, "Documents/beta.txt", "beta");
            Directory.CreateDirectory(root); File.WriteAllText(Path.Combine(root, "sentinel.txt"), "existing");
            using (RecoverySession session = catalog.Session())
            using (var cancellation = new CancellationTokenSource())
            {
                MigrationPackageResult result = new MigrationPackageExporter().Export(session, catalog.List, root, cancellation.Token, p => cancellation.Cancel());
                Require(result.Cancelled && String.IsNullOrEmpty(result.PackagePath) && result.Bytes == 0 && catalog.StoreReads == 1, "Cancelled private payloads were published or reported as success.");
                NoOutputs(root); Require(File.ReadAllText(Path.Combine(root, "sentinel.txt")) == "existing", "Cancellation deleted unrelated destination files.");
                result = new MigrationPackageExporter().Export(session, catalog.List, root, cancellation.Token, null);
                Require(result.Cancelled && String.IsNullOrEmpty(result.PackagePath), "Pre-cancelled package work reported publication.");
            }
        }
        private static void CheckPayload(XDocument manifest, string root, Node file, byte[] expected)
        {
            XElement entry = manifest.Descendants().Single(e => (e.Name == M + "File" || e.Name == M + "Attachment") && (string)e.Attribute("Id") == file.Id.ToString("D"));
            byte[] actual = File.ReadAllBytes(Path.Combine(root, (string)entry.Attribute("FileValue")));
            Require(actual.SequenceEqual(expected) && Hash(actual) == Hash(expected), "A package payload differs from independently known source bytes.");
        }
        private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        private static void NoOutputs(string root) { if (Directory.Exists(root)) Require(Directory.GetDirectories(root).Length == 0, "A failed package left published or staging directories."); }
        private static void Require(bool value, string message) { if (!value) throw new Exception(message); }
        private static void Reject<T>(Action action) where T : Exception { try { action(); } catch (T) { return; } throw new Exception("Expected " + typeof(T).Name); }

        private sealed class Catalog : ISharePointCatalog, ISharePointMigrationCatalog, ISharePointMigrationHistoryCatalog, ISharePointVersionCatalog, ISharePointLibraryCatalog, ISharePointAttachmentCatalog, IDocumentChunkStore
        {
            internal readonly Node List;
            internal readonly List<Node> Items = new(), Attachments = new();
            internal MigrationHistorySnapshot History=MigrationHistorySnapshot.Empty;
            internal readonly Dictionary<string,byte[]> VersionContent=new();
            internal Action BeforeHistory;
            public MigrationHistorySnapshot ReadMigrationHistory(MigrationListSnapshot snapshot){BeforeHistory?.Invoke();return History;}
            public List<Node> GetFileVersions(Guid site,Guid id)=>Current.TryGetValue(id,out Node current)?new[]{MigrationSnapshotCopy.Node(current)}.Concat(History.Files.Where(file=>file.Document.Id==id).Select(file=>file.Document)).ToList():new();
            public Node GetFileVersion(Node selected)=>GetFileVersions(selected.SiteId,selected.Id).SingleOrDefault(file=>file.UiVersion==selected.UiVersion && file.HistoryVersion==selected.HistoryVersion && file.Level==selected.Level && file.InternalVersion==selected.InternalVersion && file.SiteId==selected.SiteId && file.WebId==selected.WebId && file.ListId==selected.ListId);
            internal readonly List<MigrationViewSnapshot> Views = new(); internal readonly List<MigrationTemplateFile> Templates = new();
            internal readonly Dictionary<Guid, Node> Current = new(); internal readonly Dictionary<Guid, byte[]> Content = new();
            internal long Flags, Flags2; internal Guid? TemplateId;
            internal readonly Guid UserInfoListId = Guid.NewGuid(); internal Guid RootWebId; internal string StoredSiteUrl = "", StoredWebUrl = "";
            internal IEnumerable<MigrationSystemObject> SystemObjectsOverride;
            internal int ScopedReads, UnscopedReads, AttachmentReads, StoreReads; internal Action BeforeCurrent;
            internal bool HasMissingUser; internal string Title = "Source list", Build = "16.0.14326.20450", FieldsOverride, ContentTypesOverride;
            private readonly Guid titleField = Guid.NewGuid(); private readonly DateTime time = new(2024, 1, 2, 3, 4, 5, DateTimeKind.Utc);
            public string SourceName => "Synthetic package source";
            internal Catalog(bool library)
            {
                Flags = library ? 8 : 0;
                List = new Node { Kind = library ? NodeKind.Library : NodeKind.List, Id = Guid.NewGuid(), SiteId = Guid.NewGuid(), WebId = Guid.NewGuid(), ListId = Guid.NewGuid(),
                    ParentId = Guid.NewGuid(), Name = library ? "Documents" : "Items", Path = library ? "Documents" : "Lists/Items", ListBaseType = library ? 1 : 0, Created = time, Modified = time };
                RootWebId = List.WebId;
            }
            internal RecoverySession Session() => new(this, this, DocumentDecoderRegistry.CreateDefault());
            internal Node AddFile(string name, Guid parent, string path, string text)
            {
                Node file = Member(NodeKind.File, name, parent, path); Content[file.Id] = Encoding.UTF8.GetBytes(text); file.Size = Content[file.Id].Length; file.HasStream = true;
                Current[file.Id] = MigrationSnapshotCopy.Node(file); Items.Add(file); return file;
            }
            internal Node AddFolder(string name) { Node folder = Member(NodeKind.Folder, name, List.Id, List.Path + "/" + name); Items.Add(folder); return folder; }
            internal Node AddItem(string name, bool attachments) { Node item = Member(NodeKind.ListItem, name, List.Id, List.Path + "/" + name); item.HasAttachments = attachments; item.HasStream = false; Items.Add(item); return item; }
            internal Node AddAttachment(Node owner, string name, byte[] bytes)
            {
                Node file = Member(NodeKind.File, name, Guid.NewGuid(), List.Path + "/Attachments/" + owner.ListItemId + "/" + name);
                file.ListItemId = owner.ListItemId; file.ItemUniqueId = owner.ItemUniqueId; file.AttachmentOwnerId = owner.Id; file.Size = bytes.Length; file.HasStream = true;
                Content[file.Id] = bytes; Attachments.Add(file); Current[file.Id] = MigrationSnapshotCopy.Node(file); return file;
            }
            internal void AddTemplateView()
            {
                Node file = Member(NodeKind.File, "AllItems.aspx", List.Id, List.Path + "/AllItems.aspx"); file.HasStream = false;
                Templates.Add(new MigrationTemplateFile(file, @"pages\viewpage.aspx", 15, @"CONTOSO\user"));
                Views.Add(new MigrationViewSnapshot(Guid.NewGuid(), "<ViewFields><FieldRef Name='Title'/></ViewFields>", 25165829, file.Id, file.Path, "All items", 1, 0, true, 1, "Main", 0,
                    Guid.Parse("6168d343-244e-414f-7a80-aa4acab73297"), null, 1, 1));
            }
            private Node Member(NodeKind kind, string name, Guid parent, string path) => new() { Kind = kind, Id = Guid.NewGuid(), SiteId = List.SiteId, WebId = List.WebId, ListId = List.ListId, ParentId = parent,
                Name = name, Path = path, ListBaseType = List.ListBaseType, ListItemId = Items.Count + 1, ItemUniqueId = Guid.NewGuid(), Created = time, Modified = time, UiVersion = 512, InternalVersion = 1, Level = 1 };
            public MigrationListSnapshot ReadMigrationList(Node requested)
            {
                var field = new MigrationFieldDefinition(titleField, "Title", "Text", false, "<Field ID='{" + titleField + "}' Name='Title' Type='Text'/>", null, null, null);
                string fields = FieldsOverride ?? "<Fields>" + field.SchemaXml + "</Fields>";
                var metadata = new MigrationListMetadata(List, Title, "description", List.ListBaseType.Value, List.Kind == NodeKind.Library ? 101 : 100,
                    Guid.Parse(List.Kind == NodeKind.Library ? "00bfea71-e717-4e80-aa17-d0c71b360101" : "00bfea71-de22-43b2-a848-c05709900100"), TemplateId,
                    Flags, Flags2, 1, 1, 0, time, 1, 0, 0, StoredSiteUrl, StoredWebUrl, "Root web", fields, ContentTypesOverride ?? "<ContentTypes><ContentType ID='0x01' Name='Item'/></ContentTypes>", Build,
                    "/_layouts/images/itgen.png", 1, "15.0.36.0", 15, 1033);
                var snapshots = Items.Select(item => new MigrationItemSnapshot(item, item.ListItemId.Value, item.ItemUniqueId.Value, 1, 1, time, time, item.HasAttachments == true,
                    item.Kind == NodeKind.Folder ? "0x0120" : List.Kind == NodeKind.Library ? "0x0101" : "0x01",
                    new[] { new MigrationFieldValue("Title", titleField, 1, "nvarchar", "Title <& " + item.Name), new MigrationFieldValue("ContentTypeId", null, 1, "varchar", item.Kind == NodeKind.Folder ? "0x0120" : List.Kind == NodeKind.Library ? "0x0101" : "0x01") })).ToArray();
                return new MigrationListSnapshot(metadata, new[] { field, new MigrationFieldDefinition(null, "ContentTypeId", "", true, "<FieldRef Name='ContentTypeId'/>", null, null, null) }, snapshots, Items.Where(i => i.Kind == NodeKind.Folder), Items.Where(i => i.Kind == NodeKind.File),
                    HasMissingUser ? Array.Empty<MigrationUserSnapshot>() : new[] { new MigrationUserSnapshot(1, "CONTOSO\\user", "User", "user@example.test", false, false, 0, new byte[] { 1, 2, 3 }) }, Views, Templates, SystemObjectsOverride ?? SystemContext());
            }
            private IEnumerable<MigrationSystemObject> SystemContext()
            {
                yield return new MigrationSystemObject(RootWebId, "Web", StoredSiteUrl, "RootWeb");
                if (RootWebId != List.WebId) yield return new MigrationSystemObject(List.WebId, "Web", StoredWebUrl, "SelectedWeb");
                int separator = List.Path.LastIndexOf('/');
                yield return new MigrationSystemObject(List.ParentId, "Folder", separator < 0 ? "" : List.Path.Substring(0, separator), "RootParentFolder");
                yield return new MigrationSystemObject(UserInfoListId, "List", StoredSiteUrl.TrimEnd('/') + "/_catalogs/users", "UserInfoList");
            }
            public Node GetCurrentLibraryFile(Node library, Guid id) { ScopedReads++; BeforeCurrent?.Invoke(); return Current.TryGetValue(id, out Node node) ? MigrationSnapshotCopy.Node(node) : null; }
            public IEnumerable<Node> EnumerateCurrentLibraryFiles(Node library) => Items.Where(i => i.Kind == NodeKind.File).Select(MigrationSnapshotCopy.Node);
            public List<Node> GetItemAttachments(Node owner) => Attachments.Where(a => a.AttachmentOwnerId == owner.Id).Select(MigrationSnapshotCopy.Node).ToList();
            public IEnumerable<Node> EnumerateCurrentListAttachments(Node list) => Attachments.Select(MigrationSnapshotCopy.Node);
            public Node GetCurrentItemAttachment(Node owner, Guid id) { AttachmentReads++; return Current.TryGetValue(id, out Node node) ? MigrationSnapshotCopy.Node(node) : null; }
            public IList<StoredChunk> ReadChunks(Node document) { StoreReads++; return new List<StoredChunk> { new() { Partition = 0, Content = document.HistoryVersion>0 ? VersionContent[document.Id.ToString("N")+":"+document.HistoryVersion] : Content[document.Id] } }; }
            public Node GetFile(Guid site, Guid id) { UnscopedReads++; throw new Exception("Package used unscoped file lookup."); }
            public void ValidateSchema() { }
            public List<string> CheckDatabase() => new();
            public List<Node> GetRootSites() => new();
            public List<Node> GetChildren(Node parent) => throw new Exception("Package depended on visible folder navigation.");
            public IEnumerable<Node> EnumerateCurrentFiles(Guid? site) => throw new Exception("Package enumerated other lists.");
        }
    }
}