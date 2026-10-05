using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using SharePointExplorer;

namespace SharePointExplorer.Tests
{
    internal static class MigrationPackageSqlChecks
    {
        private static readonly XNamespace M = "urn:deployment-manifest-schema";
        private static readonly Guid SiteId = new("cbd6e0be-6ee8-4b9d-9b04-d9031327831b");
        private static readonly Guid WebId = new("f9c7aeec-64be-4f81-87e9-2a65625e7bc9");
        public static void Run(SqlConnectionOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            string root = Path.GetFullPath(Path.Combine(".scratch", "migration-sql-packages", Guid.NewGuid().ToString("N")));
            var failures = new List<Exception>();
            using (RecoverySession session = RecoverySession.OpenSql(options))
            {
                var catalog = session.Catalog as ISharePointMigrationCatalog;
                Require(catalog != null, "SQL source does not expose deployment metadata.");
                foreach (string name in new[] { "List1_WF", "test", "doclib1" })
                {
                    try
                    {
                    Node list = FindScope(session.Catalog, name);
                    MigrationListSnapshot snapshot = catalog.ReadMigrationList(list);
                    string destination = Path.Combine(root, name); var progress = new List<MigrationPackageProgress>();
                    MigrationPackageResult result = new MigrationPackageExporter().Export(session, list, destination, CancellationToken.None, p => progress.Add(p));
                    Require(!result.Cancelled && !String.IsNullOrEmpty(result.PackagePath) && Directory.Exists(result.PackagePath), "Live SQL package was not published: " + name);
                    MigrationPackageExporter.ValidatePackage(result.PackagePath);
                    Require(Directory.GetDirectories(destination, "*.staging").Length == 0 && result.ItemCount == snapshot.Items.Count, "Live package item count/staging mismatch.");
                    XDocument manifest = XDocument.Load(Path.Combine(result.PackagePath, "Manifest.xml"));
                    XDocument system = XDocument.Load(Path.Combine(result.PackagePath, "SystemData.xml"));
                    XElement version = system.Root.Elements().Single(e => e.Name.LocalName == "SchemaVersion");
                    Require((string)version.Attribute("Build") == snapshot.Metadata.SourceVersion && (string)version.Attribute("DatabaseVersion") == "7123", "The exact SE source build/deployment compatibility profile was lost.");
                    int realFiles = snapshot.Files.Count(file => !snapshot.TemplateFiles.Any(template => template.Document.Id == file.Id));
                    Require(result.FileCount == realFiles && manifest.Descendants(M + "ListItem").Count() == snapshot.Items.Count, "Live source and deployment item/file sets disagree.");
                    CheckListSettings(snapshot, manifest, system, result.AuditPath);
                    CheckSystemContext(snapshot, system, result.AuditPath);
                    CheckItemIdentities(snapshot, manifest);
                    CheckFields(snapshot, manifest);
                    CheckContentTypeConversion(snapshot, manifest, result.AuditPath);
                    CheckTemplatesAndViews(snapshot, manifest);
                    CheckPayloadAudit(result, manifest);
                    Require(progress.Count == result.FileCount + result.AttachmentCount && progress.All(p => p.Status == RecoveryStatus.Success), "Live package progress does not match verified payloads.");
                    if (name == "test") CheckKnownAttachment(result, manifest, snapshot, session.Catalog as ISharePointAttachmentCatalog);
                    if (name == "doclib1" && snapshot.Files.Any(file=>file.Id==Guid.Parse("f0675d2e-4225-4ea9-8358-da3f7da48220"))) CheckKnownJpeg(result, manifest);
                    if(name=="doclib1" && !snapshot.Files.Any())Require(result.FileCount==0 && !manifest.Descendants(M+"File").Any(file=>file.Attribute("FileValue")!=null),"Deleted JPEG leaked into current library package.");
                    Console.WriteLine("PASS live SQL package " + name + ": items=" + result.ItemCount + ", files=" + result.FileCount + ", attachments=" + result.AttachmentCount + ", bytes=" + result.Bytes + ", path=" + result.PackagePath);
                    }
                    catch (Exception exception)
                    {
                        failures.Add(new InvalidDataException("Live SQL deployment scope " + name + ": " + exception.Message, exception));
                        Console.WriteLine("FAIL live SQL package " + name + ": " + exception.Message);
                    }
                }
            }
            CheckKnownHistory(options,Path.Combine(root,"retained-history"));
            if (failures.Count != 0)
            {
                Directory.CreateDirectory(root);
                File.WriteAllText(Path.Combine(root, "validation-errors.txt"), String.Join(Environment.NewLine + Environment.NewLine, failures.Select(failure => failure.ToString())));
                throw new AggregateException("Live SQL deployment package checks failed; diagnostics: " + Path.Combine(root, "validation-errors.txt"), failures);
            }
        }
        private static void CheckKnownHistory(SqlConnectionOptions options,string root)
        {
            using(RecoverySession session=RecoverySession.OpenSql(options))
            {
                Node fixtureWeb=session.Catalog.GetRootSites().Single(node=>node.SiteId==SiteId && node.WebId==WebId);
                Node scope=session.Catalog.GetChildren(fixtureWeb).Single(node=>node.ListId==Guid.Parse("db32638f-d83f-4625-acd4-0d9a66af601c"));
                // Source display titles can be localized; identify this fixture by its stored list GUID.
                if(scope.ListId!=Guid.Parse("db32638f-d83f-4625-acd4-0d9a66af601c"))
                {
                    Node web=session.Catalog.GetRootSites().Single(node=>node.SiteId==SiteId && node.WebId==WebId);
                    scope=session.Catalog.GetChildren(web).Single(node=>node.ListId==Guid.Parse("db32638f-d83f-4625-acd4-0d9a66af601c"));
                }
                var migration=(ISharePointMigrationCatalog)session.Catalog;MigrationListSnapshot snapshot=migration.ReadMigrationList(scope);
                MigrationHistorySnapshot history=((ISharePointMigrationHistoryCatalog)session.Catalog).ReadMigrationHistory(snapshot);
                Guid document=Guid.Parse("44da1cf3-e59f-49bb-b618-bc69c5f5ddcc");
                Require(history.Files.Any(file=>file.Document.Id==document && file.Document.UiVersion==512),"Known retained Word version is absent from deployment history catalog.");
                MigrationPackageResult result=new MigrationPackageExporter {IncludeHistory=true}.Export(session,scope,root,CancellationToken.None,null);
                MigrationPackageExporter.ValidatePackage(result.PackagePath);
                XDocument manifest=XDocument.Load(Path.Combine(result.PackagePath,"Manifest.xml"));
                XElement file=manifest.Descendants(M+"File").Single(entry=>!entry.Ancestors(M+"Versions").Any() && (string)entry.Attribute("Id")==document.ToString("D"));
                XElement old=file.Element(M+"Versions").Elements(M+"File").Single(entry=>(string)entry.Attribute("Version")=="1.0");
                XElement current=file.Element(M+"Versions").Elements(M+"File").Single(entry=>(string)entry.Attribute("Version")=="2.0");
                Require(file.Attribute("FileValue")==null,"All-version file incorrectly duplicates its current outer payload.");
                string currentPath=Path.Combine(result.PackagePath,(string)current.Attribute("FileValue")),oldPath=Path.Combine(result.PackagePath,(string)old.Attribute("FileValue"));
                Require(new FileInfo(currentPath).Length==10886 && Hash(currentPath)=="cb5357c9daea4fe0e71d5296eb41e7df526e768e3cf236000e6a4d5bf43b8e0a","Current package Word payload changed.");
                Require(new FileInfo(oldPath).Length==10877 && Hash(oldPath)=="c4fd9e425fab2cef769c8f272a7c248ba6e9977a5f6bc9268d7eafc98dd9b362","Historical package Word payload is not the independently recorded v1.0 bytes.");
                Require(manifest.Descendants(M+"ListItem").Any(entry=>entry.Ancestors(M+"Versions").Any() && (string)entry.Attribute("Version")=="1.0"),"Historical metadata was omitted while claiming all-version coverage.");
                Console.WriteLine("PASS live SQL version-history XML package: historyFiles="+history.Files.Count+", historyItems="+history.Items.Count+", path="+result.PackagePath);
            }
        }        private static Node FindScope(ISharePointCatalog catalog, string name)
        {
            Node web = catalog.GetRootSites().Single(node => node.SiteId == SiteId && node.WebId == WebId);
            List<Node> matches = catalog.GetChildren(web).Where(node => (node.Kind == NodeKind.Library || node.Kind == NodeKind.List) && String.Equals(node.Name, name, StringComparison.Ordinal)).ToList();
            Require(matches.Count == 1, "The exact expected restored source scope is missing/ambiguous: " + name);
            return matches[0];
        }
        private static void CheckListSettings(MigrationListSnapshot snapshot, XDocument manifest, XDocument system, string auditPath)
        {
            XElement list = manifest.Root.Elements().Select(entry => entry.Elements().Single()).Single(element => element.Name == M + "List" || element.Name == M + "DocumentLibrary");
            ulong flags = unchecked((ulong)snapshot.Metadata.Flags);
            Require((string)list.Attribute("Flags") == flags.ToString(CultureInfo.InvariantCulture), "The raw source list flags changed.");
            Require((string)list.Attribute("EnableVersioning") == ((flags & 0x80UL) != 0 ? "true" : "false") &&
                (string)list.Attribute("EnableAttachments") == (snapshot.Metadata.BaseType == 0 && (flags & 8UL) == 0 ? "true" : "false") &&
                (string)list.Attribute("EnableFolderCreation") == ((flags & 0x20000000UL) != 0 ? "false" : "true") &&
                (string)list.Attribute("NavigateForFormsPages") == ((flags & 0x80000000000000UL) != 0 ? "true" : "false"), "A live list setting was replaced by a target default.");
            if (snapshot.Metadata.BaseType == 1)
                Require((string)list.Attribute("ForceCheckout") == ((flags & 0x40000UL) != 0 ? "true" : "false") &&
                    (string)list.Attribute("EnableMinorVersions") == ((flags & 0x80000UL) != 0 ? "true" : "false"), "Live library checkout/minor-version settings changed.");
            if (snapshot.Metadata.TemplateId.HasValue && snapshot.Metadata.TemplateId.Value != Guid.Empty)
                Require((string)list.Attribute("DocumentTemplateId") == snapshot.Metadata.TemplateId.Value.ToString("D"), "The live document-template identity was omitted.");
            XElement[] rootOnly = system.Descendants().Where(element => element.Name.LocalName == "RootWebOnlyLists").SelectMany(element => element.Elements()).ToArray();
            Require(rootOnly.Length == ((flags & 0x4000UL) != 0 ? 1 : 0) && rootOnly.All(element => (string)element.Attribute("Id") == snapshot.Metadata.ListId.ToString("D")), "Live root-web-only restrictions differ from the source list flag.");
            using (JsonDocument audit = JsonDocument.Parse(File.ReadAllText(auditPath)))
                Require(audit.RootElement.GetProperty("SourceListFlags").GetUInt64() == flags && audit.RootElement.GetProperty("SourceListFlags2").GetUInt64() == unchecked((ulong)snapshot.Metadata.Flags2), "Raw current/extended flags were lost from the audit.");
        }
        private static void CheckSystemContext(MigrationListSnapshot snapshot, XDocument system, string auditPath)
        {
            XElement[] declared = system.Descendants().Where(element => element.Name.LocalName == "SystemObject").ToArray();
            Require(snapshot.SystemObjects.Count >= 3 && declared.Length == snapshot.SystemObjects.Count, "Source root-web/user-info/parent-folder context was omitted.");
            foreach (MigrationSystemObject source in snapshot.SystemObjects)
            {
                XElement element = declared.Single(context => (string)context.Attribute("Id") == source.Id.ToString("D"));
                string url = String.IsNullOrEmpty(source.Url) ? "/" : "/" + source.Url.TrimStart('/');
                Require((string)element.Attribute("Type") == source.Type && (string)element.Attribute("Url") == url, "The deployment context differs from the exact current SQL identity/type/path.");
            }
            MigrationSystemObject userInfo = snapshot.SystemObjects.Single(source => source.Role == "UserInfoList");
            Require(userInfo.Id == Guid.Parse("cc697787-26ca-4937-afc5-95e876afb5c3") && userInfo.Url == "_catalogs/users", "The independently verified restored UserInfo list identity/path changed.");
            using (JsonDocument audit = JsonDocument.Parse(File.ReadAllText(auditPath)))
            {
                Require(audit.RootElement.GetProperty("SourceSystemContextComplete").GetBoolean(), "The live SQL source incorrectly reported fallback deployment context.");
                JsonElement[] stored = audit.RootElement.GetProperty("SourceStoredSystemObjects").EnumerateArray().ToArray();
                foreach (MigrationSystemObject source in snapshot.SystemObjects)
                {
                    JsonElement captured = stored.Single(context => context.GetProperty("Id").GetGuid() == source.Id);
                    Require(captured.GetProperty("Type").GetString() == source.Type && captured.GetProperty("Url").GetString() == source.Url && captured.GetProperty("Role").GetString() == source.Role,
                        "The exact raw SQL deployment context was not retained in the audit.");
                }
            }
        }
        private static void CheckItemIdentities(MigrationListSnapshot snapshot, XDocument manifest)
        {
            foreach (MigrationItemSnapshot item in snapshot.Items)
            {
                XElement record = manifest.Descendants(M + "ListItem").Single(e => (string)e.Attribute("Id") == item.ItemUniqueId.ToString("D"));
                Require((string)record.Attribute("DocId") == item.DocumentId.ToString("D") && (string)record.Attribute("IntId") == item.ItemId.ToString(CultureInfo.InvariantCulture) &&
                    (string)record.Attribute("ParentFolderId") == item.ParentFolderId.ToString("D") && (string)record.Attribute("ParentListId") == snapshot.Metadata.ListId.ToString("D") &&
                    (string)record.Attribute("ContentTypeId") == item.ContentTypeId, "A live list-item unique/backing/numeric/folder/content-type identity was substituted.");
            }
        }
        private static void CheckFields(MigrationListSnapshot snapshot, XDocument manifest)
        {
            foreach (MigrationItemSnapshot item in snapshot.Items)
            {
                XElement record = manifest.Descendants(M + "ListItem").Single(e => (string)e.Attribute("Id") == item.ItemUniqueId.ToString("D"));
                foreach (MigrationFieldValue value in item.Values.Where(value => !value.IsNull))
                {
                    XElement field = record.Element(M + "Fields").Elements(M + "Field").SingleOrDefault(e => (string)e.Attribute("Name") == value.Name);
                    if (field == null && CheckSystemValue(record, value)) continue;
                    Require(field != null, "A non-null stored field was dropped: " + value.Name);
                    if (!value.FieldId.HasValue && value.Name == "File_x0020_Type") Require((string)field.Attribute("FieldId") == "39360f11-34cf-4356-9945-25c44e68dade", "The documented inherited file-type field identity changed.");
                    if (value.FieldId.HasValue) Require((string)field.Attribute("FieldId") == value.FieldId.Value.ToString("D"), "A stored field GUID was substituted.");
                    Require((string)field.Attribute(value.Component == 2 ? "Value2" : "Value") == Text(value.Value), "A stored typed field component changed: " + value.Name);
                }
            }
        }
        private static bool CheckSystemValue(XElement item, MigrationFieldValue value)
        {
            string attribute;
            switch (value.Name)
            {
                case "ContentTypeId": attribute = "ContentTypeId"; break;
                case "FileLeafRef": attribute = "Name"; break;
                case "ID": attribute = "IntId"; break;
                case "GUID": attribute = "Id"; break;
                case "UniqueId": attribute = "DocId"; break;
                case "Author": attribute = "Author"; break;
                case "Editor": attribute = "ModifiedBy"; break;
                case "Created": attribute = "TimeCreated"; break;
                case "Modified": attribute = "TimeLastModified"; break;
                case "Attachments": Require(item.Element(M + "Attachments") != null == (bool)value.Value, "Canonical attachment presence differs from the recorded system field."); return true;
                case "FSObjType": Require(((string)item.Attribute("DocType") == "Folder" ? "1" : "0") == Text(value.Value), "Canonical folder type differs from the recorded system field."); return true;
                case "_UIVersionString": attribute = "Version"; break;
                case "_UIVersion":
                    string[] parts = ((string)item.Attribute("Version")).Split('.');
                    Require(Int32.Parse(parts[0], CultureInfo.InvariantCulture) * 512 + Int32.Parse(parts[1], CultureInfo.InvariantCulture) == Convert.ToInt32(value.Value, CultureInfo.InvariantCulture), "Canonical version differs from the recorded numeric system version."); return true;
                case "FileRef":
                    Require((string)item.Parent.Attribute("Url") == "/" + Text(value.Value).TrimStart('/'), "Canonical source URL differs from the recorded file reference."); return true;
                default: return false;
            }
            Require((string)item.Attribute(attribute) == Text(value.Value), "Canonical " + attribute + " differs from the recorded duplicate system field."); return true;
        }
        private static string Text(object value)
        {
            if (value is DateTime time) return XmlConvert.ToString(DateTime.SpecifyKind(time, DateTimeKind.Utc), XmlDateTimeSerializationMode.Utc);
            if (value is bool flag) return flag ? "1" : "0";
            if (value is double number) return XmlConvert.ToString(number);
            if (value is float single) return XmlConvert.ToString(single);
            if (value is decimal money) return XmlConvert.ToString(money);
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }
        private static void CheckContentTypeConversion(MigrationListSnapshot snapshot, XDocument manifest, string auditPath)
        {
            XElement source = XElement.Parse("<ContentTypes>" + snapshot.Metadata.ContentTypesXml + "</ContentTypes>");
            XElement[] converted = source.Descendants().Where(element => element.Name.LocalName == "ContentType" && element.Attribute("DelayActivateTemplateBinding") != null).ToArray();
            foreach (XElement original in converted)
            {
                XElement deployed = manifest.Descendants(M + "ContentType").Single(element => (string)element.Attribute("ID") == (string)original.Attribute("ID"));
                Require(deployed.Attribute("DelayActivateTemplateBinding") == null, "Source-only content-type activation marker remains outside the canonical deployment schema.");
                foreach (XAttribute attribute in original.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration && attribute.Name.LocalName != "DelayActivateTemplateBinding"))
                    Require((string)deployed.Attribute(attribute.Name) == attribute.Value, "CAML conversion changed another content-type attribute: " + attribute.Name);
            }
            using (JsonDocument audit = JsonDocument.Parse(File.ReadAllText(auditPath)))
            {
                Require(audit.RootElement.GetProperty("SourceContentTypesXml").GetString() == snapshot.Metadata.ContentTypesXml &&
                    audit.RootElement.GetProperty("SourceFieldSchemaXml").GetString() == snapshot.Metadata.FieldSchemaXml &&
                    audit.RootElement.GetProperty("CamlConversions").GetArrayLength() == converted.Length, "The original SQL schema or explicit deployment conversion was not retained in the audit.");
            }
        }
        private static void CheckTemplatesAndViews(MigrationListSnapshot snapshot, XDocument manifest)
        {
            foreach (MigrationTemplateFile template in snapshot.TemplateFiles)
            {
                XElement file = manifest.Descendants(M + "File").Single(e => (string)e.Attribute("Id") == template.Document.Id.ToString("D"));
                Require((string)file.Attribute("IsGhosted") == "true" && (string)file.Attribute("SetupPath") == template.SetupPath && file.Attribute("FileValue") == null,
                    "A live ghost/template file was omitted or given invented bytes.");
            }
            foreach (MigrationViewSnapshot view in snapshot.Views)
            {
                XElement saved = manifest.Descendants(M + "View").Single(e => (string)e.Attribute("Name") == view.Id.ToString("D"));
                Require((string)saved.Attribute("FileId") == view.FileId.Value.ToString("D") && (string)saved.Attribute("Flags") == view.Flags.Value.ToString(CultureInfo.InvariantCulture) &&
                    (string)saved.Attribute("DefaultView") == (view.PageType == 0 ? "true" : "false"), "A live public view changed its page, flags or default state.");
                if (!String.IsNullOrWhiteSpace(view.SchemaXml))
                {
                    XDocument envelope = XDocument.Parse("<View>" + view.SchemaXml + "</View>");
                    foreach (XElement source in envelope.Descendants().Where(e => e.Name.LocalName == "FieldRef"))
                        Require(saved.Descendants().Any(e => e.Name.LocalName == "FieldRef" && (string)e.Attribute("Name") == (string)source.Attribute("Name")), "A stored custom view field was lost.");
                }
            }
        }
        private static void CheckPayloadAudit(MigrationPackageResult result, XDocument manifest)
        {
            long bytes = 0;
            using (JsonDocument audit = JsonDocument.Parse(File.ReadAllText(result.AuditPath)))
            {
                Require(!audit.RootElement.GetProperty("FarmImportVerified").GetBoolean(), "The live package falsely reports a farm import.");
                foreach (JsonElement payload in audit.RootElement.GetProperty("Payloads").EnumerateArray())
                {
                    Guid id = payload.GetProperty("DocumentId").GetGuid(); string name = payload.GetProperty("Name").GetString();
                    XElement entry = manifest.Descendants().Single(e => (e.Name == M + "File" || e.Name == M + "Attachment") && (string)e.Attribute("Id") == id.ToString("D"));
                    Require((string)entry.Attribute("FileValue") == name, "Audit and live document payload reference disagree.");
                    string file = Path.Combine(result.PackagePath, name);
                    long length = new FileInfo(file).Length; bytes += length;
                    Require(length == payload.GetProperty("Bytes").GetInt64() && Hash(file) == payload.GetProperty("Sha256").GetString(), "Live recovered bytes/hash differ from the package audit.");
                }
            }
            Require(bytes == result.Bytes, "Package total bytes differ from its recovered files.");
        }
        private static void CheckKnownAttachment(MigrationPackageResult result, XDocument manifest, MigrationListSnapshot snapshot, ISharePointAttachmentCatalog catalog)
        {
            Guid attachmentId = new("046b196d-ec40-48eb-90d6-4cba4ee06810");
            XElement entry = manifest.Descendants(M + "Attachment").Single(e => (string)e.Attribute("Id") == attachmentId.ToString("D"));
            XElement owner = entry.Ancestors(M + "ListItem").Single();
            Require((string)owner.Attribute("DocId") == "701fd014-9894-426f-8f2e-bcf01130b2a8" && (string)owner.Attribute("Id") == "1b46809a-cd66-43ef-81df-3984879fefd2" && (string)owner.Attribute("IntId") == "1",
                "The independently captured attachment owner identities changed.");
            MigrationItemSnapshot sourceOwner = snapshot.Items.Single(item => item.DocumentId == Guid.Parse((string)owner.Attribute("DocId")));
            Node sourceAttachment = catalog.GetItemAttachments(sourceOwner.Document).Single(file => file.Id == attachmentId);
            Require((string)entry.Attribute("Author") == sourceOwner.AuthorId.ToString(CultureInfo.InvariantCulture) && (string)entry.Attribute("ModifiedBy") == sourceOwner.EditorId.ToString(CultureInfo.InvariantCulture) &&
                (string)entry.Attribute("TimeCreated") == (sourceAttachment.Created.HasValue ? Text(sourceAttachment.Created.Value) : null), "The live attachment creation/owner attribution was omitted or invented.");
            string file = Path.Combine(result.PackagePath, (string)entry.Attribute("FileValue"));
            Require(new FileInfo(file).Length == 454216 && Hash(file) == "18132453d1f58ee4f8a73a85f346a26956bbbac40287cc039c602b5ef63970f5", "The package PDF differs from the independently captured database attachment.");
        }
        private static void CheckKnownJpeg(MigrationPackageResult result, XDocument manifest)
        {
            XElement entry = manifest.Descendants(M + "File").Single(e => (string)e.Attribute("Id") == "f0675d2e-4225-4ea9-8358-da3f7da48220");
            string file = Path.Combine(result.PackagePath, (string)entry.Attribute("FileValue"));
            Require(new FileInfo(file).Length == 212541 && Hash(file) == "48c41a496332b7953b71429e76ad4a662210b50754626042e0c1d7b6affe64ef", "The package JPEG differs from the independently captured database file.");
        }
        private static string Hash(string file) { using (FileStream input = File.OpenRead(file)) return Convert.ToHexString(SHA256.HashData(input)).ToLowerInvariant(); }
        private static void Require(bool value, string message) { if (!value) throw new Exception(message); }
    }
}