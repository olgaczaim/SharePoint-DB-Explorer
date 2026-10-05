using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

namespace SharePointExplorer
{
    public sealed class MigrationPackageResult
    {
        public string PackagePath { get; set; }
        public string ReportPath { get; set; }
        public string AuditPath { get; set; }
        public int FileCount { get; set; }
        public int AttachmentCount { get; set; }
        public int ItemCount { get; set; }
        public long Bytes { get; set; }
        public bool Cancelled { get; set; }
    }
    public sealed class MigrationPackageProgress
    {
        public int Completed { get; set; }
        public int Total { get; set; }
        public Node Document { get; set; }
        public RecoveryStatus Status { get; set; }
        public string Message { get; set; }
        public string ExportPath { get; set; }
    }

    // Public MS-PRIMEPF content-deployment XML, without a SharePoint runtime.
    // Schema/relationship validation is not proof of a target-farm import.
    public sealed class MigrationPackageExporter
    {
        private static readonly XNamespace Manifest = "urn:deployment-manifest-schema";
        private static readonly string[] XmlFiles = { "Manifest.xml", "ExportSettings.xml", "LookupListMap.xml", "Requirements.xml", "RootObjectMap.xml", "SystemData.xml", "UserGroup.xml", "ViewFormsList.xml" };
        private static readonly string[] SchemaFiles = { "DeploymentManifest.xsd", "DeploymentExportSettings.xsd", "DeploymentLookupListMap.xsd", "DeploymentRequirements.xsd", "DeploymentRootObjectMap.xsd", "DeploymentSystemData.xsd", "DeploymentUserGroupMap.xsd", "DeploymentViewFormsList.xsd" };
        private sealed class Payload
        {
            internal Node Document, Owner;
            internal string Name, Sha256, Decoder;
            internal long Bytes;
        }
        // Optional explicitly verified target/source compatibility value. A build
        // number is not a deployment DatabaseVersion and is never substituted.
        public int? DatabaseVersion { get; set; }
        public bool IncludeHistory { get; set; }

        public MigrationPackageResult Export(RecoverySession session, Node list, string destination,
            CancellationToken token, Action<MigrationPackageProgress> progress)
        {
            if (session == null) throw new ArgumentNullException("session");
            if (list == null) throw new ArgumentNullException("list");
            if (String.IsNullOrWhiteSpace(destination)) throw new ArgumentException("Choose a package destination.", "destination");
            Node selected = MigrationSnapshotCopy.Node(list);
            if (selected.IsDeleted || (selected.Kind != NodeKind.List && selected.Kind != NodeKind.Library) || selected.HistoryVersion != 0 ||
                selected.SiteId == Guid.Empty || selected.WebId == Guid.Empty || selected.ListId == Guid.Empty || selected.Id == Guid.Empty)
                throw new ArgumentException("Select a current list or library with its complete source identity.", "list");
            ISharePointMigrationCatalog catalog = session.Catalog as ISharePointMigrationCatalog;
            if (catalog == null) throw new NotSupportedException("This source cannot read content-deployment metadata.");
            string parent = Path.GetFullPath(destination), staging = null;
            var result = new MigrationPackageResult();
            try
            {
                token.ThrowIfCancellationRequested();
                MigrationListSnapshot snapshot = catalog.ReadMigrationList(selected);
                ValidateSnapshot(snapshot, selected);
                int databaseVersion = ResolveDatabaseVersion(snapshot.Metadata);
                string snapshotHash = SnapshotHash(snapshot);
                ISharePointMigrationHistoryCatalog historyCatalog = session.Catalog as ISharePointMigrationHistoryCatalog;
                if (IncludeHistory && historyCatalog == null) throw new NotSupportedException("This source cannot read retained deployment history metadata.");
                MigrationHistorySnapshot history = IncludeHistory ? historyCatalog.ReadMigrationHistory(snapshot) : MigrationHistorySnapshot.Empty;
                ValidateHistory(snapshot, history);
                string historyHash = Convert.ToHexString(SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(history)));
                var attachments = ReadAttachments(session, snapshot, token);
                string attachmentHash = AttachmentsHash(attachments);
                var templateIds = snapshot.TemplateFiles.Select(t => t.Document.Id).ToHashSet();
                var realFiles = snapshot.Files.Where(f => !templateIds.Contains(f.Id)).ToList();
                var payloads = new List<Payload>();
                Directory.CreateDirectory(parent);
                string stem = DocumentExporter.SafeFileName(snapshot.Metadata.Title); if (stem.Length > 100) stem = stem.Substring(0, 100).TrimEnd('.', ' ');
                string packageName = stem + "-package-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N");
                string final = Path.Combine(parent, packageName);
                staging = Path.Combine(parent, "." + packageName + ".staging");
                if (Directory.Exists(staging) || File.Exists(staging)) throw new IOException("The private package staging path already exists.");
                Directory.CreateDirectory(staging);
                int total = realFiles.Count + attachments.Count + history.Files.Count;
                foreach (Node file in realFiles)
                {
                    token.ThrowIfCancellationRequested();
                    PreparedDocument prepared = session.Engine.PrepareLibraryFile(file, snapshot.List);
                    ValidatePrepared(file, prepared.Document);
                    payloads.Add(RecoverPayload(prepared, null, staging, payloads.Count + 1, token));
                    Notify(progress, payloads.Count, total, prepared.Document, payloads[payloads.Count - 1].Name);
                }
                foreach (MigrationHistoryFile retained in history.Files)
                {
                    token.ThrowIfCancellationRequested();
                    Node version = retained.Document;
                    PreparedDocument prepared = session.Engine.PrepareVersion(version);
                    ValidatePreparedVersion(version, prepared.Document);
                    payloads.Add(RecoverPayload(prepared, null, staging, payloads.Count + 1, token));
                    Notify(progress, payloads.Count, total, prepared.Document, payloads[payloads.Count - 1].Name);
                }
                foreach (Tuple<Node, Node> attachment in attachments)
                {
                    token.ThrowIfCancellationRequested();
                    PreparedDocument prepared = session.Engine.PrepareItemAttachment(attachment.Item1, attachment.Item2);
                    ValidatePrepared(attachment.Item1, prepared.Document);
                    payloads.Add(RecoverPayload(prepared, attachment.Item2, staging, payloads.Count + 1, token));
                    Notify(progress, payloads.Count, total, prepared.Document, payloads[payloads.Count - 1].Name);
                }
                token.ThrowIfCancellationRequested();
                MigrationListSnapshot finalSnapshot = catalog.ReadMigrationList(MigrationSnapshotCopy.Node(selected));
                ValidateSnapshot(finalSnapshot, selected);
                if (!String.Equals(snapshotHash, SnapshotHash(finalSnapshot), StringComparison.Ordinal)) throw new InvalidDataException("List metadata changed during package recovery. Refresh and export again.");
                if (!String.Equals(attachmentHash, AttachmentsHash(ReadAttachments(session, finalSnapshot, token)), StringComparison.Ordinal))
                    throw new InvalidDataException("The attachment set or document metadata changed during package recovery. Refresh and export again.");
                if (IncludeHistory)
                {
                    MigrationHistorySnapshot finalHistory = historyCatalog.ReadMigrationHistory(finalSnapshot);
                    ValidateHistory(finalSnapshot, finalHistory);
                    if (!String.Equals(historyHash, Convert.ToHexString(SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(finalHistory))), StringComparison.Ordinal)) throw new InvalidDataException("Retained document/item history changed during package recovery. Refresh and export again.");
                }
                WritePackage(snapshot, history, payloads, staging, final, packageName, databaseVersion, IncludeHistory);
                WriteAudit(session.Catalog.SourceName, snapshot, history, payloads, staging, IncludeHistory);
                ValidatePackage(staging);
                token.ThrowIfCancellationRequested();
                Directory.Move(staging, final); // Same-volume atomic publication, never overwrite.
                staging = null;
                result.PackagePath = final; result.ReportPath = Path.Combine(final, "export-report.csv");
                result.AuditPath = Path.Combine(final, "package-audit.json"); result.FileCount = realFiles.Count;
                result.AttachmentCount = attachments.Count; result.ItemCount = snapshot.Items.Count;
                result.Bytes = payloads.Sum(p => p.Bytes);
                return result;
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested)
            { result.Cancelled = true; return result; }
            finally
            { if (staging != null && Directory.Exists(staging)) DeleteOwnedStaging(parent, staging); }
        }
        private static string AttachmentsHash(List<Tuple<Node, Node>> attachments)
        {
            var identities = attachments.OrderBy(a => a.Item1.Id).Select(a => new { Document = a.Item1, OwnerId = a.Item2.Id, ItemId = a.Item2.ListItemId, ItemUniqueId = a.Item2.ItemUniqueId }).ToArray();
            return Convert.ToHexString(SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(identities)));
        }
        private static string SnapshotHash(MigrationListSnapshot snapshot)
        { return Convert.ToHexString(SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(snapshot))); }
        private int ResolveDatabaseVersion(MigrationListMetadata metadata)
        {
            if (DatabaseVersion.HasValue && DatabaseVersion.Value >= 0) return DatabaseVersion.Value;
            // Content-deployment compatibility identifiers per source generation
            // (fixed per generation; they are not SQL build parts).
            // The generation comes from the recorded source build.
            if (metadata.ProductVersion == 15)
                switch (SharePointBuilds.Classify(metadata.SourceVersion))
                {
                    case SharePointGeneration.SharePoint2016: return 38455;
                    case SharePointGeneration.SharePoint2019: return 12710;
                    case SharePointGeneration.SubscriptionEdition: return 7123;
                }
            throw new NotSupportedException("The source build has no content-deployment compatibility profile. SharePoint Server 2016, 2019 and Subscription Edition sources are supported; supply a verified DatabaseVersion for others.");
        }
        private static bool SupportedSource(MigrationListMetadata metadata)
        {
            return metadata.ProductVersion == 15 && SharePointBuilds.IsSupported(SharePointBuilds.Classify(metadata.SourceVersion));
        }
        // SharePoint Server 2016 and 2019 export settings also declare alert and
        // Azure container options; Subscription Edition settings omit them.
        private static IEnumerable<XAttribute> GenerationSettings(MigrationListMetadata metadata)
        {
            SharePointGeneration generation = SharePointBuilds.Classify(metadata.SourceVersion);
            if (generation != SharePointGeneration.SharePoint2016 && generation != SharePointGeneration.SharePoint2019) yield break;
            yield return A("IncludeAlerts", "false");
            yield return A("AzureContainerSourceUri", String.Empty);
            yield return A("AzureContainerManifestUri", String.Empty);
        }
        private static void ValidateSnapshot(MigrationListSnapshot snapshot, Node selected)
        {
            if (snapshot == null || snapshot.Metadata == null) throw new InvalidDataException("The source returned no list metadata.");
            Node list = snapshot.List;
            if (list.Kind != selected.Kind || list.SiteId != selected.SiteId || list.WebId != selected.WebId || list.ListId != selected.ListId || list.Id != selected.Id || list.HistoryVersion != 0)
                throw new InvalidDataException("The source returned a different list scope.");
            MigrationListMetadata metadata = snapshot.Metadata;
            if (metadata.BaseType != 0 && metadata.BaseType != 1) throw new NotSupportedException("This package writer supports ordinary lists and document libraries. This list base type requires another serializer.");
            if ((metadata.BaseType == 1) != (list.Kind == NodeKind.Library)) throw new InvalidDataException("The list and its base type disagree.");
            if (metadata.ServerTemplate <= 0 || metadata.ImageUrl == null || !metadata.Version.HasValue ||
                String.IsNullOrEmpty(metadata.FieldSchemaXml) || String.IsNullOrEmpty(metadata.ContentTypesXml))
                throw new ContentUnavailableException("Required list template, image, version, field or content-type metadata is absent.");
            if (HasListFlag(metadata, 0x400UL)) throw new NotSupportedException("Moderation-enabled lists require source approval state that this deployment catalog does not preserve.");
            if (HasListFlag(metadata, 0x100000000000UL) || HasListFlag(metadata, 0x4000000000000UL))
                throw new NotSupportedException("List validation policies require source formulas/state that this deployment catalog does not preserve.");
            DeploymentContext(snapshot);
            if (HasListFlag(metadata, 0x0000000100000000UL)) throw new NotSupportedException("The source list is marked as excluded from content-deployment migration packages.");
            ReadCaml(metadata.FieldSchemaXml, "Fields"); // Reject malformed/DTD-bearing source metadata even when definitions are pre-parsed.
            ContentTypeDefinitions(snapshot); // Validate bounded source-to-deployment conversion before recovering payloads.
            Version build;
            if (!Version.TryParse(metadata.SourceVersion, out build) || build.Revision < 0 || !metadata.ProductVersion.HasValue || metadata.ProductVersion.Value != 15)
                throw new NotSupportedException("The source build or supported SharePoint deployment schema version is unavailable.");
            if (String.IsNullOrWhiteSpace(metadata.Title) || list.ParentId == Guid.Empty || list.ParentId == list.Id)
                throw new InvalidDataException("The list title or root-folder parent identity is invalid.");
            SourcePath(list.Path);
            var folders = new Dictionary<Guid, Node> { { list.Id, list } };
            foreach (Node folder in snapshot.Folders)
            {
                ValidateMember(folder, list);
                if (folder.Kind != NodeKind.Folder || folders.ContainsKey(folder.Id)) throw new InvalidDataException("The package folder identities are duplicated or invalid.");
                folders.Add(folder.Id, folder);
            }
            foreach (Node folder in snapshot.Folders)
            {
                var visited = new HashSet<Guid>(); Node cursor = folder;
                while (cursor.Id != list.Id)
                    if (!visited.Add(cursor.Id) || !folders.TryGetValue(cursor.ParentId, out cursor)) throw new InvalidDataException("A package folder has a missing parent or a cycle.");
            }
            var ids = new HashSet<Guid>(); var itemIds = new HashSet<int>(); var backing = new Dictionary<Guid, MigrationItemSnapshot>();
            foreach (MigrationItemSnapshot item in snapshot.Items)
            {
                ValidateMember(item.Document, list);
                if (item.ItemId <= 0 || item.ItemUniqueId == Guid.Empty || !ids.Add(item.ItemUniqueId) || !itemIds.Add(item.ItemId) || backing.ContainsKey(item.DocumentId) || !folders.ContainsKey(item.ParentFolderId))
                    throw new InvalidDataException("A package list item has incomplete, duplicated or out-of-scope identities.");
                if (String.IsNullOrEmpty(item.ContentTypeId)) throw new ContentUnavailableException("A list item's content-type identity is unavailable.");
                backing.Add(item.DocumentId, item);
            }
            var files = new HashSet<Guid>();
            foreach (Node file in snapshot.Files)
            {
                ValidateMember(file, list); MigrationItemSnapshot item;
                if (metadata.BaseType != 1 || file.Kind != NodeKind.File || !files.Add(file.Id) || !backing.TryGetValue(file.Id, out item) || item.IsFolder || !folders.ContainsKey(file.ParentId))
                    throw new InvalidDataException("A package file is not bound to a current library item.");
                if (file.HasStream == false && !snapshot.TemplateFiles.Any(t => t.Document.Id == file.Id)) throw new ContentUnavailableException("A required library document has no stored payload: " + file.Path);
            }
            var templateIds = new HashSet<Guid>();
            foreach (MigrationTemplateFile template in snapshot.TemplateFiles)
            {
                Node file = template.Document; ValidateMember(file, list);
                if (file.Kind != NodeKind.File || file.HasStream != false || !templateIds.Add(file.Id) || !folders.ContainsKey(file.ParentId) || String.IsNullOrWhiteSpace(template.SetupPath))
                    throw new InvalidDataException("A template file lacks its authoritative ghost metadata or parent.");
                string setup = template.SetupPath.Replace('\\', '/');
                if (setup.StartsWith("/", StringComparison.Ordinal) || setup.Contains(":")) throw new InvalidDataException("A template setup path must be a local SharePoint-relative path.");
                SourcePath(setup);
            }
            foreach (MigrationViewSnapshot view in snapshot.Views)
                if (!view.FileId.HasValue || !templateIds.Contains(view.FileId.Value)) throw new ContentUnavailableException("A public view lacks its referenced template file.");
            if (metadata.BaseType == 1 && snapshot.Items.Count(i => !i.IsFolder) != files.Count) throw new InvalidDataException("The library item and document sets disagree.");
            var users = snapshot.Users.ToDictionary(u => u.Id);
            foreach (int id in snapshot.Items.SelectMany(i => new[] { i.AuthorId, i.EditorId }).Concat(new[] { metadata.AuthorId }).Where(id => id > 0))
                if (!users.ContainsKey(id)) throw new ContentUnavailableException("A referenced author or editor is absent from the source user map.");
            foreach (MigrationUserSnapshot user in snapshot.Users)
                if (String.IsNullOrEmpty(user.Login) || String.IsNullOrEmpty(user.Name) || user.SystemId == null || user.SystemId.Length == 0)
                    throw new ContentUnavailableException("A referenced user lacks its login, name or system identity.");
            foreach (MigrationFieldDefinition field in snapshot.Fields)
                if (field.LookupListId.HasValue && field.LookupListId.Value != Guid.Empty && field.LookupListId.Value != list.ListId && !IsUserField(field))
                    throw new NotSupportedException("External lookup-list dependencies are not included in this package: " + field.Name);
        }
        private static void ValidateMember(Node member, Node list)
        {
            if (member == null || member.SiteId != list.SiteId || member.WebId != list.WebId || member.ListId != list.ListId || member.HistoryVersion != 0 || member.Id == Guid.Empty)
                throw new InvalidDataException("A package member belongs to another source scope.");
            if (!SourcePath(member.Path).StartsWith(SourcePath(list.Path).TrimEnd('/') + "/", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("A package member lies outside the selected list path.");
        }
        private static void ValidatePrepared(Node selected, Node fresh)
        {
            if (fresh.Id != selected.Id || fresh.SiteId != selected.SiteId || fresh.WebId != selected.WebId || fresh.ListId != selected.ListId ||
                fresh.ParentId != selected.ParentId || !String.Equals(fresh.Path, selected.Path, StringComparison.Ordinal) || !String.Equals(fresh.Name, selected.Name, StringComparison.Ordinal) ||
                fresh.UiVersion != selected.UiVersion || fresh.InternalVersion != selected.InternalVersion || fresh.Level != selected.Level || fresh.Size != selected.Size)
                throw new InvalidDataException("A package document changed after metadata discovery. Refresh and export the complete list again.");
        }
        private static List<Tuple<Node, Node>> ReadAttachments(RecoverySession session, MigrationListSnapshot snapshot, CancellationToken token)
        {
            var result = new List<Tuple<Node, Node>>(); var ids = new HashSet<Guid>();
            foreach (MigrationItemSnapshot item in snapshot.Items)
            {
                if (!item.HasAttachments) continue;
                if (snapshot.Metadata.BaseType == 1 || item.IsFolder) throw new NotSupportedException("Attachments on this item kind require an additional package serializer.");
                ISharePointAttachmentCatalog catalog = session.Catalog as ISharePointAttachmentCatalog;
                if (catalog == null) throw new NotSupportedException("The source cannot recover required list attachments.");
                token.ThrowIfCancellationRequested(); Node owner = item.Document; List<Node> files = catalog.GetItemAttachments(owner);
                if (files == null || files.Count == 0) throw new ContentUnavailableException("An item reports attachments, but no attachment documents are available.");
                foreach (Node file in files)
                {
                    ValidateMember(file, snapshot.List);
                    if (file.Kind != NodeKind.File || !ids.Add(file.Id) || file.AttachmentOwnerId != owner.Id || file.ListItemId != item.ItemId || file.ItemUniqueId != item.ItemUniqueId)
                        throw new InvalidDataException("An attachment is duplicated or belongs to another list item.");
                    result.Add(Tuple.Create(MigrationSnapshotCopy.Node(file), owner));
                }
            }
            return result;
        }
        private static Payload RecoverPayload(PreparedDocument prepared, Node owner, string staging, int sequence, CancellationToken token)
        {
            if (sequence > 99999999) throw new NotSupportedException("The package exceeds the supported payload count.");
            string name = sequence.ToString("D8", CultureInfo.InvariantCulture) + ".dat"; RecoveryResult recovered;
            using (var output = new FileStream(Path.Combine(staging, name), FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                using (var cancel = new CancellationWriteStream(output, token)) recovered = prepared.Recover(cancel);
                token.ThrowIfCancellationRequested();
                if (output.Length != prepared.OutputSize) throw new InvalidDataException("Package payload length mismatch.");
                output.Flush(true);
            }
            return new Payload { Document = prepared.Document, Owner = owner, Name = name, Bytes = recovered.Bytes, Sha256 = recovered.Sha256, Decoder = recovered.Decoder };
        }
        private static void Notify(Action<MigrationPackageProgress> callback, int completed, int total, Node document, string name)
        {
            if (callback != null) callback(new MigrationPackageProgress { Completed = completed, Total = total, Document = MigrationSnapshotCopy.Node(document), Status = RecoveryStatus.Success, Message = "Verified package payload", ExportPath = name });
        }
        private static void ValidatePreparedVersion(Node selected,Node fresh)
        {
            if(fresh==null || selected.SiteId!=fresh.SiteId || selected.WebId!=fresh.WebId || selected.ListId!=fresh.ListId || selected.Id!=fresh.Id || selected.HistoryVersion!=fresh.HistoryVersion ||
                selected.UiVersion!=fresh.UiVersion || selected.Level!=fresh.Level || selected.InternalVersion!=fresh.InternalVersion || selected.Size!=fresh.Size || selected.Path!=fresh.Path)
                throw new InvalidDataException("A retained file version changed before package recovery.");
        }
        private static void ValidateHistory(MigrationListSnapshot snapshot,MigrationHistorySnapshot history)
        {
            if(history==null)throw new InvalidDataException("The source omitted its retained history collection.");
            var current=snapshot.Items.ToDictionary(item=>item.ItemUniqueId);
            var identities=new HashSet<string>(StringComparer.Ordinal);
            foreach(MigrationItemSnapshot item in history.Items)
            {
                MigrationItemSnapshot parent;Node version=item.Document;
                if(!current.TryGetValue(item.ItemUniqueId,out parent) || parent.DocumentId!=item.DocumentId || parent.ItemId!=item.ItemId ||
                    version.SiteId!=snapshot.List.SiteId || version.WebId!=snapshot.List.WebId || version.ListId!=snapshot.List.ListId ||
                    version.HistoryVersion<=0 || version.UiVersion!=version.HistoryVersion || version.UiVersion>=parent.Document.UiVersion || version.ParentId!=parent.ParentFolderId ||
                    String.IsNullOrWhiteSpace(item.ContentTypeId) || !identities.Add(item.ItemUniqueId.ToString("N")+":"+version.UiVersion))
                    throw new InvalidDataException("Retained item history disagrees with the current source scope or contains duplicate versions.");
            }
            var files=new HashSet<string>(StringComparer.Ordinal);
            foreach(MigrationHistoryFile file in history.Files)
            {
                Node version=file.Document;
                if(version.Kind!=NodeKind.File || version.HistoryVersion<=0 || version.UiVersion!=version.HistoryVersion || !files.Add(version.Id.ToString("N")+":"+version.UiVersion) ||
                    !history.Items.Any(item=>item.DocumentId==version.Id && item.Document.UiVersion==version.UiVersion && item.Document.Level==version.Level && item.Document.InternalVersion==version.InternalVersion))
                    throw new InvalidDataException("A retained file payload lacks its exact historical item metadata.");
            }
            foreach(MigrationItemSnapshot item in history.Items.Where(item=>item.Document.Kind==NodeKind.File))
                if(!files.Contains(item.DocumentId.ToString("N")+":"+item.Document.UiVersion))throw new ContentUnavailableException("Historical document metadata has no retained file bytes.");
            var users=snapshot.Users.Concat(history.Users).GroupBy(user=>user.Id);
            foreach(var group in users)
                if(group.Select(user=>JsonSerializer.Serialize(user)).Distinct(StringComparer.Ordinal).Count()!=1)throw new InvalidDataException("Current and historical user-map metadata conflict.");
            var knownUsers=users.Select(group=>group.Key).ToHashSet();
            foreach(MigrationItemSnapshot item in history.Items)
                if((item.AuthorId>0 && !knownUsers.Contains(item.AuthorId)) || (item.EditorId>0 && !knownUsers.Contains(item.EditorId)))
                    throw new ContentUnavailableException("A historical author/editor has no verified source user identity.");
        }
        private static XElement HistoricalFileElement(MigrationListSnapshot snapshot,MigrationHistorySnapshot history,Payload payload,string web)
        {
            Node file=payload.Document;MigrationItemSnapshot item=history.Items.Single(old=>old.DocumentId==file.Id && old.Document.UiVersion==file.UiVersion);
            var result=new XElement(Manifest+"File",A("Name",file.Name),A("Id",file.Id),A("Url",WebRelative(file.Path,web)),A("ParentWebId",file.WebId),A("ParentWebUrl",web),
                A("ParentId",file.ParentId),A("ListId",file.ListId),A("ListItemIntId",item.ItemId),A("FileValue",payload.Name),A("Version",VersionText(file.UiVersion)),
                A("TimeCreated",Date(item.Created)),A("TimeLastModified",Date(item.Modified)),A("IsGhosted","false"));
            if(item.AuthorId>0)result.Add(A("Author",item.AuthorId));if(item.EditorId>0)result.Add(A("ModifiedBy",item.EditorId));
            string comment=history.Files.Single(old=>old.Document.Id==file.Id && old.Document.UiVersion==file.UiVersion).CheckinComment;
            if(comment!=null)result.Add(A("CheckinComment",comment));
            return result;
        }
        private static XElement HistoricalItemElement(MigrationListSnapshot snapshot,MigrationItemSnapshot item)
        {
            Node document=item.Document;
            var result=new XElement(Manifest+"ListItem",A("Name",document.Name),A("Id",item.ItemUniqueId),A("IntId",item.ItemId),A("ParentWebId",document.WebId),A("ParentListId",document.ListId),
                A("DirName",SourcePath(ParentPath(document.Path)).TrimStart('/')),A("Version",VersionText(document.UiVersion)),A("ContentTypeId",item.ContentTypeId),A("DocType",item.IsFolder?"Folder":"File"),
                A("TimeCreated",Date(item.Created)),A("TimeLastModified",Date(item.Modified)));
            if(item.AuthorId>0)result.Add(A("Author",item.AuthorId));if(item.EditorId>0)result.Add(A("ModifiedBy",item.EditorId));
            result.Add(ItemFields(snapshot,item));return result;
        }
        private static void ValidateRetainedVersions(XDocument manifest,XDocument settings)
        {
            bool all=(string)settings.Root.Attribute("IncludeVersions")=="All";
            if(!all && manifest.Descendants(Manifest+"Versions").Any())throw new InvalidDataException("Current-version settings contain undeclared history.");
            foreach(XElement versions in manifest.Descendants(Manifest+"Versions"))
            {
                if(versions.Ancestors(Manifest+"Versions").Any())throw new InvalidDataException("Nested historical collections are invalid.");
                XElement owner=versions.Parent;var seen=new HashSet<string>(StringComparer.Ordinal);
                int current=VersionNumber((string)owner.Attribute("Version"));
                foreach(XElement item in versions.Elements())
                {
                    string label=(string)item.Attribute("Version");
                    if(item.Name!=owner.Name || (string)item.Attribute("Id")!=(string)owner.Attribute("Id") || !seen.Add(label) || VersionNumber(label)>current)
                        throw new InvalidDataException("A historical version has a duplicated, future, or unrelated source identity.");
                    foreach(string attr in owner.Name==Manifest+"File"?new[]{"ParentWebId","ParentId","ListId","ListItemIntId"}:new[]{"ParentWebId","ParentListId","IntId"})
                        if((string)item.Attribute(attr)!=(string)owner.Attribute(attr))throw new InvalidDataException("A historical version changed its source scope.");
                }
                if(versions.Elements().Count(item=>VersionNumber((string)item.Attribute("Version"))==current)!=1)throw new InvalidDataException("A version collection omits or duplicates its current source version.");
                if(owner.Name==Manifest+"File" && owner.Attribute("FileValue")!=null)throw new InvalidDataException("A history-bearing file cannot declare a second outer current payload.");
            }
        }
        private static int VersionNumber(string value)
        {
            string[] parts=(value??"").Split('.');int major,minor;
            if(parts.Length!=2 || !Int32.TryParse(parts[0],out major) || !Int32.TryParse(parts[1],out minor) || major<0 || minor<0 || minor>=512)throw new InvalidDataException("Invalid retained SharePoint UI version.");
            return checked(major*512+minor);
        }
        private static void ValidateHistoryAudit(JsonElement history,XDocument manifest)
        {
            var items=history.GetProperty("Items").EnumerateArray().ToList();
            var emitted=manifest.Descendants(Manifest+"Versions").Where(v=>v.Parent.Name==Manifest+"ListItem").SelectMany(v=>v.Elements().Where(item=>VersionNumber((string)item.Attribute("Version"))<VersionNumber((string)v.Parent.Attribute("Version")))).ToList();
            if(items.Count!=emitted.Count)throw new InvalidDataException("Historical item coverage differs from its source audit.");
            foreach(JsonElement item in items)
            {
                Guid id=item.GetProperty("ItemUniqueId").GetGuid();int ui=item.GetProperty("Document").GetProperty("UiVersion").GetInt32();
                XElement entry=emitted.SingleOrDefault(e=>(string)e.Attribute("Id")==id.ToString("D") && (string)e.Attribute("Version")==VersionText(ui));
                if(entry==null || (string)entry.Attribute("Author")!=(item.GetProperty("AuthorId").GetInt32()>0?item.GetProperty("AuthorId").GetInt32().ToString(CultureInfo.InvariantCulture):null) ||
                    (string)entry.Attribute("ModifiedBy")!=(item.GetProperty("EditorId").GetInt32()>0?item.GetProperty("EditorId").GetInt32().ToString(CultureInfo.InvariantCulture):null))
                    throw new InvalidDataException("Historical item identity or attribution changed after recovery.");
            }
            var files=history.GetProperty("Files").EnumerateArray().ToList();
            var fileEntries=manifest.Descendants(Manifest+"Versions").Where(v=>v.Parent.Name==Manifest+"File").SelectMany(v=>v.Elements().Where(item=>VersionNumber((string)item.Attribute("Version"))<VersionNumber((string)v.Parent.Attribute("Version")))).ToList();
            if(files.Count!=fileEntries.Count)throw new InvalidDataException("Historical file coverage differs from its source audit.");
            foreach(JsonElement file in files)
            {
                JsonElement document=file.GetProperty("Document");Guid id=document.GetProperty("Id").GetGuid();int ui=document.GetProperty("UiVersion").GetInt32();
                XElement entry=fileEntries.SingleOrDefault(e=>(string)e.Attribute("Id")==id.ToString("D") && (string)e.Attribute("Version")==VersionText(ui));
                if(entry==null || (string)entry.Attribute("CheckinComment")!=(file.GetProperty("CheckinComment").ValueKind==JsonValueKind.Null?null:file.GetProperty("CheckinComment").GetString()))
                    throw new InvalidDataException("Historical file identity or check-in comment changed after recovery.");
            }
        }
        private static void WritePackage(MigrationListSnapshot snapshot, MigrationHistorySnapshot history, List<Payload> payloads, string directory, string final, string packageName, int databaseVersion, bool includeHistory)
        {
            // Historical-only user fields need the same verified user map as authors/editors.
            snapshot = new MigrationListSnapshot(snapshot.Metadata,snapshot.Fields,snapshot.Items,snapshot.Folders,snapshot.Files,
                snapshot.Users.Concat(history.Users).GroupBy(user=>user.Id).Select(group=>group.First()),snapshot.Views,snapshot.TemplateFiles,snapshot.SystemObjects);
            MigrationListMetadata metadata = snapshot.Metadata; Node list = snapshot.List;
            string web = SourcePath(metadata.StoredWebUrl), root = SourcePath(list.Path);
            var objects = new XElement(Manifest + "SPObjects");
            var listElement = new XElement(Manifest + (metadata.BaseType == 1 ? "DocumentLibrary" : "List"),
                A("Id", list.ListId), A("Title", metadata.Title), A("RootFolderId", list.Id), A("RootFolderUrl", root), A("ParentWebId", list.WebId), A("ParentWebUrl", web),
                A("BaseType", metadata.BaseType == 1 ? "DocumentLibrary" : "GenericList"), A("BaseTemplate", metadata.ServerTemplate), A("ImageUrl", metadata.ImageUrl),
                A("Flags", unchecked((ulong)metadata.Flags)), A("ReadSecurity", metadata.ReadSecurity), A("WriteSecurity", metadata.WriteSecurity), A("Version", metadata.Version.Value),
                A("Description", metadata.Description ?? String.Empty), A("Direction", metadata.Direction == 1 ? "ltr" : metadata.Direction == 2 ? "rtl" : "none"),
                A("Created", Date(metadata.Created)), A("MajorVersionLimit", metadata.MaxMajorVersions), A("MajorWithMinorVersionsLimit", metadata.MaxMajorWithMinorVersions));
            listElement.Add(ListSettings(metadata));
            if (metadata.TemplateId.HasValue && metadata.TemplateId.Value != Guid.Empty) listElement.Add(A("DocumentTemplateId", metadata.TemplateId.Value));
            if (metadata.AuthorId > 0) listElement.Add(A("Author", metadata.AuthorId));
            if (metadata.TemplateFeatureId != Guid.Empty) listElement.Add(A("TemplateFeatureId", metadata.TemplateFeatureId));
            listElement.Add(FieldDefinitions(snapshot));
            listElement.Add(ContentTypeDefinitions(snapshot));
            if (snapshot.Views.Count > 0)
            {
                var views = new XElement(Manifest + "Views");
                foreach (MigrationViewSnapshot view in snapshot.Views) views.Add(ViewElement(view));
                listElement.Add(views);
            }
            objects.Add(Object(metadata.BaseType == 1 ? "SPDocumentLibrary" : "SPList", list.ListId, list.WebId, list.WebId, web, root, listElement));
            var rootFolder = MigrationSnapshotCopy.Node(list); rootFolder.Kind = NodeKind.Folder;
            objects.Add(FolderObject(rootFolder, list, web));
            var pending = snapshot.Folders.ToList(); var emitted = new HashSet<Guid> { list.Id };
            while (pending.Count > 0)
            {
                Node folder = pending.FirstOrDefault(f => emitted.Contains(f.ParentId));
                if (folder == null) throw new InvalidDataException("The package folder graph is incomplete.");
                objects.Add(FolderObject(folder, list, web)); emitted.Add(folder.Id); pending.Remove(folder);
            }
            foreach (MigrationTemplateFile template in snapshot.TemplateFiles)
            {
                Node file = template.Document;
                var element = new XElement(Manifest + "File", A("Name", file.Name), A("Id", file.Id), A("Url", WebRelative(file.Path, web)),
                    A("ParentWebId", list.WebId), A("ParentWebUrl", web), A("ParentId", file.ParentId), A("ListId", list.ListId), A("Version", VersionText(file.UiVersion)),
                    A("IsGhosted", "true"), A("SetupPath", template.SetupPath));
                if (template.SetupPathVersion.HasValue) element.Add(A("SetupPathVersion", template.SetupPathVersion.Value));
                if (!String.IsNullOrWhiteSpace(template.SetupPathUser))
                {
                    int userId; MigrationUserSnapshot user = snapshot.Users.SingleOrDefault(u => String.Equals(u.Login, template.SetupPathUser, StringComparison.OrdinalIgnoreCase));
                    if (user != null) element.Add(A("SetupPathUser", user.Id));
                    else if (Int32.TryParse(template.SetupPathUser, NumberStyles.None, CultureInfo.InvariantCulture, out userId) && snapshot.Users.Any(u => u.Id == userId)) element.Add(A("SetupPathUser", userId));
                    else throw new ContentUnavailableException("A template setup user is absent from the verified user map.");
                }
                var associated = snapshot.Views.Where(v => v.FileId == file.Id).ToList();
                if (associated.Count > 0) element.Add(new XElement(Manifest + "WebParts", associated.Select(v => WebPartElement(v, snapshot))));
                var versionPayloads = payloads.Where(p => p.Owner == null && p.Document.Id == file.Id).OrderBy(p => p.Document.UiVersion).ToList();
                if (includeHistory && versionPayloads.Count > 0)
                {
                    XElement currentVersion = new XElement(element);
                    element.Add(new XElement(Manifest + "Versions", versionPayloads.Select(p => p.Document.HistoryVersion == 0 ? new XElement(currentVersion) : HistoricalFileElement(snapshot, history, p, web))));
                    element.Attribute("FileValue").Remove();
                }
                objects.Add(Object("SPFile", file.Id, file.ParentId, list.WebId, web, SourcePath(file.Path), element));
            }
            foreach (Payload payload in payloads.Where(p => p.Owner == null && p.Document.HistoryVersion == 0))
            {
                Node file = payload.Document; MigrationItemSnapshot item = snapshot.Items.Single(i => i.DocumentId == file.Id);
                var element = new XElement(Manifest + "File", A("Name", file.Name), A("Id", file.Id), A("Url", WebRelative(file.Path, web)), A("ParentWebId", list.WebId), A("ParentWebUrl", web),
                    A("ParentId", file.ParentId), A("ListId", list.ListId), A("ListItemIntId", item.ItemId), A("FileValue", payload.Name), A("Version", VersionText(file.UiVersion)),
                    A("TimeCreated", Date(item.Created)), A("TimeLastModified", Date(item.Modified)), A("IsGhosted", "false"));
                if (item.AuthorId > 0) element.Add(A("Author", item.AuthorId));
                if (item.EditorId > 0) element.Add(A("ModifiedBy", item.EditorId));
                var versionPayloads = payloads.Where(p => p.Owner == null && p.Document.Id == file.Id).OrderBy(p => p.Document.UiVersion).ToList();
                if (includeHistory && versionPayloads.Count > 0)
                {
                    XElement currentVersion = new XElement(element);
                    element.Add(new XElement(Manifest + "Versions", versionPayloads.Select(p => p.Document.HistoryVersion == 0 ? new XElement(currentVersion) : HistoricalFileElement(snapshot, history, p, web))));
                    element.Attribute("FileValue").Remove();
                }
                objects.Add(Object("SPFile", file.Id, file.ParentId, list.WebId, web, SourcePath(file.Path), element));
            }
            foreach (MigrationItemSnapshot item in snapshot.Items)
            {
                Node document = item.Document;
                var element = new XElement(Manifest + "ListItem", A("Name", document.Name), A("Id", item.ItemUniqueId), A("IntId", item.ItemId),
                    A("ParentWebId", list.WebId), A("ParentListId", list.ListId), A("FileUrl", WebRelative(document.Path, web)), A("DocId", item.DocumentId),
                    A("ParentFolderId", item.ParentFolderId), A("DirName", SourcePath(ParentPath(document.Path)).TrimStart('/')), A("Version", VersionText(document.UiVersion)),
                    A("ContentTypeId", item.ContentTypeId), A("DocType", item.IsFolder ? "Folder" : "File"), A("TimeCreated", Date(item.Created)), A("TimeLastModified", Date(item.Modified)));
                if (item.AuthorId > 0) element.Add(A("Author", item.AuthorId));
                if (item.EditorId > 0) element.Add(A("ModifiedBy", item.EditorId));
                element.Add(ItemFields(snapshot, item));
                var retainedItems = history.Items.Where(old => old.ItemUniqueId == item.ItemUniqueId).Concat(new[]{item}).OrderBy(old => old.Document.UiVersion).ToList();
                if (includeHistory)
                    element.Add(new XElement(Manifest + "Versions", retainedItems.Select(old => HistoricalItemElement(snapshot, old))));
                var owned = payloads.Where(p => p.Owner != null && p.Owner.Id == item.DocumentId).ToList();
                if (owned.Count > 0)
                    element.Add(new XElement(Manifest + "Attachments", owned.Select(p => new XElement(Manifest + "Attachment", A("Name", p.Document.Name), A("Id", p.Document.Id),
                        A("Url", WebRelative(p.Document.Path, web)), A("ParentWebId", list.WebId), A("FileValue", p.Name), A("TimeCreated", p.Document.Created.HasValue ? Date(p.Document.Created.Value) : null),
                        A("TimeLastModified", Date(p.Document.Modified)), A("Author", item.AuthorId > 0 ? (object)item.AuthorId : null), A("ModifiedBy", item.EditorId > 0 ? (object)item.EditorId : null)))));
                objects.Add(Object("SPListItem", item.ItemUniqueId, list.ListId, list.WebId, web, SourcePath(document.Path), element));
            }
            Save(directory, "Manifest.xml", objects);
            XNamespace settings = "urn:deployment-exportsettings-schema";
            Save(directory, "ExportSettings.xml", new XElement(settings + "ExportSettings", A("SiteUrl", ExportSiteUrl(metadata.StoredSiteUrl)), A("FileLocation", final), A("BaseFileName", packageName),
                A("IncludeSecurity", "None"), A("IncludeVersions", includeHistory ? "All" : "CurrentVersion"), A("ExportMethod", "ExportAll"), A("ExportPublicSchema", "true"), A("ExcludeDependencies", "true"), A("ExportFrontEndFileStreams", "false"),
                GenerationSettings(metadata),
                new XElement(settings + "ExportObjects", new XElement(settings + "DeploymentObject", A("Id", list.ListId), A("Type", "List"), A("ParentId", list.WebId), A("Url", root), A("ExcludeChildren", "false"), A("IncludeDescendants", "All")))));
            XNamespace roots = "urn:deployment-rootobjectmap-schema";
            Save(directory, "RootObjectMap.xml", new XElement(roots + "RootObjects", new XElement(roots + "RootObject", A("Id", list.ListId), A("Type", "List"), A("ParentId", list.WebId), A("WebUrl", web), A("Url", root), A("IsDependency", "false"))));
            XNamespace system = "urn:deployment-systemdata-schema";
            var systemData = new XElement(system + "SystemData", new XElement(system + "SchemaVersion", A("Version", "15.0.0.0"), A("Build", metadata.SourceVersion),
                A("DatabaseVersion", databaseVersion), A("SiteVersion", "15"), A("ObjectsProcessed", objects.Elements().Count())), new XElement(system + "ManifestFiles", new XElement(system + "ManifestFile", A("Name", "Manifest.xml"))),
                new XElement(system + "SystemObjects", DeploymentContext(snapshot).Select(context => new XElement(system + "SystemObject", A("Id", context.Id), A("Type", context.Type), A("Url", context.Url)))));
            if (HasListFlag(metadata, 0x0000000000004000UL)) systemData.Add(new XElement(system + "RootWebOnlyLists", new XElement(system + "List", A("Id", list.ListId))));
            Save(directory, "SystemData.xml", systemData);
            XNamespace users = "urn:deployment-usergroupmap-schema";
            Save(directory, "UserGroup.xml", new XElement(users + "UserGroupMap", new XElement(users + "Users", snapshot.Users.Concat(history.Users).GroupBy(user => user.Id).Select(group => group.First()).Select(user => new XElement(users + "User", A("Id", user.Id),
                A("Name", user.Name), A("Login", user.Login), A("Email", user.Email ?? String.Empty), A("IsDomainGroup", user.IsDomainGroup ? "true" : "false"), A("IsSiteAdmin", user.IsSiteAdmin ? "true" : "false"),
                A("SystemId", Convert.ToBase64String(user.SystemId)), A("IsDeleted", user.Deleted != 0 ? "true" : "false")))), new XElement(users + "Groups")));
            XNamespace requirements = "urn:deployment-requirements-schema";
            var requirementRoot = new XElement(requirements + "Requirements");
            if (metadata.TemplateFeatureId != Guid.Empty) requirementRoot.Add(new XElement(requirements + "Requirement", A("Type", "FeatureDefinition"), A("Id", metadata.TemplateFeatureId), A("Name", metadata.TemplateFeatureId)));
            if (metadata.Language.HasValue) requirementRoot.Add(new XElement(requirements + "Requirement", A("Type", "Language"), A("Id", metadata.Language.Value), A("Name", CultureInfo.GetCultureInfo(metadata.Language.Value).EnglishName)));
            Save(directory, "Requirements.xml", requirementRoot);
            XNamespace lookups = "urn:deployment-lookuplistmap-schema";
            var lookupRoot = new XElement(lookups + "LookupLists");
            if (snapshot.Fields.Any(f => f.LookupListId == list.ListId && !IsUserField(f)))
                lookupRoot.Add(new XElement(lookups + "LookupList", A("Id", list.ListId), A("Url", root), A("Included", "true"), new XElement(lookups + "LookupItems",
                    snapshot.Items.Select(i => new XElement(lookups + "LookupItem", A("Id", i.ItemId), A("DocId", i.DocumentId), A("Url", SourcePath(i.Document.Path)), A("Included", "true"))))));
            Save(directory, "LookupListMap.xml", lookupRoot);
            XNamespace viewForms = "urn:deployment-viewformlist-schema";
            Save(directory, "ViewFormsList.xml", new XElement(viewForms + "ViewFormsList", snapshot.Views.Select(v => new XElement(viewForms + "ViewForm", A("Id", v.Id), A("Type", "View")))));
        }
        private static XElement FolderObject(Node folder, Node list, string web)
        {
            var element = new XElement(Manifest + "Folder", A("Id", folder.Id), A("Name", folder.Name), A("Url", WebRelative(folder.Path, web)), A("ParentFolderId", folder.ParentId),
                A("ParentWebId", list.WebId), A("ParentWebUrl", web), A("ContainingDocumentLibrary", list.ListId));
            if (folder.Created.HasValue) element.Add(A("TimeCreated", Date(folder.Created.Value)));
            if (folder.Modified != DateTime.MinValue) element.Add(A("TimeLastModified", Date(folder.Modified)));
            if (folder.ListItemId.HasValue) element.Add(A("ListItemIntId", folder.ListItemId.Value));
            return Object("SPFolder", folder.Id, folder.ParentId, list.WebId, web, SourcePath(folder.Path), element);
        }
        private static List<MigrationSystemObject> DeploymentContext(MigrationListSnapshot snapshot)
        {
            MigrationListMetadata metadata = snapshot.Metadata; Node list = snapshot.List;
            if (snapshot.SystemObjects.Count == 0)
                throw new ContentUnavailableException("The catalog cannot provide the required root-web, user-info-list and parent-folder deployment context.");
            var objects = snapshot.SystemObjects.Select(context => context == null ? null : new MigrationSystemObject(context.Id, context.Type, SourcePath(context.Url), context.Role)).ToList();
            var ids = new HashSet<Guid>(); var roles = new HashSet<string>(StringComparer.Ordinal);
            var exported = new HashSet<Guid> { list.Id, list.ListId };
            foreach (Node folder in snapshot.Folders) exported.Add(folder.Id);
            foreach (Node file in snapshot.Files) exported.Add(file.Id);
            foreach (MigrationTemplateFile template in snapshot.TemplateFiles) exported.Add(template.Document.Id);
            foreach (MigrationItemSnapshot item in snapshot.Items) { exported.Add(item.ItemUniqueId); exported.Add(item.DocumentId); }
            foreach (MigrationSystemObject context in objects)
            {
                if (context == null || context.Id == Guid.Empty || !ids.Add(context.Id) || !roles.Add(context.Role ?? String.Empty) || exported.Contains(context.Id))
                    throw new InvalidDataException("Deployment system context is absent, duplicated or collides with exported identities.");
                string expectedType, expectedUrl;
                switch (context.Role)
                {
                    case "RootWeb": expectedType = "Web"; expectedUrl = SourcePath(metadata.StoredSiteUrl); break;
                    case "SelectedWeb":
                        expectedType = "Web"; expectedUrl = SourcePath(metadata.StoredWebUrl);
                        if (context.Id != list.WebId) throw new InvalidDataException("The selected deployment web identity differs from the list scope."); break;
                    case "RootParentFolder":
                        expectedType = "Folder"; expectedUrl = SourcePath(ParentPath(list.Path));
                        if (context.Id != list.ParentId) throw new InvalidDataException("The root-parent deployment folder identity differs from the list scope."); break;
                    case "UserInfoList": expectedType = "List"; expectedUrl = SourcePath(metadata.StoredSiteUrl).TrimEnd('/') + "/_catalogs/users"; break;
                    default: throw new InvalidDataException("An unrelated deployment system-object role was returned.");
                }
                if (context.Type != expectedType || !String.Equals(context.Url, expectedUrl, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("A deployment system-object type or URL differs from its verified source role.");
            }
            MigrationSystemObject rootWeb = objects.SingleOrDefault(context => context.Role == "RootWeb");
            if (rootWeb == null || !roles.Contains("RootParentFolder") || !roles.Contains("UserInfoList") ||
                (rootWeb.Id == list.WebId ? roles.Contains("SelectedWeb") : !roles.Contains("SelectedWeb")) ||
                (rootWeb.Id == list.WebId && !String.Equals(rootWeb.Url, SourcePath(metadata.StoredWebUrl), StringComparison.OrdinalIgnoreCase)))
                throw new InvalidDataException("Required root/selected-web, user-info-list or parent-folder deployment context is missing or inconsistent.");
            return objects;
        }
        private static bool HasListFlag(MigrationListMetadata metadata, ulong bit)
        { return (unchecked((ulong)metadata.Flags) & bit) != 0; }
        private static IEnumerable<XAttribute> ListSettings(MigrationListMetadata metadata)
        { return ListSettings(metadata.Flags, metadata.BaseType); }
        private static IEnumerable<XAttribute> ListSettings(long rawFlags, int baseType)
        {
            Func<ulong, bool> has = bit => (unchecked((ulong)rawFlags) & bit) != 0;
            // Independent mappings from MS-WSSFO3 section2.2.2.5 List Flags to
            // the explicit SPList attributes in MS-PRIMEPF. Raw flags remain
            // recorded too; semantic settings must not rely on importer defaults.
            yield return A("Ordered", has(0x1UL) ? "true" : "false");
            yield return A("AllowDeletion", has(0x4UL) ? "false" : "true");
            yield return A("EnableAttachments", baseType == 0 && !has(0x8UL) ? "true" : "false");
            yield return A("EnableAssignToEmail", has(0x40UL) ? "true" : "false");
            yield return A("EnableVersioning", has(0x80UL) ? "true" : "false");
            yield return A("Hidden", has(0x100UL) ? "true" : "false");
            yield return A("RequestAccessEnabled", has(0x200UL) ? "true" : "false");
            yield return A("EnableModeration", has(0x400UL) ? "true" : "false");
            yield return A("ExcludeFromTemplate", has(0x2000UL) ? "true" : "false");
            yield return A("EnableContentTypes", has(0x400000UL) ? "true" : "false");
            yield return A("EnableDeployWithDependentList", has(0x8000000UL) ? "false" : "true");
            yield return A("EnableFolderCreation", has(0x20000000UL) ? "false" : "true");
            yield return A("NoCrawl", has(0x800000000UL) ? "true" : "false");
            yield return A("EnableSyndication", has(0x4000000000UL) ? "false" : "true");
            yield return A("PreserveEmptyValues", has(0x800000000000UL) ? "true" : "false");
            yield return A("ExcludeFromOfflineClient", has(0x2000000000000UL) ? "true" : "false");
            yield return A("EnforceDataValidation", has(0x4000000000000UL) ? "true" : "false");
            yield return A("DefaultItemOpenUseListSetting", has(0x8000000000000UL) ? "true" : "false");
            yield return A("DisableGridEditing", has(0x20000000000000UL) ? "true" : "false");
            yield return A("BrowserFileHandling", has(0x40000000000000UL) ? "Strict" : "Permissive");
            yield return A("NavigateForFormsPages", has(0x80000000000000UL) ? "true" : "false");
            yield return A("StrictTypeCoercion", has(0x100000000000000UL) ? "true" : "false");
            if (baseType == 1)
            {
                yield return A("ForceCheckout", has(0x40000UL) ? "true" : "false");
                yield return A("EnableMinorVersions", has(0x80000UL) ? "true" : "false");
                // Visibility restrictions apply only when drafts or moderation
                // are enabled. Keep the recorded editor/approver distinction.
                string visibility = "Reader";
                if (has(0x80000UL) || has(0x400UL))
                    visibility = has(0x100000UL) ? "Author" : has(0x200000UL) ? "Approver" : "Reader";
                yield return A("DraftVersionVisibility", visibility);
                yield return A("DefaultItemOpen", has(0x10000000UL) ? "Browser" : "PreferClient");
            }
        }
        private static readonly Dictionary<string, Guid> StandardFieldIds = new Dictionary<string, Guid>(StringComparer.Ordinal)
        {
            // Microsoft's MS-PRIMEPF3.1 Manifest example explicitly records
            // these standard field identities. Only template100/101 references
            // use this compatibility profile; custom field IDs are never made up.
            { "Title", new Guid("fa564e0f-0c70-4ab9-b863-0177e6ddd247") },
            { "_ModerationComments", new Guid("34ad21eb-75bd-4544-8c73-0e08330291fe") },
            { "Modified_x0020_By", new Guid("822c78e3-1ea9-4943-b449-57863ad33ca9") },
            { "Created_x0020_By", new Guid("4dd7e525-8d6b-4cb4-9d3e-44ee25f973eb") },
            { "File_x0020_Type", new Guid("39360f11-34cf-4356-9945-25c44e68dade") },
            { "HTML_x0020_File_x0020_Type", new Guid("0c5e0085-eb30-494b-9cdd-ece1d3c649a2") },
            { "_SourceUrl", new Guid("c63a459d-54ba-4ab7-933a-dcf1c6fadec2") },
            { "_SharedFileIndex", new Guid("034998e9-bf1c-4288-bbbd-00eacfc64410") },
            { "TemplateUrl", new Guid("4b1bf6c6-4f39-45ac-acd5-16fe7a214e5e") },
            { "xd_ProgID", new Guid("cd1ecb9f-dd4e-4f29-ab9e-e9ff40048d64") },
            { "xd_Signature", new Guid("fbf29b2d-cae5-49aa-8e0a-29955b540122") }
        };
        private static XElement ItemFields(MigrationListSnapshot snapshot, MigrationItemSnapshot item)
        {
            var fields = new XElement(Manifest + "Fields");
            foreach (var group in item.Values.GroupBy(v => v.Name, StringComparer.Ordinal))
            {
                MigrationFieldDefinition definition = snapshot.Fields.SingleOrDefault(f => String.Equals(f.Name, group.Key, StringComparison.Ordinal));
                List<MigrationFieldValue> values = group.OrderBy(v => v.Component).ToList();
                if (definition == null) throw new ContentUnavailableException("A stored item value lacks its field definition: " + group.Key);
                if (values.GroupBy(v => v.Component).Any(g => g.Count() != 1)) throw new InvalidDataException("Stored field components are duplicated: " + group.Key);
                if (definition.IsReference && !definition.Id.HasValue)
                {
                    if (values.All(v => v.IsNull)) continue;
                    object structural;
                    if (StructuralValue(item, definition.Name, out structural))
                    {
                        if (values.Count != 1 || values[0].Component != 1 || values[0].IsNull || ScalarText(values[0].Value) != ScalarText(structural))
                            throw new InvalidDataException("A duplicate system field disagrees with its canonical item metadata: " + definition.Name);
                        continue; // Already represented by exact canonical item/file attributes.
                    }
                }
                Guid fieldId;
                if (definition.Id.HasValue && definition.Id.Value != Guid.Empty) fieldId = definition.Id.Value;
                else if (!definition.IsReference || !HasStandardTemplate(snapshot.Metadata) || !StandardFieldIds.TryGetValue(definition.Name, out fieldId))
                    throw new ContentUnavailableException("A non-null field lacks a verified source/template identity: " + definition.Name);
                if (values.Any(v => v.FieldId.HasValue && v.FieldId.Value != fieldId)) throw new InvalidDataException("The stored field identity is inconsistent: " + definition.Name);
                if (values.Any(v => v.Component < 1 || v.Component > 2) || (values.Count > 1 && !String.Equals(definition.SharePointType, "URL", StringComparison.OrdinalIgnoreCase) && definition.Name != "_ShortcutUrl"))
                    throw new NotSupportedException("This field requires an unsupported compound-value serializer: " + group.Key);
                var field = new XElement(Manifest + "Field", A("Name", definition.Name), A("FieldId", fieldId));
                foreach (MigrationFieldValue value in values)
                    if (!value.IsNull) field.Add(A(value.Component == 2 ? "Value2" : "Value", FieldText(definition, value, snapshot)));
                fields.Add(field);
            }
            return fields;
        }
        private static bool HasStandardTemplate(MigrationListMetadata metadata)
        {
            return (metadata.BaseType == 0 && metadata.ServerTemplate == 100 && metadata.TemplateFeatureId == new Guid("00bfea71-de22-43b2-a848-c05709900100")) ||
                (metadata.BaseType == 1 && metadata.ServerTemplate == 101 && metadata.TemplateFeatureId == new Guid("00bfea71-e717-4e80-aa17-d0c71b360101"));
        }
        private static bool StructuralValue(MigrationItemSnapshot item, string name, out object value)
        {
            switch (name)
            {
                case "ContentTypeId": value = item.ContentTypeId; return true;
                case "FileLeafRef": value = item.Document.Name; return true;
                case "FileRef": value = item.Document.Path; return true;
                case "ID": value = item.ItemId; return true;
                case "GUID": value = item.ItemUniqueId; return true;
                case "UniqueId": value = item.DocumentId; return true;
                case "Author": value = item.AuthorId; return true;
                case "Editor": value = item.EditorId; return true;
                case "Created": value = item.Created; return true;
                case "Modified": value = item.Modified; return true;
                case "Attachments": value = item.HasAttachments; return true;
                case "FSObjType": value = item.IsFolder ? 1 : 0; return true;
                case "_UIVersion": value = item.Document.UiVersion; return true;
                case "_UIVersionString": value = VersionText(item.Document.UiVersion); return true;
                default: value = null; return false;
            }
        }
        private static string ScalarText(object value)
        {
            if (value is DateTime) return Date((DateTime)value);
            if (value is bool) return (bool)value ? "1" : "0";
            if (value is Guid) return ((Guid)value).ToString("D");
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }
        private static string FieldText(MigrationFieldDefinition field, MigrationFieldValue stored, MigrationListSnapshot snapshot)
        {
            object value = stored.Value; string type = field.SharePointType ?? String.Empty;
            if (type.Equals("Calculated", StringComparison.OrdinalIgnoreCase) || type.Equals("Computed", StringComparison.OrdinalIgnoreCase) || type.StartsWith("Taxonomy", StringComparison.OrdinalIgnoreCase) ||
                type.Equals("BusinessData", StringComparison.OrdinalIgnoreCase) || type.Equals("Geolocation", StringComparison.OrdinalIgnoreCase) || type.Equals("WorkflowStatus", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("This non-null field requires an additional migration-value serializer: " + field.Name);
            if (value is byte[]) throw new NotSupportedException("A binary item field cannot be represented by this migration serializer: " + field.Name);
            if (value is DateTime) return Date((DateTime)value);
            if (value is bool) return (bool)value ? "1" : "0";
            if (value is Guid) return ((Guid)value).ToString("D");
            if (value is double) return XmlConvert.ToString((double)value);
            if (value is float) return XmlConvert.ToString((float)value);
            if (value is decimal) return XmlConvert.ToString((decimal)value);
            string text = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (IsUserField(field)) ValidateReferenceIds(text, snapshot.Users.Select(u => u.Id).ToHashSet(), "user", field.Name);
            if (type.StartsWith("Lookup", StringComparison.OrdinalIgnoreCase))
            {
                if (field.LookupListId != snapshot.Metadata.ListId) throw new NotSupportedException("The lookup value references an external list: " + field.Name);
                ValidateReferenceIds(text, snapshot.Items.Select(i => i.ItemId).ToHashSet(), "lookup item", field.Name);
            }
            XmlConvert.VerifyXmlChars(text); return text;
        }
        private static bool IsUserField(MigrationFieldDefinition field)
        { return (field.SharePointType ?? String.Empty).StartsWith("User", StringComparison.OrdinalIgnoreCase) || field.Name == "Author" || field.Name == "Editor"; }
        private static void ValidateReferenceIds(string text, HashSet<int> available, string type, string field)
        {
            string[] parts = (text ?? String.Empty).Split(new[] { ";#" }, StringSplitOptions.None);
            for (int index = 0; index < parts.Length; index += 2)
            {
                int id;
                if (!Int32.TryParse(parts[index], NumberStyles.None, CultureInfo.InvariantCulture, out id) || id < 0 || (id > 0 && !available.Contains(id)))
                    throw new ContentUnavailableException("A field references an unavailable " + type + ": " + field);
            }
        }
        private static XElement Object(string type, Guid id, Guid parent, Guid webId, string web, string url, XElement content)
        { return new XElement(Manifest + "SPObject", A("Id", id), A("ObjectType", type), A("ParentId", parent), A("ParentWebId", webId), A("ParentWebUrl", web), A("Url", url), content); }
        private static XAttribute A(string name, object value)
        { return value == null ? null : new XAttribute(name, value is Guid ? ((Guid)value).ToString("D") : Convert.ToString(value, CultureInfo.InvariantCulture)); }
        private static string Date(DateTime value)
        { if (value == DateTime.MinValue) throw new ContentUnavailableException("A required source timestamp is absent."); return XmlConvert.ToString(DateTime.SpecifyKind(value, DateTimeKind.Utc), XmlDateTimeSerializationMode.Utc); }
        private static string VersionText(int uiVersion)
        { if (uiVersion < 0) throw new InvalidDataException("A source UI version is invalid."); return (uiVersion / 512).ToString(CultureInfo.InvariantCulture) + "." + (uiVersion % 512).ToString(CultureInfo.InvariantCulture); }
        private static string SourcePath(string path)
        {
            string value = (path ?? String.Empty).Replace('\\', '/'); Uri absolute;
            if (Uri.TryCreate(value, UriKind.Absolute, out absolute) && (absolute.Scheme == "http" || absolute.Scheme == "https")) value = absolute.AbsolutePath;
            string[] parts = value.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Any(p => p == "." || p == ".." || p.IndexOf('\0') >= 0)) throw new InvalidDataException("A source URL contains an unsafe path component.");
            return "/" + String.Join("/", parts);
        }
        private static string WebRelative(string path, string web)
        {
            string normalized = SourcePath(path), prefix = SourcePath(web).TrimEnd('/');
            if (prefix.Length > 0 && !normalized.StartsWith(prefix + "/", StringComparison.OrdinalIgnoreCase) && !String.Equals(normalized, prefix, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("The source URL is outside its parent web.");
            return normalized.Substring(prefix.Length).TrimStart('/');
        }
        private static string ExportSiteUrl(string stored)
        {
            Uri absolute;
            if (Uri.TryCreate(stored, UriKind.Absolute, out absolute) && (absolute.Scheme == "http" || absolute.Scheme == "https")) return stored;
            return SourcePath(stored); // This restored SQL source records only the relative site path.
        }
        private static string ParentPath(string path)
        { string value = SourcePath(path); int split = value.LastIndexOf('/'); return split <= 0 ? "/" : value.Substring(0, split); }
        private static XElement WebPartElement(MigrationViewSnapshot view, MigrationListSnapshot snapshot)
        {
            if (!view.Level.HasValue || (view.Level != 1 && view.Level != 2) || !view.Version.HasValue) throw new ContentUnavailableException("A public view lacks its stored publishing level/version.");
            XElement definition = ViewElement(view);
            var element = new XElement(Manifest + "WebPart", A("Name", view.Id), A("Level", view.Level.Value == 1 ? "major" : "minor"),
                A("ListId", snapshot.Metadata.ListId), A("ListRootFolderUrl", SourcePath(snapshot.Metadata.RootFolderUrl)), A("Version", view.Version.Value), A("Type", view.PageType.Value));
            var shared = new HashSet<string>(new[] { "Flags", "Personal", "Hidden", "Threaded", "FPModified", "ReadOnly", "RecurrenceRowset", "ModerationType", "OrderedView", "Scope", "DisplayName", "BaseViewID", "WebPartZoneID", "WebPartTypeId", "IsIncluded", "WebPartOrder", "FrameState", "WebPartIdProperty" });
            foreach (XAttribute attribute in definition.Attributes().Where(a => shared.Contains(a.Name.LocalName))) element.Add(new XAttribute(attribute));
            foreach (XElement child in definition.Elements()) element.Add(new XElement(child));
            if (!String.IsNullOrEmpty(view.ContentTypeId)) element.Add(A("ContentTypeId", view.ContentTypeId));
            return element;
        }
        private static XElement ViewElement(MigrationViewSnapshot view)
        {
            if (view.Id == Guid.Empty || !view.Flags.HasValue || !view.PageType.HasValue || view.PageType.Value > 1 || !view.FileId.HasValue ||
                view.FileId.Value == Guid.Empty || String.IsNullOrEmpty(view.Url)) throw new ContentUnavailableException("A public view lacks its flags, page type or file identity.");
            uint flags = view.Flags.Value;
            var element = ReadCaml(view.SchemaXml ?? String.Empty, "View");
            Set(element, "Name", view.Id); Set(element, "FileId", view.FileId.Value); Set(element, "Url", SourcePath(view.Url)); Set(element, "Flags", flags);
            Set(element, "DefaultView", view.PageType.Value == 0 ? "true" : "false");
            Set(element, "MobileView", Bit(flags, 0x00800000)); Set(element, "MobileDefaultView", Bit(flags, 0x01000000));
            Set(element, "DefaultViewForContentType", Bit(flags, 0x10000000)); Set(element, "HackLockWeb", Bit(flags, 0x00000010));
            Set(element, "FailIfEmpty", Bit(flags, 0x00000040)); Set(element, "FreeForm", Bit(flags, 0x00000080)); Set(element, "FileDialog", Bit(flags, 0x00000100));
            Set(element, "AggregateView", Bit(flags, 0x00000400)); Set(element, "IncludeRootFolder", Bit(flags, 0x08000000)); Set(element, "IncludeVersions", Bit(flags, 0x02000000));
            Set(element, "TabularView", Bit(flags, 0x00000004)); Set(element, "Hidden", Bit(flags, 0x00000008)); Set(element, "FPModified", Bit(flags, 0x00000002));
            Set(element, "ReadOnly", Bit(flags, 0x00000020)); Set(element, "RecurrenceRowset", Bit(flags, 0x00002000)); Set(element, "Threaded", Bit(flags, 0x00010000));
            Set(element, "Personal", Bit(flags, 0x00040000)); Set(element, "OrderedView", Bit(flags, 0x00400000));
            Set(element, "ModerationType", (flags & 0x00008000) != 0 ? "Moderator" : (flags & 0x00004000) != 0 ? "Contributor" : String.Empty);
            if (view.DisplayName != null) Set(element, "DisplayName", view.DisplayName);
            if (view.BaseViewId.HasValue) Set(element, "BaseViewID", view.BaseViewId.Value);
            if (view.IsIncluded.HasValue) Set(element, "IsIncluded", view.IsIncluded.Value ? "true" : "false");
            if (view.PartOrder.HasValue) Set(element, "WebPartOrder", view.PartOrder.Value);
            if (view.ZoneId != null) Set(element, "WebPartZoneID", view.ZoneId);
            if (view.FrameState.HasValue) Set(element, "FrameState", view.FrameState.Value);
            if (view.WebPartTypeId.HasValue) Set(element, "WebPartTypeId", view.WebPartTypeId.Value);
            if (!String.IsNullOrEmpty(view.ContentTypeId)) Set(element, "ContentTypeID", view.ContentTypeId);
            return element;
        }
        private static string Bit(uint flags, uint mask) { return (flags & mask) != 0 ? "true" : "false"; }
        private static void Set(XElement element, string name, object value)
        {
            string expected = value is Guid ? ((Guid)value).ToString("D") : Convert.ToString(value, CultureInfo.InvariantCulture);
            XAttribute existing = element.Attribute(name);
            if (existing != null && !String.Equals(existing.Value, expected, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Stored view XML conflicts with its source metadata: " + name);
            element.SetAttributeValue(name, expected);
        }
        private static XElement FieldDefinitions(MigrationListSnapshot snapshot)
        {
            var root = new XElement(Manifest + "Fields");
            var builtins = new HashSet<string>(new[] { "ID", "GUID", "Title", "Author", "Editor", "Created", "Modified", "Attachments", "ContentTypeId", "ContentType",
                "_ModerationComments", "_ModerationStatus", "_UIVersion", "_UIVersionString", "FileLeafRef", "FileRef", "FileDirRef", "FSObjType", "UniqueId",
                "Modified_x0020_By", "Created_x0020_By", "File_x0020_Type", "HTML_x0020_File_x0020_Type", "_SourceUrl", "_SharedFileIndex", "ComplianceAssetId",
                "TemplateUrl", "xd_ProgID", "xd_Signature", "_ShortcutUrl", "_ShortcutSiteId", "_ShortcutWebId", "_ShortcutUniqueId" }, StringComparer.Ordinal);
            foreach (MigrationFieldDefinition field in snapshot.Fields)
            {
                XElement stored = ReadCaml(field.SchemaXml, field.IsReference ? "FieldRef" : "Field");
                if (!field.IsReference) { root.Add(stored); continue; }
                // Storage-only references to built-in fields are supplied by the
                // matching list template. Custom references/overrides cannot be
                // converted into invented full field definitions.
                if (!builtins.Contains(field.Name) || stored.HasElements || stored.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name.LocalName != "Name" && a.Name.LocalName != "ID" &&
                    !a.Name.LocalName.StartsWith("ColName", StringComparison.Ordinal) && !a.Name.LocalName.StartsWith("RowOrdinal", StringComparison.Ordinal)))
                    throw new NotSupportedException("A referenced field needs a complete source definition or template override serializer: " + field.Name);
            }
            return root;
        }
        private static XElement ContentTypeDefinitions(MigrationListSnapshot snapshot)
        {
            XElement contentTypes = ReadCaml(snapshot.Metadata.ContentTypesXml, "ContentTypes");
            foreach (XElement contentType in contentTypes.Elements(Manifest + "ContentType"))
            {
                XAttribute delayed = contentType.Attribute("DelayActivateTemplateBinding");
                if (delayed == null) continue;
                bool library = snapshot.Metadata.BaseType == 1;
                string baseId = library ? "0x0101" : "0x01", id = (string)contentType.Attribute("ID") ?? String.Empty;
                bool knownId = String.Equals(id, baseId, StringComparison.OrdinalIgnoreCase) ||
                    (id.Length == baseId.Length + 34 && id.StartsWith(baseId + "00", StringComparison.OrdinalIgnoreCase) && id.Substring(baseId.Length + 2).All(Uri.IsHexDigit));
                Guid feature;
                bool standard = SupportedSource(snapshot.Metadata) && HasStandardTemplate(snapshot.Metadata) && knownId &&
                    Guid.TryParse((string)contentType.Attribute("FeatureId"), out feature) && feature == Guid.Parse("695b6570-a48b-4a8e-8ea5-26ea7fc1d162") &&
                    (string)contentType.Attribute("Name") == (library ? "$Resources:core,Document;" : "$Resources:core,Item;");
                if (!standard || delayed.Value != "GROUP,SPSPERS,SITEPAGEPUBLISHING")
                    throw new NotSupportedException("This content type has no verified source-to-deployment conversion for DelayActivateTemplateBinding: " + id);
                // The published MS-WSSCAML deployment content-type schema does
                // not define this source provisioning marker. Convert only this
                // exact standard-template case of a supported source generation;
                // retain original XML and the conversion in the audit. Other
                // markers fail, never disappear.
                delayed.Remove();
            }
            return contentTypes;
        }
        private static XElement ReadCaml(string xml, string expected)
        {
            XElement envelope;
            using (XmlReader reader = XmlReader.Create(new StringReader("<" + expected + ">" + xml + "</" + expected + ">"), SafeReader())) envelope = XElement.Load(reader, LoadOptions.PreserveWhitespace);
            XElement source = envelope.Elements().Count() == 1 && envelope.Elements().Single().Name.LocalName == expected ? envelope.Elements().Single() : envelope;
            foreach (XElement element in source.DescendantsAndSelf())
            {
                string ns = element.Name.NamespaceName;
                if (ns.Length == 0 || ns == "http://schemas.microsoft.com/sharepoint/soap/") element.Name = Manifest + element.Name.LocalName;
                foreach (XAttribute declaration in element.Attributes().Where(a => a.IsNamespaceDeclaration && (a.Value.Length == 0 || a.Value == "http://schemas.microsoft.com/sharepoint/soap/")).ToList()) declaration.Remove();
            }
            return source;
        }
        private static void Save(string directory, string name, XElement root)
        {
            using (var output = new FileStream(Path.Combine(directory, name), FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                using (XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = true, CloseOutput = false })) new XDocument(new XDeclaration("1.0", "utf-8", null), root).Save(writer);
                output.Flush(true);
            }
        }
        private static void WriteAudit(string source, MigrationListSnapshot snapshot, MigrationHistorySnapshot history, List<Payload> payloads, string directory, bool includeHistory)
        {
            var audit = new { Format = "MS-PRIMEPF content deployment", Source = source, SiteId = snapshot.Metadata.SiteId, WebId = snapshot.Metadata.WebId, ListId = snapshot.Metadata.ListId,
                CurrentItems = snapshot.Items.Count, Files = payloads.Count(p => p.Owner == null), Attachments = payloads.Count(p => p.Owner != null),
                Security = "None", Versions = includeHistory ? "All" : "CurrentVersion", RetainedHistory = history, TemplateForms = "Created by the target list template; original default form file bytes are not included.",
                SourceContentTypesXml = snapshot.Metadata.ContentTypesXml, SourceFieldSchemaXml = snapshot.Metadata.FieldSchemaXml,
                SourceListFlags = unchecked((ulong)snapshot.Metadata.Flags), SourceListFlags2 = unchecked((ulong)snapshot.Metadata.Flags2),
                SourceSystemObjects = DeploymentContext(snapshot).ToArray(), SourceStoredSystemObjects = snapshot.SystemObjects.ToArray(), SourceSystemContextComplete = snapshot.SystemObjects.Count > 0,
                CamlConversions = ReadCaml(snapshot.Metadata.ContentTypesXml, "ContentTypes").Elements(Manifest + "ContentType")
                    .Where(contentType => contentType.Attribute("DelayActivateTemplateBinding") != null).Select(contentType => new
                    {
                        ContentTypeId = (string)contentType.Attribute("ID"), Attribute = "DelayActivateTemplateBinding",
                        SourceValue = (string)contentType.Attribute("DelayActivateTemplateBinding"), DeploymentValue = "omitted",
                        Reason = "Bounded standard template100/101 conversion for a supported source generation; this source marker is absent from the published deployment ContentType schema. Raw source XML is retained; target import is unverified."
                    }).ToArray(),
                TemplateFiles = snapshot.TemplateFiles.Select(t => new { DocumentId = t.Document.Id, SourcePath = t.Document.Path, t.SetupPath, t.SetupPathVersion, t.SetupPathUser }).ToArray(),
                FarmImportVerified = false, Validation = "Published XML schemas, object relationships and exact recovered payload length/SHA256",
                Limitations = new[] { "Target farm import has not been tested.", includeHistory ? "Permissions, workflow state, historical attachment sets and external lookup dependencies are not included." : "Permissions, workflow state, historical item/file versions and external lookup dependencies are not included.", "Custom non-item pages and custom form/view behavior are not reconstructed." },
                Payloads = payloads.Select(p => new { p.Name, DocumentId = p.Document.Id, SourcePath = p.Document.Path, UiVersion = p.Document.UiVersion, HistoryVersion = p.Document.HistoryVersion, Level = p.Document.Level, InternalVersion = p.Document.InternalVersion, OwnerDocumentId = p.Owner == null ? (Guid?)null : p.Owner.Id, p.Bytes, p.Sha256, p.Decoder }).ToArray() };
            File.WriteAllText(Path.Combine(directory, "package-audit.json"), JsonSerializer.Serialize(audit, new JsonSerializerOptions { WriteIndented = true }), new UTF8Encoding(false));
            var csv = new StringBuilder("Source,SiteId,DocumentId,SourcePath,Kind,Bytes,Sha256,PackageFile,Decoder\r\n");
            foreach (Payload p in payloads) csv.AppendLine(String.Join(",", new[] { source, p.Document.SiteId.ToString("D"), p.Document.Id.ToString("D"), p.Document.Path,
                p.Owner == null ? "File" : "Attachment", p.Bytes.ToString(CultureInfo.InvariantCulture), p.Sha256, p.Name, p.Decoder }.Select(Csv)));
            File.WriteAllText(Path.Combine(directory, "export-report.csv"), csv.ToString(), new UTF8Encoding(false));
            File.WriteAllText(Path.Combine(directory, "IMPORT.txt"), "Uncompressed SharePoint content-deployment package.\r\n" +
                "On a compatible SharePoint Server target, review the package and use Import-SPWeb -Identity <target-web-url> -Path <this-folder> -NoFileCompression.\r\n" +
                "Security is not included. Current items, files and supported ordinary-list attachments only.\r\n" +
                "Schema, relationships and payload checks passed. A target-farm import has not been performed. Back up the target and review its templates/features and build compatibility before import.\r\n", new UTF8Encoding(false));
        }
        private static string Csv(string value)
        { return "\"" + (value ?? String.Empty).Replace("\"", "\"\"") + "\""; }
        private static XmlReaderSettings SafeReader()
        { return new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 128L * 1024 * 1024 }; }

        public static void ValidatePackage(string packagePath)
        {
            string root = Path.GetFullPath(packagePath);
            if (!Directory.Exists(root)) throw new DirectoryNotFoundException("The package folder is absent.");
            if ((File.GetAttributes(root) & FileAttributes.ReparsePoint) != 0) throw new InvalidDataException("A package folder cannot be a filesystem redirect.");
            var schemaErrors = new HashSet<string>(StringComparer.Ordinal);
            int schemaErrorCount = 0;
            for (int index = 0; index < XmlFiles.Length; index++)
            {
                string file = Path.Combine(root, XmlFiles[index]); RequireRegularFile(file);
                XDocument document = ReadXml(file); XmlSchemaSet schemas = LoadSchema(SchemaFiles[index]);
                document.Validate(schemas, (sender, args) =>
                {
                    schemaErrorCount++;
                    if (schemaErrors.Count < 50)
                    {
                        XObject node = sender as XObject;
                        XElement element = node as XElement ?? (node as XAttribute)?.Parent;
                        string context = element == null ? String.Empty : " [" + element.Name.LocalName + "]";
                        schemaErrors.Add(XmlFiles[index] + ": " + args.Message + context);
                    }
                }, true);
            }
            if (schemaErrorCount != 0)
                throw new InvalidDataException("Package XML schema validation failed (" + schemaErrorCount.ToString(CultureInfo.InvariantCulture) +
                    " errors; up to50 distinct diagnostics):" + Environment.NewLine + String.Join(Environment.NewLine, schemaErrors));
            XDocument manifest = ReadXml(Path.Combine(root, "Manifest.xml"));
            XDocument system = ReadXml(Path.Combine(root, "SystemData.xml"));
            var lists = manifest.Root.Elements().Select(entry => entry.Elements().Single()).Where(element => element.Name == Manifest + "List" || element.Name == Manifest + "DocumentLibrary").ToList();
            var rootOnlyExpected = new HashSet<Guid>();
            foreach (XElement list in lists)
            {
                ulong flags = UInt64.Parse((string)list.Attribute("Flags"), CultureInfo.InvariantCulture);
                int baseType = list.Name == Manifest + "DocumentLibrary" ? 1 : 0;
                foreach (XAttribute setting in ListSettings(unchecked((long)flags), baseType))
                    if ((string)list.Attribute(setting.Name) != setting.Value)
                        throw new InvalidDataException("A deployment list setting disagrees with its recorded flags: " + setting.Name);
                if ((flags & 0x4000UL) != 0) rootOnlyExpected.Add(Guid.Parse((string)list.Attribute("Id")));
            }
            XElement rootOnly = system.Root.Elements().SingleOrDefault(element => element.Name.LocalName == "RootWebOnlyLists");
            var rootOnlyActual = new HashSet<Guid>();
            if (rootOnly != null)
                foreach (XElement list in rootOnly.Elements())
                    if (!rootOnlyActual.Add(Guid.Parse((string)list.Attribute("Id")))) throw new InvalidDataException("Duplicate root-web-only list restriction.");
            if (!rootOnlyExpected.SetEquals(rootOnlyActual)) throw new InvalidDataException("Root-web-only list restrictions disagree with the source flags.");
            var known = new HashSet<Guid>();
            foreach (XElement entry in system.Descendants().Where(e => e.Name.LocalName == "SystemObject"))
                if (!known.Add(Guid.Parse((string)entry.Attribute("Id")))) throw new InvalidDataException("Duplicate package system object.");
            foreach (XElement item in manifest.Root.Elements())
            {
                Guid id = Guid.Parse((string)item.Attribute("Id")), parent = Guid.Parse((string)item.Attribute("ParentId"));
                if (!known.Contains(parent) || !known.Contains(Guid.Parse((string)item.Attribute("ParentWebId")))) throw new InvalidDataException("An exported object precedes or lacks its parent.");
                if (!known.Add(id)) throw new InvalidDataException("A package object identity is duplicated.");
                if (item.Elements().Count() != 1 || !String.Equals((string)item.Elements().Single().Attribute("Id"), id.ToString("D"), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("Object and content identities disagree.");
            }
            var referenced = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var provenance = new Dictionary<string, Tuple<Guid, Guid?, string, bool>>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement content in manifest.Descendants().Where(e => e.Name.LocalName == "File" || e.Name.LocalName == "Attachment"))
            {
                string name = (string)content.Attribute("FileValue");
                if (String.IsNullOrEmpty(name) && content.Name == Manifest + "File" && (string)content.Attribute("IsGhosted") == "true")
                {
                    string setup = (string)content.Attribute("SetupPath");
                    if (String.IsNullOrWhiteSpace(setup) || setup.StartsWith("/", StringComparison.Ordinal) || setup.Contains(":")) throw new InvalidDataException("A ghost file lacks its local template reference.");
                    SourcePath(setup);
                    continue;
                }
                if (String.IsNullOrEmpty(name) && content.Name == Manifest + "File" && content.Element(Manifest + "Versions") != null) continue;
                if (String.IsNullOrEmpty(name)) throw new InvalidDataException("A package document lacks its required payload reference.");
                if (name.Length != 12 || !name.EndsWith(".dat", StringComparison.Ordinal) || name.Substring(0, 8).Any(c => !Uri.IsHexDigit(c)) || !referenced.Add(name))
                    throw new InvalidDataException("A package payload name is unsafe or duplicated.");
                RequireRegularFile(Path.Combine(root, name));
                Guid documentId = Guid.Parse((string)content.Attribute("Id"));
                XElement owner = content.Ancestors(Manifest + "ListItem").FirstOrDefault();
                Guid? ownerId = owner == null ? (Guid?)null : Guid.Parse((string)owner.Attribute("DocId"));
                provenance.Add(name, Tuple.Create(documentId, ownerId, (string)content.Attribute("Version"), content.Ancestors(Manifest + "Versions").Any() && VersionNumber((string)content.Attribute("Version")) < VersionNumber((string)content.Parent.Parent.Attribute("Version"))));
            }
            foreach (XElement item in manifest.Descendants(Manifest + "ListItem").Where(item => !item.Ancestors(Manifest + "Versions").Any()))
            {
                if (!known.Contains(Guid.Parse((string)item.Attribute("ParentFolderId"))) || !known.Contains(Guid.Parse((string)item.Attribute("ParentListId"))) || String.IsNullOrEmpty((string)item.Attribute("DirName")))
                    throw new InvalidDataException("A list-item folder/list relationship or import-required directory is missing.");
                if ((string)item.Attribute("DocType") == "Folder" || manifest.Descendants(Manifest + "DocumentLibrary").Any())
                    if (!known.Contains(Guid.Parse((string)item.Attribute("DocId")))) throw new InvalidDataException("A library or folder item lacks its document object.");
            }
            XDocument roots = ReadXml(Path.Combine(root, "RootObjectMap.xml"));
            foreach (XElement mapping in roots.Root.Elements())
                if (!known.Contains(Guid.Parse((string)mapping.Attribute("Id"))) || !known.Contains(Guid.Parse((string)mapping.Attribute("ParentId")))) throw new InvalidDataException("The root mapping references an absent object.");
            XNamespace systemNs = "urn:deployment-systemdata-schema";
            if (system.Descendants(systemNs + "ManifestFile").Select(e => (string)e.Attribute("Name")).SingleOrDefault() != "Manifest.xml")
                throw new InvalidDataException("The system manifest map does not identify Manifest.xml.");
            if (Directory.GetFiles(root, "*.dat").Select(Path.GetFileName).Any(name => !referenced.Contains(name))) throw new InvalidDataException("An unreferenced package payload is present.");
            XDocument viewMap = ReadXml(Path.Combine(root, "ViewFormsList.xml"));
            var viewIds = viewMap.Root.Elements().Where(e => (string)e.Attribute("Type") == "View").Select(e => Guid.Parse((string)e.Attribute("Id"))).ToHashSet();
            var referencedViews = new HashSet<Guid>();
            foreach (XElement view in manifest.Descendants(Manifest + "View"))
            {
                Guid id = Guid.Parse((string)view.Attribute("Name")), fileId = Guid.Parse((string)view.Attribute("FileId"));
                XElement file = manifest.Descendants(Manifest + "File").SingleOrDefault(f => !f.Ancestors(Manifest + "Versions").Any() && Guid.Parse((string)f.Attribute("Id")) == fileId);
                if (!viewIds.Contains(id) || !referencedViews.Add(id) || file == null || !file.Descendants(Manifest + "WebPart").Any(p => Guid.Parse((string)p.Attribute("Name")) == id))
                    throw new InvalidDataException("A view lacks its matching map, file or Web Part relationship.");
            }
            if (!referencedViews.SetEquals(viewIds)) throw new InvalidDataException("The view map contains an unreferenced view.");
            XDocument settings = ReadXml(Path.Combine(root, "ExportSettings.xml"));
            if ((string)settings.Root.Attribute("IncludeSecurity") != "None" || ((string)settings.Root.Attribute("IncludeVersions") != "CurrentVersion" && (string)settings.Root.Attribute("IncludeVersions") != "All") ||
                (string)settings.Root.Attribute("ExportFrontEndFileStreams") != "false") throw new InvalidDataException("The package settings claim unsupported security/version/template-stream coverage.");
            XDocument userMap = ReadXml(Path.Combine(root, "UserGroup.xml"));
            var userIds = userMap.Descendants().Where(e => e.Name.LocalName == "User").Select(e => Int32.Parse((string)e.Attribute("Id"), CultureInfo.InvariantCulture)).ToHashSet();
            foreach (XAttribute reference in manifest.Descendants().SelectMany(e => e.Attributes()).Where(a => a.Name.LocalName == "Author" || a.Name.LocalName == "ModifiedBy" || a.Name.LocalName == "SetupPathUser"))
            {
                int id;
                if (Int32.TryParse(reference.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out id) && id > 0 && !userIds.Contains(id)) throw new InvalidDataException("An exported author/editor/setup user is absent from the user map.");
            }
            ValidateRetainedVersions(manifest, settings);
            ValidateAudit(root, referenced, provenance, system, settings, manifest);
        }
        private static void ValidateAudit(string root, HashSet<string> referenced, Dictionary<string, Tuple<Guid, Guid?, string, bool>> provenance, XDocument system, XDocument settings, XDocument manifest)
        {
            string path = Path.Combine(root, "package-audit.json"); RequireRegularFile(path);
            using (JsonDocument audit = JsonDocument.Parse(File.ReadAllText(path)))
            {
                var recordedContext = audit.RootElement.GetProperty("SourceSystemObjects").EnumerateArray().ToList();
                XElement[] declaredContext = system.Descendants().Where(element => element.Name.LocalName == "SystemObject").ToArray();
                if (recordedContext.Count != declaredContext.Length) throw new InvalidDataException("Deployment system context disagrees with the source audit.");
                foreach (JsonElement context in recordedContext)
                {
                    Guid id = context.GetProperty("Id").GetGuid();
                    XElement declared = declaredContext.SingleOrDefault(element => (string)element.Attribute("Id") == id.ToString("D"));
                    if (declared == null || (string)declared.Attribute("Type") != context.GetProperty("Type").GetString() || (string)declared.Attribute("Url") != context.GetProperty("Url").GetString())
                        throw new InvalidDataException("A verified deployment system identity/type/URL changed.");
                }
                if (audit.RootElement.GetProperty("SourceSystemContextComplete").GetBoolean())
                {
                    string[] roles = recordedContext.Select(context => context.GetProperty("Role").GetString()).ToArray();
                    if (roles.Distinct(StringComparer.Ordinal).Count() != roles.Length || !roles.Contains("RootWeb") || !roles.Contains("UserInfoList") || !roles.Contains("RootParentFolder") ||
                        roles.Any(role => role != "RootWeb" && role != "SelectedWeb" && role != "UserInfoList" && role != "RootParentFolder"))
                        throw new InvalidDataException("The source deployment context audit lacks its required roles.");
                }
                if (audit.RootElement.GetProperty("Versions").GetString() != (string)settings.Root.Attribute("IncludeVersions")) throw new InvalidDataException("Package history coverage disagrees with its source audit.");
                ValidateHistoryAudit(audit.RootElement.GetProperty("RetainedHistory"), manifest);
                var audited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (JsonElement payload in audit.RootElement.GetProperty("Payloads").EnumerateArray())
                {
                    string name = payload.GetProperty("Name").GetString();
                    if (!referenced.Contains(name) || !audited.Add(name)) throw new InvalidDataException("The package audit has an invalid payload reference.");
                    Tuple<Guid, Guid?, string, bool> identity = provenance[name];
                    JsonElement ownerElement = payload.GetProperty("OwnerDocumentId");
                    Guid? ownerId = ownerElement.ValueKind == JsonValueKind.Null ? (Guid?)null : ownerElement.GetGuid();
                    if (payload.GetProperty("DocumentId").GetGuid() != identity.Item1 || ownerId != identity.Item2)
                        throw new InvalidDataException("The manifest assigns a verified payload to a different document or attachment owner.");
                    if (identity.Item3 != null && identity.Item3 != VersionText(payload.GetProperty("UiVersion").GetInt32())) throw new InvalidDataException("A historical payload is assigned to a different UI version.");
                    if (identity.Item4 != (payload.GetProperty("HistoryVersion").GetInt32() > 0)) throw new InvalidDataException("A current and historical payload were interchanged.");
                    string file = Path.Combine(root, name);
                    if (new FileInfo(file).Length != payload.GetProperty("Bytes").GetInt64()) throw new InvalidDataException("A package payload length disagrees with its recovery audit.");
                    using (FileStream input = File.OpenRead(file))
                    {
                        string hash = Convert.ToHexString(SHA256.HashData(input)).ToLowerInvariant();
                        if (!String.Equals(hash, payload.GetProperty("Sha256").GetString(), StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("A package payload hash disagrees with its recovery audit.");
                    }
                }
                if (!audited.SetEquals(referenced)) throw new InvalidDataException("The package audit omits a referenced payload.");
            }
        }
        private static void RequireRegularFile(string file)
        {
            if (!File.Exists(file) || (File.GetAttributes(file) & (FileAttributes.ReparsePoint | FileAttributes.Directory)) != 0)
                throw new InvalidDataException("A required package file is missing or redirected: " + Path.GetFileName(file));
        }
        private static XDocument ReadXml(string path)
        { using (XmlReader reader = XmlReader.Create(path, SafeReader())) return XDocument.Load(reader, LoadOptions.None); }
        private static XmlSchemaSet LoadSchema(string name)
        {
            XmlSchema schema;
            using (Stream input = Resource(name)) schema = XmlSchema.Read(input, (sender, args) => { throw new InvalidDataException(args.Message); });
            foreach (XmlSchemaInclude include in schema.Includes)
            {
                string filename = include.SchemaLocation;
                if (filename != "wsswire_" + name) throw new InvalidDataException("Unexpected deployment schema include.");
                using (Stream input = Resource(filename)) include.Schema = XmlSchema.Read(input, (sender, args) => { throw new InvalidDataException(args.Message); });
                include.SchemaLocation = null;
            }
            var schemas = new XmlSchemaSet { XmlResolver = null }; schemas.Add(schema); schemas.Compile(); return schemas;
        }
        private static Stream Resource(string filename)
        {
            Assembly assembly = typeof(MigrationPackageExporter).Assembly;
            string name = assembly.GetManifestResourceNames().SingleOrDefault(n => n.EndsWith(".Schemas." + filename, StringComparison.Ordinal));
            if (name == null) throw new InvalidOperationException("The deployment schema resource is absent: " + filename);
            return assembly.GetManifestResourceStream(name);
        }
        private static void DeleteOwnedStaging(string parent, string staging)
        {
            string full = Path.GetFullPath(staging);
            if (!String.Equals(Path.GetDirectoryName(full), Path.TrimEndingDirectorySeparator(Path.GetFullPath(parent)), StringComparison.OrdinalIgnoreCase) ||
                !Path.GetFileName(full).StartsWith(".", StringComparison.Ordinal) || !full.EndsWith(".staging", StringComparison.Ordinal) || (File.GetAttributes(full) & FileAttributes.ReparsePoint) != 0)
                throw new IOException("The package staging path failed the cleanup boundary check.");
            var pending = new Stack<string>(); pending.Push(full);
            while (pending.Count > 0)
                foreach (string entry in Directory.EnumerateFileSystemEntries(pending.Pop()))
                {
                    FileAttributes attributes = File.GetAttributes(entry);
                    if ((attributes & FileAttributes.ReparsePoint) != 0) throw new IOException("The private package staging folder contains a filesystem redirect.");
                    if ((attributes & FileAttributes.Directory) != 0) pending.Push(entry);
                }
            Directory.Delete(full, true);
        }
        private sealed class CancellationWriteStream : Stream
        {
            private readonly Stream inner; private readonly CancellationToken token;
            internal CancellationWriteStream(Stream inner, CancellationToken token) { this.inner = inner; this.token = token; }
            public override void Write(byte[] buffer, int offset, int count) { token.ThrowIfCancellationRequested(); inner.Write(buffer, offset, count); }
            public override void Flush() { token.ThrowIfCancellationRequested(); inner.Flush(); }
            public override bool CanRead { get { return false; } }
            public override bool CanSeek { get { return false; } }
            public override bool CanWrite { get { return true; } }
            public override long Length { get { return inner.Length; } }
            public override long Position { get { return inner.Position; } set { throw new NotSupportedException(); } }
            public override int Read(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
            public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
            public override void SetLength(long value) { throw new NotSupportedException(); }
        }
    }
}