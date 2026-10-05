using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace SharePointExplorer.Tests
{
    // SharePoint generation boundaries, optional-column SQL composition and, with
    // SQL, the older-generation column profile against the restored database.
    internal static class ContentSchemaChecks
    {
        public static void Run()
        {
            var expected = new Dictionary<string, SharePointGeneration>
            {
                { "14.0.7015.1000", SharePointGeneration.SharePoint2013OrEarlier }, { "15.0.4569.1000", SharePointGeneration.SharePoint2013OrEarlier },
                { "16.0.4351.1000", SharePointGeneration.SharePoint2016 }, { "16.0.10337.12108", SharePointGeneration.SharePoint2016 },
                { "16.0.10337.12109", SharePointGeneration.SharePoint2019 }, { "16.0.14131.20291", SharePointGeneration.SharePoint2019 },
                { "16.0.14131.20292", SharePointGeneration.SubscriptionEdition }, { "16.0.14326.20450", SharePointGeneration.SubscriptionEdition },
                { "16.0.17928.20238", SharePointGeneration.SubscriptionEdition }, { "17.0.1.1", SharePointGeneration.Unknown }, { "not a build", SharePointGeneration.Unknown }
            };
            foreach (var pair in expected)
                Check(SharePointBuilds.Classify(pair.Key) == pair.Value, "Build " + pair.Key + " was classified as " + SharePointBuilds.Classify(pair.Key) + ".");
            Check(!SharePointBuilds.IsSupported(SharePointGeneration.SharePoint2013OrEarlier) && !SharePointBuilds.IsSupported(SharePointGeneration.Unknown) &&
                SharePointBuilds.IsSupported(SharePointGeneration.SharePoint2016) && SharePointBuilds.IsSupported(SharePointGeneration.SharePoint2019) &&
                SharePointBuilds.IsSupported(SharePointGeneration.SubscriptionEdition), "Supported SharePoint generations changed.");
            Check(SqlRepository.SizeExpression("d", false, false, true) == "COALESCE(CONVERT(bigint,d.Size),0)" &&
                SqlRepository.SizeExpression("d", true, true, true) == "COALESCE(CONVERT(bigint,d.SizeRead),CONVERT(bigint,d.Size),CONVERT(bigint,d.SizeWrite),0)" &&
                SqlRepository.SizeExpression("v", true, false, false) == "COALESCE(CONVERT(bigint,v.SizeRead),CONVERT(bigint,v.Size))" &&
                SqlRepository.SizeExpression("", false, false, false) == "CONVERT(bigint,Size)", "Optional size columns were composed incorrectly.");
            Console.WriteLine("PASS SharePoint generation boundaries and optional size-column SQL composition");
        }

        // Every query that adapts to optional columns runs with the detected SE
        // profile and with the older profile (no SizeRead/SizeWrite/ABS columns).
        // Results must be identical; the SE source records equal sizes.
        public static void RunSql(SqlConnectionOptions options)
        {
            var full = new SqlRepository(options); full.ValidateSchema();
            var older = new SqlRepository(options); older.ValidateSchema(); older.UseColumns(new ContentColumns());
            Check(full.Generation == SharePointGeneration.SubscriptionEdition && full.SourceBuild == "16.0.14326.20450", "The restored database generation changed.");
            ContentColumns detected = full.Columns;
            Check(detected.DocsSizeRead && detected.DocsSizeWrite && detected.VersionsSizeRead && detected.StreamCompressedSize && detected.StreamAbsId, "The SE optional-column profile was not detected.");
            int nodes = 0, files = 0, chunks = 0, versions = 0, attachments = 0, deleted = 0, packages = 0;
            Same(Keys(full.GetRootSites()), Keys(older.GetRootSites()), "root sites");
            var pending = new Queue<Node>(full.GetRootSites());
            while (pending.Count > 0)
            {
                Node parent = pending.Dequeue();
                List<Node> children = full.GetChildren(parent);
                Same(Keys(children), Keys(older.GetChildren(parent)), "children of " + parent.Path);
                nodes += children.Count;
                foreach (Node child in children)
                {
                    if (child.IsContainer) pending.Enqueue(child);
                    if (child.Kind == NodeKind.Library)
                        Same(Keys(full.EnumerateCurrentLibraryFiles(child)), Keys(older.EnumerateCurrentLibraryFiles(child)), "library " + child.Path);
                    if (child.Kind == NodeKind.List || child.Kind == NodeKind.Library)
                    {
                        Same(new[] { Outcome(() => JsonSerializer.Serialize(full.ReadMigrationList(child))) },
                            new[] { Outcome(() => JsonSerializer.Serialize(older.ReadMigrationList(child))) }, "migration snapshot of " + child.Path);
                        packages++;
                    }
                    if (child.Kind == NodeKind.ListItem && child.HasAttachments == true)
                    {
                        List<Node> owned = full.GetItemAttachments(child);
                        Same(Keys(owned), Keys(older.GetItemAttachments(child)), "attachments of " + child.Path);
                        attachments += owned.Count;
                    }
                    if (child.Kind != NodeKind.File) continue;
                    files++;
                    Same(new[] { Key(full.GetFile(child.SiteId, child.Id)) }, new[] { Key(older.GetFile(child.SiteId, child.Id)) }, "file " + child.Path);
                    Same(new[] { Outcome(() => String.Join("\n", Keys(full.GetFileVersions(child.SiteId, child.Id)))) },
                        new[] { Outcome(() => String.Join("\n", Keys(older.GetFileVersions(child.SiteId, child.Id)))) }, "versions of " + child.Path);
                    versions++;
                    if (child.HasStream == false) continue;
                    IList<StoredChunk> a = new SqlDocumentChunkStore(full).ReadChunks(child), b = new SqlDocumentChunkStore(older).ReadChunks(child);
                    Check(a.Count == b.Count && a.Zip(b, (x, y) => x.BSN == y.BSN && x.StreamId == y.StreamId && x.Type == y.Type && x.Content.AsSpan().SequenceEqual(y.Content)).All(same => same),
                        "Stored chunks differ for " + child.Path);
                    chunks += a.Count;
                }
                if (parent.Kind == NodeKind.Site)
                {
                    var site = new Node { Kind = NodeKind.Site, SiteId = parent.SiteId, WebId = parent.WebId, Id = parent.WebId };
                    List<Node> retained = full.GetDeletedItems(site);
                    Same(Keys(retained), Keys(older.GetDeletedItems(site)), "deleted items of " + parent.Path);
                    deleted += retained.Count;
                    foreach (Node item in retained.Where(item => item.Kind == NodeKind.File))
                        Same(new[] { Outcome(() => String.Join("\n", Keys(full.GetDeletedFileVersions(item)))) },
                            new[] { Outcome(() => String.Join("\n", Keys(older.GetDeletedFileVersions(item)))) }, "deleted versions of " + item.Path);
                }
            }
            Same(Keys(full.EnumerateCurrentFiles(null)), Keys(older.EnumerateCurrentFiles(null)), "all current files");
            Check(files > 0 && chunks > 0 && packages > 0, "The older-profile SQL check did not reach stored content.");
            Console.WriteLine("PASS older-generation column profile on SQL: " + nodes + " nodes, " + files + " files, " + chunks + " chunks, " + versions +
                " version lists, " + attachments + " attachments, " + deleted + " deleted items and " + packages + " migration snapshots match the detected SE profile");
        }

        private static string Outcome(Func<string> action)
        {
            try { return action(); }
            catch (Exception error) { return "ERROR " + error.GetType().Name + ": " + error.Message; }
        }
        private static List<string> Keys(IEnumerable<Node> nodes) { return nodes.Select(Key).ToList(); }
        private static string Key(Node node)
        {
            return String.Join("|", node.Kind, node.SiteId, node.Id, node.WebId, node.ListId, node.ParentId, node.Path, node.Name, node.Size, node.StreamSchema,
                node.Level, node.InternalVersion, node.UiVersion, node.HistoryVersion, node.HasStream, node.Modified.ToString("o"), node.DeletionTransactionId,
                node.ListItemId, node.ItemUniqueId, node.AttachmentOwnerId, node.HasAttachments, node.Title);
        }
        private static void Same(IList<string> expected, IList<string> actual, string scope)
        {
            if (expected.Count != actual.Count) throw new Exception("The older column profile changed the " + scope + " count: " + expected.Count + " / " + actual.Count + ".");
            for (int index = 0; index < expected.Count; index++)
                if (expected[index] != actual[index]) throw new Exception("The older column profile changed " + scope + ":\n  " + expected[index] + "\n  " + actual[index]);
        }
        private static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
    }
}
