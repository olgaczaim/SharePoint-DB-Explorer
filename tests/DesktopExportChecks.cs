using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using SharePointExplorer;
using SharePointExplorer.Desktop;

namespace SharePointExplorer.Tests
{
    public static class DesktopExportChecks
    {
        public static void Run()
        {
            string root = Path.GetFullPath(Path.Combine(".scratch", "desktop-export-checks", Guid.NewGuid().ToString("N")));
            TestMixedSelection(Path.Combine(root, "mixed"));
            TestMetadataEligibility(Path.Combine(root, "eligibility"));
            TestPlainCorruption(Path.Combine(root, "plain-corrupt"));
            TestSafePathsAndCollisions(Path.Combine(root, "paths"));
            TestCancellation(Path.Combine(root, "cancel"));
            TestScope(Path.Combine(root, "scope"));
            TestSelectionSnapshot(Path.Combine(root, "snapshot"));
            TestLifetimeAndInput(Path.Combine(root, "lifetime"));
            Console.WriteLine("PASS desktop selection, mixed outcomes, paths, reports, cancellation and lifetime");
        }

        private static void TestMixedSelection(string root)
        {
            FakeCatalog catalog = new FakeCatalog();
            Node alpha = catalog.Add("alpha,\"quoted\".txt", 0, "alpha");
            Node unavailable = catalog.Add("ghosted.aspx", 0, "none"); unavailable.HasStream = false;
            Node native = catalog.Add("unsupported.bin", 67, "native");
            Node broken = catalog.Add("broken.bin", 7, "bad");
            Node last = catalog.Add("last.txt", 0, "last");
            FakeStore store = new FakeStore(catalog);
            var registry = new DocumentDecoderRegistry(new IDocumentDecoder[] {
                DocumentDecoderRegistry.CreateDefault().Resolve(0),
                DocumentDecoderRegistry.CreateDefault().Resolve(66), new FailingDecoder() });
            using (var controller = new ExplorerController(new RecoverySession(catalog, store, registry)))
            {
                var progress = new List<DesktopExportProgress>();
                DesktopExportSummary summary = controller.ExportFiles(new Node[] {alpha, Copy(alpha), unavailable, native, broken, last},
                    root, CancellationToken.None, delegate(DesktopExportProgress item) { progress.Add(item); });
                Require(summary.Total == 3 && summary.Success == 2 && summary.Failed == 1 && summary.Skipped == 0 && !summary.Cancelled,
                    "Mixed selection totals did not deduplicate or continue past errors.");
                Require(summary.Entries.Count == 3 && progress.Count == 3 && store.Reads == 3,
                    "Duplicate or unsupported documents reached content retrieval, or progress omitted an outcome.");
                Require(summary.Entries[1].Status == RecoveryStatus.Corrupt && summary.Entries[2].Status == RecoveryStatus.Success && catalog.Lookups == 3,
                    "Desktop recovery statuses or final-file continuation changed.");
                for (int index = 0; index < progress.Count; index++)
                {
                    Require(progress[index].Completed == index + 1 && progress[index].Total == 3 &&
                        progress[index].Status == summary.Entries[index].Status && progress[index].ExportPath == summary.Entries[index].Path,
                        "Progress did not describe the completed deduplicated selection.");
                }
                CheckSuccessfulEntry(summary.Entries[0], Encoding.UTF8.GetBytes("alpha"));
                CheckSuccessfulEntry(summary.Entries[2], Encoding.UTF8.GetBytes("last"));
                CheckDirectOutput(summary, root);
                Require(Directory.GetFiles(root, "*.partial", SearchOption.AllDirectories).Length == 0,
                    "A decoder failure left partial output.");
                List<string[]> rows = ReadCsv(summary.ReportPath);
                Require(rows.Count == 4 && rows[0].Length == 9 && rows[0][0] == "Source" && rows[0][6] == "Sha256", "Report columns or row count changed.");
                for (int index = 0; index < summary.Entries.Count; index++)
                {
                    DesktopExportEntry entry = summary.Entries[index]; string[] row = rows[index + 1];
                    Require(row.Length == 9 && row[0] == catalog.SourceName && row[1] == entry.Document.SiteId.ToString("D") &&
                        row[2] == entry.Document.Id.ToString("D") && row[3] == entry.Document.Path && row[4] == entry.Status.ToString() &&
                        row[5] == entry.Bytes.ToString(CultureInfo.InvariantCulture) && row[6] == entry.Sha256 && row[7] == entry.Path && row[8] == entry.Message,
                        "CSV quoting, checksum or error text did not round-trip.");
                }
                Require(rows[2][8] == "broken, \"record\"\r\nretry", "Multiline quoted decoder errors were damaged in the report.");
                // The same document ID in another site collection is a separate selection.
                Node otherSite = Copy(alpha); otherSite.SiteId = Guid.NewGuid(); catalog.Add(otherSite, "other");
                DesktopExportSummary sites = controller.ExportFiles(new Node[] {alpha, otherSite}, Path.Combine(root, "sites"), CancellationToken.None, null);
                Require(sites.Total == 2 && sites.Success == 2, "Deduplication merged document IDs from different site collections.");
            }
        }

        private static void TestMetadataEligibility(string root)
        {
            var defaults = DocumentDecoderRegistry.CreateDefault();
            Require(defaults.Supports(0) && defaults.Supports(66) && !defaults.Supports(67) && !defaults.Supports(7),
                "Decoder support queries do not match the registered codecs.");
            IDocumentDecoder plain = defaults.Resolve(0);
            var ambiguous = new DocumentDecoderRegistry(new IDocumentDecoder[] {plain, plain});
            Throws<InvalidOperationException>(delegate {ambiguous.Supports(0);});
            Throws<InvalidOperationException>(delegate {ambiguous.Resolve(0);});
            var catalog = new FakeCatalog();
            Node stored = catalog.Add("stored.aspx", 0, "stored ASPX bytes");
            Node ghost = catalog.Add("default.aspx", 0, "template"); ghost.HasStream = false;
            Node native = catalog.Add("unsupported.bin", 67, "native");
            Node custom = catalog.Add("custom.bin", 7, "custom");
            var store = new FakeStore(catalog);
            var registry = new DocumentDecoderRegistry(new IDocumentDecoder[] {plain, new FailingDecoder()});
            var session = new RecoverySession(catalog, store, registry);
            RecoveryEngine engine = session.Engine;
            using (var controller = new ExplorerController(session))
            {
                Require(engine.CanExport(stored) && engine.CanExport(custom) && !engine.CanExport(ghost) && !engine.CanExport(native) &&
                    !engine.CanExport(new Node {Kind=NodeKind.Folder}), "Metadata eligibility ignored the actual registry or stream availability.");
                Node historical = Copy(stored); historical.HistoryVersion = 1;
                Require(engine.CanExport(historical), "Eligibility silently removed the historical-version recovery guard.");
                Node unknown = Copy(stored); unknown.HasStream = null;
                Require(engine.CanExport(unknown) && catalog.Lookups == 0 && store.Reads == 0, "Metadata eligibility fetched source data or rejected an unknown availability flag.");
                DesktopExportSummary ignored = controller.ExportFiles(new Node[] {ghost, native}, Path.Combine(root, "ignored"), CancellationToken.None, null);
                Require(ignored.Total == 0 && ignored.Success == 0 && ignored.Skipped == 0 && ignored.Failed == 0 && ignored.Entries.Count == 0 &&
                    ReadCsv(ignored.ReportPath).Count == 1 && catalog.Lookups == 0 && store.Reads == 0, "Known ineligible files were counted, reported or retrieved.");
                DesktopExportSummary saved = controller.ExportFiles(new Node[] {stored}, Path.Combine(root, "stored"), CancellationToken.None, null);
                Require(saved.Total == 1 && saved.Success == 1 && saved.Skipped == 0, "A stored ASPX file was filtered by its extension.");
                CheckSuccessfulEntry(saved.Entries[0], Encoding.UTF8.GetBytes("stored ASPX bytes"));
            }
            Throws<ObjectDisposedException>(delegate {engine.CanExport(stored);});
            // A change after selection remains a detailed runtime outcome in the report.
            for (int variant = 0; variant < 2; variant++)
            {
                var currentCatalog = new FakeCatalog(); Node selected = currentCatalog.Add("changed.aspx", 0, "stored"); Node fresh = Copy(selected);
                if (variant == 0) fresh.HasStream = false; else fresh.StreamSchema = 77;
                currentCatalog.Lookup = delegate(Guid site, Guid id) {return fresh;};
                var currentStore = new FakeStore(currentCatalog);
                using (var controller = Controller(currentCatalog, currentStore))
                {
                    DesktopExportSummary result = controller.ExportFiles(new Node[] {selected}, Path.Combine(root, "changed-" + variant), CancellationToken.None, null);
                    Require(result.Total == 1 && result.Skipped == 1 && result.Failed == 0 && result.Entries[0].Status ==
                        (variant == 0 ? RecoveryStatus.Unavailable : RecoveryStatus.Unsupported) && ReadCsv(result.ReportPath).Count == 2 &&
                        currentStore.Reads == (variant == 0 ? 1 : 0), "Freshly unavailable/unsupported content lost its runtime audit outcome.");
                }
            }
        }

        private static void TestPlainCorruption(string root)
        {
            var catalog = new FakeCatalog(); Node document = catalog.Add("wrong-size.txt", 0, "stored"); document.Size++;
            var store = new FakeStore(catalog);
            using (var controller = Controller(catalog, store))
            {
                DesktopExportSummary result = controller.ExportFiles(new Node[] {document}, root, CancellationToken.None, null);
                Require(result.Total == 1 && result.Failed == 1 && result.Skipped == 0 && result.Success == 0 &&
                    result.Entries[0].Status == RecoveryStatus.Corrupt && ReadCsv(result.ReportPath)[1][4] == "Corrupt" && store.Reads == 1,
                    "A supported resident stream with a wrong declared size was hidden as an unsupported file.");
                Require(Directory.GetFiles(root, "wrong-size.txt", SearchOption.AllDirectories).Length == 0 &&
                    Directory.GetFiles(root, "*.partial", SearchOption.AllDirectories).Length == 0, "Corrupt resident data was published or retained.");
            }
        }

        private static void TestSafePathsAndCollisions(string root)
        {
            FakeCatalog catalog = new FakeCatalog();
            Node document = catalog.Add("a:b?.txt", 0, "recovered");
            document.Path = "../CON/./bad:folder/a:b?.txt";
            Node sameName = catalog.Add("a:b?.txt", 0, "different same-name content"); sameName.Path = "other/folder/a:b?.txt";
            Node absolute = catalog.Add("absolute.bin", 0, "absolute"); absolute.Path = "/outside/../../NUL/absolute.bin";
            Node drive = catalog.Add("C:\\outside\\drive.bin", 0, "drive"); drive.Path = "C:\\outside\\..\\LPT1\\drive.bin";
            Node reserved = catalog.Add("NUL.txt", 0, "reserved"); reserved.Path = "Docs/NUL.txt";
            Node traversal = catalog.Add("../escape.txt", 0, "traversal"); traversal.Path = "../../escape.txt";
            Directory.CreateDirectory(root);
            string existing = Path.Combine(root, "a_b_.txt"); File.WriteAllText(existing, "original");
            FakeStore store = new FakeStore(catalog);
            using (var controller = Controller(catalog, store))
            {
                Node stale = Copy(document); stale.Path = "stale/metadata/a:b?.txt";
                DesktopExportSummary first = controller.ExportFiles(new Node[] {stale, sameName, absolute, drive, reserved, traversal}, root, CancellationToken.None, null);
                Require(first.Success == 6 && first.Entries[0].Path == Path.Combine(root, "a_b_ (2).txt") &&
                    first.Entries[1].Path == Path.Combine(root, "a_b_ (3).txt") && Path.GetFileName(first.Entries[4].Path) == "_NUL.txt",
                    "Flat selected exports did not preserve same-name collisions or sanitize reserved filenames.");
                CheckDirectOutput(first, root);
                Require(first.Entries[0].Document.Path == document.Path && ReadCsv(first.ReportPath)[1][3] == document.Path,
                    "Selected-file audit retained a stale source path instead of the recovered current metadata.");
                DesktopExportSummary second = controller.ExportFiles(new Node[] {document}, root, CancellationToken.None, null);
                Require(second.Success == 1 && second.Entries[0].Path == Path.Combine(root, "a_b_ (4).txt") && File.ReadAllText(existing) == "original",
                    "A repeated selected export replaced an existing file or same-name document.");
                CheckDirectOutput(second, root);
                Require(Directory.GetDirectories(root).Length == 0, "Selected exports created source hierarchy beneath the chosen folder.");
                CheckSuccessfulEntry(first.Entries[0], Encoding.UTF8.GetBytes("recovered"));
                CheckSuccessfulEntry(first.Entries[1], Encoding.UTF8.GetBytes("different same-name content"));
                CheckSuccessfulEntry(first.Entries[2], Encoding.UTF8.GetBytes("absolute"));
                CheckSuccessfulEntry(first.Entries[3], Encoding.UTF8.GetBytes("drive"));
                CheckSuccessfulEntry(first.Entries[4], Encoding.UTF8.GetBytes("reserved"));
                CheckSuccessfulEntry(first.Entries[5], Encoding.UTF8.GetBytes("traversal"));
                CheckSuccessfulEntry(second.Entries[0], Encoding.UTF8.GetBytes("recovered"));
                Require(Directory.GetFiles(root, "*.partial", SearchOption.AllDirectories).Length == 0, "Export left partial files.");
            }
        }
        private static void TestCancellation(string root)
        {
            FakeCatalog catalog = new FakeCatalog(); Node first = catalog.Add("first.txt", 0, "first"); Node second = catalog.Add("second.txt", 0, "second");
            FakeStore store = new FakeStore(catalog);
            using (var controller = Controller(catalog, store))
            using (var before = new CancellationTokenSource())
            using (var between = new CancellationTokenSource())
            {
                before.Cancel(); string untouched = Path.Combine(root, "before");
                DesktopExportSummary cancelled = controller.ExportFiles(new Node[] {first, second}, untouched, before.Token,
                    delegate(DesktopExportProgress item) { throw new Exception("Pre-cancelled selection raised progress."); });
                Require(cancelled.Cancelled && cancelled.Total == 2 && cancelled.Entries.Count == 0 && cancelled.Success == 0 &&
                    String.IsNullOrEmpty(cancelled.ReportPath) && !Directory.Exists(untouched) && store.Reads == 0 && catalog.Lookups == 0,
                    "Pre-cancellation created output or read document content.");
                int callbacks = 0;
                DesktopExportSummary partial = controller.ExportFiles(new Node[] {first, second}, Path.Combine(root, "between"), between.Token,
                    delegate(DesktopExportProgress item) { callbacks++; between.Cancel(); });
                Require(partial.Cancelled && partial.Total == 2 && partial.Success == 1 && partial.Entries.Count == 1 && callbacks == 1 && store.Reads == 1 && catalog.Lookups == 1,
                    "Cancellation between files read or exported the next document.");
                CheckSuccessfulEntry(partial.Entries[0], Encoding.UTF8.GetBytes("first"));
                CheckDirectOutput(partial, Path.Combine(root, "between"));
                Require(ReadCsv(partial.ReportPath).Count == 2 && Directory.GetFiles(partial.Directory, "second.txt", SearchOption.AllDirectories).Length == 0,
                    "Cancellation omitted the completed report or exported an unprocessed file.");
            }
        }

        private static void TestScope(string root)
        {
            for (int variant = 0; variant < 4; variant++)
            {
                FakeCatalog catalog = new FakeCatalog(); Node selected = catalog.Add("scope.txt", 0, "source"); Node returned = Copy(selected);
                if (variant == 0) returned.Id = Guid.NewGuid();
                if (variant == 1) returned.SiteId = Guid.NewGuid();
                if (variant == 2) returned.Kind = NodeKind.Folder;
                if (variant == 3) returned.HistoryVersion = 1;
                catalog.Lookup = delegate(Guid site, Guid id) { return returned; };
                FakeStore store = new FakeStore(catalog);
                using (var controller = Controller(catalog, store))
                {
                    DesktopExportSummary result = controller.ExportFiles(new Node[] {selected}, Path.Combine(root, "wrong-" + variant), CancellationToken.None, null);
                    Require(result.Failed == 1 && result.Entries[0].Status == RecoveryStatus.Corrupt && store.Reads == 0,
                        "Desktop export accepted mismatched current document metadata.");
                }
            }
            FakeCatalog missing = new FakeCatalog(); Node absent = missing.Add("missing.txt", 0, "missing");
            missing.Lookup = delegate(Guid site, Guid id) { return null; }; FakeStore missingStore = new FakeStore(missing);
            using (var controller = Controller(missing, missingStore))
            {
                DesktopExportSummary result = controller.ExportFiles(new Node[] {absent}, Path.Combine(root, "missing"), CancellationToken.None, null);
                Require(result.Skipped == 1 && result.Entries[0].Status == RecoveryStatus.Unavailable && missingStore.Reads == 0,
                    "A missing current document was not classified as unavailable.");
            }
            FakeCatalog historic = new FakeCatalog(); Node current = historic.Add("history.txt", 0, "current"); Node old = Copy(current); old.HistoryVersion = 2;
            FakeStore historicStore = new FakeStore(historic);
            using (var controller = Controller(historic, historicStore))
            {
                DesktopExportSummary result = controller.ExportFiles(new Node[] {old}, Path.Combine(root, "history"), CancellationToken.None, null);
                Require(result.Skipped == 1 && result.Entries[0].Status == RecoveryStatus.Unsupported && historicStore.Reads == 0,
                    "A historical selection silently exported current content.");
            }
            FakeCatalog changing = new FakeCatalog(); Node original = changing.Add("changing.txt", 0, "original");
            Node stale = Copy(original); stale.Name = "stale.txt"; stale.Path = "old/folder/stale.txt"; stale.Size = 999; stale.UiVersion = 512;
            Node fresh = Copy(original); fresh.Name = "fresh.txt"; fresh.Path = "Docs/NewFolder/fresh.txt"; fresh.UiVersion = 1024;
            fresh.InternalVersion = 2050; fresh.Level = 2; fresh.Modified = new DateTime(2025, 3, 4, 5, 6, 7, DateTimeKind.Utc);
            fresh.WebId = Guid.NewGuid(); fresh.ListId = Guid.NewGuid(); fresh.ParentId = Guid.NewGuid();
            Node replacement = Copy(fresh); replacement.Id = Guid.NewGuid();
            changing.Lookup = delegate(Guid site, Guid id) { return changing.Lookups == 1 ? fresh : replacement; };
            FakeStore changingStore = new FakeStore(changing);
            using (var controller = Controller(changing, changingStore))
            {
                var progress = new List<DesktopExportProgress>();
                string directory = Path.Combine(root, "changing");
                DesktopExportSummary result = controller.ExportFiles(new Node[] {stale}, directory, CancellationToken.None, progress.Add);
                DesktopExportEntry entry = result.Entries[0];
                Require(result.Success == 1 && result.Failed == 0 && changing.Lookups == 1 && changingStore.Reads == 1 &&
                    entry.Document.Name == fresh.Name && entry.Document.Path == fresh.Path && entry.Document.Size == fresh.Size &&
                    entry.Document.WebId == fresh.WebId && entry.Document.ListId == fresh.ListId && entry.Document.ParentId == fresh.ParentId &&
                    entry.Document.UiVersion == 1024 && entry.Document.InternalVersion == 2050 && entry.Document.Level == 2 && entry.Document.Modified == fresh.Modified &&
                    Path.GetFileName(entry.Path) == "fresh.txt",
                    "Selected export did not publish one authoritative current snapshot, or refreshed it a second time.");
                CheckDirectOutput(result, directory); CheckSuccessfulEntry(entry, Encoding.UTF8.GetBytes("original"));
                List<string[]> csv = ReadCsv(result.ReportPath);
                Require(csv.Count == 2 && csv[0].Length == 9 && csv[1][1] == fresh.SiteId.ToString("D") && csv[1][2] == fresh.Id.ToString("D") &&
                    csv[1][3] == fresh.Path && csv[1][4] == "Success" && csv[1][5] == fresh.Size.ToString(CultureInfo.InvariantCulture) &&
                    csv[1][6] == entry.Sha256 && csv[1][7] == entry.Path &&
                    progress.Count == 1 && progress[0].Document.Name == fresh.Name && progress[0].Document.Path == fresh.Path && progress[0].ExportPath == entry.Path,
                    "Selected-file CSV or progress lost the exact fresh current metadata and direct destination.");
            }
        }

        private static void TestSelectionSnapshot(string root)
        {
            var catalog = new FakeCatalog(); Node first = catalog.Add("first.txt", 0, "first"); Node second = catalog.Add("second.txt", 0, "second");
            Node callerSecond = Copy(second); var store = new FakeStore(catalog);
            using (var controller = Controller(catalog, store))
            {
                DesktopExportSummary result = controller.ExportFiles(new Node[] {Copy(first), callerSecond}, root, CancellationToken.None,
                    delegate(DesktopExportProgress progress) {
                        if (progress.Completed == 1) {callerSecond.Id = Guid.NewGuid(); callerSecond.Kind = NodeKind.Folder; callerSecond.HasStream = false;}
                    });
                Require(result.Success == 2 && result.Failed == 0 && result.Entries[1].Document.Id == second.Id && store.Reads == 2 && catalog.Lookups == 2,
                    "Caller mutation after export began changed the captured selected-file scope.");
                CheckDirectOutput(result, root); CheckSuccessfulEntry(result.Entries[1], Encoding.UTF8.GetBytes("second"));
            }
        }
        private static void TestLifetimeAndInput(string root)
        {
            FakeCatalog catalog = new FakeCatalog(); Node document = catalog.Add("alive.txt", 0, "alive"); FakeStore store = new FakeStore(catalog);
            var controller = Controller(catalog, store);
            Require(controller.SourceName == catalog.SourceName && controller.GetRootSites().Count == 1 && controller.GetChildren(new Node {Kind=NodeKind.Site}).Count == 1,
                "Controller navigation did not delegate to the session catalog.");
            Throws<ArgumentException>(delegate {controller.ExportFiles(new Node[0], root, CancellationToken.None, null);});
            Throws<ArgumentException>(delegate {controller.ExportFiles(new Node[] {null}, root, CancellationToken.None, null);});
            Throws<ArgumentException>(delegate {controller.ExportFiles(new Node[] {new Node {Kind=NodeKind.Folder}}, root, CancellationToken.None, null);});
            Require(!Directory.Exists(root) && store.Reads == 0, "Invalid selections created output.");
            controller.Dispose(); controller.Dispose();
            Require(catalog.Disposals == 1 && store.Disposals == 1, "Controller did not own a single session disposal.");
            Throws<ObjectDisposedException>(delegate {GC.KeepAlive(controller.SourceName);});
            Throws<ObjectDisposedException>(delegate {controller.GetRootSites();});
            Throws<ObjectDisposedException>(delegate {controller.GetChildren(new Node {Kind=NodeKind.Site});});
            Throws<ObjectDisposedException>(delegate {controller.ExportFiles(new Node[] {document}, root, CancellationToken.None, null);});
            Require(!Directory.Exists(root) && store.Reads == 0, "Disposed controller created output.");
        }

        private static ExplorerController Controller(FakeCatalog catalog, FakeStore store)
        { return new ExplorerController(new RecoverySession(catalog, store, DocumentDecoderRegistry.CreateDefault())); }
        private static void CheckSuccessfulEntry(DesktopExportEntry entry, byte[] expected)
        {
            Require(entry.Status == RecoveryStatus.Success && entry.Bytes == expected.Length && entry.Sha256 == Digest(expected), "Export byte count or SHA256 differs from source content.");
            Require(Digest(File.ReadAllBytes(entry.Path)) == entry.Sha256, "The reported digest does not describe the published file.");
        }
        private static void CheckDirectOutput(DesktopExportSummary summary, string root)
        {
            string directory = Path.GetFullPath(root);
            Require(String.Equals(summary.Directory, directory, StringComparison.OrdinalIgnoreCase) &&
                String.Equals(Path.GetDirectoryName(Path.GetFullPath(summary.ReportPath)), directory, StringComparison.OrdinalIgnoreCase) && File.Exists(summary.ReportPath),
                "The selected-file report was not saved directly in the chosen folder.");
            foreach (DesktopExportEntry entry in summary.Entries)
                if (entry.Status == RecoveryStatus.Success)
                    Require(String.Equals(Path.GetDirectoryName(Path.GetFullPath(entry.Path)), directory, StringComparison.OrdinalIgnoreCase) && File.Exists(entry.Path),
                        "A selected current file was not saved directly in the chosen folder.");
        }
        private static string Digest(byte[] bytes)
        { using (SHA256 hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant(); }
        private static Node Copy(Node source)
        { return new Node {Kind=source.Kind,SiteId=source.SiteId,WebId=source.WebId,ListId=source.ListId,Id=source.Id,Name=source.Name,Path=source.Path,Size=source.Size,StreamSchema=source.StreamSchema,HasStream=source.HasStream,HistoryVersion=source.HistoryVersion,UiVersion=source.UiVersion,InternalVersion=source.InternalVersion,Level=source.Level,Modified=source.Modified,ParentId=source.ParentId}; }
        private static string Key(Guid site, Guid id) { return site.ToString("N") + ":" + id.ToString("N"); }
        private static void Require(bool condition, string message) { if (!condition) throw new Exception(message); }
        private static void Throws<T>(Action action) where T : Exception
        { try { action(); } catch (T) { return; } throw new Exception("Expected " + typeof(T).Name + "."); }

        private static List<string[]> ReadCsv(string path)
        {
            string text = File.ReadAllText(path); var rows = new List<string[]>(); var fields = new List<string>(); var value = new StringBuilder(); bool quoted = false;
            for (int index = 0; index < text.Length; index++)
            {
                char current = text[index];
                if (current == '"')
                {
                    if (quoted && index + 1 < text.Length && text[index + 1] == '"') {value.Append('"'); index++;}
                    else quoted = !quoted;
                }
                else if (!quoted && current == ',') {fields.Add(value.ToString()); value.Length=0;}
                else if (!quoted && (current == '\r' || current == '\n'))
                {
                    fields.Add(value.ToString()); value.Length=0; rows.Add(fields.ToArray()); fields.Clear();
                    if (current == '\r' && index + 1 < text.Length && text[index + 1] == '\n') index++;
                }
                else value.Append(current);
            }
            Require(!quoted, "Report contains an unterminated quoted field.");
            if (fields.Count > 0 || value.Length > 0) {fields.Add(value.ToString()); rows.Add(fields.ToArray());}
            return rows;
        }

        private sealed class FakeCatalog : ISharePointCatalog, IDisposable
        {
            public readonly Guid Site = new Guid("dddddddd-1111-2222-3333-444444444444");
            public readonly Dictionary<string,Node> Documents = new Dictionary<string,Node>();
            public readonly Dictionary<string,byte[]> Content = new Dictionary<string,byte[]>();
            public Func<Guid,Guid,Node> Lookup; public int Lookups, Disposals;
            public string SourceName { get {return "Synthetic, \"SQL\"\r\nsource";} }
            public Node Add(string name, byte schema, string content)
            {
                Node document = new Node {Kind=NodeKind.File,SiteId=Site,Id=Guid.NewGuid(),Name=name,Path="library/" + name,Size=Encoding.UTF8.GetByteCount(content),StreamSchema=schema,HasStream=true};
                Add(document,content); return document;
            }
            public void Add(Node document, string content)
            { byte[] bytes=Encoding.UTF8.GetBytes(content); document.Size=bytes.Length; Documents.Add(Key(document.SiteId,document.Id),document); Content.Add(Key(document.SiteId,document.Id),bytes); }
            public Node GetFile(Guid site, Guid id)
            { Lookups++; if(Lookup!=null) return Lookup(site,id); Node value; return Documents.TryGetValue(Key(site,id),out value) ? value : null; }
            public void ValidateSchema() { }
            public List<string> CheckDatabase() {return new List<string>();}
            public List<Node> GetRootSites() {return new List<Node> {new Node {Kind=NodeKind.Site,SiteId=Site}};}
            public List<Node> GetChildren(Node parent) {return new List<Node> {new Node {Kind=NodeKind.Library,SiteId=Site}};}
            public IEnumerable<Node> EnumerateCurrentFiles(Guid? site) {return Documents.Values;}
            public void Dispose() {Disposals++;}
        }
        private sealed class FakeStore : IDocumentChunkStore, IDisposable
        {
            private readonly FakeCatalog catalog; public int Reads, Disposals;
            public FakeStore(FakeCatalog catalog) {this.catalog=catalog;}
            public IList<StoredChunk> ReadChunks(Node document)
            {
                Reads++; if(document.HasStream==false) throw new ContentUnavailableException("Missing template, \"provider\".");
                return new List<StoredChunk> {new StoredChunk {Partition=0,Type=0,StreamId=1,BSN=1,Content=catalog.Content[Key(document.SiteId,document.Id)]}};
            }
            public void Dispose() {Disposals++;}
        }
        private sealed class FailingDecoder : IDocumentDecoder
        {
            public string Name {get {return "Synthetic partial decoder";}}
            public bool Supports(byte schema) {return schema==7;}
            public void Write(IList<StoredChunk> chunks, long size, Stream output)
            {output.WriteByte(1); throw new InvalidDataException("broken, \"record\"\r\nretry");}
        }
    }
}
