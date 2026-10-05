#nullable enable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation.Peers;
using Microsoft.UI.Xaml.Automation.Provider;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using SharePointExplorer.Desktop;
using SharePointExplorer.Desktop.ViewModels;
using Windows.Foundation;
using Windows.Graphics.Imaging;
using Windows.Storage.Streams;

namespace SharePointExplorer.WinUI
{
    // Explicit diagnostic modes exercise the real XAML window on its dispatcher.
    // Normal startup never constructs this fixture or accesses its output files.
    public static class WinUiSmokeChecks
    {
        public static async Task RunAsync(MainWindow window, string[] args)
        {
            ArgumentNullException.ThrowIfNull(window);
            bool integration = args.Contains("--ui-integration", StringComparer.Ordinal);
            var checks = new CheckRun(window, integration, args);
            await checks.RunAsync();
        }

        private static bool SameVersion(Node first, Node second)
            => first.SiteId == second.SiteId && first.WebId == second.WebId && first.Id == second.Id && first.ListId == second.ListId &&
                first.HistoryVersion == second.HistoryVersion && first.UiVersion == second.UiVersion && first.Level == second.Level &&
                first.InternalVersion == second.InternalVersion && first.StreamSchema == second.StreamSchema && first.Size == second.Size && first.Path == second.Path;

        private sealed class CheckRun
        {
            private readonly MainWindow window;
            private readonly ExplorerViewModel model;
            private readonly FrameworkElement root;
            private readonly bool integration, skipDialogs, skipCapture;
            private readonly string workspace, logPath;
            private readonly StringBuilder log = new();
            private readonly FixtureController fixture = new();
            private int assertions;

            public CheckRun(MainWindow window, bool integration, string[] args)
            {
                this.window = window;
                this.integration = integration;
                skipDialogs = args.Contains("--ui-no-dialogs", StringComparer.Ordinal);
                skipCapture = args.Contains("--ui-no-capture", StringComparer.Ordinal);
                model = window.ViewModel;
                root = (FrameworkElement)window.Content;
                workspace = FindWorkspace();
                string scratch = Path.Combine(workspace, ".scratch");
                Directory.CreateDirectory(scratch);
                logPath = Path.Combine(scratch, integration ? "winui-runtime-check-integration.txt" : "winui-runtime-check.txt");
            }

            public async Task RunAsync()
            {
                Environment.ExitCode = 0;
                Write("START " + DateTime.UtcNow.ToString("O") + " mode=" + (integration ? "integration" : "fixture"));
                try
                {
                    await NativeChecksAsync();
                    if (!skipDialogs) await ConnectionDialogChecksAsync();
                    if (integration) await SqlChecksAsync();
                    if (!skipCapture) {
                    root.RequestedTheme = ElementTheme.Light;
                    await SettleAsync();
                    await CaptureAsync();
                    root.RequestedTheme = ElementTheme.Dark;
                    await SettleAsync();
                    Require(root.ActualTheme == ElementTheme.Dark, "The live native XAML tree applies the dark theme.");
                    await CaptureAsync("-dark");
                    }
                    Write("CLOSE starting");
                    await window.CloseSafelyAsync();
                    Write("CLOSE completed");
                    Require(fixture.Disposed, "The controller is disposed before the native window closes.");
                    Write("PASS assertions=" + assertions);
                }
                catch (Exception error)
                {
                    Environment.ExitCode = 1;
                    Write("FAIL " + error);
                    await CaptureAsync();
                }
                finally
                {
                    fixture.ReleaseSlow.Set();
                    fixture.ReleaseLibrary.Set();
                    fixture.ReleaseZip.Set();
                    fixture.ReleaseVersionRead.Set();
                    fixture.ReleaseVersionExport.Set();
                    try { await window.CloseSafelyAsync(); }
                    catch (Exception error) { Environment.ExitCode = 1; Write("CLOSE FAILED " + error); }
                    Write("END exitCode=" + Environment.ExitCode);
                }
            }

            private async Task NativeChecksAsync()
            {
                Require(root.XamlRoot != null && root.ActualWidth > 0 && root.ActualHeight > 0,
                    "The genuine WinUI XAML root is loaded and measured.");
                Require(window.AppWindow.Size.Width >= 1600 && window.AppWindow.Size.Height >= 1000,
                    "The diagnostic window uses the larger 1600 by 1000 startup bounds.");
                Require(Named<Button>("ConnectButton").XamlRoot == root.XamlRoot,
                    "The native connection button belongs to the live XAML root.");
                Require(Named<Button>("RefreshButton").KeyboardAccelerators.Any(v => v.Key == Windows.System.VirtualKey.F5),
                    "Refresh registers the native F5 keyboard accelerator.");
                Require(Named<Button>("UpButton").KeyboardAccelerators.Any(v => v.Key == Windows.System.VirtualKey.Up && v.Modifiers == Windows.System.VirtualKeyModifiers.Menu),
                    "Up registers the native Alt+Up keyboard accelerator.");
                Require(await model.SetSourceAsync(fixture), "The fixture source opens through the actual window view model.");
                await SettleAsync();
                TreeItemViewModel site = model.RootNodes.Single();
                var tree = Named<TreeView>("ExplorerTree");
                TreeViewNode? siteNode = FindNativeNode(tree.RootNodes, site);
                TreeViewItem? siteContainer = (siteNode == null ? null : tree.ContainerFromNode(siteNode)) as TreeViewItem ?? tree.ContainerFromItem(site) as TreeViewItem;
                Require(siteContainer != null, "The native tree realizes the site container.");
                new TreeViewItemAutomationPeer(siteContainer!).Expand();
                await UntilAsync(() => site.IsLoaded && !model.IsBusy, "native tree expansion");
                Require(site.Children.Count == 3 && site.Children.Any(v => v.Node.Kind == NodeKind.List),
                    "Expanding the native tree loads both a library and an ordinary list.");
                await model.NavigateAsync(site);
                TreeItemViewModel library = site.Children.Single(v => v.Name == "Documents");
                TreeItemViewModel list = site.Children.Single(v => v.Node.Kind == NodeKind.List);
                await model.NavigateAsync(library);
                TreeItemViewModel first = library.Children.Single(v => v.Name == "First folder");
                TreeItemViewModel second = library.Children.Single(v => v.Name == "Second folder");
                TreeItemViewModel broken = library.Children.Single(v => v.Name == "Unavailable folder");
                TreeItemViewModel slow = library.Children.Single(v => v.Name == "Slow folder");
                await model.NavigateAsync(first);
                await SettleAsync();
                Require(model.VisibleItems.Count == 3 && model.VisibleItems.Count(v => v.CanCheck) == 2,
                    "The native item list includes two files and a folder with file-only checkboxes.");
                ItemViewModel firstFile = model.VisibleItems.Single(v => v.Name == "first.txt");
                await ToggleAsync(firstFile);
                Require(firstFile.IsChecked && model.SelectedCount == 1 && Named<Button>("ExportButton").IsEnabled,
                    "A native checkbox selects one visible file and enables export.");
                Require(Named<TextBlock>("SelectionCountText").Text == "1 file selected",
                    "The selection label reflects the native checkbox.");
                Invoke("SelectAllButton");
                await SettleAsync();
                Require(model.SelectedCount == 2 && model.VisibleItems.Where(v => !v.CanCheck).All(v => !v.IsChecked),
                    "The native select-all button selects only files in this folder.");
                Invoke("ClearSelectionButton");
                await SettleAsync();
                Require(model.SelectedCount == 0 && !Named<Button>("ExportButton").IsEnabled,
                    "The native clear button resets the visible file selection.");
                await ToggleAsync(firstFile);
                await model.NavigateAsync(second);
                Require(model.SelectedCount == 0 && !firstFile.IsChecked,
                    "Changing folders clears the old checkboxes and export selection.");
                ItemViewModel secondFile = model.VisibleItems.Single();
                await ToggleAsync(secondFile);
                Invoke("UpButton");
                await UntilAsync(() => model.CurrentFolder == library && !model.IsBusy, "native Up navigation");
                Require(model.SelectedCount == 0 && !secondFile.IsChecked,
                    "The native Up button clears selection when the location changes.");
                await model.NavigateAsync(first);
                await ToggleAsync(model.VisibleItems.Single(v => v.Name == "first.txt"));
                await model.NavigateAsync(broken);
                Require(model.SelectedCount == 0 && model.VisibleItems.Count == 0 && model.ErrorText.Contains("fixture", StringComparison.Ordinal),
                    "A failed folder load clears selection and presents a recoverable error.");
                Require(Named<InfoBar>("ErrorInfoBar").IsOpen, "The native error bar displays the failed location load.");
                Task stale = model.NavigateAsync(slow);
                await UntilAsync(() => fixture.SlowStarted.IsSet, "delayed fixture folder");
                await model.NavigateAsync(second);
                fixture.ReleaseSlow.Set();
                await stale;
                await TreeSelectionVisibleAsync(second);
                Require(model.CurrentFolder == second && model.VisibleItems.Single().Name == "second.txt" && model.SelectedCount == 0,
                    "A stale asynchronous folder response cannot replace the new location or restore selection.");
                await model.NavigateAsync(list);
                Require(model.VisibleItems.All(v => v.Node.Kind == NodeKind.ListItem && !v.CanCheck),
                    "Ordinary list items are visible and browse-only.");
                var menu=(MenuFlyout)Named<Button>("MoreExportsButton").Flyout;
                Require(Named<Button>("MoreExportsButton").IsEnabled &&
                    menu.Items.OfType<MenuFlyoutItem>().Single(v=>v.Name=="ExportListAttachmentsMenu").IsEnabled &&
                    menu.Items.OfType<MenuFlyoutItem>().Single(v=>v.Name=="ExportPackageMenu").IsEnabled,
                    "The native More menu offers whole-list attachment and XML package exports.");
                foreach (ItemViewModel item in model.VisibleItems) item.IsChecked = true;
                Require(model.SelectedCount == 0 && model.VisibleItems.All(v => !v.IsChecked),
                    "Ordinary list items reject file selection.");
                await model.NavigateAsync(second);
                ItemViewModel selected = model.VisibleItems.Single();
                await ToggleAsync(selected);
                bool sawExporting = false, sawNativeBusyState = false;
                System.ComponentModel.PropertyChangedEventHandler exportState = (_, __) => {
                    if (model.IsExporting) { sawExporting = true; sawNativeBusyState |= !Named<Button>("ExportButton").IsEnabled && !Named<Button>("ExportLibraryButton").IsEnabled && !Named<Button>("SaveZipButton").IsEnabled && !Named<Button>("MoreExportsButton").IsEnabled && !Named<Button>("ConnectButton").IsEnabled; }
                };
                model.PropertyChanged += exportState;
                DesktopExportSummary? summary;
                try { summary = await model.ExportAsync(Path.Combine(workspace, ".scratch", "WinUiFixtureExport")); }
                finally { model.PropertyChanged -= exportState; }
                Require(summary != null && summary.Total == 1 && summary.Success == 1 && fixture.ExportedIds.SequenceEqual(new[] { selected.Node.Id }),
                    "The export command submits only the file checked in the current folder.");
                Require(sawExporting && sawNativeBusyState && !model.IsExporting && model.ExportCompleted == 1 && model.ExportTotal == 1 && model.LastExport == summary,
                    "Export progress and completion update through the actual window model.");
                Require(model.VisibleItems.Single().ExportStatus == "Exported", "The native item row receives its export result.");
                await LibraryChecksAsync(site, library, list, first, second);
                await ZipChecksAsync(first, second, list);
                await IgnoredOutcomeDialogChecksAsync(second);
                await VersionChecksAsync(first, second, list);
                await SelectedFileDialogChecksAsync(first, second);
                await BoundsAsync(1280, 850);
                await BoundsAsync(960, 720);
                window.AppWindow.Resize(new Windows.Graphics.SizeInt32(1600, 1000));
                await SettleAsync();
            }

            private async Task LibraryChecksAsync(TreeItemViewModel site, TreeItemViewModel library, TreeItemViewModel list, TreeItemViewModel first, TreeItemViewModel second)
            {
                Button libraryButton = Named<Button>("ExportLibraryButton");
                Require(libraryButton.Visibility == Visibility.Visible && new ButtonAutomationPeer(libraryButton).GetPattern(PatternInterface.Invoke) is IInvokeProvider,
                    "The distinct library-export button is visible and exposes its native Invoke pattern.");
                await model.NavigateAsync(site);
                Require(!libraryButton.IsEnabled && model.CurrentLibrary == null,
                    "A site cannot enable whole-library export.");
                await model.NavigateAsync(list);
                Require(!libraryButton.IsEnabled && model.CurrentLibrary == null,
                    "An ordinary list cannot enable document-library export.");
                await model.NavigateAsync(first);
                await SettleAsync();
                TreeItemViewModel nested = first.Children.Single();
                Require(!nested.IsLoaded && model.SelectedCount == 0 && libraryButton.IsEnabled && !Named<Button>("ExportButton").IsEnabled && ReferenceEquals(model.CurrentLibrary, library),
                    "An unopened nested folder and zero checked files still allow exporting its complete parent library.");
                fixture.HoldLibrary = true;
                fixture.LibraryStarted.Reset();
                fixture.ReleaseLibrary.Reset();
                string directory = Path.Combine(workspace, ".scratch", "WinUiLibraryFixture", Guid.NewGuid().ToString("N"));
                var dialog = new ExportProgressDialog(model, directory, library.Node) { XamlRoot = root.XamlRoot };
                var shown = dialog.ShowAsync();
                DesktopExportSummary? summary = null;
                try
                {
                    await UntilAsync(() => fixture.LibraryStarted.IsSet && model.IsExporting, "whole-library discovery");
                    Require(model.ExportIsDiscovering && ((ProgressBar)dialog.FindName("ExportProgress")).IsIndeterminate,
                        "The native library dialog shows indeterminate progress while the complete inventory is discovered.");
                    Require(dialog.Title?.ToString() == "Export library" && ((TextBlock)dialog.FindName("ExportScopeText")).Text.Contains("Documents", StringComparison.Ordinal) &&
                        ((TextBlock)dialog.FindName("ExportScopeText")).Text.Contains("nested folder", StringComparison.Ordinal),
                        "The native progress dialog identifies the entire library and its nested scope.");
                    Require(!libraryButton.IsEnabled && !Named<Button>("ExportButton").IsEnabled && !Named<Button>("ConnectButton").IsEnabled && !model.CanNavigate,
                        "Library discovery disables conflicting export, connection and navigation commands.");
                    fixture.ReleaseLibrary.Set();
                    await UntilAsync(() => dialog.Summary != null && !model.IsExporting, "whole-library export completion");
                    summary = dialog.Summary;
                    Require(((InfoBar)dialog.FindName("SummaryInfo")).IsOpen && ((ListView)dialog.FindName("ResultsList")).Items.Count == 6 &&
                        ((ProgressBar)dialog.FindName("ExportProgress")).Value == 6,
                        "The native library dialog displays all six results and completed progress.");
                    Require(((HyperlinkButton)dialog.FindName("OpenReportButton")).IsEnabled && File.Exists(summary!.ReportPath),
                        "The completed library export exposes its generated report.");
                    Require(((ListView)dialog.FindName("ResultsList")).Items.Cast<ExportResultRow>().All(row => row.Name.StartsWith("Documents/", StringComparison.Ordinal)),
                        "Library result labels distinguish documents by their complete folder paths.");
                }
                finally
                {
                    fixture.HoldLibrary = false;
                    fixture.ReleaseLibrary.Set();
                    dialog.CancelAndCloseWhenFinished();
                    await shown;
                }
                Require(summary?.Total == 6 && summary.Success == 6 && summary.Failed == 0 && summary.Skipped == 0 && fixture.CapturedLibrary?.Id == library.Node.Id,
                    "Whole-library export recovers root and unopened nested documents independently of the current folder.");
                var expected = new Dictionary<string, string> {
                    ["Documents/root.txt"] = "library root payload\n",
                    ["Documents/First/first.txt"] = "first library file\n",
                    ["Documents/First/other.txt"] = "another library file\n",
                    ["Documents/First/Nested/deep.txt"] = "deep library payload\n",
                    ["Documents/Second/second.txt"] = "second library file\n",
                    ["Documents/Slow/stale.txt"] = "unvisited library file\n"
                };
                Require(summary!.Entries.Select(entry => entry.Document.Path).ToHashSet().SetEquals(expected.Keys),
                    "The full library inventory includes every independently expected document path.");
                foreach (DesktopExportEntry entry in summary.Entries)
                {
                    byte[] expectedBytes = Encoding.UTF8.GetBytes(expected[entry.Document.Path]);
                    byte[] recovered = File.ReadAllBytes(entry.Path);
                    string expectedHash = Convert.ToHexString(SHA256.HashData(expectedBytes)).ToLowerInvariant();
                    Require(recovered.SequenceEqual(expectedBytes) && entry.Bytes == expectedBytes.Length && entry.Sha256 == expectedHash &&
                        Path.GetRelativePath(directory, entry.Path).Replace(Path.DirectorySeparatorChar, '/').EndsWith(entry.Document.Path, StringComparison.Ordinal),
                        "The nested library export preserves bytes, independent SHA-256 and hierarchy for " + entry.Document.Path + ".");
                }
                Require(ReferenceEquals(model.CurrentFolder, first) && model.SelectedCount == 0 && model.VisibleItems.All(item => !item.IsChecked) && !nested.IsLoaded && libraryButton.IsEnabled,
                    "Library export leaves the current folder, checkboxes and lazy tree loading unchanged.");
                ItemViewModel checkedFile = model.VisibleItems.Single(item => item.Name == "first.txt");
                await ToggleAsync(checkedFile);
                DesktopExportSummary? checkedSummary = await model.ExportLibraryAsync(Path.Combine(workspace, ".scratch", "WinUiLibraryFixture", Guid.NewGuid().ToString("N")));
                Require(checkedSummary?.Success == 6 && checkedFile.IsChecked && model.SelectedCount == 1,
                    "An existing single-file check neither limits nor changes an entire-library export.");
                TreeItemViewModel empty = site.Children.Single(item => item.Name == "Empty library");
                await model.NavigateAsync(empty);
                var emptyDialog = new ExportProgressDialog(model, Path.Combine(workspace, ".scratch", "WinUiLibraryFixture", Guid.NewGuid().ToString("N")), empty.Node) { XamlRoot = root.XamlRoot };
                var emptyShown = emptyDialog.ShowAsync();
                try
                {
                    await UntilAsync(() => emptyDialog.Summary != null && !model.IsExporting, "empty-library completion");
                    Require(emptyDialog.Summary!.Total == 0 && emptyDialog.Summary.Entries.Count == 0 && ((InfoBar)emptyDialog.FindName("SummaryInfo")).Title == "Nothing to export" &&
                        !((ProgressBar)emptyDialog.FindName("ExportProgress")).IsIndeterminate,
                        "An empty document library finishes with an explicit friendly empty result.");
                }
                finally { emptyDialog.CancelAndCloseWhenFinished(); await emptyShown; }
                await model.NavigateAsync(second);
                await ToggleAsync(model.VisibleItems.Single());
            }
            private async Task ZipChecksAsync(TreeItemViewModel first, TreeItemViewModel second, TreeItemViewModel list)
            {
                Button zipButton = Named<Button>("SaveZipButton");
                Require(zipButton.Visibility == Visibility.Visible && new ButtonAutomationPeer(zipButton).GetPattern(PatternInterface.Invoke) is IInvokeProvider,
                    "Save selections as ZIP exposes its distinct native button and Invoke pattern.");
                await model.NavigateAsync(first);
                Require(model.SelectedCount == 0 && !zipButton.IsEnabled,
                    "ZIP saving stays disabled before files in the current folder are checked.");
                Invoke("SelectAllButton");await SettleAsync();
                Require(model.SelectedCount == 2 && zipButton.IsEnabled,
                    "Checking the two current-folder files enables ZIP saving.");
                Guid[] selectedIds = model.VisibleItems.Where(item => item.IsChecked).Select(item => item.Node.Id).ToArray();
                string archivePath = Path.Combine(workspace, ".scratch", "WinUiZipFixture", Guid.NewGuid().ToString("N"), "selected-files.zip");
                Directory.CreateDirectory(Path.GetDirectoryName(archivePath)!);
                File.WriteAllText(archivePath, "Existing ZIP destination must remain unchanged.");
                fixture.HoldZip = true;fixture.ZipStarted.Reset();fixture.ReleaseZip.Reset();
                var dialog = new ExportProgressDialog(model, archivePath, asZip: true) { XamlRoot = root.XamlRoot };
                var shown = dialog.ShowAsync();DesktopExportSummary? summary = null;
                try
                {
                    await UntilAsync(() => fixture.ZipStarted.IsSet && model.IsExporting, "selected ZIP operation");
                    Require(!zipButton.IsEnabled && !Named<Button>("ExportButton").IsEnabled && !Named<Button>("ExportLibraryButton").IsEnabled && !Named<Button>("ConnectButton").IsEnabled && !model.CanNavigate,
                        "The ZIP operation disables conflicting native commands while preserving its selection snapshot.");
                    Require(dialog.Title?.ToString()?.Contains("ZIP", StringComparison.OrdinalIgnoreCase) == true && ((TextBlock)dialog.FindName("DestinationText")).Text.Contains(archivePath, StringComparison.Ordinal),
                        "The native ZIP progress dialog identifies the selected archive destination.");
                    fixture.ReleaseZip.Set();
                    fixture.ReleaseVersionRead.Set();
                    fixture.ReleaseVersionExport.Set();
                    await UntilAsync(() => dialog.Summary != null && !model.IsExporting, "native ZIP completion");
                    summary = dialog.Summary;
                    Require(((InfoBar)dialog.FindName("SummaryInfo")).IsOpen && ((ListView)dialog.FindName("ResultsList")).Items.Count == 2 && ((ProgressBar)dialog.FindName("ExportProgress")).Value == 2,
                        "The native ZIP dialog displays both selected-file results and completed progress.");
                    Require(((TextBlock)dialog.FindName("DestinationText")).Text.Contains(summary!.ArchivePath, StringComparison.Ordinal),
                        "The completed ZIP dialog displays the actual published filename after a destination collision.");
                    Require(((HyperlinkButton)dialog.FindName("OpenArchiveButton")).IsEnabled && !((HyperlinkButton)dialog.FindName("OpenReportButton")).IsEnabled,
                        "The completed ZIP dialog opens the archive and identifies its report as embedded.");
                }
                finally {fixture.HoldZip = false;fixture.ReleaseZip.Set();dialog.CancelAndCloseWhenFinished();await shown;}
                Require(summary?.Success == 2 && summary.Total == 2 && summary.ArchivePath != archivePath && File.Exists(summary.ArchivePath) && File.ReadAllText(archivePath) == "Existing ZIP destination must remain unchanged." && String.IsNullOrEmpty(summary.ReportPath) && summary.ArchiveReportEntry == "export-report.csv" &&
                    fixture.ZipSelectedIds.SequenceEqual(selectedIds),
                    "ZIP saving submits only checked current-folder files and publishes its actual archive and embedded report.");
                using (ZipArchive archive = ZipFile.OpenRead(summary!.ArchivePath))
                {
                    Require(archive.Entries.Count == 4 && archive.GetEntry(summary.ArchiveReportEntry) != null && archive.GetEntry("summary.txt") != null,
                        "The actual ZIP contains exactly two recovered files, its CSV report and its summary.");
                    var expected = new Dictionary<string, string> {
                        ["Documents/First/first.txt"] = "first library file\n", ["Documents/First/other.txt"] = "another library file\n"
                    };
                    foreach (DesktopExportEntry result in summary.Entries)
                    {
                        ZipArchiveEntry? entry = archive.GetEntry(result.Path);
                        Require(entry != null && result.Path.Replace(Path.DirectorySeparatorChar, '/').EndsWith(result.Document.Path, StringComparison.Ordinal),
                            "ZIP entry identity preserves the selected document hierarchy.");
                        using Stream data = entry!.Open();using var bytes = new MemoryStream();data.CopyTo(bytes);
                        byte[] expectedBytes = Encoding.UTF8.GetBytes(expected[result.Document.Path]);
                        string expectedHash = Convert.ToHexString(SHA256.HashData(expectedBytes)).ToLowerInvariant();
                        Require(bytes.ToArray().SequenceEqual(expectedBytes) && entry.Length == expectedBytes.Length && result.Bytes == expectedBytes.Length && result.Sha256 == expectedHash,
                            "ZIP bytes and SHA-256 match independently specified text for " + result.Document.Path + ".");
                    }
                    using var reportReader = new StreamReader(archive.GetEntry(summary.ArchiveReportEntry)!.Open(), Encoding.UTF8);
                    string report = reportReader.ReadToEnd();
                    Require(summary.Entries.All(result => report.Contains(result.Path, StringComparison.Ordinal) && report.Contains(result.Sha256, StringComparison.Ordinal)),
                        "The embedded ZIP report names each recovered entry and its verified hash.");
                    using var summaryReader = new StreamReader(archive.GetEntry("summary.txt")!.Open(), Encoding.UTF8);
                    string archiveSummary = summaryReader.ReadToEnd();
                    Require(archiveSummary.Contains("Total=2", StringComparison.Ordinal) && archiveSummary.Contains("Processed=2", StringComparison.Ordinal) && archiveSummary.Contains("Cancelled=False", StringComparison.Ordinal),
                        "The embedded summary records totals, completed files and cancellation state.");
                }
                Require(model.SelectedCount == 2 && model.VisibleItems.Where(item => item.CanCheck).All(item => item.IsChecked && item.ExportStatus == "Exported") && ReferenceEquals(model.CurrentFolder, first) && zipButton.IsEnabled,
                    "ZIP completion retains the user's current-folder checks and export result labels.");
                await model.NavigateAsync(list);
                Require(model.SelectedCount == 0 && !zipButton.IsEnabled,
                    "Opening an ordinary list clears ZIP checks and keeps its browse-only items out of export.");
                await model.NavigateAsync(second);await ToggleAsync(model.VisibleItems.Single());
            }
            private async Task IgnoredOutcomeDialogChecksAsync(TreeItemViewModel folder)
            {
                await model.NavigateAsync(folder);await ToggleAsync(model.VisibleItems.Single());
                Node selected = model.VisibleItems.Single().Node;
                fixture.ExtraAuditOutcomes.Clear();fixture.OutcomeOverrides.Clear();
                for (int index = 0; index < 8; index++)
                    fixture.ExtraAuditOutcomes.Add(new DesktopExportEntry {
                        Document = new Node { Kind = NodeKind.File, Id = Guid.NewGuid(), SiteId = selected.SiteId, ListId = selected.ListId,
                            Name = "Template" + index + ".aspx", Path = "Documents/Forms/Template" + index + ".aspx", HasStream = false },
                        Status = RecoveryStatus.Unavailable, Path = String.Empty, Sha256 = String.Empty, Message = "The form template has no stored content." });
                fixture.ExtraAuditOutcomes.Add(new DesktopExportEntry {
                    Document = new Node { Kind = NodeKind.File, Id = Guid.NewGuid(), SiteId = selected.SiteId, ListId = selected.ListId,
                        Name = "runtime.bin", Path = "Documents/runtime.bin", HasStream = true, StreamSchema = 66 },
                    Status = RecoveryStatus.Unsupported, Path = String.Empty, Sha256 = String.Empty, Message = "Runtime content format is unsupported." });
                string directory = Path.Combine(workspace, ".scratch", "WinUiIgnoredOutcomes", Guid.NewGuid().ToString("N"));
                var dialog = new ExportProgressDialog(model, directory) { XamlRoot = root.XamlRoot };
                var shown = dialog.ShowAsync();
                try
                {
                    await UntilAsync(() => dialog.Summary != null && !model.IsExporting, "ignored form audit outcomes");
                    InfoBar info = (InfoBar)dialog.FindName("SummaryInfo");ListView results = (ListView)dialog.FindName("ResultsList");
                    Require(dialog.Summary!.Success == 1 && dialog.Summary.Total == 10 && dialog.Summary.Skipped == 9 && dialog.Summary.Entries.Count == 10,
                        "The form and runtime unsupported outcomes remain in the complete audit summary.");
                    Require(info.Severity == InfoBarSeverity.Success && info.Message == "1 exported" && results.Items.Count == 1 && ((ExportResultRow)results.Items[0]).Status == "Exported",
                        "One recovered document with unavailable forms shows a green success summary and only its exported result.");
                    Require(!VisibleExportText(dialog).Contains("unavailable", StringComparison.OrdinalIgnoreCase) && !VisibleExportText(dialog).Contains("unsupported", StringComparison.OrdinalIgnoreCase) && !VisibleExportText(dialog).Contains("skipped", StringComparison.OrdinalIgnoreCase),
                        "Ignored templates and runtime unsupported content do not create visible warning or skipped text.");
                    Require(model.ExportCompleted == 1 && model.ExportTotal == 1 && File.ReadAllText(dialog.Summary.ReportPath).Contains("Unavailable", StringComparison.Ordinal) && File.ReadAllText(dialog.Summary.ReportPath).Contains("Unsupported", StringComparison.Ordinal),
                        "Visible progress counts only the recovered file while the CSV retains every ignored audit outcome.");
                }
                finally {dialog.CancelAndCloseWhenFinished();await shown;fixture.ExtraAuditOutcomes.Clear();}
                fixture.OutcomeOverrides[selected.Id] = RecoveryStatus.Unsupported;
                var unsupportedDialog = new ExportProgressDialog(model, Path.Combine(workspace, ".scratch", "WinUiIgnoredOutcomes", Guid.NewGuid().ToString("N"))) { XamlRoot = root.XamlRoot };
                var unsupportedShown = unsupportedDialog.ShowAsync();
                try
                {
                    await UntilAsync(() => unsupportedDialog.Summary != null && !model.IsExporting, "runtime unsupported content");
                    InfoBar info = (InfoBar)unsupportedDialog.FindName("SummaryInfo");
                    Require(info.Severity == InfoBarSeverity.Informational && info.Title == "Nothing to export" && info.Message == "No files were exported." && ((ListView)unsupportedDialog.FindName("ResultsList")).Items.Count == 0,
                        "A runtime unsupported-only result quietly reports nothing to export without warnings.");
                    Require(unsupportedDialog.Summary!.Skipped == 1 && unsupportedDialog.Summary.Entries.Single().Status == RecoveryStatus.Unsupported && File.ReadAllText(unsupportedDialog.Summary.ReportPath).Contains("Unsupported", StringComparison.Ordinal),
                        "The runtime unsupported-only operation still preserves its detailed audit record.");
                }
                finally {unsupportedDialog.CancelAndCloseWhenFinished();await unsupportedShown;}
                foreach (RecoveryStatus failure in new[] { RecoveryStatus.Corrupt, RecoveryStatus.SqlError })
                {
                    fixture.OutcomeOverrides[selected.Id] = failure;
                    var failureDialog = new ExportProgressDialog(model, Path.Combine(workspace, ".scratch", "WinUiIgnoredOutcomes", Guid.NewGuid().ToString("N"))) { XamlRoot = root.XamlRoot };
                    var failureShown = failureDialog.ShowAsync();
                    try
                    {
                        await UntilAsync(() => failureDialog.Summary != null && !model.IsExporting, "actual recovery failure");
                        InfoBar info = (InfoBar)failureDialog.FindName("SummaryInfo");ListView results = (ListView)failureDialog.FindName("ResultsList");
                        Require(failureDialog.Summary!.Failed == 1 && info.Severity == InfoBarSeverity.Warning && info.Message.Contains("1 failed", StringComparison.Ordinal) && results.Items.Count == 1,
                            "An actual " + failure + " recovery failure remains visible with warning severity.");
                        ExportResultRow row = (ExportResultRow)results.Items[0];
                        Require(row.Status == (failure == RecoveryStatus.Corrupt ? "Invalid content" : "Database error") && row.Detail.Contains(failure.ToString(), StringComparison.Ordinal),
                            "The visible failure result retains the actionable " + failure + " detail.");
                    }
                    finally {failureDialog.CancelAndCloseWhenFinished();await failureShown;}
                }
                fixture.OutcomeOverrides.Clear();await model.ExportAsync(Path.Combine(workspace, ".scratch", "WinUiIgnoredOutcomes", Guid.NewGuid().ToString("N")));
            }
            private async Task SelectedFileDialogChecksAsync(TreeItemViewModel first, TreeItemViewModel second)
            {
                await model.NavigateAsync(second);
                model.ClearSelection();
                ItemViewModel current=model.VisibleItems.Single();
                await ToggleAsync(current);
                string directory=Path.Combine(workspace,".scratch","WinUiCurrentFileFixture",Guid.NewGuid().ToString("N"));
                var dialog=new ExportProgressDialog(model,directory) {XamlRoot=root.XamlRoot};
                var shown=dialog.ShowAsync();
                try
                {
                    await UntilAsync(()=>dialog.Summary!=null && !model.IsExporting,"native checked-current export completion");
                    DesktopExportSummary summary=dialog.Summary!;
                    DesktopExportEntry entry=summary.Entries.Single();
                    byte[] expected=Encoding.UTF8.GetBytes("second library file\n");
                    Require(summary.Success==1 && entry.Document.HistoryVersion==0 &&
                        entry.Path==Path.Combine(directory,"second.txt") && Path.GetDirectoryName(summary.ReportPath)==directory &&
                        File.ReadAllBytes(entry.Path).SequenceEqual(expected) && entry.Bytes==expected.Length &&
                        entry.Sha256==Convert.ToHexString(SHA256.HashData(expected)).ToLowerInvariant(),
                        "Export selected writes the checked current document directly beside the report with verified current bytes.");
                    Require(dialog.SavedFilePath==entry.Path && ((TextBlock)dialog.FindName("DestinationText")).Text=="Saved file: "+entry.Path &&
                        ((HyperlinkButton)dialog.FindName("OpenFileButton")).Visibility==Visibility.Visible &&
                        ((HyperlinkButton)dialog.FindName("OpenFileButton")).IsEnabled &&
                        ((HyperlinkButton)dialog.FindName("OpenDestinationButton")).Content?.ToString()=="Show in folder",
                        "The normal selected-file dialog exposes the actual saved file and its Open file/Show in folder actions.");
                    Require(current.IsChecked && model.SelectedCount==1 && current.ExportStatus=="Exported" && ReferenceEquals(model.CurrentFolder,second),
                        "Current-file completion preserves the folder and its checked selection.");
                }
                finally {dialog.CancelAndCloseWhenFinished();await shown;}
                await model.NavigateAsync(first);
                model.SelectAllFiles();
                directory=Path.Combine(workspace,".scratch","WinUiCurrentSelectionFixture",Guid.NewGuid().ToString("N"));
                var multiple=new ExportProgressDialog(model,directory) {XamlRoot=root.XamlRoot};
                var multipleShown=multiple.ShowAsync();
                try
                {
                    await UntilAsync(()=>multiple.Summary!=null && !model.IsExporting,"native checked-files export completion");
                    DesktopExportSummary summary=multiple.Summary!;
                    Require(summary.Success==2 && summary.Entries.All(entry=>File.Exists(entry.Path) && Path.GetDirectoryName(entry.Path)==directory) &&
                        Path.GetDirectoryName(summary.ReportPath)==directory &&
                        ((HyperlinkButton)multiple.FindName("OpenDestinationButton")).Content?.ToString()=="Open destination" &&
                        ((HyperlinkButton)multiple.FindName("OpenFileButton")).Visibility==Visibility.Collapsed && multiple.SavedFilePath=="",
                        "Multiple checked files are visible beside their report in the selected directory without choosing an arbitrary file to open.");
                }
                finally {multiple.CancelAndCloseWhenFinished();await multipleShown;}
                model.ClearSelection();
            }
            private async Task VersionChecksAsync(TreeItemViewModel first, TreeItemViewModel second, TreeItemViewModel listFolder)
            {
                Button versionsButton = Named<Button>("VersionHistoryButton");
                await model.NavigateAsync(first);
                await SelectNativeItemAsync(Named<ListView>("ItemsList"), model.VisibleItems.Single(item => item.IsContainer));
                Require(!versionsButton.IsEnabled, "Focusing a folder cannot enable document version history.");
                await model.NavigateAsync(listFolder);
                await SelectNativeItemAsync(Named<ListView>("ItemsList"), model.VisibleItems.Single());
                Require(!versionsButton.IsEnabled, "An ordinary list item cannot enable document version history.");
                await model.NavigateAsync(second);model.ClearSelection();
                ItemViewModel currentFile = model.VisibleItems.Single();
                await SelectNativeItemAsync(Named<ListView>("ItemsList"), currentFile);
                Require(versionsButton.IsEnabled && model.SelectedCount == 0 && !Named<Button>("ExportButton").IsEnabled &&
                    new ButtonAutomationPeer(versionsButton).GetPattern(PatternInterface.Invoke) is IInvokeProvider,
                    "Focusing a document enables the native Versions command independently of file checkboxes.");
                string originalStatus = currentFile.ExportStatus;
                Node selected;
                fixture.HoldVersionRead = true;fixture.VersionReadStarted.Reset();fixture.ReleaseVersionRead.Reset();
                var dialog = new VersionHistoryDialog(model, currentFile.Node) { XamlRoot = root.XamlRoot };
                var shown = dialog.ShowAsync();
                try
                {
                    await UntilAsync(() => fixture.VersionReadStarted.IsSet && model.IsBusy, "native version history discovery");
                    Require(((ProgressRing)dialog.FindName("VersionsLoadingRing")).IsActive &&
                        !((ListView)dialog.FindName("VersionsList")).IsEnabled && !dialog.IsPrimaryButtonEnabled && !versionsButton.IsEnabled,
                        "The native version dialog displays loading progress and prevents premature export.");
                    fixture.ReleaseVersionRead.Set();
                    ListView versions = (ListView)dialog.FindName("VersionsList");
                    await UntilAsync(() => versions.Items.Count == 2 && !model.IsBusy && !((ProgressRing)dialog.FindName("VersionsLoadingRing")).IsActive,
                        "native version history rows");
                    VersionViewModel current = versions.Items.Cast<VersionViewModel>().Single(row => row.HistoryVersion == 0);
                    VersionViewModel older = versions.Items.Cast<VersionViewModel>().Single(row => row.HistoryVersion == 512);
                    Require(current.VersionText == "2.0" && current.StateText == "Current" && older.VersionText == "1.0" && older.StateText == "Historical" &&
                        current.Name == "second.txt" && older.Name == current.Name && current.CanExport && older.CanExport &&
                        current.SizeText.Length > 0 && older.SizeText.Length > 0 && current.ModifiedText != older.ModifiedText,
                        "The real version list distinguishes current 2.0 and stored 1.0 with their own size and modified metadata.");
                    Require(((TextBlock)dialog.FindName("VersionCountText")).Text == "1 older version recorded" &&
                        !((InfoBar)dialog.FindName("VersionsInfo")).IsOpen,
                        "The loaded native dialog reports the older version count without an error.");
                    await SelectNativeItemAsync(versions, current);
                    Require(dialog.IsPrimaryButtonEnabled && ((TextBlock)dialog.FindName("VersionSelectionText")).Text.Contains("2.0", StringComparison.Ordinal),
                        "Selecting the current version updates the native export choice.");
                    await SelectNativeItemAsync(versions, older);
                    DependencyObject container = versions.ContainerFromItem(older) ?? throw new InvalidOperationException("The historical row was not realized.");
                    string visibleRow = String.Join(" ", FindChildren<TextBlock>(container).Select(text => text.Text));
                    Require(visibleRow.Contains("1.0", StringComparison.Ordinal) && visibleRow.Contains("Historical", StringComparison.Ordinal) &&
                        dialog.IsPrimaryButtonEnabled && ((TextBlock)dialog.FindName("VersionSelectionText")).Text.Contains("1.0", StringComparison.Ordinal),
                        "Native row selection renders and chooses the older version for export.");
                    Button primary = FindChildren<Button>(dialog).Single(button => button.Name == "PrimaryButton");
                    IInvokeProvider? invoke = new ButtonAutomationPeer(primary).GetPattern(PatternInterface.Invoke) as IInvokeProvider;
                    Require(primary.IsEnabled && invoke != null && dialog.PrimaryButtonText == "Export version",
                        "The native version export button exposes its enabled Invoke action.");
                    invoke!.Invoke();
                    Require(await shown == ContentDialogResult.Primary && dialog.SelectedVersionNode != null,
                        "Invoking Export version closes the native dialog with the chosen version.");
                    selected = dialog.SelectedVersionNode!;
                    Require(SameVersion(selected, older.Node) && selected.HistoryVersion == 512 && selected.UiVersion == 512 && selected.InternalVersion == 513,
                        "The native dialog preserves the exact historical source identity rather than the current file identity.");
                }
                finally {fixture.HoldVersionRead = false;fixture.ReleaseVersionRead.Set();dialog.RequestClose();await shown;}
                string directory = Path.Combine(workspace, ".scratch", "WinUiVersionFixture", Guid.NewGuid().ToString("N"));
                string currentPath = Path.Combine(directory, currentFile.Node.SiteId.ToString("D"), currentFile.Path.Replace('/', Path.DirectorySeparatorChar));
                Directory.CreateDirectory(Path.GetDirectoryName(currentPath)!);
                byte[] expectedCurrent = Encoding.UTF8.GetBytes("second library file\n");
                byte[] expectedHistorical = Encoding.UTF8.GetBytes("second historical version 1.0\n");
                File.WriteAllBytes(currentPath, expectedCurrent);
                fixture.HoldVersionExport = true;fixture.VersionExportStarted.Reset();fixture.ReleaseVersionExport.Reset();
                var export = new ExportProgressDialog(model, directory, version: selected) { XamlRoot = root.XamlRoot };
                var exportShown = export.ShowAsync();
                DesktopExportSummary? summary = null;
                try
                {
                    await UntilAsync(() => fixture.VersionExportStarted.IsSet && model.IsExporting, "native historical export");
                    Require(model.SelectedCount == 0 && !versionsButton.IsEnabled && export.Title?.ToString() == "Export document version" &&
                        ((TextBlock)export.FindName("ExportScopeText")).Text.Contains("version 1.0", StringComparison.Ordinal),
                        "The normal progress dialog exports the selected old version without checked current files.");
                    fixture.ReleaseVersionExport.Set();
                    await UntilAsync(() => export.Summary != null && !model.IsExporting, "native historical export completion");
                    summary = export.Summary;
                    Require(summary!.Total == 1 && summary.Success == 1 && summary.Failed == 0 && summary.Skipped == 0 &&
                        ((InfoBar)export.FindName("SummaryInfo")).Severity == InfoBarSeverity.Success &&
                        ((ListView)export.FindName("ResultsList")).Items.Count == 1 && ((ProgressBar)export.FindName("ExportProgress")).Value == 1,
                        "The genuine version export dialog reports one completed historical file.");
                    DesktopExportEntry saved=summary!.Entries.Single();
                    var openFile=(HyperlinkButton)export.FindName("OpenFileButton");
                    var showFolder=(HyperlinkButton)export.FindName("OpenDestinationButton");
                    Require(export.SavedFilePath==saved.Path && File.Exists(saved.Path) &&
                        Path.GetDirectoryName(saved.Path)==directory && Path.GetDirectoryName(summary.ReportPath)==directory &&
                        ((TextBlock)export.FindName("DestinationText")).Text=="Saved file: "+saved.Path,
                        "The version document is saved beside its CSV directly in the selected folder and its exact path is visible.");
                    Require(openFile.Visibility==Visibility.Visible && openFile.IsEnabled && showFolder.Content?.ToString()=="Show in folder" &&
                        new HyperlinkButtonAutomationPeer(openFile).GetPattern(PatternInterface.Invoke) is IInvokeProvider &&
                        new HyperlinkButtonAutomationPeer(showFolder).GetPattern(PatternInterface.Invoke) is IInvokeProvider,
                        "A saved version exposes native Open file and Show in folder actions for the actual document.");
                }
                finally {fixture.HoldVersionExport = false;fixture.ReleaseVersionExport.Set();export.CancelAndCloseWhenFinished();await exportShown;}
                DesktopExportEntry entry = summary!.Entries.Single();
                string expectedHash = Convert.ToHexString(SHA256.HashData(expectedHistorical)).ToLowerInvariant();
                Require(File.ReadAllBytes(entry.Path).SequenceEqual(expectedHistorical) && entry.Bytes == expectedHistorical.Length && entry.Sha256 == expectedHash &&
                    Path.GetFileName(entry.Path) == "second (v1.0).txt" && SameVersion(entry.Document, selected) &&
                    fixture.CapturedVersion != null && SameVersion(fixture.CapturedVersion, selected),
                    "Historical export preserves exact version identity, version filename, independent bytes and SHA-256.");
                string report = File.ReadAllText(summary.ReportPath);
                Require(report.Contains("UiVersion,HistoryVersion,Level,InternalVersion", StringComparison.Ordinal) &&
                    report.Contains(",512,512,1,513", StringComparison.Ordinal) && report.Contains(expectedHash, StringComparison.Ordinal),
                    "The version export CSV records the older version identity and verified hash.");
                Require(File.ReadAllBytes(currentPath).SequenceEqual(expectedCurrent) && !expectedCurrent.SequenceEqual(expectedHistorical) &&
                    ReferenceEquals(model.CurrentFolder, second) && model.VisibleItems.Single() == currentFile && !currentFile.IsChecked && model.SelectedCount == 0 &&
                    currentFile.ExportStatus == originalStatus && model.ExportCompleted == 1 && model.ExportTotal == 1 && model.LastExport == summary,
                    "Exporting an older version leaves current bytes, the current row status, location and unchecked selection unchanged.");
                fixture.FailVersionRead = true;
                var failed = new VersionHistoryDialog(model, currentFile.Node) { XamlRoot = root.XamlRoot };
                var failedShown = failed.ShowAsync();
                try
                {
                    await UntilAsync(() => ((InfoBar)failed.FindName("VersionsInfo")).IsOpen && !model.IsBusy, "native version discovery failure");
                    Require(((InfoBar)failed.FindName("VersionsInfo")).Severity == InfoBarSeverity.Error &&
                        ((InfoBar)failed.FindName("VersionsInfo")).Message.Contains("fixture", StringComparison.Ordinal) &&
                        !failed.IsPrimaryButtonEnabled && ((ListView)failed.FindName("VersionsList")).Items.Count == 0,
                        "A failed version query presents an actionable native error and prevents export.");
                }
                finally {fixture.FailVersionRead = false;failed.RequestClose();await failedShown;}
                var recovered = new VersionHistoryDialog(model, currentFile.Node) { XamlRoot = root.XamlRoot };
                var recoveredShown = recovered.ShowAsync();
                try
                {
                    await UntilAsync(() => ((ListView)recovered.FindName("VersionsList")).Items.Count == 2 && !model.IsBusy, "native version discovery retry");
                    Require(recovered.IsPrimaryButtonEnabled && !((InfoBar)recovered.FindName("VersionsInfo")).IsOpen && model.ErrorText == "",
                        "Reopening version history after a failed query recovers the available choices.");
                }
                finally {recovered.RequestClose();await recoveredShown;}
            }
            private async Task SelectNativeItemAsync(ListView list, object item)
            {
                list.ScrollIntoView(item);await SettleAsync();
                ListViewItem? container = list.ContainerFromItem(item) as ListViewItem;
                Require(container != null && container.IsEnabled, "The requested row has a realized enabled native list container.");
                // Selection belongs to the native data-item peer supplied by
                // the parent selector, rather than its visual-container peer.
                ListViewBaseAutomationPeer parentPeer = FrameworkElementAutomationPeer.CreatePeerForElement(list) as ListViewBaseAutomationPeer
                    ?? new ListViewAutomationPeer(list);
                var children = parentPeer.GetChildren();
                ListViewItemDataAutomationPeer? dataPeer = children?.OfType<ListViewItemDataAutomationPeer>()
                    .SingleOrDefault(peer => ReferenceEquals(peer.Item, item));
                if (dataPeer == null && children != null)
                    Write("SELECTION peers=" + String.Join(",", children.Select(peer => peer.GetType().Name + ":" +
                        (peer is ItemAutomationPeer data ? data.Item?.GetType().Name ?? "null" : "container"))));
                Require(dataPeer != null, "The native ListView exposes the requested data-item automation peer.");
                ISelectionItemProvider? selection = dataPeer!.GetPattern(PatternInterface.SelectionItem) as ISelectionItemProvider;
                Require(selection != null, "The native data row exposes its SelectionItem automation pattern.");
                selection!.Select();await SettleAsync();
                Require(selection.IsSelected && container!.IsSelected && ReferenceEquals(list.SelectedItem, item),
                    "The native selection action chooses and selects the requested realized row.");
            }

            private static string VisibleExportText(ExportProgressDialog dialog)
                => ((InfoBar)dialog.FindName("SummaryInfo")).Message + " " + ((TextBlock)dialog.FindName("ProgressText")).Text + " " + ((TextBlock)dialog.FindName("CurrentFileText")).Text;
            private async Task ConnectionDialogChecksAsync()
            {
                SqlConnectionOptions? captured = null;
                int discoveryCalls = 0;
                var connected = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
                var dialog = new ConnectionDialog(new SqlConnectionOptions(), options => {
                    captured = options.Clone();
                    return connected.Task;
                }, _ => {
                    System.Threading.Interlocked.Increment(ref discoveryCalls);
                    return new List<string> { "master", "WSS_Content" };
                }) { XamlRoot = root.XamlRoot };
                var shown = dialog.ShowAsync();
                try
                {
                    var server = (TextBox)dialog.FindName("ServerBox");
                    var database = (ComboBox)dialog.FindName("DatabaseBox");
                    await SettleAsync();
                    Require(server.Text == "" && database.Text == "" && database.Items.Count == 0 && discoveryCalls == 0 &&
                        !((InfoBar)dialog.FindName("ConnectionError")).IsOpen,
                        "A new connection opens blank without discovery or a validation error.");
                    server.Text = "SQL";
                    await SettleAsync();
                    var loadDatabases = (Button)dialog.FindName("RefreshDatabasesButton");
                    var load = new ButtonAutomationPeer(loadDatabases).GetPattern(PatternInterface.Invoke) as IInvokeProvider;
                    Require(load != null, "The Load databases button exposes its Invoke pattern.");
                    load!.Invoke();
                    await SettleAsync();
                    Write("DISCOVERY calls=" + discoveryCalls + " choices=" + database.Items.Count + " status=" +
                        ((TextBlock)dialog.FindName("DiscoveryText")).Text);
                    await UntilAsync(() => database.Items.Count == 2, "native dialog database discovery");
                    Require(database.Text == "" && database.SelectedIndex < 0 && database.Items.Cast<string>().Contains("WSS_Content") && discoveryCalls == 1,
                        "Loading databases uses the entered server and leaves the database choice to the user.");
                    var credentials = (Grid)dialog.FindName("CredentialsPanel");
                    var currentUser = (TextBlock)dialog.FindName("CurrentUserText");
                    Require(credentials.Visibility == Visibility.Collapsed && currentUser.Visibility == Visibility.Visible,
                        "Windows authentication displays the current account and hides SQL credentials.");
                    var authentication = (ComboBox)dialog.FindName("AuthenticationBox");
                    var authenticationPeer = new ComboBoxAutomationPeer(authentication);
                    authenticationPeer.Expand();
                    await SettleAsync();
                    Require(authentication.IsDropDownOpen, "The native authentication ComboBox opens through its automation peer.");
                    authentication.SelectedIndex = 1;
                    authenticationPeer.Collapse();
                    await SettleAsync();
                    Require(authentication.SelectedIndex == 1 && credentials.Visibility == Visibility.Visible && currentUser.Visibility == Visibility.Collapsed,
                        "Choosing SQL authentication reveals native username and password fields.");
                    ((TextBox)dialog.FindName("UsernameBox")).Text = "fixture-user";
                    ((PasswordBox)dialog.FindName("PasswordBox")).Password = "fixture-password";
                    database.Text = "WSS_Content";
                    Button? primary = FindChildren<Button>(dialog).FirstOrDefault(v => v.Name == "PrimaryButton");
                    Require(primary != null && primary.IsEnabled, "The native connection dialog realizes its Connect button.");
                    IInvokeProvider? invoke = new ButtonAutomationPeer(primary!).GetPattern(PatternInterface.Invoke) as IInvokeProvider;
                    Require(invoke != null, "The native Connect button exposes its Invoke pattern.");
                    invoke!.Invoke();
                    await UntilAsync(() => captured != null, "native Connect callback");
                    Require(!((TextBox)dialog.FindName("ServerBox")).IsEnabled && !dialog.IsPrimaryButtonEnabled && ((StackPanel)dialog.FindName("ConnectingPanel")).Visibility == Visibility.Visible,
                        "The connection dialog disables editing and displays native connection progress.");
                    connected.TrySetResult("");
                    ContentDialogResult result = await shown;
                    Require(result == ContentDialogResult.Primary && captured!.Server == "SQL" && captured.Database == "WSS_Content" &&
                        captured.Authentication == SqlAuthenticationMode.SqlLogin && captured.Username == "fixture-user" && captured.Password == "fixture-password",
                        "The native Connect action submits the chosen options to the injected callback.");
                    Require(dialog.Options.Password == "" && ((PasswordBox)dialog.FindName("PasswordBox")).Password == "",
                        "Closing the native connection dialog clears its password and retained options.");
                }
                finally { connected.TrySetResult(""); dialog.RequestClose(); await shown; }
            }

            private static IEnumerable<T> FindChildren<T>(DependencyObject value) where T : DependencyObject
            {
                if (value is T item) yield return item;
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(value); i++)
                    foreach (T child in FindChildren<T>(VisualTreeHelper.GetChild(value, i))) yield return child;
            }
            private async Task SqlChecksAsync()
            {
                Require(await model.ConnectAsync(new SqlConnectionOptions { Server = "SQL", Database = "WSS_Content" }), "Windows authentication opens SQL / WSS_Content.");
                Require(model.RootNodes.Count > 0, "The live database exposes site collections through WinUI.");
                Guid jpegId = Guid.Parse("44da1cf3-e59f-49bb-b618-bc69c5f5ddcc");
                Guid textId = Guid.Parse("6afefc32-a65e-4c67-82b6-8145cec99fbd");
                (TreeItemViewModel Jpeg, TreeItemViewModel Text, TreeItemViewModel List) locations = await LocateFixturesAsync(jpegId, textId);
                Require(locations.Jpeg != locations.Text, "The real Word document and text fixtures are in different folders.");
                await model.NavigateAsync(locations.Jpeg);
                ItemViewModel jpeg = model.VisibleItems.Single(v => v.Node.Id == jpegId);
                await ToggleAsync(jpeg);
                await model.NavigateAsync(locations.Text);
                Require(model.SelectedCount == 0 && !jpeg.IsChecked, "Real SQL folder navigation clears the Word document selection.");
                ItemViewModel text = model.VisibleItems.Single(v => v.Node.Id == textId);
                await ToggleAsync(text);
                await model.NavigateAsync(locations.Jpeg);
                Require(model.SelectedCount == 0 && !text.IsChecked, "Returning to the real Word document folder clears the text selection.");
                jpeg = model.VisibleItems.Single(v => v.Node.Id == jpegId);
                await ToggleAsync(jpeg);
                string directory = Path.Combine(workspace, ".scratch", "WinUiIntegration", DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
                var exportDialog = new ExportProgressDialog(model, directory) { XamlRoot = root.XamlRoot };
                var exportShown = exportDialog.ShowAsync();
                DesktopExportSummary? summary;
                try
                {
                    await UntilAsync(() => exportDialog.Summary != null && !model.IsExporting, "native SQL export dialog completion");
                    summary = exportDialog.Summary;
                    Require(((InfoBar)exportDialog.FindName("SummaryInfo")).IsOpen && ((ListView)exportDialog.FindName("ResultsList")).Items.Count == 1,
                        "The genuine export dialog displays the live export summary and one result.");
                    Require(((ProgressBar)exportDialog.FindName("ExportProgress")).Value == 1,
                        "The genuine export dialog completes its native progress bar.");
                }
                finally { exportDialog.CancelAndCloseWhenFinished(); await exportShown; }
                Require(summary != null && summary.Total == 1 && summary.Success == 1 && summary.Failed == 0 && summary.Skipped == 0,
                    "The WinUI export operation recovers exactly one selected live Word document.");
                DesktopExportEntry entry = summary!.Entries.Single();
                Require(entry.Document.Id == jpegId && File.Exists(entry.Path) && entry.Bytes == 10886,
                    "The recovered Word document has the known identity and 10886-byte length.");
                using (FileStream file = File.OpenRead(entry.Path))
                {
                    string hash = Convert.ToHexString(SHA256.HashData(file)).ToLowerInvariant();
                    Require(hash == "cb5357c9daea4fe0e71d5296eb41e7df526e768e3cf236000e6a4d5bf43b8e0a" && entry.Sha256.Equals(hash, StringComparison.OrdinalIgnoreCase),
                        "The recovered Word document matches the independently recorded SHA-256.");
                    Write("EXPORT " + entry.Path + " SHA256=" + hash);
                }
                Require(model.ExportCompleted == 1 && model.ExportTotal == 1 && !model.IsExporting && model.LastExport == summary,
                    "The live SQL export finishes through the ordinary progress state.");
                await model.NavigateAsync(locations.List);
                Require(model.VisibleItems.Any(v => v.Node.Kind == NodeKind.ListItem) && model.VisibleItems.Where(v => v.Node.Kind == NodeKind.ListItem).All(v => !v.CanCheck),
                    "The actual ordinary SharePoint list exposes browse-only items.");
                model.SelectAllFiles();
                Require(model.SelectedCount == 0, "List items cannot enter the live export selection.");
                Write("SQL roots=" + model.RootNodes.Count + " listItems=" + model.VisibleItems.Count);
                await ExtendedSqlDialogsAsync(locations.Jpeg);
                await BulkAndDeletedSqlDialogsAsync(locations.Jpeg,locations.Text);
                await model.NavigateAsync(locations.Jpeg);
                await ToggleAsync(model.VisibleItems.Single(v => v.Node.Id == jpegId));
                Require(Named<ListView>("ItemsList").ActualHeight >= 48, "The final SQL Word document folder has a visible native file viewport.");
                await TreeSelectionVisibleAsync(locations.Jpeg);
            }

            private async Task BulkAndDeletedSqlDialogsAsync(TreeItemViewModel documentLocation,TreeItemViewModel textLocation)
            {
                string destination=Path.Combine(workspace,".scratch","WinUiBulkIntegration",Guid.NewGuid().ToString("N"));
                await model.NavigateAsync(documentLocation);var menu=(MenuFlyout)Named<Button>("MoreExportsButton").Flyout;
                Require(menu.Items.OfType<MenuFlyoutItem>().Single(item=>item.Name=="SaveLibraryZipMenu").IsEnabled && menu.Items.OfType<MenuFlyoutItem>().Single(item=>item.Name=="ExportHistoryPackageMenu").IsEnabled,"Native More menu enables library ZIP and package history.");
                var zip=new ExportProgressDialog(model,Path.Combine(destination,"library.zip"),asZip:true,zipScope:documentLocation.Node){XamlRoot=root.XamlRoot};var shown=zip.ShowAsync();
                try{await UntilAsync(()=>zip.Summary!=null && !model.IsExporting,"native whole-library ZIP");Require(zip.Summary.Success>0 && File.Exists(zip.Summary.ArchivePath),"Native whole-library ZIP publishes all retained current files.");}
                finally{zip.CancelAndCloseWhenFinished();await shown;}
                await model.NavigateAsync(textLocation);
                Require(menu.Items.OfType<MenuFlyoutItem>().Single(item=>item.Name=="SaveFolderZipMenu").IsEnabled,"Native More menu enables nested folder ZIP.");
                var folderZip=new ExportProgressDialog(model,Path.Combine(destination,"folder.zip"),asZip:true,zipScope:textLocation.Node){XamlRoot=root.XamlRoot};shown=folderZip.ShowAsync();
                try{await UntilAsync(()=>folderZip.Summary!=null && !model.IsExporting,"native whole-folder ZIP");Require(folderZip.Summary.Success>0 && File.Exists(folderZip.Summary.ArchivePath),"Native folder ZIP succeeds independently of tree expansion.");}
                finally{folderZip.CancelAndCloseWhenFinished();await shown;}
                await model.NavigateAsync(documentLocation);ItemViewModel word=model.VisibleItems.Single(item=>item.Node.Id==Guid.Parse("44da1cf3-e59f-49bb-b618-bc69c5f5ddcc"));await ToggleAsync(word);
                Require(menu.Items.OfType<MenuFlyoutItem>().Single(item=>item.Name=="ExportVersionsMenu").IsEnabled,"Checking a document enables native batch version recovery.");
                var versions=new ExportProgressDialog(model,Path.Combine(destination,"versions"),allVersions:true){XamlRoot=root.XamlRoot};shown=versions.ShowAsync();
                try{await UntilAsync(()=>versions.Summary!=null && !model.IsExporting,"native batch retained versions");Require(versions.Summary.Success==2 && versions.Summary.Entries.Any(entry=>entry.Document.HistoryVersion==512 && entry.Bytes==10877),"Native batch recovery exports current and exact old Word bytes.");}
                finally{versions.CancelAndCloseWhenFinished();await shown;}
                var history=new ExportProgressDialog(model,Path.Combine(destination,"history"),packageScope:documentLocation.Node,includeHistory:true){XamlRoot=root.XamlRoot};shown=history.ShowAsync();
                try{await UntilAsync(()=>!model.IsExporting && (history.Summary!=null || !String.IsNullOrEmpty(model.ErrorText)),"native history XML package");Require(history.Summary?.Success==1 && File.Exists(Path.Combine(history.Summary.PackagePath,"Manifest.xml")),"Native XML package includes retained history: "+model.ErrorText);}
                finally{history.CancelAndCloseWhenFinished();await shown;}
                TreeItemViewModel site=model.RootNodes.Single(root=>root.Node.SiteId==Guid.Parse("cbd6e0be-6ee8-4b9d-9b04-d9031327831b"));
                await model.NavigateAsync(site);TreeItemViewModel deleted=site.Children.Single(child=>child.Node.Kind==NodeKind.DeletedItems);await model.NavigateAsync(deleted);
                ItemViewModel jpeg=model.VisibleItems.Single(item=>item.Node.Id==Guid.Parse("f0675d2e-4225-4ea9-8358-da3f7da48220") && item.Node.IsDeleted);await ToggleAsync(jpeg);
                Require(model.CanExport && jpeg.KindText.Contains("Deleted",StringComparison.Ordinal),"Deleted JPEG is visible and exportable through native checkboxes.");
                var recovery=new ExportProgressDialog(model,Path.Combine(destination,"deleted")){XamlRoot=root.XamlRoot};shown=recovery.ShowAsync();
                try{await UntilAsync(()=>recovery.Summary!=null && !model.IsExporting,"native deleted JPEG recovery");DesktopExportEntry entry=recovery.Summary.Entries.Single();Require(entry.Status==RecoveryStatus.Success && entry.Bytes==212541 && entry.Sha256=="48c41a496332b7953b71429e76ad4a662210b50754626042e0c1d7b6affe64ef","Native deleted JPEG export preserves known bytes and digest.");}
                finally{recovery.CancelAndCloseWhenFinished();await shown;}
                ItemViewModel deletedItem=model.VisibleItems.Single(item=>item.Node.IsDeleted && item.Node.Kind==NodeKind.ListItem && item.Node.ListItemId==2);Named<ListView>("ItemsList").SelectedItem=deletedItem;await SettleAsync();
                Require(menu.Items.OfType<MenuFlyoutItem>().Single(item=>item.Name=="ExportDeletedItemMenu").IsEnabled,"Native More menu enables deleted ordinary-item metadata recovery.");
                var metadata=new ExportProgressDialog(model,Path.Combine(destination,"metadata"),deletedItem:deletedItem.Node){XamlRoot=root.XamlRoot};shown=metadata.ShowAsync();
                try{await UntilAsync(()=>metadata.Summary!=null && !model.IsExporting,"native deleted-item XML metadata");Require(metadata.Summary.Success==1 && File.ReadAllText(metadata.Summary.Entries.Single().Path).Contains("bbb",StringComparison.Ordinal),"Native deleted-item recovery exports retained source values.");}
                finally{metadata.CancelAndCloseWhenFinished();await shown;}
            }            private async Task ExtendedSqlDialogsAsync(TreeItemViewModel jpegLocation)
            {
                IEnumerable<TreeItemViewModel> Descendants(TreeItemViewModel item)
                {
                    yield return item;
                    foreach(var child in item.Children)
                        foreach(var descendant in Descendants(child)) yield return descendant;
                }
                TreeItemViewModel list=model.RootNodes.SelectMany(Descendants).Single(v=>v.Node.Kind==NodeKind.List &&
                    v.Node.ListId==Guid.Parse("cb9a0845-257b-40e1-8379-7468dcce5500"));
                await model.NavigateAsync(list);
                ItemViewModel item=model.VisibleItems.Single(v=>v.Node.Id==Guid.Parse("701fd014-9894-426f-8f2e-bcf01130b2a8"));
                Named<ListView>("ItemsList").SelectedItem=item;
                await SettleAsync();
                var menu=(MenuFlyout)Named<Button>("MoreExportsButton").Flyout;
                Require(menu.Items.OfType<MenuFlyoutItem>().Single(v=>v.Name=="ExportItemAttachmentsMenu").IsEnabled &&
                    item.ExportStatus=="Has attachments" && !item.CanCheck,
                    "Focusing a real item offers attachment export without enabling document checkboxes.");
                string destination=Path.Combine(workspace,".scratch","WinUiExtendedIntegration",Guid.NewGuid().ToString("N"));
                var attachments=new ExportProgressDialog(model,Path.Combine(destination,"attachments"),attachmentScope:item.Node){XamlRoot=root.XamlRoot};
                var shown=attachments.ShowAsync();
                try
                {
                    await UntilAsync(()=>!model.IsExporting && (attachments.Summary!=null || !String.IsNullOrEmpty(model.ErrorText)),"native SQL attachments");
                    Require(attachments.Summary?.Success==1 && attachments.Summary.Failed==0 && attachments.Summary.Entries.Single().Bytes==454216 &&
                        attachments.Summary.Entries.Single().Document.AttachmentOwnerId==item.Node.Id &&
                        File.Exists(attachments.Summary.Entries.Single().Path),
                        "The native attachment dialog exports the actual PDF bound to its current list item.");
                    Require(((InfoBar)attachments.FindName("SummaryInfo")).Severity==InfoBarSeverity.Success &&
                        ((TextBlock)attachments.FindName("DestinationText")).Text.Contains(attachments.Summary!.Entries.Single().Path,StringComparison.Ordinal),
                        "Attachment completion displays the actual saved file with success.");
                }
                finally{attachments.CancelAndCloseWhenFinished();await shown;}
                foreach(Node scope in new[]{list.Node,jpegLocation.Node.Kind==NodeKind.Library ? jpegLocation.Node :
                    throw new InvalidOperationException("JPEG library fixture moved.")})
                {
                    await model.NavigateAsync(scope.Kind==NodeKind.List ? list : jpegLocation);
                    var package=new ExportProgressDialog(model,Path.Combine(destination,"packages"),packageScope:scope){XamlRoot=root.XamlRoot};
                    var packageShown=package.ShowAsync();
                    try
                    {
                        await UntilAsync(()=>!model.IsExporting && (package.Summary!=null || !String.IsNullOrEmpty(model.ErrorText)),"native SQL XML package");
                        Require(package.Summary?.Success==1 && !String.IsNullOrEmpty(package.Summary.PackagePath) &&
                            File.Exists(Path.Combine(package.Summary.PackagePath,"Manifest.xml")),
                            "The native XML package dialog publishes a complete list/library package: "+model.ErrorText);
                        Require(package.SavedFilePath=="" && ((HyperlinkButton)package.FindName("OpenFileButton")).Visibility==Visibility.Collapsed &&
                            ((HyperlinkButton)package.FindName("OpenDestinationButton")).Content.ToString()=="Open package folder" &&
                            ((TextBlock)package.FindName("DestinationText")).Text=="Package: "+package.Summary!.PackagePath,
                            "Package completion offers the actual package folder without treating it as a document.");
                    }
                    finally{package.CancelAndCloseWhenFinished();await packageShown;}
                }
            }
            private async Task<(TreeItemViewModel Jpeg, TreeItemViewModel Text, TreeItemViewModel List)> LocateFixturesAsync(Guid jpegId, Guid textId)
            {
                TreeItemViewModel? jpeg = null, text = null, list = null;
                var queue = new Queue<TreeItemViewModel>(model.RootNodes);
                int visited = 0;
                while (queue.Count > 0 && (jpeg == null || text == null || list == null))
                {
                    TreeItemViewModel folder = queue.Dequeue();
                    if (++visited > 250) throw new InvalidOperationException("The known SQL fixtures were not found within 250 locations.");
                    await model.NavigateAsync(folder);
                    if (!String.IsNullOrWhiteSpace(model.ErrorText)) throw new InvalidOperationException(model.ErrorText);
                    if (model.VisibleItems.Any(v => v.Node.Id == jpegId)) jpeg = folder;
                    if (model.VisibleItems.Any(v => v.Node.Id == textId)) text = folder;
                    if (folder.Node.Kind == NodeKind.List && folder.Name == "List1_WF") list = folder;
                    foreach (TreeItemViewModel child in folder.Children)
                    {
                        // Ordinary lists are needed only for the known browse-only fixture.
                        if (child.Node.Kind != NodeKind.List || child.Name == "List1_WF") queue.Enqueue(child);
                    }
                }
                if (jpeg == null || text == null || list == null) throw new InvalidOperationException("The known JPEG, text, or list fixture is absent from the source.");
                Write("SQL locations=" + visited + " JPEG=" + jpeg.Path + " TEXT=" + text.Path + " LIST=" + list.Path);
                return (jpeg, text, list);
            }

            private static TreeViewNode? FindNativeNode(IEnumerable<TreeViewNode> nodes, TreeItemViewModel folder)
            {
                foreach (TreeViewNode node in nodes)
                {
                    if (ReferenceEquals(node.Content, folder)) return node;
                    TreeViewNode? child = FindNativeNode(node.Children, folder);
                    if (child != null) return child;
                }
                return null;
            }
            private async Task TreeSelectionVisibleAsync(TreeItemViewModel folder)
            {
                await SettleAsync();
                TreeView tree = Named<TreeView>("ExplorerTree");
                TreeViewNode? nativeNode = FindNativeNode(tree.RootNodes, folder);
                ListViewBase? nativeList = FindChild<ListViewBase>(tree);
                int flattenedIndex = -1;
                if (nativeList != null)
                    for (int index = 0; index < nativeList.Items.Count; index++)
                    {
                        object value = nativeList.Items[index];
                        if (ReferenceEquals(value, folder) || value is TreeViewNode node && ReferenceEquals(node.Content, folder)) { flattenedIndex = index; break; }
                    }
                FrameworkElement? container = (nativeNode == null ? null : tree.ContainerFromNode(nativeNode)) as FrameworkElement ??
                    tree.ContainerFromItem(folder) as FrameworkElement ??
                    (flattenedIndex < 0 ? null : nativeList!.ContainerFromIndex(flattenedIndex)) as FrameworkElement;
                Write("TREE target=" + folder.Name + " selected=" + (tree.SelectedItem as TreeItemViewModel)?.Name +
                    " nativeRoots=" + tree.RootNodes.Count + " container=" + (container?.GetType().Name ?? "none") +
                    " list=" + nativeList?.GetType().Name + " items=" + nativeList?.Items.Count + " contains=" + nativeList?.Items.Contains(folder));
                if (nativeList != null)
                {
                    foreach (object value in nativeList.Items)
                    {
                        TreeItemViewModel? item = value as TreeItemViewModel ?? (value as TreeViewNode)?.Content as TreeItemViewModel;
                        if (item?.Name == folder.Name)
                            Write("TREE matchingItem=" + value.GetType().FullName + " same=" + ReferenceEquals(item, folder) + " nativeContent=" + (value as TreeViewNode)?.Content?.GetType().FullName);
                    }
                }
                Require((ReferenceEquals(tree.SelectedItem, folder) || nativeNode != null && ReferenceEquals(tree.SelectedNode, nativeNode)) && container != null && container.ActualHeight > 0,
                    "The current folder has a realized native navigation-tree container.");
                Point point = container!.TransformToVisual(tree).TransformPoint(new Point(0, 0));
                Require(point.Y >= -1 && point.Y + container.ActualHeight <= tree.ActualHeight + 1,
                    "The selected native navigation-tree folder is within its visible viewport.");
            }
            private async Task BoundsAsync(int width, int height)
            {
                window.AppWindow.Resize(new Windows.Graphics.SizeInt32(width, height));
                await SettleAsync();
                Require(Named<ListView>("ItemsList").ActualHeight >= 48, "A native file row remains visible at " + width + "x" + height + ".");
                foreach (string name in new[] { "ConnectButton", "RefreshButton", "UpButton", "SelectAllButton", "ClearSelectionButton", "ExportButton", "SaveZipButton", "ExportLibraryButton", "VersionHistoryButton", "MoreExportsButton" })
                {
                    FrameworkElement control = Named<FrameworkElement>(name);
                    Point point = control.TransformToVisual(root).TransformPoint(new Point(0, 0));
                    Write($"BOUNDS {width}x{height} {name} x={point.X:F1} y={point.Y:F1} width={control.ActualWidth:F1} height={control.ActualHeight:F1} root={root.ActualWidth:F1}x{root.ActualHeight:F1}");
                    Require(control.ActualWidth > 0 && control.ActualHeight > 0 && point.X >= -1 && point.Y >= -1 &&
                        point.X + control.ActualWidth <= root.ActualWidth + 1 && point.Y + control.ActualHeight <= root.ActualHeight + 1,
                        name + " remains within the native window at " + width + "x" + height + ".");
                }
            }

            private async Task ToggleAsync(ItemViewModel item)
            {
                Require(item.CanCheck, "The requested row is an exportable file.");
                var list = Named<ListView>("ItemsList");
                list.ScrollIntoView(item);
                await SettleAsync();
                DependencyObject? container = list.ContainerFromItem(item);
                CheckBox? checkbox = container == null ? null : FindChild<CheckBox>(container);
                Require(checkbox != null && checkbox.Visibility == Visibility.Visible && checkbox.IsEnabled,
                    "The visible file row has an enabled native checkbox.");
                var peer = new CheckBoxAutomationPeer(checkbox!);
                IToggleProvider? toggle = peer.GetPattern(PatternInterface.Toggle) as IToggleProvider;
                Require(toggle != null, "The native checkbox exposes its Toggle automation pattern.");
                toggle!.Toggle();
                await SettleAsync();
            }

            private void Invoke(string name)
            {
                Button button = Named<Button>(name);
                Require(button.IsEnabled, name + " is enabled for its native action.");
                IInvokeProvider? invoke = new ButtonAutomationPeer(button).GetPattern(PatternInterface.Invoke) as IInvokeProvider;
                Require(invoke != null, name + " exposes its Invoke automation pattern.");
                invoke!.Invoke();
            }

            private async Task CaptureAsync(string suffix = "")
            {
                try
                {
                    var bitmap = new RenderTargetBitmap();
                    await bitmap.RenderAsync(root);
                    if (bitmap.PixelWidth == 0 || bitmap.PixelHeight == 0) throw new NotSupportedException("The offscreen composition root returned no pixels.");
                    IBuffer buffer = await bitmap.GetPixelsAsync();
                    byte[] pixels = new byte[buffer.Length];
                    using (DataReader reader = DataReader.FromBuffer(buffer)) reader.ReadBytes(pixels);
                    using var stream = new InMemoryRandomAccessStream();
                    BitmapEncoder encoder = await BitmapEncoder.CreateAsync(BitmapEncoder.PngEncoderId, stream);
                    encoder.SetPixelData(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied,
                        (uint)bitmap.PixelWidth, (uint)bitmap.PixelHeight, 96, 96, pixels);
                    await encoder.FlushAsync();
                    byte[] png = new byte[checked((int)stream.Size)];
                    using (var reader = new DataReader(stream.GetInputStreamAt(0)))
                    {
                        await reader.LoadAsync((uint)png.Length);
                        reader.ReadBytes(png);
                    }
                    string path = Path.Combine(workspace, ".scratch", (integration ? "winui-runtime-check-integration" : "winui-runtime-check") + suffix + ".png");
                    File.WriteAllBytes(path, png);
                    Write("CAPTURE " + path + " " + bitmap.PixelWidth + "x" + bitmap.PixelHeight);
                }
                catch (Exception error) { Write("CAPTURE unavailable: " + error.GetBaseException().Message); }
            }

            private T Named<T>(string name) where T : FrameworkElement
            {
                if (root.FindName(name) is T value) return value;
                throw new InvalidOperationException("The live XAML control " + name + " is missing.");
            }
            private static T? FindChild<T>(DependencyObject value) where T : DependencyObject
            {
                if (value is T found) return found;
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(value); i++)
                {
                    T? child = FindChild<T>(VisualTreeHelper.GetChild(value, i));
                    if (child != null) return child;
                }
                return null;
            }
            private async Task SettleAsync() { root.UpdateLayout(); await Task.Delay(60); root.UpdateLayout(); }
            private static async Task UntilAsync(Func<bool> condition, string action)
            {
                var elapsed = Stopwatch.StartNew();
                while (!condition())
                {
                    if (elapsed.Elapsed > TimeSpan.FromSeconds(15)) throw new TimeoutException("Timed out waiting for " + action + ".");
                    await Task.Delay(20);
                }
            }
            private void Require(bool condition, string description)
            {
                if (!condition) throw new InvalidOperationException(description);
                assertions++;
                Write("OK " + description);
            }
            private void Write(string text)
            {
                log.AppendLine(text);
                File.WriteAllText(logPath, log.ToString(), new UTF8Encoding(false));
            }
            private static string FindWorkspace()
            {
                DirectoryInfo? directory = new(AppContext.BaseDirectory);
                while (directory != null)
                {
                    if (File.Exists(Path.Combine(directory.FullName, "SharePointExplorer.Recovery.csproj"))) return directory.FullName;
                    directory = directory.Parent;
                }
                return Path.GetFullPath(Environment.CurrentDirectory);
            }
        }

        private sealed class FixtureController : IExplorerController
        {
            private readonly Guid siteId = Guid.Parse("5e8a324e-77a7-4d49-9bba-4d304bdca034");
            private readonly Dictionary<Guid, List<Node>> children = new();
            private readonly Dictionary<Guid, byte[]> payloads = new();
            private readonly Node site, library, broken, slow;
            public readonly ManualResetEventSlim SlowStarted = new(false), ReleaseSlow = new(false), LibraryStarted = new(false), ReleaseLibrary = new(false), ZipStarted = new(false), ReleaseZip = new(false);
            public readonly ManualResetEventSlim VersionReadStarted = new(false), ReleaseVersionRead = new(false), VersionExportStarted = new(false), ReleaseVersionExport = new(false);
            public readonly List<Guid> ExportedIds = new();
            public readonly List<Guid> ZipSelectedIds = new();
            public readonly Dictionary<Guid, RecoveryStatus> OutcomeOverrides = new();
            // Legacy/runtime audit summaries may include candidates rejected during recovery.
            public readonly List<DesktopExportEntry> ExtraAuditOutcomes = new();
            public Node? CapturedLibrary;
            public Node? CapturedVersion;
            private readonly Node versionedDocument, historicalVersion;
            private readonly byte[] historicalPayload = Encoding.UTF8.GetBytes("second historical version 1.0\n");
            public bool HoldVersionRead, HoldVersionExport, FailVersionRead;
            public bool HoldLibrary, HoldZip;
            public bool Disposed { get; private set; }
            public string SourceName => "WinUI diagnostic fixture";
            public FixtureController()
            {
                site = Node(NodeKind.Site, "Test site", "Test site");
                library = Node(NodeKind.Library, "Documents", "Documents"); library.ListId = library.Id;
                Node empty = Node(NodeKind.Library, "Empty library", "Empty"); empty.ListId = empty.Id;
                Node list = Node(NodeKind.List, "Tasks", "Lists/Tasks"); list.ListId = list.Id;
                Node first = LibraryNode(NodeKind.Folder, "First folder", "Documents/First");
                Node second = LibraryNode(NodeKind.Folder, "Second folder", "Documents/Second");
                broken = LibraryNode(NodeKind.Folder, "Unavailable folder", "Documents/Unavailable");
                slow = LibraryNode(NodeKind.Folder, "Slow folder", "Documents/Slow");
                children[site.Id] = new() { library, empty, list };
                children[empty.Id] = new();
                children[library.Id] = new() { first, second, broken, slow, Document("root.txt", "Documents/root.txt", "library root payload\n") };
                Node nested = LibraryNode(NodeKind.Folder, "Nested folder", "Documents/First/Nested");
                children[first.Id] = new() { Document("first.txt", "Documents/First/first.txt", "first library file\n"), Document("other.txt", "Documents/First/other.txt", "another library file\n"), nested };
                children[nested.Id] = new() { Document("deep.txt", "Documents/First/Nested/deep.txt", "deep library payload\n") };
                versionedDocument = Document("second.txt", "Documents/Second/second.txt", "second library file\n");
                versionedDocument.UiVersion = 1024;versionedDocument.InternalVersion = 1025;
                historicalVersion = ExplorerNode.Copy(versionedDocument);historicalVersion.HistoryVersion = 512;historicalVersion.UiVersion = 512;
                historicalVersion.InternalVersion = 513;historicalVersion.Size = historicalPayload.Length;
                historicalVersion.Modified = new DateTime(2026, 10, 2, 0, 0, 0, DateTimeKind.Utc);
                children[second.Id] = new() { versionedDocument };
                children[broken.Id] = new();
                children[slow.Id] = new() { Document("stale.txt", "Documents/Slow/stale.txt", "unvisited library file\n") };
                Node task = Node(NodeKind.ListItem, "Task 1", "Lists/Tasks/1_.000"); task.ListId = list.Id;
                task.Title = "Task 1"; task.ListItemId = 1; task.HasAttachments = true;
                children[list.Id] = new() { task };
            }
            private Node Node(NodeKind kind, string name, string path) => new() {
                Kind = kind, SiteId = siteId, WebId = siteId, Id = Guid.NewGuid(), ListId = siteId,
                Name = name, Path = path, HasStream = kind == NodeKind.File,
                Modified = new DateTime(2026, 10, 3, 0, 0, 0, DateTimeKind.Utc), Level = 1
            };
            private Node LibraryNode(NodeKind kind, string name, string path)
            {
                Node value = Node(kind, name, path); value.ListId = library.Id; return value;
            }
            private Node Document(string name, string path, string text)
            {
                Node value = LibraryNode(NodeKind.File, name, path);
                byte[] bytes = Encoding.UTF8.GetBytes(text); value.Size = bytes.Length; payloads[value.Id] = bytes; return value;
            }
            public List<Node> GetRootSites() => new() { site };
            public List<Node> GetChildren(Node parent)
            {
                if (parent.Id == broken.Id) throw new IOException("The fixture location is unavailable.");
                if (parent.Id == slow.Id)
                {
                    SlowStarted.Set();
                    if (!ReleaseSlow.Wait(TimeSpan.FromSeconds(15))) throw new TimeoutException("The delayed fixture was not released.");
                }
                return new List<Node>(children[parent.Id]);
            }
            public DesktopExportSummary ExportFiles(IEnumerable<Node> documents, string directory, CancellationToken cancellationToken, Action<DesktopExportProgress> progress)
                => ExportDocuments(documents.ToArray(), directory, cancellationToken, progress, false);
            public List<Node> GetFileVersions(Node document)
            {
                VersionReadStarted.Set();
                if (HoldVersionRead && !ReleaseVersionRead.Wait(TimeSpan.FromSeconds(15))) throw new TimeoutException("The version history fixture was not released.");
                if (FailVersionRead) throw new IOException("The fixture version history is temporarily unavailable.");
                if (document.Id != versionedDocument.Id || document.SiteId != versionedDocument.SiteId || document.HistoryVersion != 0)
                    throw new InvalidOperationException("The fixture version query requires its current document.");
                return new() { ExplorerNode.Copy(versionedDocument), ExplorerNode.Copy(historicalVersion) };
            }
            public DesktopExportSummary ExportVersion(Node version, string directory, CancellationToken cancellationToken, Action<DesktopExportProgress> progress)
            {
                CapturedVersion = ExplorerNode.Copy(version);VersionExportStarted.Set();
                if (HoldVersionExport && !ReleaseVersionExport.Wait(TimeSpan.FromSeconds(15))) throw new TimeoutException("The version export fixture was not released.");
                Node known = version.HistoryVersion == 0 ? versionedDocument : historicalVersion;
                if (!SameVersion(version, known)) throw new InvalidOperationException("The fixture export requires an exact recorded version.");
                var summary = new DesktopExportSummary { Total = 1, Directory = directory, ReportPath = String.Empty };
                if (cancellationToken.IsCancellationRequested) { summary.Cancelled = true;return summary; }
                byte[] bytes = version.HistoryVersion == 0 ? payloads[versionedDocument.Id] : historicalPayload;
                string path = Path.Combine(directory, Path.GetFileNameWithoutExtension(version.Name) +
                    " (v" + VersionViewModel.FormatVersion(version.UiVersion) + ")" + Path.GetExtension(version.Name));
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);File.WriteAllBytes(path, bytes);
                string hash = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
                summary.Success = 1;summary.ReportPath = Path.Combine(directory, "version-export-report.csv");
                summary.Entries.Add(new DesktopExportEntry { Document = ExplorerNode.Copy(version), Status = RecoveryStatus.Success, Bytes = bytes.Length, Path = path, Sha256 = hash, Message = "" });
                File.WriteAllText(summary.ReportPath, "Document,Status,Bytes,Sha256,UiVersion,HistoryVersion,Level,InternalVersion\n" +
                    version.Path + ",Success," + bytes.Length + "," + hash + "," + version.UiVersion + "," + version.HistoryVersion + "," + version.Level + "," + version.InternalVersion + "\n", new UTF8Encoding(false));
                progress(new DesktopExportProgress { Document = ExplorerNode.Copy(version), Completed = 1, Total = 1, Status = RecoveryStatus.Success, Message = "", ExportPath = path });
                return summary;
            }

            public DesktopExportSummary ExportFilesAsZip(IEnumerable<Node> documents, string archivePath, CancellationToken cancellationToken, Action<DesktopExportProgress> progress)
            {
                Node[] selected = documents.ToArray();ZipSelectedIds.Clear();ZipSelectedIds.AddRange(selected.Select(document => document.Id));ZipStarted.Set();
                if (HoldZip && !ReleaseZip.Wait(TimeSpan.FromSeconds(15))) throw new TimeoutException("The ZIP fixture was not released.");
                string directory = Path.GetDirectoryName(Path.GetFullPath(archivePath)) ?? throw new ArgumentException("The ZIP fixture requires a directory.", nameof(archivePath));
                var summary = new DesktopExportSummary { Total = selected.Length, Directory = directory, ReportPath = String.Empty, ArchivePath = String.Empty, ArchiveReportEntry = String.Empty };
                if (cancellationToken.IsCancellationRequested) {summary.Cancelled = true;return summary;}
                Directory.CreateDirectory(directory);
                string finalPath = File.Exists(archivePath) ? Path.Combine(directory, Path.GetFileNameWithoutExtension(archivePath) + " (1)" + Path.GetExtension(archivePath)) : archivePath;
                using (var output = new FileStream(finalPath, FileMode.CreateNew, FileAccess.Write))
                using (var archive = new ZipArchive(output, ZipArchiveMode.Create))
                {
                    foreach (Node document in selected)
                    {
                        if (cancellationToken.IsCancellationRequested) {summary.Cancelled = true;break;}
                        byte[] bytes = payloads[document.Id];string path = document.SiteId.ToString("D") + "/" + document.Path;
                        using (Stream data = archive.CreateEntry(path).Open()) data.Write(bytes, 0, bytes.Length);
                        string hash = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();summary.Success++;
                        summary.Entries.Add(new DesktopExportEntry { Document = document, Status = RecoveryStatus.Success, Bytes = bytes.Length, Path = path, Sha256 = hash, Message = "" });
                        progress(new DesktopExportProgress { Document = document, Completed = summary.Success, Total = selected.Length, Status = RecoveryStatus.Success, Message = "", ExportPath = path });
                    }
                    using (var report = new StreamWriter(archive.CreateEntry("export-report.csv").Open(), new UTF8Encoding(false)))
                    {
                        report.WriteLine("Document,Path,Status,Bytes,Sha256");
                        foreach (DesktopExportEntry entry in summary.Entries) report.WriteLine(entry.Document.Path + "," + entry.Path + ",Success," + entry.Bytes + "," + entry.Sha256);
                    }
                    using (var text = new StreamWriter(archive.CreateEntry("summary.txt").Open(), new UTF8Encoding(false)))
                        text.WriteLine("Cancelled=" + summary.Cancelled + "\nTotal=" + summary.Total + "\nProcessed=" + summary.Entries.Count);
                }
                if (summary.Success > 0) {summary.ArchivePath = finalPath;summary.ArchiveReportEntry = "export-report.csv";}
                else File.Delete(finalPath);
                return summary;
            }
            public DesktopExportSummary ExportLibrary(Node selectedLibrary, string directory, CancellationToken cancellationToken, Action<DesktopExportProgress> progress)
            {
                CapturedLibrary = selectedLibrary;
                LibraryStarted.Set();
                if (HoldLibrary && !ReleaseLibrary.Wait(TimeSpan.FromSeconds(15))) throw new TimeoutException("The library discovery fixture was not released.");
                Node[] all = children.Values.SelectMany(value => value).Where(document => document.Kind == NodeKind.File && document.SiteId == selectedLibrary.SiteId && document.ListId == selectedLibrary.Id).OrderBy(document => document.Path, StringComparer.Ordinal).ToArray();
                return ExportDocuments(all, directory, cancellationToken, progress);
            }
            private DesktopExportSummary ExportDocuments(Node[] documents, string directory, CancellationToken cancellationToken, Action<DesktopExportProgress> progress, bool preserveHierarchy=true)
            {
                Directory.CreateDirectory(directory);
                var summary = new DesktopExportSummary { Total = documents.Length + ExtraAuditOutcomes.Count, Directory = directory, ReportPath = Path.Combine(directory, "export-report.csv") };
                foreach (Node file in documents)
                {
                    if (cancellationToken.IsCancellationRequested) { summary.Cancelled = true; break; }
                    RecoveryStatus outcome = OutcomeOverrides.TryGetValue(file.Id, out RecoveryStatus configured) ? configured : RecoveryStatus.Success;
                    string path = String.Empty, hash = String.Empty, message = String.Empty;long length = 0;
                    if (outcome == RecoveryStatus.Success)
                    {
                        byte[] bytes = payloads[file.Id];
                        path = preserveHierarchy ? Path.Combine(directory, file.SiteId.ToString("D"), file.Path.Replace('/', Path.DirectorySeparatorChar))
                            : Path.Combine(directory, file.Name);
                        Directory.CreateDirectory(Path.GetDirectoryName(path)!);File.WriteAllBytes(path, bytes);
                        hash = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();length = bytes.Length;ExportedIds.Add(file.Id);summary.Success++;
                    }
                    else
                    {
                        message = "Configured " + outcome + " recovery failure.";
                        if (outcome == RecoveryStatus.Unsupported || outcome == RecoveryStatus.Unavailable) summary.Skipped++;else summary.Failed++;
                    }
                    summary.Entries.Add(new DesktopExportEntry { Document = file, Status = outcome, Bytes = length, Path = path, Sha256 = hash, Message = message });
                    progress(new DesktopExportProgress { Document = file, Completed = summary.Entries.Count, Total = summary.Total, Status = outcome, Message = message, ExportPath = path });
                }
                foreach (DesktopExportEntry audit in ExtraAuditOutcomes)
                {
                    summary.Entries.Add(audit);
                    if (audit.Status == RecoveryStatus.Unsupported || audit.Status == RecoveryStatus.Unavailable) summary.Skipped++;else if (audit.Status != RecoveryStatus.Success) summary.Failed++;
                    progress(new DesktopExportProgress { Document = audit.Document, Completed = summary.Entries.Count, Total = summary.Total, Status = audit.Status, Message = audit.Message, ExportPath = audit.Path });
                }
                File.WriteAllText(summary.ReportPath, "Document,Status\n" + String.Join("\n", summary.Entries.Select(entry => entry.Document.Path + "," + entry.Status)), new UTF8Encoding(false));
                return summary;
            }
            public void Dispose() { Disposed = true; }
        }
    }
}