#nullable enable
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace SharePointExplorer.Desktop.ViewModels
{
    public sealed class ExplorerViewModel : ObservableViewModel, IAsyncDisposable
    {
        private readonly object stateGate = new(), lifetimeGate = new();
        private readonly Guid owner = Guid.NewGuid();
        private readonly Func<Action,Task> dispatch;
        private readonly Func<SqlConnectionOptions,IExplorerController> connect;
        private readonly HashSet<Task> pending = new();
        private readonly List<ControllerSource> sources = new();
        private ControllerSource? source;
        private CancellationTokenSource? exportCancellation;
        private volatile bool disposeRequested;
        private Task? disposeTask;
        private long sourceVersion, navigationVersion, versionRequest;
        private readonly Dictionary<string,Node> availableVersions = new();
        private bool connectionBusy, navigationBusy, isExporting, updatingChecks;
        private int expanding, versionReads, selectedCount, exportCompleted, exportTotal;
        private string sourceName = "", status = "Connect a restored database to get started.", errorText = "", exportCurrentFile = "";
        private TreeItemViewModel? currentFolder;
        private DesktopExportSummary? lastExport;

        public ExplorerViewModel(Func<Action,Task>? dispatchAsync = null,
            Func<SqlConnectionOptions,IExplorerController>? connect = null)
        {
            dispatch = dispatchAsync ?? CaptureDispatcher();
            this.connect = connect ?? ExplorerController.Connect;
        }
        public ObservableCollection<TreeItemViewModel> RootNodes { get; } = new();
        public ObservableCollection<ItemViewModel> VisibleItems { get; } = new();
        public TreeItemViewModel? CurrentFolder => currentFolder;
        public TreeItemViewModel? CurrentLibrary
        {
            get
            {
                for (TreeItemViewModel? item = currentFolder; item != null; item = item.Parent)
                    if (item.Node.Kind == NodeKind.Library && IsCurrent(item)) return item;
                return null;
            }
        }
        public TreeItemViewModel? CurrentContentList
        {
            get
            {
                for(TreeItemViewModel? item=currentFolder;item!=null;item=item.Parent)
                    if((item.Node.Kind==NodeKind.List || item.Node.Kind==NodeKind.Library) && IsCurrent(item)) return item;
                return null;
            }
        }
        public bool CanExportPackage => IsConnected && !IsBusy && CurrentContentList!=null && !CurrentContentList.Node.IsDeleted;
        public bool CanExportScopeZip => IsConnected && !IsBusy && CurrentFolder!=null && !CurrentFolder.Node.IsDeleted && (CurrentFolder.Node.Kind==NodeKind.Library || (CurrentFolder.Node.Kind==NodeKind.Folder && CurrentLibrary!=null));
        public bool CanExportVersions => CanExport && VisibleItems.Any(item=>item.IsChecked && item.Node.Kind==NodeKind.File);
        public bool CanExportListAttachments => IsConnected && !IsBusy && CurrentContentList?.Node.Kind==NodeKind.List && !CurrentContentList.Node.IsDeleted;
        public string SourceName => sourceName;
        public bool IsConnected => source != null && !disposeRequested;
        public bool IsBusy => connectionBusy || navigationBusy || expanding > 0 || versionReads > 0 || isExporting;
        public bool IsLoadingVersions => versionReads > 0;
        public bool IsExporting => isExporting;
        public bool CanNavigate => IsConnected && !connectionBusy && !isExporting;
        public bool CanExport => IsConnected && !IsBusy && selectedCount > 0;
        public bool CanExportLibrary => IsConnected && !IsBusy && CurrentLibrary != null && !CurrentLibrary.Node.IsDeleted;
        public int SelectedCount => selectedCount;
        public string SelectionText => selectedCount == 1 ? "1 file selected" : selectedCount + " files selected";
        public string Status => status;
        public string ErrorText => errorText;
        public string LocationTitle => currentFolder?.Name ?? "Site collections";
        public string Breadcrumb => currentFolder?.Path ?? sourceName;
        public string EmptyMessage => IsBusy ? "Loading content..." : !IsConnected ? "Connect a restored database to browse its content." : "No items in this location.";
        public int ExportCompleted => exportCompleted;
        public int ExportTotal => exportTotal;
        public bool ExportIsDiscovering => isExporting && lastExport == null && exportTotal == 0;
        public string ExportCurrentFile => exportCurrentFile;
        public DesktopExportSummary? LastExport => lastExport;

        public Task<bool> ConnectAsync(SqlConnectionOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            SqlConnectionOptions snapshot = options.Clone();
            return Track(() => OpenSourceAsync(() => Task.Run(() => connect(snapshot)), snapshot.Password ?? ""));
        }
        public Task<bool> SetSourceAsync(IExplorerController controller)
        {
            ArgumentNullException.ThrowIfNull(controller);
            return Track(() => OpenSourceAsync(() => Task.FromResult(controller), ""));
        }
        public Task NavigateAsync(TreeItemViewModel? folder)
        { return Track(() => NavigateCoreAsync(folder, false)); }
        public Task NavigateItemAsync(ItemViewModel item)
        {
            ArgumentNullException.ThrowIfNull(item);
            return Track(async () => {
                TreeItemViewModel? tree = null;
                await UiAsync(() => {if (VisibleItems.Contains(item) && item.IsContainer) tree = item.TreeNode;});
                if (tree != null) await NavigateCoreAsync(tree, false);
            });
        }
        public Task UpAsync()
        {
            return Track(async () => {
                TreeItemViewModel? parent = null; bool canGo = false;
                await UiAsync(() => {canGo = CanNavigate && currentFolder != null; parent = currentFolder?.Parent;});
                if (canGo) await NavigateCoreAsync(parent, false);
            });
        }
        public Task RefreshAsync()
        {
            return Track(async () => {
                TreeItemViewModel? folder = null; bool canGo = false;
                await UiAsync(() => {canGo = CanNavigate; folder = currentFolder;});
                if (!canGo) return;
                if (folder != null) await NavigateCoreAsync(folder, true);
                else await ReloadRootsAsync();
            });
        }
        public Task ExpandAsync(TreeItemViewModel folder)
        {
            ArgumentNullException.ThrowIfNull(folder);
            return Track(async () => {
                ControllerSource? active = null; long version = 0;
                await UiAsync(() => {
                    if (!CanNavigate || !IsCurrent(folder)) return;
                    active = source; version = sourceVersion; expanding++; errorText = ""; PublishState();
                });
                if (active == null) return;
                try {await LoadChildrenAsync(active, folder);}
                catch (Exception error) {await UiAsync(() => {if (IsCurrent(active, version)) {errorText = error.Message; status = "Could not load this location."; PublishState();}});}
                finally {await UiAsync(() => {if (IsCurrent(active, version)) {expanding--; PublishState();}});}
            });
        }

        public void SelectAllFiles()
        {
            lock (stateGate)
            {
                EnsureOpen();
                if (IsBusy || !IsConnected) return;
                updatingChecks = true;
                try {foreach (ItemViewModel item in VisibleItems) if (item.CanCheck) item.IsChecked = true;}
                finally {updatingChecks = false;}
                UpdateSelectionCount();
            }
        }
        public void ClearSelection()
        {
            lock (stateGate) {EnsureOpen(); ClearChecks();}
        }
        public void CancelExport()
        {
            lock (stateGate)
            {
                if (exportCancellation == null) return;
                exportCancellation.Cancel(); status = "Cancelling after the current file..."; Notify(nameof(Status));
            }
        }
        public Task<DesktopExportSummary?> ExportAsync(string directory, CancellationToken cancellationToken = default)
        {
            if (String.IsNullOrWhiteSpace(directory)) throw new ArgumentException("Choose an export folder.", nameof(directory));
            return Track(() => ExportCoreAsync(directory, cancellationToken, false));
        }
        public Task<DesktopExportSummary?> ExportLibraryAsync(string directory, CancellationToken cancellationToken = default)
        {
            if (String.IsNullOrWhiteSpace(directory)) throw new ArgumentException("Choose an export folder.", nameof(directory));
            return Track(() => ExportCoreAsync(directory, cancellationToken, true));
        }

        public Task<DesktopExportSummary?> ExportAttachmentsAsync(Node scope,string directory,CancellationToken cancellationToken=default)
        {
            ArgumentNullException.ThrowIfNull(scope);
            if(String.IsNullOrWhiteSpace(directory)) throw new ArgumentException("Choose an export folder.",nameof(directory));
            Node snapshot=ExplorerNode.Copy(scope);
            return Track(()=>ExportCoreAsync(directory,cancellationToken,false,extendedScope:snapshot));
        }
        public Task<DesktopExportSummary?> ExportPackageAsync(Node scope,string directory,CancellationToken cancellationToken=default,bool includeHistory=false)
        {
            ArgumentNullException.ThrowIfNull(scope);
            if(String.IsNullOrWhiteSpace(directory)) throw new ArgumentException("Choose an export folder.",nameof(directory));
            Node snapshot=ExplorerNode.Copy(scope);
            return Track(()=>ExportCoreAsync(directory,cancellationToken,false,extendedScope:snapshot,asPackage:true,includeHistory:includeHistory));
        }

        public Task<DesktopExportSummary?> ExportZipAsync(string archivePath, CancellationToken cancellationToken = default)
        {
            if (String.IsNullOrWhiteSpace(archivePath)) throw new ArgumentException("Choose a ZIP filename.", nameof(archivePath));
            return Track(() => ExportCoreAsync(archivePath, cancellationToken, false, true));
        }

        public Task<DesktopExportSummary?> ExportScopeZipAsync(Node scope,string archivePath,CancellationToken cancellationToken=default)
        {
            ArgumentNullException.ThrowIfNull(scope);Node snapshot=ExplorerNode.Copy(scope);
            return Track(()=>ExportCoreAsync(archivePath,cancellationToken,false,true,bulkScope:snapshot));
        }
        public Task<DesktopExportSummary?> ExportAllVersionsAsync(string directory,CancellationToken cancellationToken=default)
        {return Track(()=>ExportCoreAsync(directory,cancellationToken,false,batchVersions:true));}
        public Task<DesktopExportSummary?> ExportDeletedItemAsync(Node item,string directory,CancellationToken cancellationToken=default)
        {
            ArgumentNullException.ThrowIfNull(item);Node snapshot=ExplorerNode.Copy(item);
            return Track(()=>ExportCoreAsync(directory,cancellationToken,false,deletedItemScope:snapshot));
        }
        public Task<List<Node>> GetFileVersionsAsync(Node document)
        {
            ArgumentNullException.ThrowIfNull(document);
            if (document.Kind != NodeKind.File || document.HistoryVersion != 0 || document.SiteId == Guid.Empty || document.Id == Guid.Empty)
                throw new ArgumentException("Select a current document to view its versions.", nameof(document));
            Node snapshot = ExplorerNode.Copy(document);
            return Track(() => GetFileVersionsCoreAsync(snapshot));
        }
        public Task<DesktopExportSummary?> ExportVersionAsync(Node version, string directory, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(version);
            if (String.IsNullOrWhiteSpace(directory)) throw new ArgumentException("Choose an export folder.", nameof(directory));
            Node snapshot = ExplorerNode.Copy(version);
            return Track(() => ExportCoreAsync(directory, cancellationToken, false, false, snapshot));
        }

        private async Task<List<Node>> GetFileVersionsCoreAsync(Node document)
        {
            ControllerSource? active = null;long sourceGeneration = 0, navigation = 0, request = 0;
            await UiAsync(() => {
                if (!CanNavigate || connectionBusy || navigationBusy || expanding > 0) return;
                if (!VisibleItems.Any(item => item.Node.Kind == NodeKind.File && SameFile(item.Node, document)))
                {
                    errorText = "Select a file in the current folder to view its versions.";Notify(nameof(ErrorText));return;
                }
                active = source;sourceGeneration = sourceVersion;navigation = navigationVersion;request = ++versionRequest;
                availableVersions.Clear();versionReads++;errorText = "";status = "Loading versions of " + document.Name + "...";PublishState();
            });
            if (active == null) return new List<Node>();
            try
            {
                List<Node> versions = await active.CallAsync(controller => Snapshot(controller.GetFileVersions(document))).ConfigureAwait(false);
                List<Node> result = new();
                await UiAsync(() => {
                    if (!IsCurrent(active, sourceGeneration) || navigation != navigationVersion || request != versionRequest) return;
                    foreach (Node version in versions) availableVersions.Add(VersionIdentity(version), ExplorerNode.Copy(version));
                    result = versions.Select(ExplorerNode.Copy).ToList();
                    status = result.Count == 0 ? "No document versions are available." : result.Count + (result.Count == 1 ? " document version" : " document versions");PublishState();
                });
                return result;
            }
            catch (Exception error)
            {
                await UiAsync(() => {
                    if (!IsCurrent(active, sourceGeneration) || navigation != navigationVersion || request != versionRequest) return;
                    availableVersions.Clear();errorText = error.Message;status = "Could not load document versions.";PublishState();
                });
                return new List<Node>();
            }
            finally
            {
                await UiAsync(() => {if (IsCurrent(active, sourceGeneration)) {versionReads = Math.Max(0, versionReads - 1);PublishState();}});
            }
        }
        private async Task<bool> OpenSourceAsync(Func<Task<IExplorerController>> factory, string password)
        {
            long version = 0; ControllerSource? candidate = null; bool adopted = false;
            await UiAsync(() => {
                if (disposeRequested) return;
                version = ++sourceVersion; navigationVersion++;
                ResetContent(); source = null; sourceName = ""; connectionBusy = true; navigationBusy = false; expanding = 0;
                errorText = ""; status = "Connecting to database..."; lastExport = null; PublishState();
            });
            if (version == 0) return false;
            try
            {
                IExplorerController controller = await factory().ConfigureAwait(false);
                if (controller == null) throw new InvalidOperationException("The connection did not return a content source.");
                candidate = new ControllerSource(controller);
                await UiAsync(() => sources.Add(candidate));
                if (!IsVersionCurrent(version)) return false;
                var roots = await candidate.CallAsync(ReadRoots).ConfigureAwait(false);
                await UiAsync(() => {
                    if (!IsVersionCurrent(version)) return;
                    source = candidate; sourceName = roots.Name;
                    SetRoots(roots.Nodes); connectionBusy = false;
                    status = RootNodes.Count == 0 ? "Connected. No site collections are available." : "Connected. Choose a site collection.";
                    adopted = true; PublishState();
                });
                return adopted;
            }
            catch (Exception error)
            {
                await UiAsync(() => {
                    if (!IsVersionCurrent(version)) return;
                    connectionBusy = false; errorText = Redact(error.Message, password); status = "Could not connect to this database."; PublishState();
                });
                return false;
            }
            finally
            {
                if (candidate != null && !adopted) await RetireQuietlyAsync(candidate).ConfigureAwait(false);
            }
        }
        private async Task ReloadRootsAsync()
        {
            ControllerSource? active = null; long version = 0;
            await UiAsync(() => {
                if (!CanNavigate) return;
                active = source; version = ++sourceVersion; navigationVersion++;
                versionReads = 0;versionRequest++;availableVersions.Clear();
                ResetRows(); DetachRoots(); RootNodes.Clear(); currentFolder = null;
                navigationBusy = true; expanding = 0; errorText = ""; status = "Refreshing site collections..."; PublishState();
            });
            if (active == null) return;
            try
            {
                var roots = await active.CallAsync(ReadRoots).ConfigureAwait(false);
                await UiAsync(() => {
                    if (!IsCurrent(active, version)) return;
                    sourceName = roots.Name; SetRoots(roots.Nodes); status = RootNodes.Count == 0 ? "No site collections are available." : "Site collections refreshed.";
                });
            }
            catch (Exception error) {await UiAsync(() => {if (IsCurrent(active, version)) {errorText = error.Message; status = "Could not refresh site collections.";}});}
            finally {await UiAsync(() => {if (IsCurrent(active, version)) {navigationBusy = false; PublishState();}});}
        }
        private async Task NavigateCoreAsync(TreeItemViewModel? folder, bool refresh)
        {
            ControllerSource? active = null; long version = 0, navigation = 0;
            await UiAsync(() => {
                if (!CanNavigate || (folder != null && !IsCurrent(folder))) return;
                active = source; version = sourceVersion; navigation = ++navigationVersion;
                versionRequest++;availableVersions.Clear();
                // Row clearing publishes selection changes. Expose the target
                // first so native selection bindings cannot restore the old node.
                currentFolder = folder; navigationBusy = true; errorText = ""; status = "Loading " + (folder?.Name ?? "site collections") + "...";
                ResetRows();
                if (refresh && folder != null) {DetachChildren(folder); folder.Children.Clear(); folder.LoadedItems = null; folder.LoadTask = null; folder.IsLoaded = false;}
                PublishState();
            });
            if (active == null) return;
            try
            {
                if (folder != null) await LoadChildrenAsync(active, folder).ConfigureAwait(false);
                await UiAsync(() => {
                    if (!IsCurrent(active, version) || navigation != navigationVersion) return;
                    if (folder == null) SetRows(RootNodes.Select(tree => new ItemViewModel(tree.Node, tree)));
                    else if (folder.LoadedItems != null) SetRows(folder.LoadedItems.Select(node => new ItemViewModel(node, node.IsContainer ? FindChild(folder, node) : null)));
                    status = VisibleItems.Count == 0 ? "No items in this location." : VisibleItems.Count + (VisibleItems.Count == 1 ? " item" : " items");
                });
            }
            catch (Exception error)
            {
                await UiAsync(() => {if (IsCurrent(active, version) && navigation == navigationVersion) {errorText = error.Message; status = "Could not load this location.";}});
            }
            finally {await UiAsync(() => {if (IsCurrent(active, version) && navigation == navigationVersion) {navigationBusy = false; PublishState();}});}
        }
        private async Task LoadChildrenAsync(ControllerSource active, TreeItemViewModel folder)
        {
            Task<List<Node>>? load = null;
            await UiAsync(() => {
                if (!IsCurrent(folder) || !ReferenceEquals(source, active)) return;
                if (folder.LoadedItems != null) {load = Task.FromResult(folder.LoadedItems); return;}
                if (folder.LoadTask == null)
                {
                    Node parent = folder.Node;
                    folder.IsLoading = true; folder.ErrorText = "";
                    folder.LoadTask = active.CallAsync(controller => Snapshot(controller.GetChildren(parent)));
                }
                load = folder.LoadTask;
            });
            if (load == null) return;
            try
            {
                List<Node> children = await load.ConfigureAwait(false);
                await UiAsync(() => {
                    if (!IsCurrent(folder) || !ReferenceEquals(folder.LoadTask, load) || folder.IsLoaded) return;
                    folder.LoadedItems = children; folder.Children.Clear();
                    foreach (Node node in children.Where(node => node.IsContainer))
                        folder.Children.Add(new TreeItemViewModel(node, folder, owner, sourceVersion));
                    folder.IsLoaded = true; folder.IsLoading = false; folder.ErrorText = "";
                });
            }
            catch (Exception error)
            {
                await UiAsync(() => {
                    if (!IsCurrent(folder) || !ReferenceEquals(folder.LoadTask, load)) return;
                    folder.LoadTask = null; folder.IsLoading = false; folder.ErrorText = error.Message;
                });
                throw;
            }
        }
        private async Task<DesktopExportSummary?> ExportCoreAsync(string directory, CancellationToken cancellationToken, bool wholeLibrary, bool asZip = false, Node? exactVersion = null, Node? extendedScope=null, bool asPackage=false,Node? bulkScope=null,bool batchVersions=false,bool includeHistory=false,Node? deletedItemScope=null)
        {
            ControllerSource? active = null; CancellationTokenSource? cancellation = null; Node[] selected = Array.Empty<Node>(); Node? library = null, historical = null; long version = 0;
            await UiAsync(() => {
                if(bulkScope!=null)
                {
                    if(!CanExportScopeZip || CurrentFolder==null || (ExplorerNode.Identity(CurrentFolder.Node)!=ExplorerNode.Identity(bulkScope) && (CurrentLibrary==null || ExplorerNode.Identity(CurrentLibrary.Node)!=ExplorerNode.Identity(bulkScope))))
                    {errorText="Open the library or folder again before saving its ZIP.";Notify(nameof(ErrorText));return;}
                    bulkScope=ExplorerNode.Copy(bulkScope);
                }
                else if(deletedItemScope!=null)
                {
                    if(!IsConnected || IsBusy || !deletedItemScope.IsDeleted || deletedItemScope.Kind!=NodeKind.ListItem ||
                        !VisibleItems.Any(item=>ExplorerNode.Identity(item.Node)==ExplorerNode.Identity(deletedItemScope)))
                    {errorText="Select the deleted list item again before exporting.";Notify(nameof(ErrorText));return;}
                    deletedItemScope=ExplorerNode.Copy(deletedItemScope);
                }
                else if(extendedScope!=null)
                {
                    Node? current=CurrentContentList?.Node;
                    bool listScope=current!=null && current.Kind==extendedScope.Kind && current.SiteId==extendedScope.SiteId &&
                        current.WebId==extendedScope.WebId && current.ListId==extendedScope.ListId && current.Id==extendedScope.Id;
                    bool itemScope=!asPackage && extendedScope.Kind==NodeKind.ListItem && CurrentContentList?.Node.Kind==NodeKind.List &&
                        VisibleItems.Any(row=>row.Node.Kind==NodeKind.ListItem && row.Node.SiteId==extendedScope.SiteId &&
                        row.Node.WebId==extendedScope.WebId && row.Node.ListId==extendedScope.ListId && row.Node.Id==extendedScope.Id &&
                        row.Node.ListItemId==extendedScope.ListItemId && row.Node.ItemUniqueId==extendedScope.ItemUniqueId);
                    if(!IsConnected || IsBusy || !(listScope || itemScope) || (!asPackage && listScope && current!.Kind!=NodeKind.List))
                    {
                        errorText="Open the list, library, or item again before exporting.";Notify(nameof(ErrorText));return;
                    }
                    extendedScope=ExplorerNode.Copy(extendedScope);
                }
                else if (exactVersion != null)
                {
                    if (!IsConnected || IsBusy || !availableVersions.TryGetValue(VersionIdentity(exactVersion), out Node? known) || known == null || !SameVersionMetadata(known, exactVersion))
                    {
                        errorText = "Open version history again before exporting this version.";Notify(nameof(ErrorText));return;
                    }
                    if (!new VersionViewModel(known).CanExport)
                    {
                        errorText = "Choose a version with stored, supported content.";Notify(nameof(ErrorText));return;
                    }
                    historical = ExplorerNode.Copy(known);selected = new[] { historical };
                }
                else
                {
                    if (wholeLibrary ? !CanExportLibrary : !CanExport)
                    {
                        errorText = wholeLibrary ? "Open a document library or one of its folders before exporting the library." : "Check at least one file in the current folder before exporting.";
                        Notify(nameof(ErrorText));return;
                    }
                    if (wholeLibrary) library = ExplorerNode.Copy(CurrentLibrary!.Node);
                    else selected = VisibleItems.Where(item => item.CanCheck && item.IsChecked).Select(item => item.Node).ToArray();
                }
                active = source; version = sourceVersion;
                cancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken); exportCancellation = cancellation;
                isExporting = true; exportCompleted = 0; exportTotal = selected.Length; exportCurrentFile = ""; lastExport = null;
                errorText = ""; status = bulkScope!=null ? "Finding every file in "+bulkScope.Name+"..." : batchVersions ? "Finding retained versions of selected documents..." : deletedItemScope!=null ? "Recovering deleted list item metadata..." : extendedScope!=null ? asPackage ? "Preparing XML package..." : "Finding list attachments..." : historical != null ? "Exporting version " + VersionViewModel.FormatVersion(historical.UiVersion) + "..." : wholeLibrary ? "Finding files in " + library!.Name + " and all its folders..." : asZip ? "Saving selected files as ZIP..." : "Exporting selected files..."; PublishState();
            });
            if (active == null || cancellation == null) return null;
            var updates = new List<Task>();
            try
            {
                Action<DesktopExportProgress> reportProgress = progress => {
                    Task update = UiAsync(() => {
                        if (!IsCurrent(active, version) || !ReferenceEquals(exportCancellation, cancellation)) return;
                        exportCompleted = progress.Completed; exportTotal = progress.Total; exportCurrentFile = progress.Document?.Name ?? "";
                        foreach (ItemViewModel item in VisibleItems.Where(item => extendedScope==null && SameFile(item.Node, progress.Document) && (historical == null || progress.Document != null && VersionIdentity(item.Node) == VersionIdentity(progress.Document))))
                            item.ExportStatus = ExportStatusText(progress.Status);
                        status = cancellation.IsCancellationRequested ? asPackage ? "Cancelling and discarding the incomplete package..." : "Cancelling after the current file..." : "Processed " + progress.Completed + " of " + progress.Total + " files.";
                        PublishState();
                    });
                    lock (updates) updates.Add(update);
                };
                DesktopExportSummary summary = await active.CallAsync(controller => bulkScope!=null
                    ? controller is IExplorerBulkExportController bulkZip ? bulkZip.ExportScopeAsZip(bulkScope,directory,cancellation.Token,reportProgress) : throw new NotSupportedException("This source cannot save entire folders or libraries as ZIP.")
                    : batchVersions
                    ? controller is IExplorerBulkExportController bulkVersions ? bulkVersions.ExportVersions(selected,directory,cancellation.Token,reportProgress) : throw new NotSupportedException("This source cannot export versions in a batch.")
                    : deletedItemScope!=null
                    ? controller is IExplorerDeletedExportController deletedItem ? deletedItem.ExportDeletedItem(deletedItemScope,directory,cancellation.Token,reportProgress) : throw new NotSupportedException("This source cannot recover deleted item metadata.")
                    : extendedScope!=null
                    ? controller is IExplorerExtendedExportController extended
                        ? asPackage ? includeHistory ? controller is IExplorerBulkExportController historyExport ? historyExport.ExportMigrationPackage(extendedScope,directory,true,cancellation.Token,reportProgress) : throw new NotSupportedException("This source cannot include retained history in packages.") : extended.ExportMigrationPackage(extendedScope,directory,cancellation.Token,reportProgress)
                            : extended.ExportAttachments(extendedScope,directory,cancellation.Token,reportProgress)
                        : throw new NotSupportedException("This source does not support attachment or XML package export.")
                    : historical != null
                    ? controller.ExportVersion(historical, directory, cancellation.Token, reportProgress)
                    : library != null
                    ? controller.ExportLibrary(library, directory, cancellation.Token, reportProgress)
                    : asZip ? controller.ExportFilesAsZip(selected, directory, cancellation.Token, reportProgress)
                    : selected.Any(file=>file.IsDeleted) ? controller is IExplorerDeletedExportController deletedExport ? deletedExport.ExportDeletedFiles(selected,directory,cancellation.Token,reportProgress) : throw new NotSupportedException("This source cannot recover deleted files.") : controller.ExportFiles(selected, directory, cancellation.Token, reportProgress)).ConfigureAwait(false);
                Task[] queued; lock (updates) queued = updates.ToArray();
                await Task.WhenAll(queued).ConfigureAwait(false);
                await UiAsync(() => {
                    if (!IsCurrent(active, version) || !ReferenceEquals(exportCancellation, cancellation)) return;
                    lastExport = summary;
                    exportCompleted = summary.Entries.Count(entry => entry.Status != RecoveryStatus.Unsupported && entry.Status != RecoveryStatus.Unavailable);
                    exportTotal = exportCompleted + Math.Max(0, summary.Total - summary.Entries.Count);
                    status = !String.IsNullOrEmpty(summary.PackagePath) ? "XML package exported. " + summary.PackageItemCount + " items, " + summary.PackageFileCount + " files, " + summary.PackageAttachmentCount + " attachments." : !summary.Cancelled && summary.Success == 0 && summary.Failed == 0 ? "No files were exported."
                        : (summary.Cancelled ? "Export cancelled. " : "Export complete. ") + summary.Success + " exported" +
                            (summary.Failed > 0 ? ", " + summary.Failed + " failed" : "") + ".";
                    PublishState();
                });
                return summary;
            }
            catch (Exception error)
            {
                await UiAsync(() => {
                    if (!IsCurrent(active, version) || !ReferenceEquals(exportCancellation, cancellation)) return;
                    errorText = error is OperationCanceledException ? "" : error.Message;
                    status = error is OperationCanceledException ? "Export cancelled." : "Export could not complete."; PublishState();
                });
                return null;
            }
            finally
            {
                await UiAsync(() => {
                    if (!ReferenceEquals(exportCancellation, cancellation)) return;
                    exportCancellation = null; isExporting = false; PublishState();
                });
                cancellation.Dispose();
            }
        }

        private void SetRoots(List<Node> nodes)
        {
            DetachRoots(); RootNodes.Clear();
            foreach (Node node in nodes.Where(node => node.IsContainer)) RootNodes.Add(new TreeItemViewModel(node, null, owner, sourceVersion));
            currentFolder = null; SetRows(RootNodes.Select(tree => new ItemViewModel(tree.Node, tree)));
        }
        private void SetRows(IEnumerable<ItemViewModel> items)
        {
            ResetRows();
            foreach (ItemViewModel item in items) {item.PropertyChanged += OnItemChanged; VisibleItems.Add(item);}
            PublishState();
        }
        private void ClearChecks()
        {
            updatingChecks = true;
            try {foreach (ItemViewModel item in VisibleItems) item.IsChecked = false;}
            finally {updatingChecks = false;}
            selectedCount = 0; Notify(nameof(SelectedCount)); Notify(nameof(SelectionText)); Notify(nameof(CanExport));
        }
        private void ResetRows()
        {
            ClearChecks();
            foreach (ItemViewModel item in VisibleItems) item.PropertyChanged -= OnItemChanged;
            VisibleItems.Clear();
        }
        private void ResetContent()
        {
            exportCancellation?.Cancel(); exportCancellation = null; isExporting = false;
            if (source != null) _ = RetireQuietlyAsync(source);
            // Selection notifications must clear native tree selection while its
            // old nodes still exist, before the bound collections are reset.
            currentFolder = null;
            ResetRows(); DetachRoots(); RootNodes.Clear();
            versionReads = 0;versionRequest++;availableVersions.Clear();
            exportCompleted = 0; exportTotal = 0; exportCurrentFile = "";
        }
        private void OnItemChanged(object? sender, PropertyChangedEventArgs args)
        {
            if (args.PropertyName != nameof(ItemViewModel.IsChecked)) return;
            lock (stateGate)
            {
                if (updatingChecks || sender is not ItemViewModel item || !VisibleItems.Contains(item)) return;
                UpdateSelectionCount();
            }
        }
        private void UpdateSelectionCount()
        {
            selectedCount = VisibleItems.Count(row => row.CanCheck && row.IsChecked);
            Notify(nameof(SelectedCount)); Notify(nameof(SelectionText)); Notify(nameof(CanExport));
        }
        private bool IsCurrent(TreeItemViewModel tree) => !disposeRequested && tree.IsAttached && tree.Owner == owner && tree.SourceVersion == sourceVersion;
        private bool IsVersionCurrent(long version) => !disposeRequested && version == Interlocked.Read(ref sourceVersion);
        private bool IsCurrent(ControllerSource active, long version) => IsVersionCurrent(version) && ReferenceEquals(source, active);
        private static string ExportStatusText(RecoveryStatus status) => status switch {
            RecoveryStatus.Success => "Exported",
            RecoveryStatus.Unsupported => "Not exported",
            RecoveryStatus.Unavailable => "Not exported",
            RecoveryStatus.Corrupt => "Invalid content",
            RecoveryStatus.SqlError => "Database error",
            _ => "Failed"
        };
        private static bool SameFile(Node left, Node? right) => right != null && left.SiteId == right.SiteId && left.Id == right.Id && left.DeletionTransactionId==right.DeletionTransactionId;
        private static string VersionIdentity(Node version)
            => version.SiteId.ToString("N") + ":" + version.Id.ToString("N") + ":" + version.WebId.ToString("N") + ":" + version.ListId.ToString("N") + ":" +
                version.HistoryVersion + ":" + version.UiVersion + ":" + version.Level + ":" + version.InternalVersion + ":" + version.DeletionTransactionId;
        private static bool SameVersionMetadata(Node first, Node second)
            => VersionIdentity(first) == VersionIdentity(second) && first.Kind == second.Kind && first.Name == second.Name && first.Path == second.Path &&
                first.Size == second.Size && first.StreamSchema == second.StreamSchema && first.HasStream == second.HasStream && first.Modified == second.Modified && first.ParentId == second.ParentId;
        private static TreeItemViewModel? FindChild(TreeItemViewModel parent, Node node)
        { string identity = ExplorerNode.Identity(node); return parent.Children.FirstOrDefault(child => ExplorerNode.Identity(child.Node) == identity); }
        private void DetachRoots() {foreach (TreeItemViewModel tree in RootNodes) Detach(tree);}
        private static void DetachChildren(TreeItemViewModel parent) {foreach (TreeItemViewModel tree in parent.Children) Detach(tree);}
        private static void Detach(TreeItemViewModel tree) {tree.IsAttached = false; foreach (TreeItemViewModel child in tree.Children) Detach(child);}
        private void PublishState()
        {
            foreach (string name in new[] {nameof(CurrentFolder),nameof(CurrentLibrary),nameof(CurrentContentList),nameof(CanExportPackage),nameof(CanExportScopeZip),nameof(CanExportVersions),nameof(CanExportListAttachments),nameof(SourceName),nameof(IsConnected),nameof(IsBusy),nameof(IsLoadingVersions),nameof(IsExporting),nameof(CanNavigate),nameof(CanExport),nameof(CanExportLibrary),
                nameof(SelectedCount),nameof(SelectionText),nameof(Status),nameof(ErrorText),nameof(LocationTitle),nameof(Breadcrumb),nameof(EmptyMessage),
                nameof(ExportCompleted),nameof(ExportTotal),nameof(ExportIsDiscovering),nameof(ExportCurrentFile),nameof(LastExport)}) Notify(name);
        }
        private Task UiAsync(Action action) => dispatch(() => {lock (stateGate) action();});
        private static string Redact(string message, string password) => String.IsNullOrEmpty(password) ? message : message.Replace(password, "[redacted]");
        private static List<Node> Snapshot(List<Node> nodes)
        {
            if (nodes == null || nodes.Any(node => node == null)) throw new InvalidDataException("The content source returned invalid navigation data.");
            return nodes.Select(ExplorerNode.Copy).ToList();
        }
        private static (string Name, List<Node> Nodes) ReadRoots(IExplorerController controller)
        { return (controller.SourceName, Snapshot(controller.GetRootSites())); }
        private static Func<Action,Task> CaptureDispatcher()
        {
            SynchronizationContext? context = SynchronizationContext.Current;
            return action => {
                if (context == null || ReferenceEquals(SynchronizationContext.Current, context)) {action(); return Task.CompletedTask;}
                var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
                context.Post(_ => {try {action(); completion.SetResult();} catch (Exception error) {completion.SetException(error);}}, null);
                return completion.Task;
            };
        }

        private Task Track(Func<Task> work) => Track(async () => {await work().ConfigureAwait(false); return true;});
        private Task<T> Track<T>(Func<Task<T>> work)
        {
            var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            lock (lifetimeGate) {EnsureOpen(); pending.Add(completion.Task);}
            _ = CompleteTrackedAsync(work, completion);
            return completion.Task;
        }
        private async Task CompleteTrackedAsync<T>(Func<Task<T>> work, TaskCompletionSource<T> completion)
        {
            try {completion.TrySetResult(await work().ConfigureAwait(false));}
            catch (Exception error) {completion.TrySetException(error);}
            finally {lock (lifetimeGate) pending.Remove(completion.Task);}
        }
        private void EnsureOpen() {if (disposeRequested) throw new ObjectDisposedException(nameof(ExplorerViewModel));}
        public ValueTask DisposeAsync()
        {
            TaskCompletionSource? completion = null; Task[] running = Array.Empty<Task>();
            lock (lifetimeGate)
            {
                if (disposeTask != null) return new ValueTask(disposeTask);
                disposeRequested = true; Interlocked.Increment(ref sourceVersion); Interlocked.Increment(ref navigationVersion);
                completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously); disposeTask = completion.Task; running = pending.ToArray();
            }
            _ = DisposeCoreAsync(running, completion);
            return new ValueTask(disposeTask);
        }
        private async Task DisposeCoreAsync(Task[] running, TaskCompletionSource completion)
        {
            try
            {
                await UiAsync(() => {ResetContent(); source = null; sourceName = ""; connectionBusy = false; navigationBusy = false; expanding = 0; status = "Closed."; PublishState();});
                try {await Task.WhenAll(running).ConfigureAwait(false);} catch { /* Operations already expose their own outcome. */ }
                ControllerSource[] owned = Array.Empty<ControllerSource>(); await UiAsync(() => owned = sources.ToArray());
                await Task.WhenAll(owned.Select(RetireQuietlyAsync)).ConfigureAwait(false);
                completion.TrySetResult();
            }
            catch (Exception error) {completion.TrySetException(error);}
        }
        private async Task RetireQuietlyAsync(ControllerSource retired)
        {
            try {await retired.RetireAsync().ConfigureAwait(false);}
            catch { /* Closing an old source cannot replace a new source's UI state. */ }
            finally
            {
                // Keep pending retirements visible to DisposeAsync, then release
                // the old controller/session and its connection credentials.
                // This private ownership list has no UI notifications.
                lock (stateGate) sources.Remove(retired);
            }
        }

        // A replaced source is retired only after its synchronous SQL/export calls
        // finish. Disposal therefore never closes a source beneath an active call.
        private sealed class ControllerSource
        {
            private readonly object gate = new();
            private readonly IExplorerController controller;
            private readonly TaskCompletionSource idle = new(TaskCreationOptions.RunContinuationsAsynchronously);
            private int calls;
            private bool retired;
            private Task? retirement;
            internal ControllerSource(IExplorerController controller) {this.controller = controller;}
            internal async Task<T> CallAsync<T>(Func<IExplorerController,T> work)
            {
                lock (gate) {if (retired) throw new ObjectDisposedException(nameof(IExplorerController)); calls++;}
                try {return await Task.Run(() => work(controller)).ConfigureAwait(false);}
                finally {lock (gate) {calls--; if (retired && calls == 0) idle.TrySetResult();}}
            }
            internal Task RetireAsync()
            {
                lock (gate)
                {
                    if (retirement != null) return retirement;
                    retired = true; if (calls == 0) idle.TrySetResult();
                    retirement = CloseWhenIdleAsync(); return retirement;
                }
            }
            private async Task CloseWhenIdleAsync()
            {await idle.Task.ConfigureAwait(false); await Task.Run(controller.Dispose).ConfigureAwait(false);}
        }
    }
}
