using System;
using System.ComponentModel;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using SharePointExplorer.Desktop;
using SharePointExplorer.Desktop.ViewModels;
using Microsoft.Windows.Storage.Pickers;
using Windows.System;

namespace SharePointExplorer.WinUI
{
    public sealed partial class MainWindow : Window
    {
        public ExplorerViewModel ViewModel { get; }
        private readonly IExplorerController initialController;
        private SqlConnectionOptions lastOptions=new SqlConnectionOptions();
        private bool initialized,dialogOpen,closingRequested,modelDisposed,closed,bindingsDetached,synchronizingTree;
        private Task disposeAndCloseTask;
        private string uiError="";
        private TreeItemViewModel revealedFolder;
        private bool treeRevealQueued;
        private readonly Dictionary<TreeItemViewModel,TreeBinding> treeBindings=new Dictionary<TreeItemViewModel,TreeBinding>();
        private readonly bool suppressStartupConnection;
        private readonly TaskCompletionSource<bool> startup=new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        public Task StartupTask => startup.Task;
        // Complete on the owning dispatcher so final-close continuations finish before its message loop stops.
        private readonly TaskCompletionSource<bool> closeCompletion=new TaskCompletionSource<bool>();
        private ConnectionDialog connectionDialog;
        private ExportProgressDialog exportDialog;
        private VersionHistoryDialog versionDialog;

        public MainWindow() : this(null) { }
        public MainWindow(IExplorerController controller) : this(controller,false) { }
        public MainWindow(IExplorerController controller,bool suppressStartupConnection)
        {
            InitializeComponent();
            initialController=controller; this.suppressStartupConnection=suppressStartupConnection;
            ViewModel=new ExplorerViewModel(DispatchAsync);
            ItemsList.ItemsSource=ViewModel.VisibleItems;
            ViewModel.RootNodes.CollectionChanged+=RootNodesChanged;
            ViewModel.PropertyChanged+=StateChanged; ViewModel.VisibleItems.CollectionChanged+=ItemsChanged;
            // Title bar and taskbar icon; the same icon is embedded in the executable.
            string icon=System.IO.Path.Combine(AppContext.BaseDirectory,"Assets","AppIcon.ico");
            if(System.IO.File.Exists(icon)) AppWindow.SetIcon(icon);
            if(AppWindow.Presenter is OverlappedPresenter presenter) { presenter.PreferredMinimumWidth=960; presenter.PreferredMinimumHeight=720; }
            AppWindow.Resize(new Windows.Graphics.SizeInt32(1600,1000));
            if(!suppressStartupConnection && AppWindow.Presenter is OverlappedPresenter startupPresenter) startupPresenter.Maximize();
            AppWindow.Closing+=WindowClosing;
            Closed+=WindowClosed;
            UpdateState();
        }
        private void ExplorerLayoutSizeChanged(object sender,SizeChangedEventArgs args)
        {
            // XAML layout uses effective pixels, which can differ substantially
            // from the physical AppWindow size on a high-DPI display.
            bool compact=ShellGrid.ActualWidth<1000;
            NavigationColumn.Width=new GridLength(compact ? 200 : 270);
            HeaderGrid.Padding=compact ? new Thickness(16,12,16,12) : new Thickness(28,22,28,16);
            HeaderTitleText.FontSize=compact ? 20 : 28;
            HeaderSubtitleText.Visibility=compact ? Visibility.Collapsed : Visibility.Visible;
            SourceBadge.MaxWidth=compact ? 200 : 310;
            SourceBadge.Padding=compact ? new Thickness(10,8,10,8) : new Thickness(16,10,16,10);
            NavigationHint.Visibility=compact ? Visibility.Collapsed : Visibility.Visible;
            HeaderActions.Padding=compact ? new Thickness(16,0,16,12) : new Thickness(28,0,28,16);
            ExplorerLayout.Margin=compact ? new Thickness(16,0,16,12) : new Thickness(20,0,20,16);
            ContentCard.Padding=compact ? new Thickness(16,12,16,12) : new Thickness(22,18,22,18);
            ContentGrid.RowSpacing=compact ? 8 : 12;
            StatusBorder.Padding=compact ? new Thickness(16,8,16,8) : new Thickness(28,12,28,12);
            Grid.SetRow(SelectedExportActions,compact ? 1 : 0); Grid.SetColumn(SelectedExportActions,compact ? 0 : 2);
            Grid.SetColumnSpan(SelectedExportActions,compact ? 4 : 2);
            Grid.SetRow(ExportLibraryButton,compact ? 0 : 1); Grid.SetColumn(ExportLibraryButton,compact ? 3 : 0);
            Grid.SetColumnSpan(ExportLibraryButton,compact ? 1 : 2);
            ExportLibraryButton.HorizontalAlignment=compact ? HorizontalAlignment.Right : HorizontalAlignment.Left;
            LibraryExportHint.Visibility=compact ? Visibility.Collapsed : Visibility.Visible;
            double width=Math.Max(0,args.NewSize.Width-NavigationColumn.Width.Value-ExplorerLayout.ColumnSpacing);
            ContentCard.MaxWidth=width;
            TableScrollViewer.MaxWidth=Math.Max(0,width-ContentCard.Padding.Left-ContentCard.Padding.Right);
        }        private Task DispatchAsync(Action action)
        {
            if(closed) return Task.FromException(new ObjectDisposedException(nameof(MainWindow)));
            if(DispatcherQueue.HasThreadAccess) { action(); return Task.CompletedTask; }
            TaskCompletionSource<bool> completion=new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            if(!DispatcherQueue.TryEnqueue(()=>{ try { if(!closed) action(); completion.SetResult(true); } catch(Exception error) { completion.SetException(error); } }))
                completion.SetException(new InvalidOperationException("The window dispatcher is unavailable."));
            return completion.Task;
        }
        private async void ShellLoaded(object sender,RoutedEventArgs args)
        {
            if(initialized) return; initialized=true;
            if(initialController!=null) await ViewModel.SetSourceAsync(initialController);
            else if(!suppressStartupConnection) await ShowConnection();
            UpdateState(); startup.TrySetResult(true);
        }
        private void StateChanged(object sender,PropertyChangedEventArgs args) { if(ViewModel.IsBusy) uiError=""; if(!closed) UpdateState(); }
        private void ItemsChanged(object sender,NotifyCollectionChangedEventArgs args) { UpdateState(); }
        private void UpdateState()
        {
            if(closed || modelDisposed || bindingsDetached || ViewModel==null) return;
            bool idle=!ViewModel.IsBusy && !ViewModel.IsExporting && !dialogOpen && !closingRequested;
            bool connected=!String.IsNullOrEmpty(ViewModel.SourceName);
            SourceText.Text=connected ? ViewModel.SourceName : "No database connected";
            LocationText.Text=ViewModel.CurrentFolder==null ? connected ? "Site collections" : "Your SharePoint content" : ViewModel.CurrentFolder.Name;
            BreadcrumbText.Text=ViewModel.CurrentFolder==null ? connected ? ViewModel.SourceName : "Connect a restored content database to get started." : ViewModel.CurrentFolder.Node.Path;
            ConnectButton.IsEnabled=idle; RefreshButton.IsEnabled=idle && connected;
            UpButton.IsEnabled=idle && ViewModel.CurrentFolder!=null;
            ExplorerTree.IsEnabled=ViewModel.CanNavigate && !dialogOpen && !closingRequested; ItemsList.IsEnabled=idle && connected;
            SelectAllButton.IsEnabled=idle && ViewModel.VisibleItems.Any(item=>item.CanCheck);
            ClearSelectionButton.IsEnabled=idle && ViewModel.SelectedCount>0;
            ExportButton.IsEnabled=idle && ViewModel.CanExport;
            SaveZipButton.IsEnabled=idle && ViewModel.CanExport;
            VersionHistoryButton.IsEnabled=idle && ItemsList.SelectedItem is ItemViewModel focused && focused.Node.Kind==NodeKind.File;
            ExportLibraryButton.IsEnabled=idle && ViewModel.CanExportLibrary;
            MoreExportsButton.IsEnabled=idle && ViewModel.IsConnected;
            SaveLibraryZipMenu.IsEnabled=idle && ViewModel.CanExportLibrary;
            SaveFolderZipMenu.IsEnabled=idle && ViewModel.CanExportScopeZip && ViewModel.CurrentFolder.Node.Kind==NodeKind.Folder;
            ExportVersionsMenu.IsEnabled=idle && ViewModel.CanExportVersions;
            ExportDeletedItemMenu.IsEnabled=idle && ItemsList.SelectedItem is ItemViewModel deleted && deleted.Node.IsDeleted && deleted.Node.Kind==NodeKind.ListItem;
            ExportItemAttachmentsMenu.IsEnabled=idle && ViewModel.CanExportListAttachments &&
                ItemsList.SelectedItem is ItemViewModel item && item.Node.Kind==NodeKind.ListItem && item.Node.HasAttachments!=false;
            ExportListAttachmentsMenu.IsEnabled=idle && ViewModel.CanExportListAttachments;
            ExportPackageMenu.IsEnabled=idle && ViewModel.CanExportPackage;
            ExportHistoryPackageMenu.IsEnabled=idle && ViewModel.CanExportPackage;
            TreeItemViewModel library=ViewModel.CurrentLibrary;
            LibraryExportHint.Text=library!=null ? "All exportable files and subfolders in " + library.Name + "." : "Open a library to export all its files.";
            string libraryHelp=library!=null ? "Export every current file in " + library.Name + ", including all nested folders." : "Open a document library or one of its folders.";
            ToolTipService.SetToolTip(ExportLibraryButton,libraryHelp);
            AutomationProperties.SetHelpText(ExportLibraryButton,libraryHelp);
            SelectionCountText.Text=ViewModel.SelectedCount==1 ? "1 file selected" : ViewModel.SelectedCount+" files selected";
            ExportButton.Content=ViewModel.SelectedCount>0 ? "Export selected ("+ViewModel.SelectedCount+")" : "Export selected";
            LoadingRing.IsActive=ViewModel.IsBusy; LoadingRing.Visibility=ViewModel.IsBusy ? Visibility.Visible : Visibility.Collapsed;
            string errorText=String.IsNullOrWhiteSpace(ViewModel.ErrorText) ? uiError : ViewModel.ErrorText;
            ErrorInfoBar.Message=errorText ?? ""; ErrorInfoBar.IsOpen=!String.IsNullOrWhiteSpace(errorText) && !dialogOpen;
            StatusText.Text=ViewModel.Status ?? "Ready";
            EmptyState.Visibility=ViewModel.VisibleItems.Count==0 ? Visibility.Visible : Visibility.Collapsed;
            EmptyStateText.Text=ViewModel.IsBusy ? "Loading content..." : !connected ? "Connect a database to browse its content." : "This location is empty.";
            SynchronizeTreeSelection();
            if(closingRequested) StatusText.Text="Closing the database connection after the current operation...";
        }
        private sealed class TreeBinding
        {
            public TreeItemViewModel Item;
            public TreeViewNode Node;
            public NotifyCollectionChangedEventHandler ChildrenChanged;
            public PropertyChangedEventHandler StateChanged;
        }
        private void RootNodesChanged(object sender,NotifyCollectionChangedEventArgs args)
        {
            if(!bindingsDetached) SynchronizeTreeNodes(ExplorerTree.RootNodes,ViewModel.RootNodes);
        }
        private TreeViewNode CreateTreeNode(TreeItemViewModel item)
        {
            TreeBinding binding=new TreeBinding { Item=item, Node=new TreeViewNode { Content=item, HasUnrealizedChildren=item.HasUnrealizedChildren } };
            treeBindings.Add(item,binding);
            binding.ChildrenChanged=(sender,args)=>{
                if(!bindingsDetached) SynchronizeTreeNodes(binding.Node.Children,item.Children);
            };
            binding.StateChanged=(sender,args)=>{
                if(!bindingsDetached && args.PropertyName==nameof(TreeItemViewModel.HasUnrealizedChildren)) binding.Node.HasUnrealizedChildren=item.HasUnrealizedChildren;
            };
            item.Children.CollectionChanged+=binding.ChildrenChanged;
            item.PropertyChanged+=binding.StateChanged;
            SynchronizeTreeNodes(binding.Node.Children,item.Children);
            return binding.Node;
        }
        private void SynchronizeTreeNodes(IList<TreeViewNode> nodes,IList<TreeItemViewModel> items)
        {
            bool previous=synchronizingTree; synchronizingTree=true;
            try
            {
                for(int index=nodes.Count-1; index>=0; index--)
                {
                    TreeViewNode node=nodes[index];
                    if(node.Content is TreeItemViewModel item && items.Any(value=>Object.ReferenceEquals(value,item))) continue;
                    ClearSelectionWithin(node); nodes.RemoveAt(index); DetachTreeNode(node);
                }
                for(int index=0; index<items.Count; index++)
                {
                    TreeItemViewModel item=items[index];
                    TreeViewNode node=treeBindings.TryGetValue(item,out TreeBinding binding) ? binding.Node : CreateTreeNode(item);
                    int existing=-1;
                    for(int position=0; position<nodes.Count; position++) if(Object.ReferenceEquals(nodes[position],node)) { existing=position; break; }
                    if(existing==index) continue;
                    if(existing>=0) { ClearSelectionWithin(node); nodes.RemoveAt(existing); }
                    nodes.Insert(index,node);
                }
            }
            finally { synchronizingTree=previous; }
        }
        private void ClearSelectionWithin(TreeViewNode node)
        {
            for(TreeViewNode selected=ExplorerTree.SelectedItem as TreeViewNode; selected!=null; selected=selected.Parent)
                if(Object.ReferenceEquals(selected,node)) { ExplorerTree.SelectedItem=null; break; }
        }
        private void DetachTreeNode(TreeViewNode node)
        {
            foreach(TreeViewNode child in node.Children) DetachTreeNode(child);
            if(node.Content is TreeItemViewModel item && treeBindings.TryGetValue(item,out TreeBinding binding))
            {
                item.Children.CollectionChanged-=binding.ChildrenChanged;
                item.PropertyChanged-=binding.StateChanged;
                treeBindings.Remove(item);
            }
        }
        private void SynchronizeTreeSelection()
        {
            TreeItemViewModel folder=ViewModel.CurrentFolder;
            TreeViewNode target=folder!=null && treeBindings.TryGetValue(folder,out TreeBinding binding) ? binding.Node : null;
            bool previous=synchronizingTree; synchronizingTree=true;
            try
            {
                if(target!=null)
                {
                    var ancestors=new Stack<TreeViewNode>();
                    for(TreeViewNode parent=target.Parent; parent!=null; parent=parent.Parent) ancestors.Push(parent);
                    while(ancestors.Count>0) ancestors.Pop().IsExpanded=true;
                }
                if(!Object.ReferenceEquals(ExplorerTree.SelectedItem,target)) ExplorerTree.SelectedItem=target;
            }
            finally { synchronizingTree=previous; }
            if(folder==null) revealedFolder=null;
            else if(target!=null && !Object.ReferenceEquals(revealedFolder,folder)) QueueTreeReveal(folder);
        }
        private void QueueTreeReveal(TreeItemViewModel folder)
        {
            if(treeRevealQueued || closed || closingRequested || bindingsDetached) return;
            treeRevealQueued=true;
            if(!DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low,()=>{
                treeRevealQueued=false;
                if(closed || closingRequested || bindingsDetached) return;
                if(!Object.ReferenceEquals(ViewModel.CurrentFolder,folder)) { SynchronizeTreeSelection(); return; }
                if(!treeBindings.TryGetValue(folder,out TreeBinding binding)) return;
                ExplorerTree.UpdateLayout();
                if(FindVisualChild<ListViewBase>(ExplorerTree) is ListViewBase list)
                {
                    // Explicit native nodes are the actual list items, so normal
                    // virtualized scrolling has a stable identity at every depth.
                    list.ScrollIntoView(binding.Node,ScrollIntoViewAlignment.Default);
                    list.UpdateLayout();
                }
                if(ExplorerTree.ContainerFromNode(binding.Node) is FrameworkElement container)
                {
                    AutomationProperties.SetName(container,folder.Name);
                    container.StartBringIntoView(new BringIntoViewOptions { AnimationDesired=false });
                    revealedFolder=folder;
                }
            })) treeRevealQueued=false;
        }
        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for(int index=0; index<VisualTreeHelper.GetChildrenCount(parent); index++)
            {
                DependencyObject child=VisualTreeHelper.GetChild(parent,index);
                if(child is T match) return match;
                T nested=FindVisualChild<T>(child);
                if(nested!=null) return nested;
            }
            return null;
        }
        private async Task ShowConnection()
        {
            if(dialogOpen || ViewModel.IsBusy || ViewModel.IsExporting || closed) return;
            dialogOpen=true; uiError=""; UpdateState();
            try
            {
                ConnectionDialog dialog=new ConnectionDialog(lastOptions,async options=>{
                    bool success=await ViewModel.ConnectAsync(options);
                    return success ? "" : String.IsNullOrWhiteSpace(ViewModel.ErrorText) ? "Could not open this database." : ViewModel.ErrorText;
                }) { XamlRoot=ShellGrid.XamlRoot };
                connectionDialog=dialog;
                ContentDialogResult result=await dialog.ShowAsync(); connectionDialog=null;
                if(result==ContentDialogResult.Primary && dialog.Options!=null) { lastOptions=dialog.Options.Clone(); lastOptions.Password=""; }
            }
            catch(Exception error) { ShowUiError(error); }
            finally { connectionDialog=null; dialogOpen=false; UpdateState(); if(closingRequested) await DisposeAndCloseAsync(); }
        }
        private async void ConnectClicked(object sender,RoutedEventArgs args) { await ShowConnection(); }
        private async void RefreshClicked(object sender,RoutedEventArgs args) { if(ViewModel.CanNavigate) await ViewModel.RefreshAsync(); }
        private async void UpClicked(object sender,RoutedEventArgs args) { if(ViewModel.CanNavigate) await ViewModel.UpAsync(); }
        private async void RefreshAccelerator(KeyboardAccelerator sender,KeyboardAcceleratorInvokedEventArgs args) { args.Handled=true; if(RefreshButton.IsEnabled) await ViewModel.RefreshAsync(); }
        private async void UpAccelerator(KeyboardAccelerator sender,KeyboardAcceleratorInvokedEventArgs args) { args.Handled=true; if(UpButton.IsEnabled) await ViewModel.UpAsync(); }
        private void SelectAllClicked(object sender,RoutedEventArgs args) { ViewModel.SelectAllFiles(); }
        private void ClearSelectionClicked(object sender,RoutedEventArgs args) { ViewModel.ClearSelection(); }
        private async void TreeExpanding(TreeView sender,TreeViewExpandingEventArgs args)
        {
            TreeItemViewModel item=args.Item as TreeItemViewModel ?? args.Node?.Content as TreeItemViewModel;
            if(!closingRequested && !synchronizingTree && item!=null && ViewModel.CanNavigate) await ViewModel.ExpandAsync(item);
        }
        private async void TreeItemInvoked(TreeView sender,TreeViewItemInvokedEventArgs args)
        {
            TreeItemViewModel item=args.InvokedItem as TreeItemViewModel ?? (args.InvokedItem as TreeViewNode)?.Content as TreeItemViewModel;
            if(!closingRequested && !synchronizingTree && item!=null && ViewModel.CanNavigate && !Object.ReferenceEquals(item,ViewModel.CurrentFolder)) await ViewModel.NavigateAsync(item);
        }
        private async void TreeSelectionChanged(TreeView sender,TreeViewSelectionChangedEventArgs args)
        {
            object selected=args.AddedItems.FirstOrDefault();
            TreeItemViewModel item=selected as TreeItemViewModel ?? (selected as TreeViewNode)?.Content as TreeItemViewModel;
            if(!closingRequested && !synchronizingTree && item!=null && ViewModel.CanNavigate && !Object.ReferenceEquals(item,ViewModel.CurrentFolder)) await ViewModel.NavigateAsync(item);
        }
        private void ItemsContainerChanging(ListViewBase sender,ContainerContentChangingEventArgs args)
        {
            if(args.Item is ItemViewModel item)
            {
                AutomationProperties.SetName(args.ItemContainer,item.Name);
                AutomationProperties.SetItemType(args.ItemContainer,item.KindText);
                AutomationProperties.SetHelpText(args.ItemContainer,item.Path);
            }
        }
        private async void ItemsDoubleTapped(object sender,DoubleTappedRoutedEventArgs args)
        {
            if(IsFromCheckbox(args.OriginalSource as DependencyObject)) return;
            ItemViewModel item=FindItem(args.OriginalSource as DependencyObject);
            if(item!=null && item.IsContainer && ViewModel.CanNavigate) { args.Handled=true; await ViewModel.NavigateItemAsync(item); }
        }
        private async void ItemsKeyDown(object sender,KeyRoutedEventArgs args)
        {
            if(IsFromCheckbox(args.OriginalSource as DependencyObject) || !ViewModel.CanNavigate) return;
            ItemViewModel item=ItemsList.SelectedItem as ItemViewModel;
            if(args.Key==VirtualKey.Enter && item!=null && item.IsContainer) { args.Handled=true; await ViewModel.NavigateItemAsync(item); }
            else if(args.Key==VirtualKey.Space && item!=null && item.CanCheck) { args.Handled=true; item.IsChecked=!item.IsChecked; }
        }
        private void FileCheckDoubleTapped(object sender,DoubleTappedRoutedEventArgs args) { args.Handled=true; }
        private void FileCheckChanged(object sender,RoutedEventArgs args) { UpdateState(); }
        private static bool IsFromCheckbox(DependencyObject element)
        {
            while(element!=null) { if(element is CheckBox) return true; if(element is ListViewItem) return false; element=VisualTreeHelper.GetParent(element); }
            return false;
        }
        private static ItemViewModel FindItem(DependencyObject element)
        {
            while(element!=null) { if(element is FrameworkElement view && view.DataContext is ItemViewModel item) return item; element=VisualTreeHelper.GetParent(element); }
            return null;
        }
        private void ItemsSelectionChanged(object sender,SelectionChangedEventArgs args) { if(ViewModel!=null && !closed && !bindingsDetached) UpdateState(); }
        private async void VersionHistoryClicked(object sender,RoutedEventArgs args)
        {
            ItemViewModel item=ItemsList.SelectedItem as ItemViewModel;
            if(item==null || item.Node.Kind!=NodeKind.File || !ViewModel.CanNavigate || ViewModel.IsBusy || dialogOpen || closingRequested || closed) return;
            Node document=ExplorerNode.Copy(item.Node);
            dialogOpen=true; uiError=""; UpdateState();
            try
            {
                versionDialog=new VersionHistoryDialog(ViewModel,document) { XamlRoot=ShellGrid.XamlRoot };
                await versionDialog.ShowAsync();
                Node selected=versionDialog.SelectedVersionNode; versionDialog=null;
                if(selected==null || closingRequested || closed) return;
                FolderPicker picker=new FolderPicker(AppWindow.Id) { SuggestedStartLocation=PickerLocationId.DocumentsLibrary, CommitButtonText="Export version here" };
                PickFolderResult folder=await picker.PickSingleFolderAsync();
                if(folder==null || closingRequested || closed) return;
                exportDialog=new ExportProgressDialog(ViewModel,folder.Path,version:selected) { XamlRoot=ShellGrid.XamlRoot };
                await exportDialog.ShowAsync(); exportDialog=null;
            }
            catch(Exception error) { ShowUiError(error); }
            finally { versionDialog=null; exportDialog=null; dialogOpen=false; UpdateState(); if(closingRequested) await DisposeAndCloseAsync(); }
        }
        private async void ExportItemAttachmentsClicked(object sender,RoutedEventArgs args)
        {
            if(!ExportItemAttachmentsMenu.IsEnabled || ItemsList.SelectedItem is not ItemViewModel item) return;
            await ShowExtendedExportAsync(item.Node,false);
        }
        private async void ExportListAttachmentsClicked(object sender,RoutedEventArgs args)
        {
            if(!ViewModel.CanExportListAttachments) return;
            await ShowExtendedExportAsync(ViewModel.CurrentContentList.Node,false);
        }
        private async void ExportPackageClicked(object sender,RoutedEventArgs args)
        {
            if(!ViewModel.CanExportPackage) return;
            await ShowExtendedExportAsync(ViewModel.CurrentContentList.Node,true);
        }
        private async Task ShowExtendedExportAsync(Node scope,bool asPackage,bool includeHistory=false)
        {
            if(!ViewModel.IsConnected || ViewModel.IsBusy || dialogOpen || closingRequested || closed) return;
            Node snapshot=ExplorerNode.Copy(scope);
            dialogOpen=true;uiError="";UpdateState();
            try
            {
                FolderPicker picker=new FolderPicker(AppWindow.Id) { SuggestedStartLocation=PickerLocationId.DocumentsLibrary,
                    CommitButtonText=asPackage ? "Save package here" : "Export attachments here" };
                PickFolderResult folder=await picker.PickSingleFolderAsync();
                if(folder==null || closingRequested || closed) return;
                exportDialog=new ExportProgressDialog(ViewModel,folder.Path,
                    attachmentScope:asPackage ? null : snapshot,packageScope:asPackage ? snapshot : null,includeHistory:includeHistory) {XamlRoot=ShellGrid.XamlRoot};
                await exportDialog.ShowAsync();exportDialog=null;
            }
            catch(Exception error) {ShowUiError(error);}
            finally {exportDialog=null;dialogOpen=false;UpdateState();if(closingRequested)await DisposeAndCloseAsync();}
        }
        private async void ExportHistoryPackageClicked(object sender,RoutedEventArgs args)
        {if(ViewModel.CanExportPackage)await ShowExtendedExportAsync(ViewModel.CurrentContentList.Node,true,true);}
        private async void SaveLibraryZipClicked(object sender,RoutedEventArgs args)
        {if(ViewModel.CanExportLibrary)await ShowScopeZipAsync(ViewModel.CurrentLibrary.Node);}
        private async void SaveFolderZipClicked(object sender,RoutedEventArgs args)
        {if(ViewModel.CanExportScopeZip)await ShowScopeZipAsync(ViewModel.CurrentFolder.Node);}
        private async Task ShowScopeZipAsync(Node scope)
        {
            if(dialogOpen || closingRequested || closed || ViewModel.IsBusy)return;
            Node snapshot=ExplorerNode.Copy(scope);dialogOpen=true;uiError="";UpdateState();
            try
            {
                FileSavePicker picker=new FileSavePicker(AppWindow.Id) {SuggestedStartLocation=PickerLocationId.DocumentsLibrary,SuggestedFileName=DocumentExporter.SafeFileName(snapshot.Name)+"-"+DateTime.Now.ToString("yyyyMMdd-HHmmss"),DefaultFileExtension=".zip",CommitButtonText="Save ZIP"};
                picker.FileTypeChoices.Add("ZIP archive",new List<string>{".zip"});PickFileResult file=await picker.PickSaveFileAsync();
                if(file==null || closingRequested || closed)return;
                exportDialog=new ExportProgressDialog(ViewModel,file.Path,asZip:true,zipScope:snapshot){XamlRoot=ShellGrid.XamlRoot};
                await exportDialog.ShowAsync();exportDialog=null;
            }
            catch(Exception error){ShowUiError(error);}
            finally{exportDialog=null;dialogOpen=false;UpdateState();if(closingRequested)await DisposeAndCloseAsync();}
        }
        private async void ExportVersionsClicked(object sender,RoutedEventArgs args)
        {if(ViewModel.CanExportVersions)await ShowAdvancedFolderExportAsync(null,true);}
        private async void ExportDeletedItemClicked(object sender,RoutedEventArgs args)
        {if(ExportDeletedItemMenu.IsEnabled && ItemsList.SelectedItem is ItemViewModel item)await ShowAdvancedFolderExportAsync(item.Node,false);}
        private async Task ShowAdvancedFolderExportAsync(Node deletedItem,bool versions)
        {
            if(dialogOpen || closingRequested || closed || ViewModel.IsBusy)return;
            Node snapshot=deletedItem!=null?ExplorerNode.Copy(deletedItem):null;dialogOpen=true;uiError="";UpdateState();
            try
            {
                FolderPicker picker=new FolderPicker(AppWindow.Id){SuggestedStartLocation=PickerLocationId.DocumentsLibrary,CommitButtonText=versions?"Export versions here":"Recover item here"};
                PickFolderResult folder=await picker.PickSingleFolderAsync();if(folder==null || closingRequested || closed)return;
                exportDialog=new ExportProgressDialog(ViewModel,folder.Path,allVersions:versions,deletedItem:snapshot){XamlRoot=ShellGrid.XamlRoot};
                await exportDialog.ShowAsync();exportDialog=null;
            }
            catch(Exception error){ShowUiError(error);}
            finally{exportDialog=null;dialogOpen=false;UpdateState();if(closingRequested)await DisposeAndCloseAsync();}
        }        private async void ExportClicked(object sender,RoutedEventArgs args) { await ShowExportAsync(false); }
        private async void ExportLibraryClicked(object sender,RoutedEventArgs args) { await ShowExportAsync(true); }
        private async void SaveZipClicked(object sender,RoutedEventArgs args) { await ShowZipExportAsync(); }
        private async Task ShowZipExportAsync()
        {
            if(!ViewModel.CanExport || dialogOpen || closingRequested || closed) return;
            dialogOpen=true; uiError=""; UpdateState();
            try
            {
                FileSavePicker picker=new FileSavePicker(AppWindow.Id) {
                    SuggestedStartLocation=PickerLocationId.DocumentsLibrary,
                    SuggestedFileName=DocumentExporter.SafeFileName(ViewModel.CurrentFolder?.Name ?? "SharePoint files")+"-"+DateTime.Now.ToString("yyyyMMdd-HHmmss"),
                    DefaultFileExtension=".zip", CommitButtonText="Save ZIP" };
                picker.FileTypeChoices.Add("ZIP archive",new List<string> { ".zip" });
                PickFileResult file=await picker.PickSaveFileAsync();
                if(file==null || closingRequested || closed) return;
                exportDialog=new ExportProgressDialog(ViewModel,file.Path,asZip:true) { XamlRoot=ShellGrid.XamlRoot };
                await exportDialog.ShowAsync(); exportDialog=null;
            }
            catch(Exception error) { ShowUiError(error); }
            finally { exportDialog=null; dialogOpen=false; UpdateState(); if(closingRequested) await DisposeAndCloseAsync(); }
        }
        private async Task ShowExportAsync(bool wholeLibrary)
        {
            if((wholeLibrary ? !ViewModel.CanExportLibrary : !ViewModel.CanExport) || dialogOpen || closingRequested || closed) return;
            Node library=wholeLibrary ? ExplorerNode.Copy(ViewModel.CurrentLibrary.Node) : null;
            dialogOpen=true; uiError=""; UpdateState();
            try
            {
                FolderPicker picker=new FolderPicker(AppWindow.Id) { SuggestedStartLocation=PickerLocationId.DocumentsLibrary, CommitButtonText="Export here" };
                PickFolderResult folder=await picker.PickSingleFolderAsync();
                if(folder==null || closingRequested || closed) return;
                exportDialog=new ExportProgressDialog(ViewModel,folder.Path,library) { XamlRoot=ShellGrid.XamlRoot };
                await exportDialog.ShowAsync(); exportDialog=null;
            }
            catch(Exception error) { ShowUiError(error); }
            finally { exportDialog=null; dialogOpen=false; UpdateState(); if(closingRequested) await DisposeAndCloseAsync(); }
        }
        private void ShowUiError(Exception error)
        {
            uiError=error.GetBaseException().Message; ErrorInfoBar.Message=uiError; ErrorInfoBar.IsOpen=true;
        }
        private void WindowClosing(AppWindow sender,AppWindowClosingEventArgs args)
        {
            if(modelDisposed) return;
            args.Cancel=true; closingRequested=true; UpdateState();
            if(exportDialog!=null) exportDialog.CancelAndCloseWhenFinished();
            else if(versionDialog!=null) versionDialog.RequestClose();
            else if(connectionDialog!=null) connectionDialog.RequestClose();
            else if(ViewModel.IsExporting) ViewModel.CancelExport();
            if(!dialogOpen) _=DisposeAndCloseAsync();
        }
        private Task DisposeAndCloseAsync()
        {
            if(disposeAndCloseTask!=null) return disposeAndCloseTask;
            disposeAndCloseTask=DisposeAndCloseCoreAsync();
            return disposeAndCloseTask;
        }
        public Task CloseSafelyAsync()
        {
            closingRequested=true; UpdateState();
            if(dialogOpen)
            {
                if(exportDialog!=null) exportDialog.CancelAndCloseWhenFinished();
                else if(versionDialog!=null) versionDialog.RequestClose();
                else connectionDialog?.RequestClose();
                return closeCompletion.Task;
            }
            return DisposeAndCloseAsync();
        }
        private async Task DisposeAndCloseCoreAsync()
        {
            // Stop native TreeView selection and x:Bind collection reactions
            // before the model clears its observable collections. The dispatcher
            // remains alive while outstanding reads/exports and sessions finish.
            bindingsDetached=true; synchronizingTree=true;
            ViewModel.PropertyChanged-=StateChanged;
            ViewModel.VisibleItems.CollectionChanged-=ItemsChanged;
            ItemsList.SelectionChanged-=ItemsSelectionChanged;
            ExplorerTree.SelectionChanged-=TreeSelectionChanged;
            ExplorerTree.ItemInvoked-=TreeItemInvoked;
            ExplorerTree.Expanding-=TreeExpanding;
            ExplorerTree.SelectedItem=null;
            ItemsList.SelectedItem=null;
            ViewModel.RootNodes.CollectionChanged-=RootNodesChanged;
            foreach(TreeViewNode root in ExplorerTree.RootNodes.ToArray()) DetachTreeNode(root);
            ExplorerTree.RootNodes.Clear();
            ItemsList.ItemsSource=null;
            Exception disposeError=null;
            try { await ViewModel.DisposeAsync(); }
            catch(Exception error) { disposeError=error; }
            modelDisposed=true;
            if(DispatcherQueue.HasThreadAccess) CompleteNativeClose();
            else if(!DispatcherQueue.TryEnqueue(CompleteNativeClose))
                closeCompletion.TrySetException(new InvalidOperationException("The window dispatcher is unavailable during shutdown."));
            await closeCompletion.Task;
            if(disposeError!=null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(disposeError).Throw();
        }
        private void CompleteNativeClose()
        {
            if(closed) { closeCompletion.TrySetResult(true); return; }
            try { Close(); }
            catch(Exception error) { closeCompletion.TrySetException(error); }
        }
        private void WindowClosed(object sender,WindowEventArgs args)
        {
            closed=true; closeCompletion.TrySetResult(true); startup.TrySetCanceled();
        }
    }
    public sealed class NodeGlyphConverter : IValueConverter
    {
        public object Convert(object value,Type targetType,object parameter,string language)
        {
            return (value?.ToString() ?? "") switch { "Deleted items"=>"\uE74D", "Site"=>"\uE774", "Library"=>"\uE8F1", "List"=>"\uE8FD", "Folder"=>"\uF12B", "ListItem"=>"\uE8A5", _=>"\uE8A5" };
        }
        public object ConvertBack(object value,Type targetType,object parameter,string language) { throw new NotSupportedException(); }
    }
    public sealed class BooleanVisibilityConverter : IValueConverter
    {
        public object Convert(object value,Type targetType,object parameter,string language) { return value is bool visible && visible ? Visibility.Visible : Visibility.Collapsed; }
        public object ConvertBack(object value,Type targetType,object parameter,string language) { throw new NotSupportedException(); }
    }
}



















