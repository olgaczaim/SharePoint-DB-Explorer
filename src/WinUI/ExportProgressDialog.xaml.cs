using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Automation;
using SharePointExplorer.Desktop;
using SharePointExplorer.Desktop.ViewModels;
using Windows.Storage;
using Windows.System;

namespace SharePointExplorer.WinUI
{
    public sealed class ExportResultRow
    {
        public string Name { get; set; }
        public string Status { get; set; }
        public string Detail { get; set; }
    }
    public sealed partial class ExportProgressDialog : ContentDialog
    {
        private readonly ExplorerViewModel viewModel;
        private readonly string directory;
        private readonly string destination;
        private readonly bool asZip;
        private readonly Node library;
        private readonly Node version,attachmentScope,packageScope,zipScope,deletedItem;
        private readonly bool allVersions,includeHistory;
        private readonly CancellationTokenSource cancellation=new CancellationTokenSource();
        private readonly ObservableCollection<ExportResultRow> rows=new ObservableCollection<ExportResultRow>();
        private bool started,finished,closed,closeAfterFinish;
        public DesktopExportSummary Summary { get; private set; }
        internal string SavedFilePath { get; private set; }=String.Empty;

        public ExportProgressDialog(ExplorerViewModel viewModel,string destination,Node library=null,bool asZip=false,Node version=null,Node attachmentScope=null,Node packageScope=null,Node zipScope=null,bool allVersions=false,bool includeHistory=false,Node deletedItem=null)
        {
            this.viewModel=viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            this.destination=destination ?? throw new ArgumentNullException(nameof(destination));
            this.asZip=asZip;this.zipScope=zipScope!=null ? ExplorerNode.Copy(zipScope) : null;this.allVersions=allVersions;this.includeHistory=includeHistory;this.deletedItem=deletedItem!=null?ExplorerNode.Copy(deletedItem):null;
            this.directory=asZip ? Path.GetDirectoryName(Path.GetFullPath(destination)) : destination;
            if(version!=null && (library!=null || asZip)) throw new ArgumentException("A version export requires its own folder destination.",nameof(version));
            int scopes=(library!=null ? 1 : 0)+(asZip ? 1 : 0)+(version!=null ? 1 : 0)+(attachmentScope!=null ? 1 : 0)+(packageScope!=null ? 1 : 0)+(allVersions?1:0)+(deletedItem!=null?1:0);
            if(scopes>1) throw new ArgumentException("Choose a single export operation.");
            this.attachmentScope=attachmentScope!=null ? ExplorerNode.Copy(attachmentScope) : null;
            this.packageScope=packageScope!=null ? ExplorerNode.Copy(packageScope) : null;
            this.library=library!=null ? ExplorerNode.Copy(library) : null;
            this.version=version!=null ? ExplorerNode.Copy(version) : null;
            InitializeComponent(); DestinationText.Text=(asZip ? "ZIP: " : "Destination: ")+destination;
            if(asZip)
            {
                Title=zipScope!=null ? "Save "+(zipScope.Kind==NodeKind.Library?"library":"folder")+" as ZIP" : "Save selected files as ZIP";
                if(zipScope!=null)ExportScopeText.Text=zipScope.Name+". Every exportable file, including nested folders.";
                CancellationNote.Text="Cancel stops after the current file. Completed files are kept in a valid ZIP marked as cancelled.";
            }
            if(this.library!=null)
            {
                Title="Export library";
                ExportScopeText.Text="Library: "+this.library.Name+". All exportable files, including every nested folder.";
                ProgressText.Text="Finding files in the library...";
                ExportProgress.IsIndeterminate=true;
            }
            if(this.version!=null)
            {
                Title="Export document version";
                ExportScopeText.Text=this.version.Path+"; version "+new VersionViewModel(this.version).VersionText;
            }
            if(this.attachmentScope!=null)
            {
                Title="Export list attachments";
                ExportScopeText.Text=this.attachmentScope.Kind==NodeKind.ListItem
                    ? "Attachments for item "+this.attachmentScope.ListItemId+": "+this.attachmentScope.Name
                    : "List: "+this.attachmentScope.Name+". Attachments from all current items, including nested folders. Files are grouped by item ID.";
                ProgressText.Text="Finding attachments...";ExportProgress.IsIndeterminate=true;
            }
            if(this.packageScope!=null)
            {
                Title="Export XML package";
                ExportScopeText.Text=this.packageScope.Name+(includeHistory ? ". Current items and all retained document/item versions, with current attachments." : ". Current items, folders, documents, and attachments.");
                CancellationNote.Text="Cancel discards the incomplete package. The package is saved after all content and XML validation completes.";
                ProgressText.Text="Reading list metadata...";ExportProgress.IsIndeterminate=true;
            }
            if(allVersions){Title="Export document versions";ExportScopeText.Text="All retained versions of the checked documents, including current versions. Files keep their version labels and source folders.";ExportProgress.IsIndeterminate=true;}
            if(deletedItem!=null){Title="Recover deleted list item";ExportScopeText.Text=deletedItem.Name+". Source fields and retained metadata are saved as XML.";ExportProgress.IsIndeterminate=true;}
            ResultsList.ItemsSource=rows;
            viewModel.PropertyChanged+=ProgressChanged;
        }
        private async void DialogOpened(ContentDialog sender,ContentDialogOpenedEventArgs args)
        {
            if(started) return; started=true;
            try
            {
                Summary=zipScope!=null ? await viewModel.ExportScopeZipAsync(zipScope,destination,cancellation.Token)
                    : allVersions ? await viewModel.ExportAllVersionsAsync(directory,cancellation.Token)
                    : deletedItem!=null ? await viewModel.ExportDeletedItemAsync(deletedItem,directory,cancellation.Token)
                    : packageScope!=null ? await viewModel.ExportPackageAsync(packageScope,directory,cancellation.Token,includeHistory)
                    : attachmentScope!=null ? await viewModel.ExportAttachmentsAsync(attachmentScope,directory,cancellation.Token)
                    : version!=null ? await viewModel.ExportVersionAsync(version,directory,cancellation.Token)
                    : library!=null ? await viewModel.ExportLibraryAsync(directory,cancellation.Token)
                    : asZip ? await viewModel.ExportZipAsync(destination,cancellation.Token)
                    : await viewModel.ExportAsync(directory,cancellation.Token);
                if(Summary==null)
                {
                    SummaryInfo.Title="Export could not finish"; SummaryInfo.Message=viewModel.ErrorText;
                    SummaryInfo.Severity=InfoBarSeverity.Error; SummaryInfo.IsOpen=true;
                }
                else RenderSummary();
            }
            catch(Exception error)
            {
                SummaryInfo.Title="Export could not finish"; SummaryInfo.Message=error.GetBaseException().Message;
                SummaryInfo.Severity=InfoBarSeverity.Error; SummaryInfo.IsOpen=true;
            }
            finally
            {
                finished=true; ExportProgress.IsIndeterminate=false; CloseButtonText="Close"; CancellationNote.Visibility=Visibility.Collapsed;
                if(closeAfterFinish && !closed) DispatcherQueue.TryEnqueue(()=>{ if(!closed) Hide(); });
            }
        }
        private void ProgressChanged(object sender,PropertyChangedEventArgs args)
        {
            if(closed || finished) return;
            ExportProgress.IsIndeterminate=viewModel.ExportIsDiscovering;
            ExportProgress.Maximum=Math.Max(1,viewModel.ExportTotal);
            ExportProgress.Value=Math.Min(ExportProgress.Maximum,Math.Max(0,viewModel.ExportCompleted));
            ProgressText.Text=viewModel.ExportIsDiscovering ? packageScope!=null ? "Preparing and validating the package..." : attachmentScope!=null ? "Finding attachments..." : "Finding files in the library..."
                : viewModel.ExportCompleted+" of "+viewModel.ExportTotal+" files processed";
            CurrentFileText.Text=viewModel.ExportCurrentFile;
        }
        private void Cancel()
        {
            if(finished || closed) return;
            cancellation.Cancel(); viewModel.CancelExport();
            ProgressText.Text=packageScope!=null ? "Cancelling and discarding the incomplete package..." : "Cancelling after the current file..."; CloseButtonText="Cancelling...";
        }
        public void CancelAndCloseWhenFinished()
        {
            if(closed) return; closeAfterFinish=true; if(finished) Hide(); else Cancel();
        }
        private void CloseClicked(ContentDialog sender,ContentDialogButtonClickEventArgs args)
        {
            if(!finished) { args.Cancel=true; Cancel(); }
        }
        private void DialogClosing(ContentDialog sender,ContentDialogClosingEventArgs args)
        {
            if(!finished) { args.Cancel=true; Cancel(); }
        }
        private void DialogClosed(ContentDialog sender,ContentDialogClosedEventArgs args)
        {
            closed=true; viewModel.PropertyChanged-=ProgressChanged; cancellation.Dispose();
        }
        private void RenderSummary()
        {
            bool hasArchive=asZip && !String.IsNullOrEmpty(Summary.ArchivePath);
            bool hasPackage=!String.IsNullOrEmpty(Summary.PackagePath);
            DesktopExportEntry savedFile=!asZip && !hasPackage && Summary.Success==1 ? Summary.Entries.Find(entry=>entry.Status==RecoveryStatus.Success && !String.IsNullOrEmpty(entry.Path)) : null;
            SavedFilePath=savedFile?.Path ?? String.Empty;
            bool hasFile=!String.IsNullOrEmpty(SavedFilePath);
            bool nothingExported=!Summary.Cancelled && Summary.Success==0 && Summary.Failed==0;
            SummaryInfo.Title=Summary.Cancelled ? hasArchive ? "Partial ZIP saved" : "Export cancelled"
                : nothingExported ? "Nothing to export" : hasPackage ? "XML package saved" : asZip ? "ZIP saved" : hasFile && version!=null ? "Version saved" : "Export complete";
            SummaryInfo.Message=nothingExported ? "No files were exported."
                : Summary.Success+" exported"+(Summary.Failed>0 ? ", "+Summary.Failed+" failed" : "");
            if(Summary.Cancelled && asZip)
                SummaryInfo.Message+=hasArchive ? ". Contains completed files; the remaining selection was cancelled." : ". No ZIP was saved.";
            if(hasPackage)
            {
                DestinationText.Text="Package: "+Summary.PackagePath;
                SummaryInfo.Message=Summary.PackageItemCount+" items, "+Summary.PackageFileCount+" files, "+Summary.PackageAttachmentCount+" attachments";
            }
            if(hasArchive) DestinationText.Text="ZIP: "+Summary.ArchivePath;
            if(hasFile) DestinationText.Text="Saved file: "+SavedFilePath;
            SummaryInfo.Severity=Summary.Failed>0 || Summary.Cancelled ? InfoBarSeverity.Warning
                : nothingExported ? InfoBarSeverity.Informational : InfoBarSeverity.Success;
            SummaryInfo.IsOpen=true; ProgressText.Text=SummaryInfo.Message; CurrentFileText.Text="";
            ExportProgress.IsIndeterminate=false; if(!Summary.Cancelled) ExportProgress.Value=ExportProgress.Maximum;
            rows.Clear(); foreach(DesktopExportEntry entry in Summary.Entries)
            {
                if(entry.Status==RecoveryStatus.Unsupported || entry.Status==RecoveryStatus.Unavailable) continue;
                rows.Add(new ExportResultRow { Name=entry.Document?.Path ?? entry.Document?.Name ?? "Document", Status=Friendly(entry.Status),
                    Detail=entry.Status==RecoveryStatus.Success ? entry.Path : entry.Message });
            }
            ResultsList.Visibility=rows.Count>0 ? Visibility.Visible : Visibility.Collapsed;
            ResultActions.Visibility=Visibility.Visible; OpenReportButton.IsEnabled=!String.IsNullOrEmpty(Summary.ReportPath);
            OpenArchiveButton.Visibility=hasArchive ? Visibility.Visible : Visibility.Collapsed;
            OpenFileButton.Visibility=hasFile ? Visibility.Visible : Visibility.Collapsed;
            OpenFileButton.IsEnabled=hasFile;
            OpenDestinationButton.Content=hasFile ? "Show in folder" : hasPackage ? "Open package folder" : "Open destination";
            AutomationProperties.SetName(OpenDestinationButton,hasFile ? "Show exported document in its folder" : "Open export destination");
            ArchiveReportNote.Visibility=hasArchive ? Visibility.Visible : Visibility.Collapsed;
        }
        private async void OpenFileClicked(object sender,RoutedEventArgs args)
        {
            if(String.IsNullOrEmpty(SavedFilePath)) return;
            try { if(!await Launcher.LaunchFileAsync(await StorageFile.GetFileFromPathAsync(SavedFilePath))) ShowOpenError(); }
            catch(Exception error) { ShowOpenError(error.Message); }
        }
        private async void OpenDestinationClicked(object sender,RoutedEventArgs args)
        {
            try
            {
                if(!String.IsNullOrEmpty(SavedFilePath))
                {
                    StorageFile file=await StorageFile.GetFileFromPathAsync(SavedFilePath);
                    StorageFolder folder=await StorageFolder.GetFolderFromPathAsync(Path.GetDirectoryName(SavedFilePath));
                    var options=new FolderLauncherOptions();
                    options.ItemsToSelect.Add(file);
                    if(!await Launcher.LaunchFolderAsync(folder,options)) ShowOpenError();
                }
                else if(!await Launcher.LaunchFolderAsync(await StorageFolder.GetFolderFromPathAsync(!String.IsNullOrEmpty(Summary?.PackagePath) ? Summary.PackagePath : directory))) ShowOpenError();
            }
            catch(Exception error) { ShowOpenError(error.Message); }
        }
        private async void OpenArchiveClicked(object sender,RoutedEventArgs args)
        {
            if(String.IsNullOrEmpty(Summary?.ArchivePath)) return;
            try { if(!await Launcher.LaunchFileAsync(await StorageFile.GetFileFromPathAsync(Summary.ArchivePath))) ShowOpenError(); }
            catch(Exception error) { ShowOpenError(error.Message); }
        }
        private async void OpenReportClicked(object sender,RoutedEventArgs args)
        {
            if(String.IsNullOrEmpty(Summary?.ReportPath)) return;
            try { if(!await Launcher.LaunchFileAsync(await StorageFile.GetFileFromPathAsync(Summary.ReportPath))) ShowOpenError(); }
            catch(Exception error) { ShowOpenError(error.Message); }
        }
        private void ShowOpenError(string message="Windows could not open this location.")
        {
            SummaryInfo.Title="Could not open the export location"; SummaryInfo.Message=message; SummaryInfo.Severity=InfoBarSeverity.Warning; SummaryInfo.IsOpen=true;
        }
        private static string Friendly(RecoveryStatus status)
        {
            return status switch { RecoveryStatus.Success=>"Exported", RecoveryStatus.Unsupported=>"Unsupported format",
                RecoveryStatus.Unavailable=>"Content unavailable", RecoveryStatus.Corrupt=>"Invalid content", RecoveryStatus.SqlError=>"Database error", _=>"Failed" };
        }
    }
}

