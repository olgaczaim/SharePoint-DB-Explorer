using System;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using SharePointExplorer.Desktop.ViewModels;

namespace SharePointExplorer.WinUI
{
    public sealed partial class VersionHistoryDialog : ContentDialog
    {
        private readonly ExplorerViewModel viewModel;
        private readonly Node document;
        private readonly ObservableCollection<VersionViewModel> versions=new ObservableCollection<VersionViewModel>();
        private bool started,closed;
        public Node SelectedVersionNode { get; private set; }

        public VersionHistoryDialog(ExplorerViewModel viewModel,Node document)
        {
            this.viewModel=viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            if(document==null || document.Kind!=NodeKind.File || document.HistoryVersion!=0)
                throw new ArgumentException("Select a current document to browse its versions.",nameof(document));
            this.document=ExplorerNode.Copy(document);
            InitializeComponent();
            DocumentNameText.Text=this.document.Name;
            DocumentPathText.Text=this.document.Path;
            VersionsList.ItemsSource=versions;
        }
        private async void DialogOpened(ContentDialog sender,ContentDialogOpenedEventArgs args)
        {
            if(started) return; started=true;
            try
            {
                var rows=await viewModel.GetFileVersionsAsync(document);
                if(closed) return;
                foreach(Node row in rows) versions.Add(new VersionViewModel(row));
                int historical=versions.Count(row=>row.HistoryVersion!=0);
                VersionCountText.Text=historical==0 ? "No older versions are recorded for this document."
                    : historical+" older "+(historical==1 ? "version" : "versions")+" recorded";
                VersionsList.IsEnabled=versions.Count>0;
                VersionsList.SelectedItem=versions.FirstOrDefault(row=>row.HistoryVersion!=0 && row.CanExport)
                    ?? versions.FirstOrDefault(row=>row.CanExport) ?? versions.FirstOrDefault();
                if(versions.Count==0)
                {
                    VersionsInfo.Title=String.IsNullOrEmpty(viewModel.ErrorText) ? "No versions found" : "Could not load versions";
                    VersionsInfo.Message=String.IsNullOrEmpty(viewModel.ErrorText) ? "Refresh the folder and open version history again." : viewModel.ErrorText;
                    VersionsInfo.Severity=String.IsNullOrEmpty(viewModel.ErrorText) ? InfoBarSeverity.Informational : InfoBarSeverity.Error;
                    VersionsInfo.IsOpen=true;
                }
            }
            catch(Exception error)
            {
                if(closed) return;
                VersionsInfo.Title="Could not load versions"; VersionsInfo.Message=error.GetBaseException().Message;
                VersionsInfo.Severity=InfoBarSeverity.Error; VersionsInfo.IsOpen=true;
                VersionCountText.Text="Version history could not be loaded.";
            }
            finally { if(!closed) { VersionsLoadingRing.IsActive=false; VersionsLoadingRing.Visibility=Visibility.Collapsed; } }
        }
        private void VersionSelectionChanged(object sender,SelectionChangedEventArgs args)
        {
            if(closed) return;
            VersionViewModel selected=VersionsList.SelectedItem as VersionViewModel;
            IsPrimaryButtonEnabled=selected?.CanExport==true;
            VersionSelectionText.Text=selected==null ? "Choose a version, then select an export folder."
                : selected.CanExport ? "Export version "+selected.VersionText+" to a folder you choose." : "This version is browse only.";
        }
        private void ExportVersionClicked(ContentDialog sender,ContentDialogButtonClickEventArgs args)
        {
            VersionViewModel selected=VersionsList.SelectedItem as VersionViewModel;
            if(selected?.CanExport!=true) {args.Cancel=true;return;}
            SelectedVersionNode=ExplorerNode.Copy(selected.Node);
        }
        public void RequestClose() { if(!closed) Hide(); }
        private void DialogClosed(ContentDialog sender,ContentDialogClosedEventArgs args) { closed=true; }
    }
}