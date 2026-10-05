#nullable enable
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using SharePointExplorer.Desktop;
using SharePointExplorer.Desktop.ViewModels;

namespace SharePointExplorer.Tests
{
    public static class ExplorerViewModelChecks
    {
        public static async Task RunAsync()
        {
            OneNoteEligibility();
            await ExportableFilesAsync();
            await NavigationAndSelectionAsync();
            await StaleNavigationAsync();
            await SourceReplacementAsync();
            await ConnectionRacesAsync();
            await ExportSnapshotAndCancellationAsync();
            await VersionBrowsingAndExportAsync();
            await VersionReadRacesAndLifetimeAsync();
            await ZipSelectionSnapshotAsync();
            await ZipCancellationAndLifetimeAsync();
            await LibraryScopeAndExportAsync();
            await LibraryCancellationAndLifetimeAsync();
            await ErrorAndRetryAsync();
            await DisposalWaitsAsync();
            await RetiredSourcesAreReleasedAsync();
            Console.WriteLine("PASS WinUI-independent MVVM navigation, folder-local checks, stale results, export and async disposal");
        }

        private static void OneNoteEligibility()
        {
            foreach(byte schema in new byte[] {82,98})
            {
                Node file=File(schema==82?"section.one":"notebook.onetoc2");file.StreamSchema=schema;file.HasStream=true;
                var row=new ItemViewModel(file);row.IsChecked=true;
                Check(row.CanCheck && row.IsChecked && new VersionViewModel(file).CanExport,"A supported native OneNote section or TOC cannot be selected for recovery.");
                file.HistoryVersion=512;file.UiVersion=512;
                Check(!new ItemViewModel(file).CanCheck && new VersionViewModel(file).CanExport,"Native OneNote history entered current selection or lost version recovery eligibility.");
                file.HasStream=false;
                Check(!new ItemViewModel(file).CanCheck && !new VersionViewModel(file).CanExport,"Native OneNote eligibility concealed missing stored content.");
            }
        }
        private static async Task ExportableFilesAsync()
        {
            Node storedText=File("stored.txt"), storedAspx=File("custom-page.aspx"), template=File("DispForm.aspx"), unsupported77=File("unsupported77.bin"), unsupported67=File("unsupported67.bin"), unknown=File("unknown.bin"), history=File("old.txt");
            storedAspx.StreamSchema=66;template.HasStream=false;unsupported77.StreamSchema=77;unsupported67.StreamSchema=67;unknown.StreamSchema=255;history.HistoryVersion=1;
            Node folder=Container(NodeKind.Folder,"Folder"), item=new(){Kind=NodeKind.ListItem,Id=Guid.NewGuid(),Name="Task"};
            foreach(Node document in new[]{template,unsupported77,unsupported67,unknown,history,folder,item})
            {
                var row=new ItemViewModel(document);row.IsChecked=true;
                Check(!row.CanCheck && !row.IsChecked,"A template, unsupported schema, historical version or browse-only item accepted an export checkbox.");
                if(document.Kind==NodeKind.File)Check(row.ExportStatus=="Browse only","A non-exportable file initially displayed a recovery warning.");
            }
            foreach(Node document in new[]{storedText,storedAspx})
            {
                var row=new ItemViewModel(document);row.IsChecked=true;
                Check(row.CanCheck && row.IsChecked,"A current supported stored document, including ASPX, was excluded by extension.");
            }
            Node legacy=File("legacy.txt");legacy.HasStream=null;
            Check(new ItemViewModel(legacy).CanCheck,"Unknown legacy HasStream metadata excluded a supported storage schema.");
            var controller=new FakeController("Exportable files");Node site=Container(NodeKind.Site,"Mixed files");
            controller.Roots.Add(site);controller.Children[site.Id]=new(){storedText,storedAspx,template,unsupported77,unsupported67,unknown,history,folder,item};
            await using var harness=new Harness(controller);var vm=harness.Model;
            await vm.SetSourceAsync(controller);await vm.NavigateAsync(vm.RootNodes.Single());await harness.Ui.InvokeAsync(vm.SelectAllFiles);
            Check(vm.SelectedCount==2 && vm.VisibleItems.Where(row=>row.IsChecked).Select(row=>row.Node.Id).ToHashSet().SetEquals(new[]{storedText.Id,storedAspx.Id}),
                "Select all included unavailable forms, unsupported storage or historical rows, or omitted a stored ASPX file.");
            Check((await vm.ExportAsync("unused-fake-destination"))?.Success==2 && controller.Captured.Select(document=>document.Id).ToHashSet().SetEquals(new[]{storedText.Id,storedAspx.Id}),
                "Browse-only templates reached the selection export controller.");
            await harness.Ui.InvokeAsync(vm.ClearSelection);ItemViewModel storedRow=vm.VisibleItems.Single(row=>row.Node.Id==storedText.Id);
            await harness.Ui.InvokeAsync(()=>storedRow.IsChecked=true);
            foreach(RecoveryStatus ignored in new[]{RecoveryStatus.Unsupported,RecoveryStatus.Unavailable})
            {
                controller.Outcomes[storedText.Id]=ignored;DesktopExportSummary? summary=await vm.ExportAsync("unused-fake-destination");
                Check(summary?.Skipped==1 && summary.Entries.Single().Status==ignored && vm.LastExport==summary,
                    "A runtime unsupported/unavailable result was removed from the audit summary.");
                Check(storedRow.ExportStatus=="Not exported" && !vm.Status.Contains("skipped",StringComparison.OrdinalIgnoreCase) && !vm.Status.Contains("unsupported",StringComparison.OrdinalIgnoreCase) && !vm.Status.Contains("unavailable",StringComparison.OrdinalIgnoreCase),
                    "A runtime unsupported/unavailable result exposed a warning in the visible status.");
                Check(vm.ExportCompleted==0 && vm.ExportTotal==0 && !vm.IsExporting && storedRow.IsChecked,
                    "Ignored runtime outcomes entered visible completion totals or changed selection.");
            }
            foreach(RecoveryStatus failure in new[]{RecoveryStatus.Corrupt,RecoveryStatus.SqlError})
            {
                controller.Outcomes[storedText.Id]=failure;DesktopExportSummary? summary=await vm.ExportAsync("unused-fake-destination");
                Check(summary?.Failed==1 && summary.Entries.Single().Status==failure && vm.Status.Contains("1 failed",StringComparison.Ordinal) && storedRow.ExportStatus== (failure==RecoveryStatus.Corrupt ? "Invalid content" : "Database error"),
                    "A real corrupt-content or SQL failure was hidden with unsupported forms.");
            }
            harness.CheckAffinity();
        }
        private static async Task NavigationAndSelectionAsync()
        {
            var controller = new FakeController("Navigation");
            Node site = Container(NodeKind.Site,"Site"), library = Container(NodeKind.Library,"Documents"), list = Container(NodeKind.List,"Tasks");
            Node a = Container(NodeKind.Folder,"A"), b = Container(NodeKind.Folder,"B"), top = File("top.txt"), first = File("first.txt"), second = File("second.txt");
            Node item = new() {Kind=NodeKind.ListItem,Id=Guid.NewGuid(),Name="Task",Title="Read-only task",ListItemId=7,HasAttachments=true};
            controller.Roots.Add(site); controller.Children[site.Id] = new() {library,list};
            controller.Children[library.Id] = new() {a,b,top}; controller.Children[a.Id] = new() {first}; controller.Children[b.Id] = new() {second}; controller.Children[list.Id] = new() {item};
            await using var harness = new Harness(controller); var vm = harness.Model;
            Check(await vm.SetSourceAsync(controller),"Source did not open.");
            TreeItemViewModel root = vm.RootNodes.Single();
            Check(root.HasUnrealizedChildren && !root.IsLoaded && controller.ChildCalls.Count == 0,"Opening a source eagerly loaded the tree.");
            await vm.ExpandAsync(root); await vm.NavigateAsync(root);
            Check(controller.ChildCalls[site.Id] == 1 && root.Children.Count == 2,"Expansion and navigation did not share their cached query.");
            TreeItemViewModel libraryTree = root.Children.Single(tree => tree.Node.Id == library.Id), listTree = root.Children.Single(tree => tree.Node.Id == list.Id);
            await vm.NavigateAsync(libraryTree); await harness.Ui.InvokeAsync(vm.SelectAllFiles);
            ItemViewModel old = vm.VisibleItems.Single(row => row.CanCheck);
            Check(vm.SelectedCount == 1 && vm.CanExport,"File selection was not counted.");
            using var gate = new Gate(); controller.Gates[a.Id] = gate;
            TreeItemViewModel aTree = libraryTree.Children.Single(tree => tree.Node.Id == a.Id), bTree = libraryTree.Children.Single(tree => tree.Node.Id == b.Id);
            Task navigation = vm.NavigateAsync(aTree); await gate.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            await harness.Ui.InvokeAsync(() => Check(!old.IsChecked && vm.SelectedCount == 0 && vm.VisibleItems.Count == 0 && !vm.CanExport,"Previous folder checks survived the start of navigation."));
            gate.Release(); await navigation;
            await harness.Ui.InvokeAsync(() => vm.VisibleItems.Single().IsChecked = true);
            ItemViewModel previous = vm.VisibleItems.Single();
            await vm.NavigateAsync(bTree);
            Check(!previous.IsChecked && vm.SelectedCount == 0 && !vm.VisibleItems.Single().IsChecked,"Folder changes carried checks into the next folder.");
            await vm.NavigateAsync(aTree); await harness.Ui.InvokeAsync(vm.SelectAllFiles); await vm.RefreshAsync();
            Check(controller.ChildCalls[a.Id] == 2 && vm.SelectedCount == 0 && vm.VisibleItems.All(row => !row.IsChecked),"Refresh reused stale data or retained checks.");
            await vm.UpAsync(); Check(ReferenceEquals(vm.CurrentFolder,libraryTree),"Up did not navigate to the actual parent.");
            await vm.NavigateAsync(listTree);
            await harness.Ui.InvokeAsync(() => vm.VisibleItems.Single().IsChecked = true);
            ItemViewModel task = vm.VisibleItems.Single();
            Check(!task.CanCheck && !task.IsChecked && vm.SelectedCount == 0 && task.ItemIdText == "7" && task.Title == "Read-only task" && task.HasAttachmentsText == "Yes","Ordinary list metadata or read-only selection changed.");
            await vm.NavigateAsync(null); await vm.RefreshAsync();
            TreeItemViewModel fresh = vm.RootNodes.Single(); Check(!ReferenceEquals(fresh,root),"Source-root refresh retained stale tree models.");
            await vm.NavigateAsync(root); Check(vm.CurrentFolder == null,"A detached tree from a previous refresh was accepted.");
            harness.CheckAffinity();
        }

        private static async Task StaleNavigationAsync()
        {
            var controller = new FakeController("Stale navigation"); Node site=Container(NodeKind.Site,"Site"), a=Container(NodeKind.Folder,"A"), b=Container(NodeKind.Folder,"B");
            controller.Roots.Add(site); controller.Children[site.Id]=new(){a,b}; controller.Children[a.Id]=new(){File("old.txt")}; controller.Children[b.Id]=new(){File("new.txt")};
            await using var harness = new Harness(controller); var vm=harness.Model; await vm.SetSourceAsync(controller); await vm.ExpandAsync(vm.RootNodes.Single());
            TreeItemViewModel at=vm.RootNodes.Single().Children[0], bt=vm.RootNodes.Single().Children[1];
            using var gate = new Gate(); controller.Gates[a.Id]=gate;
            Task older=vm.NavigateAsync(at); await gate.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            await vm.NavigateAsync(bt); await harness.Ui.InvokeAsync(() => vm.VisibleItems.Single().IsChecked=true);
            gate.Release(); await older;
            Check(ReferenceEquals(vm.CurrentFolder,bt) && vm.VisibleItems.Single().Name=="new.txt" && vm.SelectedCount==1 && !vm.IsBusy,"An older folder completion replaced the current folder or its selection.");
        }

        private static async Task SourceReplacementAsync()
        {
            var oldSource = new FakeController("Old source"); Node oldRoot=Container(NodeKind.Site,"Old"), slow=Container(NodeKind.Folder,"Slow");
            oldSource.Roots.Add(oldRoot); oldSource.Children[oldRoot.Id]=new(){slow,File("old.txt")}; oldSource.Children[slow.Id]=new(){File("delayed.txt")};
            var newSource = new FakeController("New source"); Node newRoot=Container(NodeKind.Site,"New"); newSource.Roots.Add(newRoot);
            await using var harness = new Harness(oldSource,newSource); var vm=harness.Model; await vm.SetSourceAsync(oldSource); await vm.NavigateAsync(vm.RootNodes.Single());
            ItemViewModel oldFile=vm.VisibleItems.Single(row=>row.CanCheck); await harness.Ui.InvokeAsync(()=>oldFile.IsChecked=true);
            using var gate = new Gate(); oldSource.Gates[slow.Id]=gate;
            Task delayed=vm.ExpandAsync(vm.RootNodes.Single().Children.Single()); await gate.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Check(await vm.SetSourceAsync(newSource),"Replacement source failed to open.");
            Check(!oldFile.IsChecked && vm.SelectedCount==0 && vm.SourceName=="New source" && oldSource.Disposals==0,"Source replacement retained checks or disposed an active read.");
            gate.Release(); await delayed; await oldSource.Disposed.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Check(vm.RootNodes.Single().Name=="New" && vm.VisibleItems.Single().Name=="New" && vm.ErrorText=="" && oldSource.Disposals==1 && !oldSource.DisposedWhileActive,"A replaced source completion changed the new source or closed an active operation.");
        }

        private static async Task ConnectionRacesAsync()
        {
            FakeController first=new("First"), second=new("Second"); first.Roots.Add(Container(NodeKind.Site,"First site")); second.Roots.Add(Container(NodeKind.Site,"Second site"));
            using var gate=new Gate();
            await using var harness=new Harness(new[]{first,second}, options=> {
                if(options.Server=="first") {gate.Wait(); return first;}
                if(options.Server=="error") throw new IOException("Connection rejected private-password.");
                return second;
            });
            var vm=harness.Model;
            Task<bool> older=vm.ConnectAsync(new SqlConnectionOptions {Server="first"}); await gate.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Check(await vm.ConnectAsync(new SqlConnectionOptions {Server="second"}),"Newer connection failed.");
            gate.Release(); Check(!await older,"An obsolete connection reported itself current."); await first.Disposed.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Check(vm.SourceName=="Second" && vm.RootNodes.Single().Name=="Second site","An older connect completion replaced the newer source.");
            Check(!await vm.ConnectAsync(new SqlConnectionOptions {Server="error",Password="private-password"}),"A failed connection reported success.");
            Check(!vm.IsConnected && vm.RootNodes.Count==0 && vm.VisibleItems.Count==0 && !vm.IsBusy && vm.ErrorText.Contains("[redacted]") && !vm.ErrorText.Contains("private-password"),"A connection error left source state or revealed a password.");
        }

        private static async Task ExportSnapshotAndCancellationAsync()
        {
            var controller=new FakeController("Exports"); Node root=Container(NodeKind.Site,"Files"), first=File("one.txt"), second=File("two.txt");
            controller.Roots.Add(root); controller.Children[root.Id]=new(){first,second};
            await using var harness=new Harness(controller); var vm=harness.Model; await vm.SetSourceAsync(controller); await vm.NavigateAsync(vm.RootNodes.Single()); await harness.Ui.InvokeAsync(vm.SelectAllFiles);
            using var exportGate=new Gate(); controller.ExportGate=exportGate;
            Task<DesktopExportSummary?> export=vm.ExportAsync("unused-fake-destination"); await exportGate.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            await harness.Ui.InvokeAsync(vm.ClearSelection); first.Name="mutated source";
            Check(vm.IsExporting && !vm.CanNavigate && !vm.CanExport && controller.Captured.Count==2 && controller.Captured[0].Name=="one.txt","Export did not freeze the checked selection before background work.");
            await vm.NavigateAsync(null); Check(vm.CurrentFolder!=null,"Navigation was accepted during export.");
            exportGate.Release(); DesktopExportSummary? complete=await export;
            Check(complete?.Success==2 && vm.LastExport==complete && vm.ExportCompleted==2 && vm.ExportTotal==2 && !vm.IsExporting && vm.VisibleItems.All(item=>item.ExportStatus=="Exported"),"Export progress/summary did not reach the view model.");
            controller.ExportGate=null; await harness.Ui.InvokeAsync(vm.SelectAllFiles);
            using var pre=new CancellationTokenSource(); pre.Cancel(); DesktopExportSummary? cancelled=await vm.ExportAsync("unused-fake-destination",pre.Token);
            Check(cancelled?.Cancelled==true && cancelled.Entries.Count==0 && !vm.IsExporting,"Pre-cancellation was not passed to the controller.");
            using var between=new Gate(); controller.AfterFirst=between;
            Task<DesktopExportSummary?> partial=vm.ExportAsync("unused-fake-destination"); await between.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            await harness.Ui.InvokeAsync(vm.CancelExport); between.Release(); DesktopExportSummary? result=await partial;
            Check(result?.Cancelled==true && result.Success==1 && vm.LastExport==result && !vm.IsExporting,"Cancellation between files did not preserve the completed result.");
            controller.AfterFirst=null; controller.ExportError=new InvalidDataException("Decoder failed.");
            Check(await vm.ExportAsync("unused-fake-destination")==null && vm.ErrorText=="Decoder failed." && !vm.IsExporting && vm.CanNavigate,"Export errors did not restore navigation and a friendly error state.");
            harness.CheckAffinity();
        }

        private static async Task VersionBrowsingAndExportAsync()
        {
            var controller=new FakeController("Version source");Node site=Container(NodeKind.Site,"Site"), current=File("report.txt"), other=File("other.txt");
            current.UiVersion=1024;current.InternalVersion=2049;current.Level=1;
            Node old=CopyVersion(current);old.HistoryVersion=512;old.UiVersion=512;old.InternalVersion=1025;
            Node draft=CopyVersion(current);draft.HistoryVersion=513;draft.UiVersion=513;draft.InternalVersion=1026;draft.Level=2;
            controller.Roots.Add(site);controller.Children[site.Id]=new(){current,other};controller.Versions[current.Id]=new(){current,draft,old};
            await using var harness=new Harness(controller);var vm=harness.Model;
            await vm.SetSourceAsync(controller);await vm.NavigateAsync(vm.RootNodes.Single());
            List<Node> versions=await vm.GetFileVersionsAsync(current);Node historical=versions.Single(version=>version.HistoryVersion==512);
            var row=new VersionViewModel(historical);var currentVersion=new VersionViewModel(versions.Single(version=>version.HistoryVersion==0));
            Check(versions.Count==3 && row.VersionText=="1.0" && row.StateText=="Historical" && row.CanExport && !new ItemViewModel(historical).CanCheck && currentVersion.VersionText=="2.0" && currentVersion.StateText=="Current",
                "Version history did not label current/historical states or keep history outside ordinary checkboxes.");
            historical.Name="Changed returned row";
            Check(row.Node.Name=="report.txt","A version row retained mutable source metadata.");
            ItemViewModel visible=vm.VisibleItems.Single(item=>item.Node.Id==current.Id);
            Check(vm.SelectedCount==0 && !vm.CanExport,"Reading version history changed the checked-file selection.");
            DesktopExportSummary? saved=await vm.ExportVersionAsync(row.Node,"unused-fake-destination");
            Check(saved?.Success==1 && controller.VersionExports==1 && controller.FilesExports==0 && controller.LibraryExports==0 && controller.ZipExports==0 && controller.CapturedVersion?.HistoryVersion==512,
                "Exporting a historical version required checks or routed through current-file recovery.");
            Check(visible.ExportStatus=="Ready" && !visible.IsChecked && vm.SelectedCount==0 && saved!.Entries.Single().Document.UiVersion==512 && vm.LastExport==saved,
                "Historical recovery marked the current file exported or changed current-folder checks.");
            await harness.Ui.InvokeAsync(()=>visible.IsChecked=true);using var gate=new Gate();controller.ExportGate=gate;
            Task<DesktopExportSummary?> exporting=vm.ExportVersionAsync(row.Node,"unused-fake-destination");await gate.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Check(vm.IsExporting && !vm.CanNavigate && !vm.CanExport && controller.CapturedVersion?.InternalVersion==1025,
                "Exact-version export did not publish busy state or freeze the stored identity.");
            Node changed=row.Node;changed.UiVersion=1024;await vm.NavigateAsync(null);
            Check(vm.CurrentFolder!=null && visible.IsChecked && vm.SelectedCount==1,"Version recovery allowed navigation or changed checked current files.");
            gate.Release();await exporting;controller.ExportGate=null;
            Check(visible.ExportStatus=="Ready" && visible.IsChecked && vm.SelectedCount==1,"Version completion changed current-file export status or checks.");
            int calls=controller.VersionExports;
            Check(await vm.ExportVersionAsync(changed,"unused-fake-destination")==null && controller.VersionExports==calls,
                "A mutated version identity reached the controller.");
            controller.VersionsError=new IOException("Version metadata is unavailable.");
            Check((await vm.GetFileVersionsAsync(current)).Count==0 && vm.ErrorText=="Version metadata is unavailable." && !vm.IsLoadingVersions && !vm.IsBusy,
                "A failed version query left stale choices or blocked the view model.");
            Check(await vm.ExportVersionAsync(row.Node,"unused-fake-destination")==null,"A failed version refresh left old choices exportable.");
            controller.VersionsError=null;versions=await vm.GetFileVersionsAsync(current);
            Node fresh=versions.Single(version=>version.HistoryVersion==512);
            using var cancelled=new CancellationTokenSource();cancelled.Cancel();DesktopExportSummary? cancelledResult=await vm.ExportVersionAsync(fresh,"unused-fake-destination",cancelled.Token);
            Check(cancelledResult?.Cancelled==true && cancelledResult.Entries.Count==0 && !vm.IsExporting && visible.IsChecked,
                "A pre-cancelled version export recovered bytes or altered checks.");
            controller.ExportError=new IOException("Version destination rejected.");
            Check(await vm.ExportVersionAsync(fresh,"unused-fake-destination")==null && vm.ErrorText=="Version destination rejected." && vm.CanNavigate,
                "A version export error failed to restore navigation.");
            controller.ExportError=null;
            Check((await vm.ExportVersionAsync(currentVersion.Node,"unused-fake-destination"))?.Success==1 && visible.ExportStatus=="Exported",
                "An explicitly selected current version did not use the exact version route.");
            harness.CheckAffinity();
        }

        private static async Task VersionReadRacesAndLifetimeAsync()
        {
            var controller=new FakeController("Version races");Node site=Container(NodeKind.Site,"Site"), folder=Container(NodeKind.Folder,"Other folder"), current=File("report.txt");
            current.UiVersion=1024;current.InternalVersion=2049;current.Level=1;Node old=CopyVersion(current);old.HistoryVersion=512;old.UiVersion=512;old.InternalVersion=1025;
            controller.Roots.Add(site);controller.Children[site.Id]=new(){current,folder};controller.Children[folder.Id]=new(){File("different.txt")};controller.Versions[current.Id]=new(){current,old};
            var replacement=new FakeController("New version source");Node replacementSite=Container(NodeKind.Site,"Replacement");replacement.Roots.Add(replacementSite);replacement.Children[replacementSite.Id]=new(){CopyVersion(current)};replacement.Versions[current.Id]=new(){CopyVersion(current),CopyVersion(old)};
            await using var harness=new Harness(controller,replacement);var vm=harness.Model;await vm.SetSourceAsync(controller);await vm.NavigateAsync(vm.RootNodes.Single());
            TreeItemViewModel root=vm.CurrentFolder!, other=root.Children.Single();await harness.Ui.InvokeAsync(()=>vm.VisibleItems.Single(row=>row.CanCheck).IsChecked=true);
            using var delayed=new Gate();controller.VersionsGate=delayed;
            Task<List<Node>> read=vm.GetFileVersionsAsync(current);await delayed.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Check(vm.IsLoadingVersions && vm.IsBusy && !vm.CanExport && vm.SelectedCount==1,"A version metadata query did not expose read-busy state independently of checks.");
            await vm.NavigateAsync(other);delayed.Release();Check((await read).Count==0 && ReferenceEquals(vm.CurrentFolder,other) && vm.SelectedCount==0 && !vm.IsLoadingVersions && vm.ErrorText=="",
                "An older version query changed a new folder or returned detached choices.");
            controller.VersionsGate=null;await vm.NavigateAsync(root);
            using var rootRefresh=new Gate();controller.VersionsGate=rootRefresh;
            Task<List<Node>> beforeRefresh=vm.GetFileVersionsAsync(current);await rootRefresh.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            await vm.NavigateAsync(null);await vm.RefreshAsync();
            Check(!vm.IsLoadingVersions && !vm.IsBusy,"Root refresh retained the busy state of an invalidated version query.");
            rootRefresh.Release();Check((await beforeRefresh).Count==0,"A version query survived a root refresh generation.");controller.VersionsGate=null;
            root=vm.RootNodes.Single();await vm.NavigateAsync(root);other=root.Children.Single();
            List<Node> choices=await vm.GetFileVersionsAsync(current);Node selected=choices.Single(version=>version.HistoryVersion==512);
            await vm.NavigateAsync(other);int exports=controller.VersionExports;
            Check(await vm.ExportVersionAsync(selected,"unused-fake-destination")==null && controller.VersionExports==exports,"A version choice survived a folder change.");
            await vm.NavigateAsync(root);using var sourceRead=new Gate();controller.VersionsGate=sourceRead;
            Task<List<Node>> pending=vm.GetFileVersionsAsync(current);await sourceRead.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Check(await vm.SetSourceAsync(replacement),"Replacement source failed to open during a version read.");
            Check(controller.Disposals==0,"A source was disposed beneath its active version metadata call.");
            sourceRead.Release();Check((await pending).Count==0,"A replaced source returned version choices into the new session.");await controller.Disposed.Task.WaitAsync(TimeSpan.FromSeconds(10));
            await vm.NavigateAsync(vm.RootNodes.Single());
            Check(await vm.ExportVersionAsync(selected,"unused-fake-destination")==null && replacement.VersionExports==0,"A previous source version matched coincidentally equal document IDs in a new source.");
            List<Node> newChoices=await vm.GetFileVersionsAsync(current);Node fresh=newChoices.Single(version=>version.HistoryVersion==512);
            using var pendingExport=new Gate();replacement.ExportGate=pendingExport;
            Task<DesktopExportSummary?> active=vm.ExportVersionAsync(fresh,"unused-fake-destination");await pendingExport.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Task closing=vm.DisposeAsync().AsTask();await harness.Ui.InvokeAsync(()=>{});
            Check(!closing.IsCompleted && replacement.Disposals==0,"Shutdown disposed a source while version recovery was active.");
            pendingExport.Release();await active;await closing;
            Check(replacement.Disposals==1 && !replacement.DisposedWhileActive && !vm.IsLoadingVersions && !vm.IsExporting,"Version shutdown did not await and release its source exactly once.");
            harness.CheckAffinity();
        }
        private static Node CopyVersion(Node node)=>new() {Kind=node.Kind,SiteId=node.SiteId,WebId=node.WebId,ListId=node.ListId,Id=node.Id,Name=node.Name,Path=node.Path,Size=node.Size,
            StreamSchema=node.StreamSchema,HasStream=node.HasStream,HistoryVersion=node.HistoryVersion,UiVersion=node.UiVersion,InternalVersion=node.InternalVersion,Level=node.Level,Modified=node.Modified,ParentId=node.ParentId};
        private static async Task ZipSelectionSnapshotAsync()
        {
            var controller=new FakeController("ZIP selections");
            Node site=Container(NodeKind.Site,"Site"), firstFolder=Container(NodeKind.Folder,"First"), secondFolder=Container(NodeKind.Folder,"Second"), list=Container(NodeKind.List,"Tasks");
            Node first=File("first.txt"), uncheckedFile=File("unchecked.txt"), other=File("other.txt");
            controller.Roots.Add(site);controller.Children[site.Id]=new(){firstFolder,secondFolder,list};
            controller.Children[firstFolder.Id]=new(){first,uncheckedFile};controller.Children[secondFolder.Id]=new(){other};
            controller.Children[list.Id]=new(){new Node{Kind=NodeKind.ListItem,Id=Guid.NewGuid(),Name="Task"}};
            await using var harness=new Harness(controller);var vm=harness.Model;
            await vm.SetSourceAsync(controller);await vm.NavigateAsync(vm.RootNodes.Single());
            TreeItemViewModel firstTree=vm.RootNodes.Single().Children.Single(tree=>tree.Node.Id==firstFolder.Id), secondTree=vm.RootNodes.Single().Children.Single(tree=>tree.Node.Id==secondFolder.Id), listTree=vm.RootNodes.Single().Children.Single(tree=>tree.Node.Id==list.Id);
            await vm.NavigateAsync(firstTree);
            string archivePath=Path.Combine(Path.GetTempPath(),"selected-files.zip");
            Check(await vm.ExportZipAsync(archivePath)==null && controller.ZipExports==0,"ZIP export accepted an empty selection.");
            ItemViewModel checkedRow=vm.VisibleItems.Single(row=>row.Name=="first.txt");
            await harness.Ui.InvokeAsync(()=>checkedRow.IsChecked=true);
            using var gate=new Gate();controller.ExportGate=gate;
            Task<DesktopExportSummary?> saving=vm.ExportZipAsync(archivePath);await gate.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Check(vm.IsExporting && !vm.CanNavigate && !vm.CanExport && controller.ZipExports==1 && controller.FilesExports==0 && controller.LibraryExports==0,
                "Save as ZIP did not route the selection exclusively through the archive controller operation.");
            Check(controller.RequestedArchivePath==archivePath && controller.Captured.Count==1 && controller.Captured[0].Id==first.Id && controller.Captured[0].Name=="first.txt",
                "ZIP export did not capture exactly the checked current-folder file and requested destination.");
            first.Name="Changed source name";await harness.Ui.InvokeAsync(vm.ClearSelection);
            await vm.NavigateAsync(secondTree);
            Check(ReferenceEquals(vm.CurrentFolder,firstTree) && controller.Captured.Count==1 && controller.Captured[0].Name=="first.txt",
                "An active ZIP operation changed scope after navigation or selection/source mutations.");
            gate.Release();DesktopExportSummary? result=await saving;controller.ExportGate=null;
            Check(result?.Success==1 && result.Total==1 && result.ArchivePath==archivePath && result.ArchiveReportEntry=="export-report.csv" && String.IsNullOrEmpty(result.ReportPath) && vm.LastExport==result && vm.ExportCompleted==1 && vm.ExportTotal==1 && !vm.IsExporting,
                "The ZIP summary, embedded report and progress were not published to the view model.");
            Check(ReferenceEquals(vm.CurrentFolder,firstTree) && checkedRow.ExportStatus=="Exported" && vm.SelectedCount==0 && !checkedRow.IsChecked,
                "ZIP completion restored old checks or changed the current folder.");
            await vm.NavigateAsync(secondTree);
            Check(vm.SelectedCount==0 && !vm.CanExport,"Folder navigation left a previous ZIP selection enabled.");
            await harness.Ui.InvokeAsync(()=>vm.VisibleItems.Single().IsChecked=true);
            controller.ExportError=new IOException("ZIP destination unavailable.");
            Check(await vm.ExportZipAsync(archivePath)==null && vm.ErrorText=="ZIP destination unavailable." && !vm.IsExporting && vm.CanNavigate,
                "A ZIP error did not restore navigation and the readable failure state.");
            controller.ExportError=null;await vm.NavigateAsync(listTree);int calls=controller.ZipExports;
            await harness.Ui.InvokeAsync(vm.SelectAllFiles);
            Check(!vm.CanExport && await vm.ExportZipAsync(archivePath)==null && controller.ZipExports==calls,
                "Ordinary list items reached ZIP recovery.");
            harness.CheckAffinity();
        }

        private static async Task ZipCancellationAndLifetimeAsync()
        {
            var controller=new FakeController("ZIP cancellation");Node site=Container(NodeKind.Site,"Site"), first=File("first.txt"), second=File("second.txt");
            controller.Roots.Add(site);controller.Children[site.Id]=new(){first,second};
            await using var harness=new Harness(controller);var vm=harness.Model;
            await vm.SetSourceAsync(controller);await vm.NavigateAsync(vm.RootNodes.Single());await harness.Ui.InvokeAsync(vm.SelectAllFiles);
            string archivePath=Path.Combine(Path.GetTempPath(),"cancelled-selections.zip");
            using var pre=new CancellationTokenSource();pre.Cancel();DesktopExportSummary? cancelled=await vm.ExportZipAsync(archivePath,pre.Token);
            Check(cancelled?.Cancelled==true && cancelled.Entries.Count==0 && String.IsNullOrEmpty(cancelled.ArchivePath) && !vm.IsExporting,
                "A pre-cancelled ZIP command claimed a published archive.");
            using var between=new Gate();controller.AfterFirst=between;
            Task<DesktopExportSummary?> saving=vm.ExportZipAsync(archivePath);await between.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            await harness.Ui.InvokeAsync(vm.CancelExport);between.Release();DesktopExportSummary? partial=await saving;controller.AfterFirst=null;
            Check(partial?.Cancelled==true && partial.Success==1 && partial.Entries.Count==1 && partial.ArchivePath==archivePath && vm.LastExport==partial && !vm.IsExporting && vm.SelectedCount==2,
                "ZIP cancellation between files lost its partial archive result or changed the checked selection.");
            using var pending=new Gate();controller.ExportGate=pending;
            Task<DesktopExportSummary?> active=vm.ExportZipAsync(archivePath);await pending.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Task closing=vm.DisposeAsync().AsTask();await harness.Ui.InvokeAsync(()=>{});
            Check(!closing.IsCompleted && controller.Disposals==0,"A source was closed while its ZIP operation was active.");
            pending.Release();await active;await closing;
            Check(controller.Disposals==1 && !controller.DisposedWhileActive && !vm.IsExporting,
                "ZIP shutdown did not cancel and await controller work exactly once.");
            harness.CheckAffinity();
        }
        private static async Task LibraryScopeAndExportAsync()
        {
            var controller=new FakeController("Library export");
            Node site=Container(NodeKind.Site,"Site"), library=Container(NodeKind.Library,"Documents"), empty=Container(NodeKind.Library,"Empty"), list=Container(NodeKind.List,"Tasks");
            Node folder=Container(NodeKind.Folder,"First"), nested=Container(NodeKind.Folder,"Unopened nested folder"), listFolder=Container(NodeKind.Folder,"List folder");
            Node top=File("top.txt"), visible=File("visible.txt"), deep=File("deep.txt");
            library.ListId=library.Id; empty.ListId=empty.Id;
            foreach(Node document in new[]{top,visible,deep}) {document.SiteId=library.SiteId;document.ListId=library.Id;}
            controller.Roots.Add(site); controller.Children[site.Id]=new(){library,empty,list};
            controller.Children[library.Id]=new(){top,folder}; controller.Children[folder.Id]=new(){visible,nested}; controller.Children[nested.Id]=new(){deep};
            controller.Children[list.Id]=new(){listFolder}; controller.LibraryFiles[library.Id]=new(){top,visible,deep};
            await using var harness=new Harness(controller); var vm=harness.Model;
            await vm.SetSourceAsync(controller);
            Check(vm.CurrentLibrary==null && !vm.CanExportLibrary,"The source-root view allowed whole-library export.");
            TreeItemViewModel siteTree=vm.RootNodes.Single(); await vm.NavigateAsync(siteTree);
            Check(vm.CurrentLibrary==null && !vm.CanExportLibrary,"A site was mistaken for a document library.");
            Check(await vm.ExportLibraryAsync("unused-fake-destination")==null && controller.LibraryExports==0,"A site reached the library export controller.");
            TreeItemViewModel libraryTree=siteTree.Children.Single(item=>item.Node.Id==library.Id), emptyTree=siteTree.Children.Single(item=>item.Node.Id==empty.Id), listTree=siteTree.Children.Single(item=>item.Node.Id==list.Id);
            await vm.NavigateAsync(libraryTree);
            Check(ReferenceEquals(vm.CurrentLibrary,libraryTree) && vm.CanExportLibrary && !vm.CanExport && vm.SelectedCount==0,"A library required checked files before it could be exported.");
            TreeItemViewModel folderTree=libraryTree.Children.Single(); await vm.NavigateAsync(folderTree);
            TreeItemViewModel nestedTree=folderTree.Children.Single();
            Check(ReferenceEquals(vm.CurrentLibrary,libraryTree) && vm.CanExportLibrary && !nestedTree.IsLoaded,"A subfolder did not retain its nearest library scope or eagerly loaded nested content.");
            ItemViewModel row=vm.VisibleItems.Single(item=>item.CanCheck);
            using var discovering=new Gate(); controller.LibraryDiscoveryGate=discovering;
            Task<DesktopExportSummary?> exporting=vm.ExportLibraryAsync("unused-fake-destination"); await discovering.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Check(vm.IsExporting && vm.ExportIsDiscovering && !vm.CanExportLibrary && !vm.CanExport && !vm.CanNavigate,"Library discovery did not expose progress and block conflicting commands.");
            Check(controller.CapturedLibrary?.Id==library.Id && controller.CapturedLibrary.Name=="Documents" && vm.SelectedCount==0 && !row.IsChecked,"The entire-library command did not freeze the current library independently of file checks.");
            library.Name="Changed after snapshot";
            await vm.NavigateAsync(siteTree); Check(ReferenceEquals(vm.CurrentFolder,folderTree),"Library export accepted navigation while its source was active.");
            discovering.Release(); DesktopExportSummary? result=await exporting; controller.LibraryDiscoveryGate=null;
            Check(result?.Success==3 && result.Total==3 && result.Entries.Select(entry=>entry.Document.Id).ToHashSet().SetEquals(new[]{top.Id,visible.Id,deep.Id}),"Whole-library results omitted a root or unopened nested document.");
            Check(controller.CapturedLibrary?.Name=="Documents" && !controller.ChildCalls.ContainsKey(nested.Id),"Library export mutated its scope or traversed the UI tree to discover documents.");
            Check(vm.LastExport==result && vm.ExportCompleted==3 && vm.ExportTotal==3 && !vm.ExportIsDiscovering && !vm.IsExporting && vm.CanExportLibrary,"Library progress and completion were not published.");
            Check(ReferenceEquals(vm.CurrentFolder,folderTree) && vm.SelectedCount==0 && !row.IsChecked && row.ExportStatus=="Exported","Whole-library export changed navigation/checks or failed to update its visible file status.");
            await harness.Ui.InvokeAsync(()=>row.IsChecked=true);
            Check((await vm.ExportLibraryAsync("unused-fake-destination"))?.Success==3 && row.IsChecked && vm.SelectedCount==1,"A checked subset limited whole-library scope or its checks were changed.");
            controller.ExportError=new IOException("Library inventory unavailable.");
            Check(await vm.ExportLibraryAsync("unused-fake-destination")==null && vm.ErrorText=="Library inventory unavailable." && !vm.IsExporting && vm.CanNavigate,"A library error left navigation/progress locked.");
            controller.ExportError=null;
            await vm.NavigateAsync(emptyTree); DesktopExportSummary? emptyResult=await vm.ExportLibraryAsync("unused-fake-destination");
            Check(emptyResult?.Total==0 && emptyResult.Entries.Count==0 && !vm.IsExporting && !vm.ExportIsDiscovering,"An empty library remained in discovery or failed to return a summary.");
            await vm.NavigateAsync(listTree); int calls=controller.LibraryExports;
            Check(vm.CurrentLibrary==null && !vm.CanExportLibrary && await vm.ExportLibraryAsync("unused-fake-destination")==null && controller.LibraryExports==calls,"An ordinary list was passed to document-library export.");
            await vm.NavigateAsync(listTree.Children.Single());
            Check(vm.CurrentLibrary==null && !vm.CanExportLibrary,"An ordinary-list folder was mistaken for a document-library folder.");
            await vm.NavigateAsync(null); Check(vm.CurrentLibrary==null && !vm.CanExportLibrary,"Returning to source roots retained a library export scope.");
            harness.CheckAffinity();
        }
        private static async Task LibraryCancellationAndLifetimeAsync()
        {
            var controller=new FakeController("Library cancellation"); Node site=Container(NodeKind.Site,"Site"), library=Container(NodeKind.Library,"Documents"), first=File("first.txt"), second=File("second.txt");
            library.ListId=library.Id; controller.Roots.Add(site); controller.Children[site.Id]=new(){library}; controller.Children[library.Id]=new(){first,second}; controller.LibraryFiles[library.Id]=new(){first,second};
            await using var harness=new Harness(controller); var vm=harness.Model; await vm.SetSourceAsync(controller); await vm.NavigateAsync(vm.RootNodes.Single()); await vm.NavigateAsync(vm.RootNodes.Single().Children.Single());
            using var pre=new CancellationTokenSource(); pre.Cancel(); DesktopExportSummary? cancelled=await vm.ExportLibraryAsync("unused-fake-destination",pre.Token);
            Check(cancelled?.Cancelled==true && cancelled.Entries.Count==0 && !vm.IsExporting,"A pre-cancelled whole-library command exported files.");
            using var discovery=new Gate(); controller.LibraryDiscoveryGate=discovery;
            Task<DesktopExportSummary?> discovering=vm.ExportLibraryAsync("unused-fake-destination"); await discovery.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            await harness.Ui.InvokeAsync(vm.CancelExport); discovery.Release(); DesktopExportSummary? discoveryResult=await discovering; controller.LibraryDiscoveryGate=null;
            Check(discoveryResult?.Cancelled==true && discoveryResult.Entries.Count==0 && !vm.IsExporting && !vm.ExportIsDiscovering,"Cancelling library discovery left work or a busy state behind.");
            using var between=new Gate(); controller.AfterFirst=between;
            Task<DesktopExportSummary?> exporting=vm.ExportLibraryAsync("unused-fake-destination"); await between.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            await harness.Ui.InvokeAsync(vm.CancelExport); between.Release(); DesktopExportSummary? partial=await exporting; controller.AfterFirst=null;
            Check(partial?.Cancelled==true && partial.Success==1 && partial.Entries.Count==1 && vm.LastExport==partial && !vm.IsExporting,"Library cancellation between files discarded the completed result.");
            using var pending=new Gate(); controller.LibraryDiscoveryGate=pending;
            Task<DesktopExportSummary?> active=vm.ExportLibraryAsync("unused-fake-destination"); await pending.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Task closing=vm.DisposeAsync().AsTask(); await harness.Ui.InvokeAsync(()=>{});
            Check(!closing.IsCompleted && controller.Disposals==0,"A source was disposed while library inventory was still active.");
            pending.Release(); await active; await closing;
            Check(controller.Disposals==1 && !controller.DisposedWhileActive && !vm.IsExporting,"Library shutdown did not await controller work exactly once.");
            harness.CheckAffinity();
        }
        private static async Task ErrorAndRetryAsync()
        {
            var controller=new FakeController("Retry"); Node root=Container(NodeKind.Site,"Site"); controller.Roots.Add(root); controller.Children[root.Id]=new(){File("recovered.txt")};
            controller.ChildError=new IOException("Temporary query failure.");
            await using var harness=new Harness(controller); var vm=harness.Model; await vm.SetSourceAsync(controller); TreeItemViewModel tree=vm.RootNodes.Single();
            await vm.NavigateAsync(tree);
            Check(vm.ErrorText=="Temporary query failure." && vm.VisibleItems.Count==0 && !tree.IsLoaded && !tree.IsLoading && tree.HasUnrealizedChildren && !vm.IsBusy,"A failed navigation did not remain retryable.");
            controller.ChildError=null; await vm.RefreshAsync();
            Check(vm.ErrorText=="" && vm.VisibleItems.Single().Name=="recovered.txt" && tree.IsLoaded && controller.ChildCalls[root.Id]==2,"Refresh did not retry a failed query.");
        }

        private static async Task DisposalWaitsAsync()
        {
            var controller=new FakeController("Dispose"); Node root=Container(NodeKind.Site,"Site"); controller.Roots.Add(root); controller.Children[root.Id]=new(){File("late.txt")};
            await using var harness=new Harness(controller); var vm=harness.Model; await vm.SetSourceAsync(controller);
            using var gate=new Gate(); controller.Gates[root.Id]=gate;
            Task read=vm.NavigateAsync(vm.RootNodes.Single()); await gate.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Task close=vm.DisposeAsync().AsTask(); await harness.Ui.InvokeAsync(()=>{});
            Check(!close.IsCompleted && controller.Disposals==0 && vm.SelectedCount==0 && vm.RootNodes.Count==0,"Dispose completed while a database read was still active.");
            gate.Release(); await read; await close; await vm.DisposeAsync();
            Check(controller.Disposals==1 && !controller.DisposedWhileActive && vm.VisibleItems.Count==0 && !vm.IsBusy,"Dispose did not wait exactly once or accepted a late completion.");
            Throws<ObjectDisposedException>(()=>vm.RefreshAsync());
            var pendingSource=new FakeController("Pending connect"); pendingSource.Roots.Add(Container(NodeKind.Site,"Pending")); using var connectGate=new Gate();
            await using var connecting=new Harness(new[]{pendingSource},_=>{connectGate.Wait(); return pendingSource;});
            Task<bool> open=connecting.Model.ConnectAsync(new SqlConnectionOptions()); await connectGate.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            Task shutdown=connecting.Model.DisposeAsync().AsTask(); Check(!shutdown.IsCompleted,"Shutdown ignored a pending source factory.");
            connectGate.Release(); Check(!await open,"A connection completed after disposal."); await shutdown;
            Check(pendingSource.Disposals==1 && connecting.Model.RootNodes.Count==0,"A late-created connection was leaked during shutdown.");
        }

        private static async Task RetiredSourcesAreReleasedAsync()
        {
            var current = new FakeController("Current source"); current.Roots.Add(Container(NodeKind.Site,"Current"));
            await using var harness = new Harness(current);
            WeakReference retired = await OpenAndRetireSourceAsync(harness.Model,current);
            // Advance the dispatcher past the source callback before checking
            // ownership; an idle native/thread dispatcher may retain its last delegate.
            await harness.Ui.InvokeAsync(static()=>{});
            await RequireCollectedAsync(retired,"The view model retained a disposed source/controller after replacement.");
            WeakReference failed = await OpenFailedSourceAsync(harness.Model);
            await harness.Ui.InvokeAsync(static()=>{});
            await RequireCollectedAsync(failed,"The view model retained a failed source/controller after disposal.");
        }
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static async Task<WeakReference> OpenAndRetireSourceAsync(ExplorerViewModel model,FakeController replacement)
        {
            var old = new FakeController("Retired source"); old.Roots.Add(Container(NodeKind.Site,"Old"));
            var weak = new WeakReference(old);
            Check(await model.SetSourceAsync(old),"Retirement fixture did not open.");
            Check(await model.SetSourceAsync(replacement),"Replacement fixture did not open.");
            await old.Disposed.Task.WaitAsync(TimeSpan.FromSeconds(10));
            return weak;
        }
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static async Task<WeakReference> OpenFailedSourceAsync(ExplorerViewModel model)
        {
            var failed = new FakeController("Failed source") {RootError=new IOException("Unavailable source.")};
            var weak = new WeakReference(failed);
            Check(!await model.SetSourceAsync(failed),"Failed-source fixture unexpectedly opened.");
            Check(failed.Disposals==1,"Failed source was not disposed.");
            return weak;
        }
        private static async Task RequireCollectedAsync(WeakReference reference,string message)
        {
            // Yield completed async state machines before checking actual object
            // reachability; a retained ownership-list entry keeps it alive.
            for (int attempt=0;attempt<100;attempt++)
            {
                await Task.Delay(10);
                GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
                if (!reference.IsAlive) return;
            }
            throw new Exception(message);
        }
        private static Node Container(NodeKind kind,string name) => new() {Kind=kind,SiteId=Guid.NewGuid(),Id=Guid.NewGuid(),Name=name,Path=name};
        private static Node File(string name) => new() {Kind=NodeKind.File,SiteId=Guid.NewGuid(),Id=Guid.NewGuid(),Name=name,Path=name,Size=4,HasStream=true};
        private static void Check(bool condition,string message) {if(!condition) throw new Exception(message);}
        private static void Throws<T>(Action action) where T:Exception {try{action();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name);}

        private sealed class Gate : IDisposable
        {
            private readonly ManualResetEventSlim release=new(false);
            internal readonly TaskCompletionSource Started=new(TaskCreationOptions.RunContinuationsAsynchronously);
            internal void Wait() {Started.TrySetResult(); if(!release.Wait(TimeSpan.FromSeconds(15))) throw new TimeoutException("Test gate did not release.");}
            internal void Release(){try{release.Set();}catch(ObjectDisposedException){}}
            public void Dispose(){release.Set(); release.Dispose();}
        }
        private sealed class UiThread : IDisposable
        {
            private readonly BlockingCollection<(Action Work,TaskCompletionSource Completion)> queue=new();
            private readonly Thread thread;
            internal int Id=>thread.ManagedThreadId;
            internal UiThread() {thread=new Thread(Run){IsBackground=true,Name="MVVM test dispatcher"};thread.Start();}
            internal Task InvokeAsync(Action work)
            {
                if(Environment.CurrentManagedThreadId==Id){work();return Task.CompletedTask;}
                var completion=new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);queue.Add((work,completion));return completion.Task;
            }
            private void Run()
            {
                while (queue.TryTake(out var item,Timeout.Infinite))
                {
                    try {item.Work();item.Completion.SetResult();}
                    catch (Exception error) {item.Completion.SetException(error);}
                    // A dispatcher must release the callback it has finished,
                    // rather than retain its captured objects while idle.
                    item=default;
                }
            }
            public void Dispose(){queue.CompleteAdding();if(!thread.Join(TimeSpan.FromSeconds(10)))throw new TimeoutException("UI test dispatcher did not stop.");queue.Dispose();}
        }
        private sealed class Harness : IAsyncDisposable
        {
            internal readonly UiThread Ui=new();
            internal readonly ExplorerViewModel Model;
            private readonly FakeController[] controllers;
            private int wrongThread;
            internal Harness(params FakeController[] controllers):this(controllers,null){}
            internal Harness(FakeController[] controllers,Func<SqlConnectionOptions,IExplorerController>? connector)
            {
                this.controllers=controllers;Model=new ExplorerViewModel(Ui.InvokeAsync,connector);
                Model.PropertyChanged+=(_,_)=>{if(Environment.CurrentManagedThreadId!=Ui.Id)Interlocked.Increment(ref wrongThread);};
                Model.RootNodes.CollectionChanged+=(_,_)=>{if(Environment.CurrentManagedThreadId!=Ui.Id)Interlocked.Increment(ref wrongThread);};
                Model.VisibleItems.CollectionChanged+=(_,_)=>{if(Environment.CurrentManagedThreadId!=Ui.Id)Interlocked.Increment(ref wrongThread);};
            }
            internal void CheckAffinity(){Check(wrongThread==0,"View-model state changed outside its injected dispatcher.");foreach(var controller in controllers)Check(controller.WorkerThreads.All(id=>id!=Ui.Id),"A synchronous controller call blocked the UI dispatcher.");}
            public async ValueTask DisposeAsync(){foreach(var controller in controllers)controller.ReleaseAll();await Model.DisposeAsync();CheckAffinity();Ui.Dispose();}
        }
        private sealed class FakeController : IExplorerController
        {
            internal readonly List<Node> Roots=new();
            internal readonly Dictionary<Guid,List<Node>> Children=new();
            internal readonly ConcurrentDictionary<Guid,int> ChildCalls=new();
            internal readonly ConcurrentDictionary<Guid,Gate> Gates=new();
            internal readonly ConcurrentBag<int> WorkerThreads=new();
            internal readonly TaskCompletionSource Disposed=new(TaskCreationOptions.RunContinuationsAsynchronously);
            internal Gate? ExportGate,AfterFirst,LibraryDiscoveryGate,VersionsGate;
            internal readonly Dictionary<Guid,List<Node>> LibraryFiles=new();
            internal readonly Dictionary<Guid,RecoveryStatus> Outcomes=new();
            internal readonly Dictionary<Guid,List<Node>> Versions=new();
            internal Node? CapturedLibrary,CapturedVersion;
            internal int LibraryExports,FilesExports,ZipExports,VersionExports,VersionQueries;
            internal string? RequestedArchivePath;
            internal Exception? ChildError,ExportError,RootError,VersionsError;
            internal List<Node> Captured=new();
            internal int Disposals;
            internal bool DisposedWhileActive;
            private readonly string name;
            private int active;
            internal FakeController(string name){this.name=name;}
            public string SourceName {get{WorkerThreads.Add(Environment.CurrentManagedThreadId);return name;}}
            public List<Node> GetRootSites(){WorkerThreads.Add(Environment.CurrentManagedThreadId);if(RootError!=null)throw RootError;return new(Roots);}
            public List<Node> GetChildren(Node parent)
            {
                WorkerThreads.Add(Environment.CurrentManagedThreadId);Interlocked.Increment(ref active);
                try{ChildCalls.AddOrUpdate(parent.Id,1,(_,old)=>old+1);if(Gates.TryGetValue(parent.Id,out var gate))gate.Wait();if(ChildError!=null)throw ChildError;return Children.TryGetValue(parent.Id,out var children)?new(children):new();}
                finally{Interlocked.Decrement(ref active);}
            }
            public List<Node> GetFileVersions(Node document)
            {
                WorkerThreads.Add(Environment.CurrentManagedThreadId);Interlocked.Increment(ref active);
                try
                {
                    VersionQueries++;VersionsGate?.Wait();if(VersionsError!=null)throw VersionsError;
                    return Versions.TryGetValue(document.Id,out var versions)?versions.Select(CopyVersion).ToList():new();
                }
                finally{Interlocked.Decrement(ref active);}
            }
            public DesktopExportSummary ExportVersion(Node version,string directory,CancellationToken token,Action<DesktopExportProgress> progress)
            {
                WorkerThreads.Add(Environment.CurrentManagedThreadId);Interlocked.Increment(ref active);
                try
                {
                    VersionExports++;CapturedVersion=CopyVersion(version);ExportGate?.Wait();if(ExportError!=null)throw ExportError;
                    var result=new DesktopExportSummary{Total=1,Directory=directory,ReportPath=Path.Combine(directory,"version-report.csv")};
                    if(token.IsCancellationRequested){result.Cancelled=true;return result;}
                    result.Success=1;result.Entries.Add(new DesktopExportEntry{Document=CapturedVersion,Status=RecoveryStatus.Success,Path=CapturedVersion.Name,Bytes=CapturedVersion.Size});
                    progress(new DesktopExportProgress{Document=CapturedVersion,Completed=1,Total=1,Status=RecoveryStatus.Success,ExportPath=CapturedVersion.Name});return result;
                }
                finally{Interlocked.Decrement(ref active);}
            }
            public DesktopExportSummary ExportFiles(IEnumerable<Node> documents,string directory,CancellationToken token,Action<DesktopExportProgress> progress)
            {
                WorkerThreads.Add(Environment.CurrentManagedThreadId);Interlocked.Increment(ref active);
                try
                {
                    FilesExports++;Captured=documents.ToList();ExportGate?.Wait();if(ExportError!=null)throw ExportError;
                    var result=new DesktopExportSummary{Total=Captured.Count,Directory=directory,ReportPath=Path.Combine(directory,"report.csv")};
                    foreach(Node document in Captured)
                    {
                        if(token.IsCancellationRequested){result.Cancelled=true;break;}
                        RecoveryStatus outcome=Outcomes.TryGetValue(document.Id,out RecoveryStatus configured)?configured:RecoveryStatus.Success;
                        if(outcome==RecoveryStatus.Success)result.Success++;
                        else if(outcome==RecoveryStatus.Unsupported || outcome==RecoveryStatus.Unavailable)result.Skipped++;
                        else result.Failed++;
                        result.Entries.Add(new DesktopExportEntry{Document=document,Status=outcome,Path=outcome==RecoveryStatus.Success?document.Name:String.Empty,Bytes=outcome==RecoveryStatus.Success?document.Size:0,Message="Configured recovery outcome."});
                        progress(new DesktopExportProgress{Document=document,Completed=result.Entries.Count,Total=result.Total,Status=outcome,ExportPath=document.Name});
                        if(result.Entries.Count==1)AfterFirst?.Wait();
                    }
                    return result;
                }
                finally{Interlocked.Decrement(ref active);}
            }
            public DesktopExportSummary ExportFilesAsZip(IEnumerable<Node> documents,string archivePath,CancellationToken token,Action<DesktopExportProgress> progress)
            {
                WorkerThreads.Add(Environment.CurrentManagedThreadId);Interlocked.Increment(ref active);
                try
                {
                    ZipExports++;RequestedArchivePath=archivePath;Captured=documents.ToList();ExportGate?.Wait();if(ExportError!=null)throw ExportError;
                    var result=new DesktopExportSummary{Total=Captured.Count,Directory=Path.GetDirectoryName(archivePath) ?? String.Empty,ReportPath=String.Empty,ArchivePath=String.Empty,ArchiveReportEntry=String.Empty};
                    foreach(Node document in Captured)
                    {
                        if(token.IsCancellationRequested){result.Cancelled=true;break;}
                        result.Success++;result.Entries.Add(new DesktopExportEntry{Document=document,Status=RecoveryStatus.Success,Path=document.Path,Bytes=document.Size});
                        progress(new DesktopExportProgress{Document=document,Completed=result.Success,Total=result.Total,Status=RecoveryStatus.Success,ExportPath=document.Path});
                        if(result.Success==1)AfterFirst?.Wait();
                    }
                    if(result.Success>0){result.ArchivePath=archivePath;result.ArchiveReportEntry="export-report.csv";}
                    return result;
                }
                finally{Interlocked.Decrement(ref active);}
            }
            public DesktopExportSummary ExportLibrary(Node library,string directory,CancellationToken token,Action<DesktopExportProgress> progress)
            {
                WorkerThreads.Add(Environment.CurrentManagedThreadId);Interlocked.Increment(ref active);
                try
                {
                    CapturedLibrary=library;LibraryExports++;LibraryDiscoveryGate?.Wait();
                    if(token.IsCancellationRequested)return new DesktopExportSummary{Directory=directory,Cancelled=true};
                    if(ExportError!=null)throw ExportError;
                    Captured=LibraryFiles.TryGetValue(library.Id,out var documents)?new(documents):new();
                    var result=new DesktopExportSummary{Total=Captured.Count,Directory=directory,ReportPath=Path.Combine(directory,"report.csv")};
                    foreach(Node document in Captured)
                    {
                        if(token.IsCancellationRequested){result.Cancelled=true;break;}
                        RecoveryStatus outcome=Outcomes.TryGetValue(document.Id,out RecoveryStatus configured)?configured:RecoveryStatus.Success;
                        if(outcome==RecoveryStatus.Success)result.Success++;
                        else if(outcome==RecoveryStatus.Unsupported || outcome==RecoveryStatus.Unavailable)result.Skipped++;
                        else result.Failed++;
                        result.Entries.Add(new DesktopExportEntry{Document=document,Status=outcome,Path=outcome==RecoveryStatus.Success?document.Name:String.Empty,Bytes=outcome==RecoveryStatus.Success?document.Size:0,Message="Configured recovery outcome."});
                        progress(new DesktopExportProgress{Document=document,Completed=result.Entries.Count,Total=result.Total,Status=outcome,ExportPath=document.Name});
                        if(result.Entries.Count==1)AfterFirst?.Wait();
                    }
                    return result;
                }
                finally{Interlocked.Decrement(ref active);}
            }
            internal void ReleaseAll(){foreach(Gate gate in Gates.Values)gate.Release();ExportGate?.Release();AfterFirst?.Release();LibraryDiscoveryGate?.Release();VersionsGate?.Release();}
            public void Dispose(){WorkerThreads.Add(Environment.CurrentManagedThreadId);if(Volatile.Read(ref active)>0)DisposedWhileActive=true;Interlocked.Increment(ref Disposals);Disposed.TrySetResult();}
        }
    }
}

