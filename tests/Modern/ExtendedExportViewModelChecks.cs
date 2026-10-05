#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using SharePointExplorer.Desktop;
using SharePointExplorer.Desktop.ViewModels;

namespace SharePointExplorer.Tests
{
    internal static class ExtendedExportViewModelChecks
    {
        internal static async Task RunAsync()
        {
            var backend=new Source();
            await using var model=new ExplorerViewModel();
            Check(await model.SetSourceAsync(backend),"Source did not open.");
            Check(!model.CanExportPackage && !model.CanExportListAttachments,"Site roots enabled content export.");
            await model.NavigateAsync(model.RootNodes.Single());
            TreeItemViewModel list=model.CurrentFolder!.Children.Single(n=>n.Node.Kind==NodeKind.List);
            TreeItemViewModel library=model.CurrentFolder.Children.Single(n=>n.Node.Kind==NodeKind.Library);
            await model.NavigateAsync(list);
            Check(model.CanExportPackage && model.CanExportListAttachments && model.CurrentContentList==list,"Ordinary list export scope missing.");
            Node item=model.VisibleItems.Single(n=>n.Node.Kind==NodeKind.ListItem).Node;
            Check((await model.ExportAttachmentsAsync(item,"attachment-output"))?.Success==1 &&
                backend.Captured!.Id==item.Id && backend.Captured.ItemUniqueId==item.ItemUniqueId,"Item attachment export lost ownership.");
            Check(model.SelectedCount==0 && !model.IsExporting,"Attachment export changed file checks or left busy state.");
            Node forged=ExplorerNode.Copy(item);forged.ItemUniqueId=Guid.NewGuid();
            Check(await model.ExportAttachmentsAsync(forged,"attachment-output")==null && backend.AttachmentCalls==1,"Forged item identity reached backend.");
            Check((await model.ExportAttachmentsAsync(list.Node,"list-output"))?.Success==1 && backend.Captured!.Kind==NodeKind.List,"Whole-list attachments did not use list scope.");
            var package=await model.ExportPackageAsync(list.Node,"package-output");
            Check(package?.PackageItemCount==2 && model.Status.StartsWith("XML package exported.",StringComparison.Ordinal),"Published package status is misleading.");
            TreeItemViewModel listFolder=list.Children.Single();
            await model.NavigateAsync(listFolder);
            Check(model.CurrentContentList==list && model.CanExportListAttachments,"Nested list folder lost ancestor scope.");
            Check((await model.ExportPackageAsync(list.Node,"package-output"))?.Success==1,"Nested list did not allow whole-list package.");
            Check(await model.ExportAttachmentsAsync(item,"attachment-output")==null,"An item from another folder reached attachment export.");
            await model.NavigateAsync(library);
            TreeItemViewModel folder=library.Children.Single();
            await model.NavigateAsync(folder);
            Check(model.CurrentContentList==library && model.CanExportPackage && !model.CanExportListAttachments,"Library folder enabled ordinary-list attachment export.");
            Check(await model.ExportAttachmentsAsync(library.Node,"attachment-output")==null,"Library was accepted as ordinary attachment list.");
            Check((await model.ExportPackageAsync(library.Node,"package-output"))?.Success==1 && backend.Captured!.Id==library.Node.Id,"Whole-library package used folder GUID.");
            await model.NavigateAsync(list);
            backend.Block=true;
            Node scope=list.Node;
            Task<DesktopExportSummary?> pending=model.ExportPackageAsync(scope,"package-output");
            await backend.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Check(model.IsExporting && !model.CanNavigate && !model.CanExportPackage,"Package export did not protect navigation.");
            scope.Id=Guid.NewGuid();scope.Name="mutated caller";
            model.CancelExport();backend.Release.Set();
            Check((await pending)?.Cancelled==true && backend.Captured!.Id==list.Node.Id && backend.Captured.Name==list.Name,
                "Cancellation or immutable package scope failed.");
            Check(!model.IsExporting && model.CanExportPackage,"Cancel did not restore controls.");
            Console.WriteLine("PASS extended MVVM attachment ownership, nested list/library scope, snapshot, stale item and cancellation checks");
        }
        private static void Check(bool value,string message){if(!value)throw new Exception(message);}
        private sealed class Source : IExplorerController,IExplorerExtendedExportController
        {
            private readonly Node site,list,library,listFolder,libraryFolder,item;
            public Node? Captured;
            public int AttachmentCalls;
            public bool Block;
            public readonly TaskCompletionSource Started=new(TaskCreationOptions.RunContinuationsAsynchronously);
            public readonly ManualResetEventSlim Release=new(false);
            public Source()
            {
                Guid siteId=Guid.NewGuid(),webId=Guid.NewGuid();
                Node Make(NodeKind kind,string name,Guid listId,Guid parentId)=>new(){Kind=kind,SiteId=siteId,WebId=webId,
                    Id=Guid.NewGuid(),ListId=listId,Name=name,Path=name,ParentId=parentId,HistoryVersion=0,ListBaseType=kind==NodeKind.Library?1:0};
                site=Make(NodeKind.Site,"Site",Guid.Empty,Guid.Empty);
                list=Make(NodeKind.List,"Tasks",Guid.NewGuid(),site.Id);
                library=Make(NodeKind.Library,"Documents",Guid.NewGuid(),site.Id);
                listFolder=Make(NodeKind.Folder,"Nested tasks",list.ListId,list.Id);
                libraryFolder=Make(NodeKind.Folder,"Nested documents",library.ListId,library.Id);libraryFolder.ListBaseType=1;
                item=Make(NodeKind.ListItem,"Task",list.ListId,list.Id);item.ListItemId=1;item.ItemUniqueId=Guid.NewGuid();item.HasAttachments=true;
            }
            public string SourceName=>"fixture";
            public List<Node> GetRootSites()=>new(){site};
            public List<Node> GetChildren(Node parent)=>parent.Id==site.Id ? new(){list,library} : parent.Id==list.Id ? new(){listFolder,item} : parent.Id==library.Id ? new(){libraryFolder} : new();
            public List<Node> GetFileVersions(Node document)=>new();
            public DesktopExportSummary ExportAttachments(Node scope,string directory,CancellationToken token,Action<DesktopExportProgress> progress)
            {
                AttachmentCalls++;Captured=ExplorerNode.Copy(scope);
                var summary=new DesktopExportSummary{Total=1,Success=1,Directory=directory};
                summary.Entries.Add(new(){Document=Captured,Status=RecoveryStatus.Success,Path=Path.Combine(directory,"attachment.txt")});
                progress(new(){Completed=1,Total=1,Status=RecoveryStatus.Success});return summary;
            }
            public DesktopExportSummary ExportMigrationPackage(Node scope,string directory,CancellationToken token,Action<DesktopExportProgress> progress)
            {
                Captured=ExplorerNode.Copy(scope);
                if(Block){Started.TrySetResult();if(!Release.Wait(TimeSpan.FromSeconds(8)))throw new TimeoutException();}
                if(token.IsCancellationRequested)return new(){Cancelled=true,Directory=directory};
                string output=Path.Combine(directory,scope.Name+" package");
                var summary=new DesktopExportSummary{Total=1,Success=1,Directory=directory,PackagePath=output,PackageItemCount=2};
                summary.Entries.Add(new(){Document=Captured,Status=RecoveryStatus.Success,Path=output});
                progress(new(){Completed=1,Total=1,Document=scope,Status=RecoveryStatus.Success});return summary;
            }
            public DesktopExportSummary ExportVersion(Node node,string dir,CancellationToken token,Action<DesktopExportProgress> progress)=>throw new NotSupportedException();
            public DesktopExportSummary ExportFiles(IEnumerable<Node> nodes,string dir,CancellationToken token,Action<DesktopExportProgress> progress)=>throw new NotSupportedException();
            public DesktopExportSummary ExportFilesAsZip(IEnumerable<Node> nodes,string dir,CancellationToken token,Action<DesktopExportProgress> progress)=>throw new NotSupportedException();
            public DesktopExportSummary ExportLibrary(Node node,string dir,CancellationToken token,Action<DesktopExportProgress> progress)=>throw new NotSupportedException();
            public void Dispose(){Release.Set();Release.Dispose();}
        }
    }
}