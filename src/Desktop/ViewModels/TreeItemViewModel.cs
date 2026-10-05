#nullable enable
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace SharePointExplorer.Desktop.ViewModels
{
    public sealed class TreeItemViewModel : ObservableViewModel
    {
        private readonly Node node;
        private bool isLoaded, isLoading;
        private string errorText = "";
        internal TreeItemViewModel(Node node, TreeItemViewModel? parent, Guid owner, long sourceVersion)
        {
            this.node = ExplorerNode.Copy(node);
            Parent = parent;
            Owner = owner;
            SourceVersion = sourceVersion;
        }
        public Node Node => ExplorerNode.Copy(node);
        public string Name => node.Name ?? "";
        public string Path => node.Path ?? "";
        public string IconKey => ExplorerNode.Icon(node);
        public TreeItemViewModel? Parent { get; }
        public ObservableCollection<TreeItemViewModel> Children { get; } = new();
        public bool HasUnrealizedChildren => !IsLoaded && node.IsContainer;
        public bool IsLoaded
        {
            get => isLoaded;
            internal set { if (Set(ref isLoaded, value)) Notify(nameof(HasUnrealizedChildren)); }
        }
        public bool IsLoading {get => isLoading; internal set => Set(ref isLoading, value);}
        public string ErrorText {get => errorText; internal set => Set(ref errorText, value);}
        internal bool IsAttached { get; set; } = true;
        internal Guid Owner { get; }
        internal long SourceVersion { get; }
        internal List<Node>? LoadedItems { get; set; }
        internal Task<List<Node>>? LoadTask { get; set; }
    }
}
