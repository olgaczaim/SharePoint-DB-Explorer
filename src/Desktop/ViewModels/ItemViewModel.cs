#nullable enable
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace SharePointExplorer.Desktop.ViewModels
{
    public abstract class ObservableViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        protected bool Set<T>(ref T field, T value, [CallerMemberName] string? property = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            Notify(property);
            return true;
        }
        protected void Notify([CallerMemberName] string? property = null)
        { PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property)); }
    }

    internal static class ExplorerNode
    {
        internal static Node Copy(Node node)
        {
            return new Node {
                Kind=node.Kind,SiteId=node.SiteId,WebId=node.WebId,Id=node.Id,ListId=node.ListId,
                Name=node.Name,Path=node.Path,Size=node.Size,StreamSchema=node.StreamSchema,Level=node.Level,
                InternalVersion=node.InternalVersion,HistoryVersion=node.HistoryVersion,UiVersion=node.UiVersion,
                ParentId=node.ParentId,HasStream=node.HasStream,Modified=node.Modified,ListBaseType=node.ListBaseType,
                ListItemId=node.ListItemId,ItemUniqueId=node.ItemUniqueId,Title=node.Title,Created=node.Created,
                HasAttachments=node.HasAttachments,AttachmentOwnerId=node.AttachmentOwnerId,DeletionTransactionId=node.DeletionTransactionId,DeletedAt=node.DeletedAt
            };
        }
        internal static string Identity(Node node)
        { return node.SiteId.ToString("N") + ":" + node.Kind + ":" + node.WebId.ToString("N") + ":" + node.Id.ToString("N") + ":" + node.ListId.ToString("N") + ":" + node.Path + ":" + node.DeletionTransactionId; }
        internal static string Icon(Node node)
        {
            return node.Kind switch {
                NodeKind.DeletedItems => "Deleted items",
                NodeKind.Site => "Site",
                NodeKind.Library => "Library",
                NodeKind.List => "List",
                NodeKind.Folder => "Folder",
                NodeKind.ListItem => "ListItem",
                _ => "File"
            };
        }
        internal static string Kind(Node node)
        {
            return node.Kind switch {
                NodeKind.DeletedItems => "Deleted items",
                NodeKind.Site => "Site",
                NodeKind.Library => "Document library",
                NodeKind.List => "List",
                NodeKind.Folder => "Folder",
                NodeKind.ListItem => "List item",
                _ => "File"
            };
        }
        internal static string Size(long bytes)
        {
            if (bytes < 0) return "";
            string[] units = {"B", "KiB", "MiB", "GiB", "TiB"};
            double value = bytes; int unit = 0;
            while (value >= 1024 && unit < units.Length - 1) {value /= 1024; unit++;}
            return value.ToString(unit == 0 ? "0" : "0.##", CultureInfo.CurrentCulture) + " " + units[unit];
        }
    }

    public sealed class ItemViewModel : ObservableViewModel
    {
        private static readonly DocumentDecoderRegistry exportDecoders=DocumentDecoderRegistry.CreateDefault();
        private readonly Node node;
        private readonly bool canCheck;
        private bool isChecked;
        private string exportStatus = "";
        public ItemViewModel(Node node, TreeItemViewModel? treeNode = null)
        {
            ArgumentNullException.ThrowIfNull(node);
            this.node = ExplorerNode.Copy(node);
            TreeNode = treeNode;
            canCheck = node.Kind == NodeKind.File && node.HistoryVersion == 0 && node.HasStream != false && exportDecoders.Supports(node.StreamSchema);
            exportStatus = node.IsContainer ? "Open to browse" : node.Kind==NodeKind.ListItem && node.HasAttachments==true ? "Has attachments" : canCheck ? "Ready" : "Browse only";
        }
        public Node Node => ExplorerNode.Copy(node);
        public TreeItemViewModel? TreeNode { get; }
        public string Name => node.Name ?? "";
        public string Title => node.Title ?? "";
        public string Path => node.Path ?? "";
        public string IconKey => ExplorerNode.Icon(node);
        public string KindText => (node.IsDeleted ? "Deleted " : "") + ExplorerNode.Kind(node);
        public bool IsContainer => node.IsContainer;
        public bool CanCheck => canCheck;
        public string SizeText => node.Kind == NodeKind.File ? ExplorerNode.Size(node.Size) : "";
        public string ModifiedText => node.Modified == DateTime.MinValue ? "" : node.Modified.ToString("g", CultureInfo.CurrentCulture);
        public string ItemIdText => node.ListItemId?.ToString(CultureInfo.CurrentCulture) ?? "";
        public string HasAttachmentsText => node.HasAttachments.HasValue ? (node.HasAttachments.Value ? "Yes" : "No") : "";
        public bool IsChecked
        {
            get => isChecked;
            set => Set(ref isChecked, CanCheck && value);
        }
        public string ExportStatus
        {
            get => exportStatus;
            internal set => Set(ref exportStatus, value);
        }
    }
}
