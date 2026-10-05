using System;

namespace SharePointExplorer
{
    public enum NodeKind
    {
        Site = 0,
        Library = 1,
        Folder = 2,
        File = 3,
        List = 4,
        ListItem = 5,
        DeletedItems = 6
    }

    public sealed class Node
    {
        public NodeKind Kind { get; set; }
        public Guid SiteId { get; set; }
        public Guid WebId { get; set; }
        public Guid Id { get; set; }
        public Guid ListId { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public long Size { get; set; }
        public byte StreamSchema { get; set; }
        public byte Level { get; set; }
        public int InternalVersion { get; set; }
        public int HistoryVersion { get; set; }
        public int UiVersion { get; set; }
        public Guid ParentId { get; set; }
        public bool? HasStream { get; set; }
        public DateTime Modified { get; set; }
        public int? ListBaseType { get; set; }
        public int? ListItemId { get; set; }
        public Guid? ItemUniqueId { get; set; }
        public string Title { get; set; }
        public DateTime? Created { get; set; }
        public bool? HasAttachments { get; set; }
        public Guid? AttachmentOwnerId { get; set; }
        public string DeletionTransactionId { get; set; }
        public DateTime? DeletedAt { get; set; }
        public bool IsDeleted { get { return !String.IsNullOrEmpty(DeletionTransactionId); } }
        public bool IsContainer { get { return Kind == NodeKind.Site || Kind == NodeKind.Library || Kind == NodeKind.Folder || Kind == NodeKind.List || Kind == NodeKind.DeletedItems || (Kind == NodeKind.ListItem && IsDeleted); } }
    }
}


