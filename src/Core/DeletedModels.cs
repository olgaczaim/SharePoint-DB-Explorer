using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;

namespace SharePointExplorer
{
    public interface ISharePointDeletedCatalog
    {
        List<Node> GetDeletedItems(Node site);
        List<Node> GetDeletedChildren(Node container);
        Node GetDeletedFile(Node selected);
        List<Node> GetDeletedFileVersions(Node selected);
        DeletedListItemSnapshot GetDeletedListItem(Node selected);
    }

    public sealed class DeletedListItemValue
    {
        public int RowOrdinal { get; private set; }
        public string Column { get; private set; }
        public string ValueType { get; private set; }
        public string Value { get; private set; }
        public bool IsNull { get; private set; }
        public DeletedListItemValue(int rowOrdinal,string column,string valueType,string value,bool isNull)
        { RowOrdinal=rowOrdinal;Column=column;ValueType=valueType;Value=value;IsNull=isNull; }
    }
    public sealed class DeletedListItemSnapshot
    {
        private readonly Node item;
        public Node Item { get { return DeletedIdentity.Copy(item); } }
        public string StoredFieldsXml { get; private set; }
        public ReadOnlyCollection<DeletedListItemValue> Values { get; private set; }
        public DeletedListItemSnapshot(Node item,string storedFieldsXml,IEnumerable<DeletedListItemValue> values)
        {
            DeletedIdentity.Validate(item,false);
            if(item.Kind!=NodeKind.ListItem)throw new ArgumentException("Select a deleted ordinary list item.","item");
            if(values==null)throw new ArgumentNullException("values");
            this.item=DeletedIdentity.Copy(item);StoredFieldsXml=storedFieldsXml ?? String.Empty;
            Values=new List<DeletedListItemValue>(values).AsReadOnly();
        }
    }

    // A recycle identity is a binary transaction, not a GUID whose byte order
    // can be reformatted. Keeping exact hex also makes reports reproducible.
    internal static class DeletedIdentity
    {
        internal static byte[] Parse(string identity)
        {
            if(identity==null || identity.Length!=32)throw new ArgumentException("A deleted item requires its exact 16-byte deletion transaction identity.");
            var bytes=new byte[16];
            for(int index=0;index<bytes.Length;index++)
            {
                byte value;
                if(!Byte.TryParse(identity.Substring(index*2,2),NumberStyles.AllowHexSpecifier,CultureInfo.InvariantCulture,out value))
                    throw new ArgumentException("The deletion transaction identity contains invalid hexadecimal data.");
                bytes[index]=value;
            }
            return bytes;
        }
        internal static string Format(byte[] value)
        {
            if(value==null || value.Length!=16)throw new InvalidDataException("The source returned an incomplete deletion transaction identity.");
            return BitConverter.ToString(value).Replace("-",String.Empty);
        }
        internal static void Validate(Node item,bool fileOnly)
        {
            if(item==null)throw new ArgumentNullException("selected");
            Parse(item.DeletionTransactionId);
            if(item.SiteId==Guid.Empty || item.WebId==Guid.Empty || item.Id==Guid.Empty ||
                item.InternalVersion<0 || item.UiVersion<0 || item.HistoryVersion<0 ||
                (item.HistoryVersion>0 && item.HistoryVersion!=item.UiVersion) ||
                (fileOnly && item.Kind!=NodeKind.File) ||
                (!fileOnly && item.Kind!=NodeKind.File && item.Kind!=NodeKind.Folder && item.Kind!=NodeKind.Library && item.Kind!=NodeKind.List && item.Kind!=NodeKind.ListItem))
                throw new ArgumentException("Select a retained deleted object with its complete source and version identity.","selected");
        }
        internal static bool Same(Node left,Node right)
        {
            return left!=null && right!=null && left.Kind==right.Kind && left.SiteId==right.SiteId && left.WebId==right.WebId &&
                left.ListId==right.ListId && left.Id==right.Id && left.HistoryVersion==right.HistoryVersion &&
                left.UiVersion==right.UiVersion && left.Level==right.Level && left.InternalVersion==right.InternalVersion &&
                String.Equals(left.DeletionTransactionId,right.DeletionTransactionId,StringComparison.OrdinalIgnoreCase) &&
                left.ListItemId==right.ListItemId && left.ItemUniqueId==right.ItemUniqueId && left.AttachmentOwnerId==right.AttachmentOwnerId;
        }
        internal static Node Copy(Node source)
        {
            return new Node {Kind=source.Kind,SiteId=source.SiteId,WebId=source.WebId,ListId=source.ListId,Id=source.Id,
                ParentId=source.ParentId,Name=source.Name,Path=source.Path,Size=source.Size,StreamSchema=source.StreamSchema,
                Level=source.Level,InternalVersion=source.InternalVersion,UiVersion=source.UiVersion,HistoryVersion=source.HistoryVersion,
                HasStream=source.HasStream,Modified=source.Modified,Created=source.Created,ListBaseType=source.ListBaseType,
                ListItemId=source.ListItemId,ItemUniqueId=source.ItemUniqueId,Title=source.Title,HasAttachments=source.HasAttachments,
                AttachmentOwnerId=source.AttachmentOwnerId,DeletionTransactionId=source.DeletionTransactionId,DeletedAt=source.DeletedAt};
        }
    }
}