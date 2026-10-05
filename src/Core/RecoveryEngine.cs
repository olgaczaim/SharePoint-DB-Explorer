using System;
using System.Collections.Generic;
using System.IO;

namespace SharePointExplorer
{
    public sealed class RecoveryEngine
    {
        private readonly ISharePointCatalog catalog;
        private readonly IDocumentChunkStore store;
        private readonly DocumentDecoderRegistry registry;
        private bool closed;
        public RecoveryEngine(ISharePointCatalog catalog,IDocumentChunkStore store,DocumentDecoderRegistry registry)
        {
            if(catalog==null) throw new ArgumentNullException("catalog");
            if(store==null) throw new ArgumentNullException("store");
            if(registry==null) throw new ArgumentNullException("registry");
            this.catalog=catalog; this.store=store; this.registry=registry;
        }
        // Metadata eligibility does not fetch content or replace recovery validation.
        public bool CanExport(Node document)
        {
            EnsureOpen();
            if(document==null) throw new ArgumentNullException("document");
            return document.Kind==NodeKind.File && document.HasStream!=false && registry.Supports(document.StreamSchema);
        }
        public PreparedDocument Prepare(Node selected)
        {
            ValidateSelection(selected);
            return PrepareResolved(selected,catalog.GetFile(selected.SiteId,selected.Id),null);
        }
        public PreparedDocument PrepareLibraryFile(Node selected,Node library)
        {
            ValidateSelection(selected);
            if(library==null) throw new ArgumentNullException("library");
            if(library.Kind!=NodeKind.Library || library.IsDeleted || library.HistoryVersion!=0 || library.SiteId==Guid.Empty ||
                library.WebId==Guid.Empty || library.ListId==Guid.Empty || library.Id==Guid.Empty)
                throw new ArgumentException("Select a current document library with its complete source scope.","library");
            var scope=new Node { Kind=library.Kind,SiteId=library.SiteId,WebId=library.WebId,ListId=library.ListId,Id=library.Id };
            if(selected.SiteId!=scope.SiteId) throw new InvalidDataException("The selected file belongs to another library scope.");
            ISharePointLibraryCatalog libraryCatalog=catalog as ISharePointLibraryCatalog;
            if(libraryCatalog==null) throw new NotSupportedException("This source cannot recover a file within a document library scope.");
            var query=new Node { Kind=scope.Kind,SiteId=scope.SiteId,WebId=scope.WebId,ListId=scope.ListId,Id=scope.Id };
            return PrepareResolved(selected,libraryCatalog.GetCurrentLibraryFile(query,selected.Id),scope);
        }
        public PreparedDocument PrepareDeleted(Node selected)
        {
            EnsureOpen();DeletedIdentity.Validate(selected,true);
            Node identity=DeletedIdentity.Copy(selected);
            ISharePointDeletedCatalog deleted=catalog as ISharePointDeletedCatalog;
            if(deleted==null)throw new NotSupportedException("This source does not expose retained deleted document recovery.");
            Node document=deleted.GetDeletedFile(DeletedIdentity.Copy(identity));
            if(document==null)throw new ContentUnavailableException("The selected deleted document or version is no longer retained with this identity.");
            if(!DeletedIdentity.Same(document,identity))throw new InvalidDataException("The source returned another deleted document, transaction or version.");
            Node snapshot=DeletedIdentity.Copy(document);
            PreparedDocument prepared=PrepareContent(snapshot);
            // Stream maps do not contain a deletion transaction. Revalidate the
            // immutable metadata after reading them so a restore or purge cannot
            // silently substitute a current state for the selected recycled file.
            Node confirmed=deleted.GetDeletedFile(DeletedIdentity.Copy(identity));
            if(!DeletedIdentity.Same(confirmed,snapshot) || confirmed.Size!=snapshot.Size || confirmed.StreamSchema!=snapshot.StreamSchema ||
                confirmed.HasStream!=snapshot.HasStream || confirmed.ParentId!=snapshot.ParentId || confirmed.Path!=snapshot.Path || confirmed.Name!=snapshot.Name)
                throw new ContentUnavailableException("The deleted document changed or was purged while its content was being read. Refresh the deleted items.");
            return prepared;
        }
        public PreparedDocument PrepareVersion(Node selectedVersion)
        {
            EnsureOpen();
            if(selectedVersion==null) throw new ArgumentNullException("selectedVersion");
            if(selectedVersion.IsDeleted)throw new NotSupportedException("Deleted versions require the deleted recovery operation.");
            if(selectedVersion.Kind!=NodeKind.File || selectedVersion.SiteId==Guid.Empty || selectedVersion.Id==Guid.Empty ||
                selectedVersion.HistoryVersion<0 || selectedVersion.UiVersion<0 || selectedVersion.InternalVersion<0 ||
                (selectedVersion.HistoryVersion>0 && selectedVersion.HistoryVersion!=selectedVersion.UiVersion))
                throw new ArgumentException("Select a complete document version identity.","selectedVersion");
            ISharePointVersionCatalog versions=catalog as ISharePointVersionCatalog;
            if(versions==null) throw new NotSupportedException("This content source does not support document version recovery.");
            Node selected=CopyVersionIdentity(selectedVersion);
            Node document=versions.GetFileVersion(CopyVersionIdentity(selected));
            if(document==null) throw new ContentUnavailableException("The selected document version is no longer available in this source.");
            if(document.Kind!=NodeKind.File || document.IsDeleted || document.SiteId!=selected.SiteId || document.Id!=selected.Id ||
                document.WebId!=selected.WebId || document.ListId!=selected.ListId ||
                document.HistoryVersion!=selected.HistoryVersion || document.UiVersion!=selected.UiVersion ||
                document.Level!=selected.Level || document.InternalVersion!=selected.InternalVersion)
                throw new InvalidDataException("The catalog returned a different document version.");
            return PrepareContent(document);
        }
        public PreparedDocument PrepareItemAttachment(Node selected,Node ownerItem)
        {
            ValidateSelection(selected);
            if(selected.Id==Guid.Empty) throw new ArgumentException("Select an attachment with its document identity.","selected");
            // Keep the validated identities separate from the mutable catalog request.
            var file=new Node {Kind=selected.Kind,SiteId=selected.SiteId,WebId=selected.WebId,ListId=selected.ListId,Id=selected.Id,
                AttachmentOwnerId=selected.AttachmentOwnerId,ListItemId=selected.ListItemId,ItemUniqueId=selected.ItemUniqueId};
            if(ownerItem==null) throw new ArgumentNullException("ownerItem");
            if(ownerItem.Kind!=NodeKind.ListItem || ownerItem.IsDeleted || ownerItem.HistoryVersion!=0 || ownerItem.SiteId==Guid.Empty ||
                ownerItem.WebId==Guid.Empty || ownerItem.ListId==Guid.Empty || ownerItem.Id==Guid.Empty ||
                !ownerItem.ListItemId.HasValue || ownerItem.ListItemId.Value<=0 || !ownerItem.ItemUniqueId.HasValue ||
                ownerItem.ItemUniqueId.Value==Guid.Empty || ownerItem.ListBaseType==1)
                throw new ArgumentException("Select a current ordinary list item with its complete source identity.","ownerItem");
            var owner=new Node {Kind=ownerItem.Kind,SiteId=ownerItem.SiteId,WebId=ownerItem.WebId,ListId=ownerItem.ListId,
                Id=ownerItem.Id,ListItemId=ownerItem.ListItemId,ItemUniqueId=ownerItem.ItemUniqueId,ListBaseType=ownerItem.ListBaseType,
                ParentId=ownerItem.ParentId,HistoryVersion=0};
            if(file.SiteId!=owner.SiteId || file.WebId!=owner.WebId || file.ListId!=owner.ListId ||
                file.AttachmentOwnerId!=owner.Id || file.ListItemId!=owner.ListItemId || file.ItemUniqueId!=owner.ItemUniqueId)
                throw new InvalidDataException("The selected attachment belongs to a different list item.");
            ISharePointAttachmentCatalog attachments=catalog as ISharePointAttachmentCatalog;
            if(attachments==null) throw new NotSupportedException("This source does not support list attachment recovery.");
            var query=new Node {Kind=owner.Kind,SiteId=owner.SiteId,WebId=owner.WebId,ListId=owner.ListId,Id=owner.Id,
                ListItemId=owner.ListItemId,ItemUniqueId=owner.ItemUniqueId,ListBaseType=owner.ListBaseType,ParentId=owner.ParentId,HistoryVersion=0};
            Node document=attachments.GetCurrentItemAttachment(query,file.Id);
            if(document==null) throw new ContentUnavailableException("The selected attachment is no longer present on this list item.");
            if(document.Kind!=NodeKind.File || document.IsDeleted || document.HistoryVersion!=0 || document.SiteId!=file.SiteId || document.Id!=file.Id ||
                document.WebId!=owner.WebId || document.ListId!=owner.ListId || document.AttachmentOwnerId!=owner.Id ||
                document.ListItemId!=owner.ListItemId || document.ItemUniqueId!=owner.ItemUniqueId)
                throw new InvalidDataException("The source returned an attachment from a different list item.");
            return PrepareContent(document);
        }
        private static Node CopyVersionIdentity(Node selected)
        {
            return new Node {Kind=selected.Kind,SiteId=selected.SiteId,Id=selected.Id,WebId=selected.WebId,ListId=selected.ListId,
                HistoryVersion=selected.HistoryVersion,UiVersion=selected.UiVersion,Level=selected.Level,InternalVersion=selected.InternalVersion};
        }
        private void ValidateSelection(Node selected)
        {
            EnsureOpen();
            if(selected==null) throw new ArgumentNullException("selected");
            if(selected.Kind!=NodeKind.File) throw new ArgumentException("Select a file to recover.","selected");
            if(selected.IsDeleted)throw new NotSupportedException("Deleted files require the deleted recovery operation.");
            if(selected.HistoryVersion!=0) throw new NotSupportedException("Historical file versions require the version recovery operation. Select the current document for ordinary export.");
        }
        private PreparedDocument PrepareResolved(Node selected,Node document,Node library)
        {
            if(document==null) throw new ContentUnavailableException("The selected file is no longer present in the content source.");
            if(document.Kind!=NodeKind.File || document.IsDeleted || document.HistoryVersion!=0 || document.SiteId!=selected.SiteId || document.Id!=selected.Id) throw new InvalidDataException("The catalog returned a different document scope.");
            if(library!=null && (document.SiteId!=library.SiteId || document.WebId!=library.WebId || document.ListId!=library.ListId))
                throw new InvalidDataException("The source returned a file outside the selected document library.");
            return PrepareContent(document);
        }
        private PreparedDocument PrepareContent(Node document)
        {
            if(document.Size<0) throw new InvalidDataException("The document has an invalid declared size.");
            IDocumentDecoder decoder=registry.Resolve(document.StreamSchema);
            IList<StoredChunk> chunks=store.ReadChunks(document);
            if(chunks==null) throw new InvalidDataException("The content source returned no storage collection.");
            if(document.Size>0 && chunks.Count==0) throw new ContentUnavailableException("The document's bytes are absent from this content source.");
            IReconstructedDocumentDecoder reconstructed=decoder as IReconstructedDocumentDecoder;
            ReconstructedContent content=null;
            if(reconstructed!=null)
                using(RecoveryLog.ForDocument(document)) content=reconstructed.Reconstruct(chunks,document.Size);
            return new PreparedDocument(document,chunks,decoder,EnsureOpen,content);
        }
        public RecoveryResult Recover(Node selected,Stream output) { return Prepare(selected).Recover(output); }
        public RecoveryResult Verify(Node selected) { return Prepare(selected).Recover(Stream.Null); }
        internal void Close() { closed=true; }
        private void EnsureOpen() { if(closed) throw new ObjectDisposedException("RecoverySession"); }
    }
    public sealed class PreparedDocument
    {
        private readonly Node document;
        private readonly IList<StoredChunk> chunks;
        private readonly IDocumentDecoder decoder;
        private readonly Action ensureOpen;
        private readonly ReconstructedContent reconstructed;
        internal PreparedDocument(Node document,IList<StoredChunk> chunks,IDocumentDecoder decoder,Action ensureOpen,ReconstructedContent reconstructed=null)
        {
            this.document=Copy(document);
            this.chunks=new List<StoredChunk>(chunks).AsReadOnly();
            this.decoder=decoder;
            this.ensureOpen=ensureOpen;
            this.reconstructed=reconstructed;
        }
        public long OutputSize { get { ensureOpen();return reconstructed==null?document.Size:reconstructed.Length; } }
        public Node Document { get { ensureOpen(); return Copy(document); } }
        public RecoveryResult Recover(Stream output)
        {
            ensureOpen();
            using(var counted=new HashingWriteStream(output,OutputSize,true))
            {
                if(reconstructed==null)decoder.Write(chunks,document.Size,counted);
                else reconstructed.WriteTo(counted);
                string digest=counted.Complete();
                return new RecoveryResult { Document=Copy(document),Bytes=counted.BytesWritten,Sha256=digest,Decoder=decoder.Name };
            }
        }
        private static Node Copy(Node source)
        {
            return new Node {
                Kind=source.Kind,SiteId=source.SiteId,WebId=source.WebId,Id=source.Id,ListId=source.ListId,
                Name=source.Name,Path=source.Path,Size=source.Size,StreamSchema=source.StreamSchema,Level=source.Level,
                InternalVersion=source.InternalVersion,HistoryVersion=source.HistoryVersion,UiVersion=source.UiVersion,
                ParentId=source.ParentId,HasStream=source.HasStream,Modified=source.Modified,
                ListBaseType=source.ListBaseType,ListItemId=source.ListItemId,ItemUniqueId=source.ItemUniqueId,AttachmentOwnerId=source.AttachmentOwnerId,
                Created=source.Created,Title=source.Title,HasAttachments=source.HasAttachments,
                DeletionTransactionId=source.DeletionTransactionId,DeletedAt=source.DeletedAt
            };
        }
    }
}