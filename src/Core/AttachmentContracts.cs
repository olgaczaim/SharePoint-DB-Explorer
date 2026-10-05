using System;
using System.Collections.Generic;

namespace SharePointExplorer
{
    // Attachment discovery and reload always bind the current ordinary-list item
    // identity. File nodes preserve their real parent folder and carry the owner.
    public interface ISharePointAttachmentCatalog
    {
        List<Node> GetItemAttachments(Node listItem);
        IEnumerable<Node> EnumerateCurrentListAttachments(Node list);
        Node GetCurrentItemAttachment(Node listItem,Guid attachmentDocId);
    }
}