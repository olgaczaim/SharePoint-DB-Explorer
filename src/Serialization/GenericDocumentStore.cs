using System;
using System.Collections.Generic;
using System.IO;

namespace SharePointExplorer
{
    // Nodes and leaves have independent state. Sequence comes from the contained
    // blob, not SQL BSN. Fragment updates are applied in parse order.
    internal sealed class GenericDocumentStore
    {
        private sealed class TreeNode
        {
            internal ulong Sequence;
            internal IList<string> References;
        }
        private sealed class Leaf
        {
            internal ulong Sequence;
            internal ArraySegment<byte>? Data;
            internal List<Fragment> Fragments;
        }
        private sealed class BlobFallback
        {
            internal ulong Sequence;
            internal Func<IList<ArraySegment<byte>>> Read;
        }
        internal struct Fragment
        {
            internal ulong Start;
            internal ArraySegment<byte> Data;
        }
        private readonly Dictionary<string,TreeNode> nodes=new Dictionary<string,TreeNode>(StringComparer.Ordinal);
        private readonly Dictionary<string,Leaf> leaves=new Dictionary<string,Leaf>(StringComparer.Ordinal);
        // Resident object-data BLOB recovery is an additional format. A BLOB
        // reference never replaces an inline or partial leaf state.
        private readonly Dictionary<string,BlobFallback> blobs=new Dictionary<string,BlobFallback>(StringComparer.Ordinal);

        internal void AddInline(ulong sequence,string id,IList<string> references,ArraySegment<byte> data)
        {
            if(references.Count>0)
            {
                TreeNode node;
                if(!nodes.TryGetValue(id,out node) || sequence>node.Sequence)
                    nodes[id]=new TreeNode { Sequence=sequence,References=references };
                return;
            }
            Leaf leaf;
            if(!leaves.TryGetValue(id,out leaf))
            {
                leaf=new Leaf();
                leaves.Add(id,leaf);
            }
            else if(sequence<=leaf.Sequence) return;
            leaf.Sequence=sequence;
            leaf.Data=data;
            // Existing fragments remain the preferred leaf representation.
        }
        internal void AddFragment(ulong sequence,string id,ulong start,ArraySegment<byte> data)
        {
            Leaf leaf;
            if(!leaves.TryGetValue(id,out leaf))
            {
                leaf=new Leaf { Sequence=sequence };
                leaves.Add(id,leaf);
            }
            else if(sequence<leaf.Sequence) return;
            // Accepted fragments do not advance an existing threshold.
            if(leaf.Fragments==null) leaf.Fragments=new List<Fragment>();
            leaf.Fragments.Add(new Fragment { Start=start,Data=data });
            leaf.Data=null;
        }
        internal void AddBlobFallback(ulong sequence,string id,Func<IList<ArraySegment<byte>>> read)
        {
            BlobFallback blob;
            if(!blobs.TryGetValue(id,out blob) || sequence>blob.Sequence)
                blobs[id]=new BlobFallback { Sequence=sequence,Read=read };
        }
        internal IList<string> References(string id)
        {
            TreeNode node;
            return nodes.TryGetValue(id,out node) ? node.References : null;
        }
        internal IList<ArraySegment<byte>> Data(string id)
        {
            Leaf leaf;
            if(leaves.TryGetValue(id,out leaf))
            {
                if(leaf.Fragments!=null) return JoinFragments(leaf.Fragments);
                if(leaf.Data.HasValue) return new[] { leaf.Data.Value };
                throw new InvalidDataException("A stored file leaf has no data.");
            }
            BlobFallback blob;
            if(blobs.TryGetValue(id,out blob)) return blob.Read();
            throw new InvalidDataException("A file object is absent from the backup.");
        }
        internal static IList<ArraySegment<byte>> JoinFragments(IEnumerable<Fragment> fragments)
        {
            var offsets=new SortedDictionary<ulong,ArraySegment<byte>>();
            foreach(Fragment fragment in fragments) offsets[fragment.Start]=fragment.Data;
            var parts=new List<ArraySegment<byte>>(offsets.Count);
            ulong written=0;
            foreach(var fragment in offsets)
            {
                if(fragment.Key>written)
                    throw new ContentUnavailableException("A fragmented stored object is missing a byte range.");
                // Distinct offsets contribute complete bytes, including
                // overlaps. The last accepted fragment at each offset wins.
                parts.Add(fragment.Value);
                written=checked(written+(ulong)fragment.Value.Count);
            }
            return parts;
        }
    }
}