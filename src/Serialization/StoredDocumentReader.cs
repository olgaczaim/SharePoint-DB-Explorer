using System;
using System.Collections.Generic;
using System.IO;

namespace SharePointExplorer
{
    // Generic reconstruction follows fixed byte rules (see README). The
    // file root is the one named by the last stored revision read, and
    // nodes and leaves have independent sequence state. Partial leaf fragments
    // follow their stateful append rules. File order comes from the reference
    // tree, never from physical blob ordering. Publication uses the resulting
    // length; stored node sizes, leaf hashes and logical size are not compared.
    internal static partial class StoredDocumentReader
    {
        private const string IndexRoot = ShreddedStore.IndexRoot;
        private const string PrimaryRoot = "G:84defab9aaa34a0da3a8520c77ac7073:2";
        // Bounds a corrupt graph that repeats shared subtrees; real files visit each object once.
        private const long VisitBudget = 100000000;
        private sealed class Level
        {
            internal string Id;
            internal IList<string> References;
            internal int Next;
        }

        // The file's bytes as ordered segments of the stored buffers.
        internal static List<ArraySegment<byte>> Reconstruct(IList<StoredChunk> chunks)
        {
            if (chunks == null) throw new ArgumentNullException("chunks");
            ShreddedStore store = ShreddedStore.Load(chunks);
            string root = store.LastRoot(PrimaryRoot);
            if (root == null) throw new InvalidDataException("No stored revision names the primary file root.");
            GenericDocumentStore objects = store.GenericObjects;
            var parts = new List<ArraySegment<byte>>();
            // Depth-first walk with an explicit stack. An object already on the
            // current path would repeat forever, so it stops the walk.
            var stack = new Stack<Level>();
            var path = new HashSet<string>(StringComparer.Ordinal);
            var consumed = new HashSet<string>(StringComparer.Ordinal);
            long visits = 0;
            Visit(objects, root, stack, path, consumed, parts);
            while (stack.Count > 0)
            {
                Level level = stack.Peek();
                if (level.Next == level.References.Count)
                {
                    stack.Pop();
                    path.Remove(level.Id);
                    continue;
                }
                if (++visits > VisitBudget) throw new InvalidDataException("The stored file tree exceeds the supported traversal budget.");
                Visit(objects, level.References[level.Next++], stack, path, consumed, parts);
            }
            return parts;
        }

        internal static void Write(IList<StoredChunk> chunks, Stream output)
        {
            if (output == null || !output.CanWrite) throw new ArgumentException("A writable export stream is required.", "output");
            foreach (ArraySegment<byte> part in Reconstruct(chunks)) output.Write(part.Array, part.Offset, part.Count);
        }

        private static void Visit(GenericDocumentStore objects, string id, Stack<Level> stack, HashSet<string> path, HashSet<string> consumed, List<ArraySegment<byte>> parts)
        {
            string key = ShreddedStore.ObjectKey(id);
            IList<string> references=objects.References(key);
            // A node takes precedence over leaf data at the same identity.
            if(references==null)
            {
                if(!consumed.Add(key)) throw new InvalidDataException("The stored file references a consumed leaf object more than once.");
                parts.AddRange(objects.Data(key));
                return;
            }
            if(!path.Add(key)) throw new InvalidDataException("The stored file tree is cyclic.");
            stack.Push(new Level { Id=key,References=references });
        }
    }
}
