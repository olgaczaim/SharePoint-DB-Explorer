using System;
using System.Collections.Generic;
using System.IO;
using SharePointExplorer;

namespace SharePointExplorer.Tests
{
    // Synthetic persisted stores in the structured host-blob layout, assembled from
    // invented content and element IDs. Only the format's well-known root, cell
    // and index GUIDs are shared with real stores.
    internal static class StorageFixture
    {
        private const uint RootObject = 0x20000006;
        private const uint AlphaLeaf = 0x2000000A;
        private const uint BetaLeaf = 0x2000000E;
        private const uint AlphaRaw = 0x20000012;
        private const uint BetaRaw = 0x20000016;
        private const uint AlphaGroup = 0x20000022;
        private const uint BetaGroup = 0x20000026;
        private const uint TreeGroup = 0x2000002A;
        private const uint CellElement = 0x2000002E;
        private const uint RevisionElement = 0x20000032;
        private const uint RevisionId = 0x20000036;
        private const uint AlphaBlobElement = 0x2000003A;
        private const uint BaseRevisionElement = 0x2000003E;
        private const uint BaseRevisionId = 0x20000042;
        private const uint TreeGroup2 = 0x20000046;
        private const uint BetaGroup2 = 0x2000004A;
        private const uint BetaBlob = 0x2000004E;
        private const uint BlobElement = 0x20000052;
        private const uint FragmentA = 0x20000056;
        private const uint FragmentB = 0x2000005A;
        private const uint ReorderedRoot = 0x2000005E;
        private const uint LaterRevisionElement = 0x20000062;
        private const uint LaterRevisionId = 0x20000066;
        private const uint DamagedElement = 0x2000006A;
        private const string DictionaryGuid = "11111111-2222-3333-4444-555555555555";
        private const string OtherGuid = "99999999-8888-7777-6666-555555555555";
        private const string PrimaryRoot = "84defab9-aaa3-4a0d-a3a8-520c77ac7073";
        private const string CellB = "6f2a4665-42c8-46c7-bab4-e28fdce1e32b";
        private const string IndexGuid = "fb737e23-2247-4ec8-8d39-bda7f137215e";

        internal static List<StoredChunk> Create(byte[] alpha, byte[] beta)
        {
            return Scenario("plain", alpha, beta);
        }

        internal static List<StoredChunk> CreateAliasedBlobGroup(byte[] alpha, byte[] beta, bool wrongGuid = false, bool wrongValue = false)
        {
            return Scenario(wrongGuid ? "aliased-wrong-guid" : wrongValue ? "aliased-wrong-value" : "aliased", alpha, beta);
        }

        // Named layouts exercising each persisted representation. Rows are read
        // main row first, then the second and third rows. The expected file is
        // beta followed by alpha, except: "inherited" (and its cycle/missing-base
        // variants), whose newer copies with higher sequence numbers replace beta
        // with its reversed bytes; and "last-root-wins", whose later revision
        // names a root ordering alpha before beta.
        internal static List<StoredChunk> Scenario(string name, byte[] alpha, byte[] beta)
        {
            if (alpha == null) throw new ArgumentNullException("alpha");
            if (beta == null) throw new ArgumentNullException("beta");
            bool unhashed = name == "unhashed-leaves";
            byte[] tree = Join(
                NodeRecord(AlphaLeaf, new[] { AlphaRaw }, unhashed ? UnhashedLeaf(alpha) : LeafPayload(alpha)),
                NodeRecord(BetaLeaf, new[] { BetaRaw }, unhashed ? UnhashedLeaf(beta) : LeafPayload(beta)),
                NodeRecord(RootObject, new[] { BetaLeaf, AlphaLeaf }, ParentPayload((long)alpha.Length + beta.Length), 3));
            var main = new List<byte[]>();
            var second = new List<byte[]>();
            var third = new List<byte[]>();
            var fourth = new List<byte[]>();
            var index = new List<byte[]> { CellEntry(CellElement), RevisionEntry(Id(RevisionId), Id(RevisionElement)) };
            main.Add(DataBlob(CellElement, 3, Node(1, Id(RevisionId))));
            bool thirdTable = false, repeatedSlot = false;
            switch (name)
            {
                case "plain":
                case "unhashed-leaves":
                case "duplicate-object":
                    main.Add(DataBlob(TreeGroup, 5, GroupContent(tree, name == "duplicate-object" ? NodeRecord(RootObject, new[] { BetaLeaf, AlphaLeaf }, ParentPayload((long)alpha.Length + beta.Length), 3) : new byte[0])));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, AlphaGroup, BetaGroup, TreeGroup)));
                    second.Add(DataBlob(AlphaGroup, 5, GroupContent(RawRecord(AlphaRaw, alpha))));
                    third.Add(DataBlob(BetaGroup, 5, GroupContent(RawRecord(BetaRaw, beta))));
                    break;
                case "last-root-wins":
                {
                    // A revision read later names another root over the same leaves.
                    main.Add(DataBlob(TreeGroup, 5, GroupContent(tree)));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, AlphaGroup, BetaGroup, TreeGroup)));
                    second.Add(DataBlob(AlphaGroup, 5, GroupContent(RawRecord(AlphaRaw, alpha))));
                    third.Add(DataBlob(BetaGroup, 5, GroupContent(RawRecord(BetaRaw, beta),
                        NodeRecord(ReorderedRoot, new[] { AlphaLeaf, BetaLeaf }, ParentPayload((long)alpha.Length + beta.Length), 3))));
                    third.Add(DataBlob(LaterRevisionElement, 4, Revision(LaterRevisionId, null, ReorderedRoot, BetaGroup)));
                    break;
                }
                case "tied-conflicting-copy":
                case "newer-damaged-copy":
                {
                    // A fourth row holds another copy of alpha's group with altered
                    // bytes, at the same or a higher sequence number.
                    byte[] altered = (byte[])alpha.Clone(); altered[0] ^= 1;
                    main.Add(DataBlob(TreeGroup, 5, GroupContent(tree)));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, AlphaGroup, BetaGroup, TreeGroup)));
                    second.Add(DataBlob(AlphaGroup, 5, GroupContent(RawRecord(AlphaRaw, alpha))));
                    third.Add(DataBlob(BetaGroup, 5, GroupContent(RawRecord(BetaRaw, beta))));
                    fourth.Add(DataBlob(AlphaGroup, 5, GroupContent(RawRecord(AlphaRaw, altered)), name == "newer-damaged-copy" ? 2UL : 1UL));
                    break;
                }
                case "lenient-layout":
                    // Tolerated layouts: a repeated main-table slot (the
                    // later GUID wins), an element identity different from its blob,
                    // two element GUID tables and bytes after an element's footers.
                    main.Add(DataBlob(TreeGroup, 5, GroupContent(tree)));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, AlphaGroup, BetaGroup, TreeGroup)));
                    second.Add(Blob(5, Id(DamagedElement), Join(DataElementBlob(Id(AlphaGroup), 5, GroupContent(RawRecord(AlphaRaw, alpha))), new byte[] { 0xEE, 0xEE })));
                    third.Add(Blob(5, Id(BetaGroup), DataElementBlob(Id(BetaGroup), 5, GroupContent(RawRecord(BetaRaw, beta)), OtherGuid, null, 2)));
                    repeatedSlot = true;
                    break;
                case "damaged-after-beta":
                case "damaged-before-beta":
                {
                    // A damaged data element ends the third row; beta's group is read
                    // only when it comes first.
                    byte[] damaged = Blob(5, Id(DamagedElement), new byte[] { 1, 2, 3, 4, 5, 6 });
                    main.Add(DataBlob(TreeGroup, 5, GroupContent(tree)));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, AlphaGroup, BetaGroup, TreeGroup)));
                    second.Add(DataBlob(AlphaGroup, 5, GroupContent(RawRecord(AlphaRaw, alpha))));
                    if (name == "damaged-before-beta") third.Add(damaged);
                    third.Add(DataBlob(BetaGroup, 5, GroupContent(RawRecord(BetaRaw, beta))));
                    if (name == "damaged-after-beta") third.Add(damaged);
                    break;
                }
                case "aliased":
                case "aliased-wrong-guid":
                case "aliased-wrong-value":
                {
                    main.Add(DataBlob(TreeGroup, 5, GroupContent(tree)));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, AlphaGroup, BetaGroup, TreeGroup)));
                    // The outer blob carries the canonical GUID form; the fragment's
                    // group reference uses the compact alias resolved by the table.
                    byte value = (byte)((AlphaGroup & 0x0FFFFFFFU) >> 2);
                    if (name == "aliased-wrong-value") value++;
                    byte[] canonical = GuidId(name == "aliased-wrong-guid" ? "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" : DictionaryGuid, value);
                    second.Add(Blob(5, canonical, DataElementBlob(Id(AlphaBlobElement), 11, FragmentContent(Id(AlphaGroup), FragmentState(AlphaRaw, alpha)))));
                    third.Add(DataBlob(BetaGroup, 5, GroupContent(RawRecord(BetaRaw, beta))));
                    break;
                }
                case "fragmented-tree":
                {
                    // Real stores declare objects in one fragment and include them in
                    // another; descriptions and exclusions never hide included states.
                    byte[] fragmentA = FragmentContent(Id(TreeGroup),
                        FragmentState(RootObject, ParentPayload((long)alpha.Length + beta.Length), BetaLeaf, AlphaLeaf),
                        FragmentState(AlphaLeaf, LeafPayload(alpha), AlphaRaw),
                        Declaration(7, BetaLeaf, 0));
                    byte[] fragmentB = FragmentContent(Id(TreeGroup),
                        Declaration(5, RootObject, 0),
                        FragmentState(BetaLeaf, LeafPayload(beta), BetaRaw));
                    main.Add(Blob(5, Id(TreeGroup), DataElementBlob(Id(FragmentA), 11, fragmentA)));
                    second.Add(Blob(5, Id(TreeGroup), DataElementBlob(Id(FragmentB), 11, fragmentB)));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, AlphaGroup, BetaGroup, TreeGroup)));
                    second.Add(DataBlob(AlphaGroup, 5, GroupContent(RawRecord(AlphaRaw, alpha))));
                    third.Add(DataBlob(BetaGroup, 5, GroupContent(RawRecord(BetaRaw, beta))));
                    break;
                }
                case "partial-alpha":
                case "partial-gap":
                case "partial-overlap-conflict":
                {
                    // Alpha's data object is split into partial fragments in two group fragments.
                    int half = alpha.Length / 2;
                    byte[] first = Slice(alpha, 0, name == "partial-overlap-conflict" ? half + 1 : half);
                    byte[] rest = Slice(alpha, name == "partial-gap" ? half + 1 : half, alpha.Length - (name == "partial-gap" ? half + 1 : half));
                    if (name == "partial-overlap-conflict") { rest = (byte[])rest.Clone(); rest[0] ^= 0x20; }
                    int restStart = name == "partial-gap" ? half + 1 : half;
                    main.Add(DataBlob(TreeGroup, 5, GroupContent(tree)));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, AlphaGroup, BetaGroup, TreeGroup)));
                    second.Add(Blob(5, Id(AlphaGroup), DataElementBlob(Id(FragmentA), 11, FragmentContent(Id(AlphaGroup),
                        Declaration(5, AlphaRaw, (ulong)alpha.Length), Partial(AlphaRaw, (ulong)alpha.Length, 0, first)))));
                    third.Add(Blob(5, Id(AlphaGroup), DataElementBlob(Id(FragmentB), 11, FragmentContent(Id(AlphaGroup),
                        Partial(AlphaRaw, (ulong)alpha.Length, (ulong)restStart, rest)))));
                    third.Add(DataBlob(BetaGroup, 5, GroupContent(RawRecord(BetaRaw, beta))));
                    break;
                }
                case "blob-beta":
                case "blob-missing":
                {
                    // Beta's data object refers to an object data BLOB stored in two pieces.
                    main.Add(DataBlob(TreeGroup, 5, GroupContent(tree)));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, AlphaGroup, BetaGroup, TreeGroup)));
                    second.Add(DataBlob(AlphaGroup, 5, GroupContent(RawRecord(AlphaRaw, alpha))));
                    third.Add(DataBlob(BetaGroup, 5, GroupContent(BlobRecord(BetaRaw, BetaBlob))));
                    if (name == "blob-beta")
                    {
                        int split = beta.Length / 2;
                        third.Add(Blob(5, Id(BlobElement), DataElementBlob(Id(BlobElement), 12, Join(
                            BlobPiece(BetaBlob, (ulong)beta.Length, 0, Slice(beta, 0, split)),
                            BlobPiece(BetaBlob, (ulong)beta.Length, (ulong)split, Slice(beta, split, beta.Length - split))))));
                    }
                    break;
                }
                case "inherited":
                case "inherited-cycle":
                case "inherited-missing-base":
                case "stale-higher-sequence":
                {
                    // The base revision holds the original tree and groups. The current
                    // revision adds newer beta leaf, raw and root states only, stored at
                    // a higher sequence number. Revision links are not followed, so a
                    // cyclic or missing base changes nothing. In "stale-higher-sequence"
                    // the original copies carry the higher sequence number and win.
                    byte[] newBeta = (byte[])beta.Clone(); Array.Reverse(newBeta);
                    byte[] newerTree = Join(
                        NodeRecord(BetaLeaf, new[] { BetaRaw }, LeafPayload(newBeta)),
                        NodeRecord(RootObject, new[] { BetaLeaf, AlphaLeaf }, ParentPayload((long)alpha.Length + newBeta.Length), 3));
                    ulong original = name == "stale-higher-sequence" ? 3UL : 1UL;
                    main.Add(DataBlob(TreeGroup, 5, GroupContent(tree), original));
                    main.Add(DataBlob(TreeGroup2, 5, GroupContent(newerTree), 2));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, name == "inherited-missing-base" ? 0x2000007EU : BaseRevisionId, null, TreeGroup2, BetaGroup2)));
                    main.Add(DataBlob(BaseRevisionElement, 4, Revision(BaseRevisionId, name == "inherited-cycle" ? RevisionId : (uint?)null, RootObject, AlphaGroup, BetaGroup, TreeGroup)));
                    index.Add(RevisionEntry(Id(BaseRevisionId), Id(BaseRevisionElement)));
                    second.Add(DataBlob(AlphaGroup, 5, GroupContent(RawRecord(AlphaRaw, alpha))));
                    third.Add(DataBlob(BetaGroup, 5, GroupContent(RawRecord(BetaRaw, beta)), original));
                    third.Add(DataBlob(BetaGroup2, 5, GroupContent(RawRecord(BetaRaw, newBeta)), 2));
                    break;
                }
                case "element-guid-table":
                    // Every data element declares its own GUID table remapping slot 2.
                    // Those tables are ignored: identifiers resolve against the host
                    // table, so applying them would change every element identity.
                    main[0] = DataBlob(CellElement, 3, Node(1, Id(RevisionId)), 1, OtherGuid);
                    main.Add(DataBlob(TreeGroup, 5, GroupContent(tree), 1, OtherGuid));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, AlphaGroup, BetaGroup, TreeGroup), 1, OtherGuid));
                    second.Add(DataBlob(AlphaGroup, 5, GroupContent(RawRecord(AlphaRaw, alpha)), 1, OtherGuid));
                    third.Add(DataBlob(BetaGroup, 5, GroupContent(RawRecord(BetaRaw, beta)), 1, OtherGuid));
                    break;
                case "secondary-table":
                case "undeclared-slot":
                {
                    // The third host blob declares its own table (slot 5); its IDs use
                    // that slot, which the main table does not declare.
                    main.Add(DataBlob(TreeGroup, 5, GroupContent(tree)));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, AlphaGroup, BetaGroup, TreeGroup)));
                    second.Add(DataBlob(AlphaGroup, 5, GroupContent(RawRecord(AlphaRaw, alpha))));
                    byte[] betaGroup = Slot5(BetaGroup), betaRaw = Slot5(BetaRaw);
                    third.Add(Blob(5, betaGroup, DataElementBlob(betaGroup, 5, GroupContent(Node(2, Join(betaRaw, Compact(1), Compact(0), Compact(0), Binary(beta), Compact(0)))), null, Serial4(5, 1))));
                    thirdTable = name == "secondary-table";
                    break;
                }
                case "non-aggregatable":
                    main.Add(DataBlob(TreeGroup, 5, GroupContent(tree)));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, AlphaGroup, BetaGroup, TreeGroup)));
                    second.Add(DataBlob(AlphaGroup, 5, GroupContent(RawRecord(AlphaRaw, alpha))));
                    third.Add(Blob(6, Id(BetaGroup), DataElementBlob(Id(BetaGroup), 5, GroupContent(RawRecord(BetaRaw, beta)))));
                    break;
                case "compact-forms":
                {
                    // Two-byte and six-byte table-indexed IDs plus six-byte and full
                    // serial numbers, mixed with the four-byte forms used elsewhere.
                    main.Add(DataBlob(TreeGroup, 5, GroupContent(tree)));
                    main.Add(DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, AlphaGroup, BetaGroup, TreeGroup)));
                    byte[] alphaGroup = Id2(8), alphaRaw = Id6(4);
                    second.Add(Blob(5, alphaGroup, DataElementBlob(alphaGroup, 5, GroupContent(Node(2, Join(alphaRaw, Compact(1), Compact(0), Compact(0), Binary(alpha), Compact(0)))), null, Serial6(2, 7))));
                    third.Add(Blob(5, Id(BetaGroup), DataElementBlob(Id(BetaGroup), 5, GroupContent(RawRecord(BetaRaw, beta)), null, SerialFull(DictionaryGuid, 9))));
                    break;
                }
                default: throw new ArgumentException("Unknown storage fixture scenario.", "name");
            }
            main.Insert(0, IndexBlob(index.ToArray()));
            var rows = new List<StoredChunk>
            {
                new StoredChunk { Partition = 0, Type = 10, BSN = 1, StreamId = 2, Content = HostBlob(false, second.ToArray()) },
                new StoredChunk { Partition = 0, Type = 10, BSN = 2, StreamId = 3, Content = HostBlob(false, third.ToArray(), thirdTable) },
                new StoredChunk { Partition = 0, Type = 11, BSN = 3, StreamId = 1, Content = HostBlob(true, main.ToArray(), false, repeatedSlot) }
            };
            if (fourth.Count > 0) rows.Add(new StoredChunk { Partition = 0, Type = 10, BSN = 4, StreamId = 4, Content = HostBlob(false, fourth.ToArray()) });
            return rows;
        }

        // End-to-end inputs for the generic reader's state transitions.
        // Expected bytes belong to the checks, not this serializer. Each operation
        // occupies its own row, and deliberately decreasing SQL BSNs prevent the
        // fixtures from confusing physical row identity with embedded sequence.
        internal static List<StoredChunk> GenericReaderScenario(string name)
        {
            var operations = new List<byte[]>();
            switch (name)
            {
                case "node-precedes-newer-inline":
                    operations.Add(ManagedGroup(AlphaGroup, 10, NodeRecord(RootObject, new[] { AlphaRaw, BetaRaw }, new byte[0])));
                    operations.Add(ManagedGroup(BetaGroup, 20, RawRecord(RootObject, Bytes("SHADOW"))));
                    break;
                case "node-precedes-inline-read-first":
                    operations.Add(ManagedGroup(AlphaGroup, 20, RawRecord(RootObject, Bytes("SHADOW"))));
                    operations.Add(ManagedGroup(BetaGroup, 10, NodeRecord(RootObject, new[] { AlphaRaw, BetaRaw }, new byte[0])));
                    break;
                case "node-tie-keeps-first":
                    operations.Add(ManagedGroup(AlphaGroup, 10, NodeRecord(RootObject, new[] { BetaRaw, AlphaRaw }, new byte[0])));
                    operations.Add(ManagedGroup(BetaGroup, 10, NodeRecord(RootObject, new[] { AlphaRaw, BetaRaw }, new byte[0])));
                    break;
                case "newer-node-replaces-older":
                    operations.Add(ManagedGroup(AlphaGroup, 10, NodeRecord(RootObject, new[] { BetaRaw, AlphaRaw }, new byte[0])));
                    operations.Add(ManagedGroup(BetaGroup, 20, NodeRecord(RootObject, new[] { AlphaRaw, BetaRaw }, new byte[0])));
                    operations.Add(ManagedGroup(TreeGroup2, 15, NodeRecord(RootObject, new[] { BetaRaw, AlphaRaw }, new byte[0])));
                    break;
                case "inline-tie-keeps-first":
                    operations.Add(ManagedGroup(AlphaGroup, 10, RawRecord(RootObject, Bytes("FIRST"))));
                    operations.Add(ManagedGroup(BetaGroup, 10, RawRecord(RootObject, Bytes("SECOND"))));
                    break;
                case "newer-inline-replaces-older":
                    operations.Add(ManagedGroup(AlphaGroup, 10, RawRecord(RootObject, Bytes("OLD"))));
                    operations.Add(ManagedGroup(BetaGroup, 20, RawRecord(RootObject, Bytes("NEW"))));
                    operations.Add(ManagedGroup(TreeGroup2, 15, RawRecord(RootObject, Bytes("STALE"))));
                    break;
                case "partial-with-references-is-leaf":
                    operations.Add(ManagedFragments(AlphaGroup, FragmentA, 10,
                        ManagedPartial(RootObject, 4, 0, 4, Bytes("PART"), AlphaRaw)));
                    break;
                case "fragments-across-groups":
                    operations.Add(ManagedFragments(AlphaGroup, FragmentA, 10, ManagedPartial(RootObject, 2, 0, 1, Bytes("A"))));
                    operations.Add(ManagedFragments(BetaGroup, FragmentB, 10, ManagedPartial(RootObject, 2, 1, 2, Bytes("B"))));
                    break;
                case "fragment-threshold-does-not-advance":
                    operations.Add(ManagedFragments(AlphaGroup, FragmentA, 10, ManagedPartial(RootObject, 3, 0, 1, Bytes("A"))));
                    operations.Add(ManagedFragments(BetaGroup, FragmentB, 20, ManagedPartial(RootObject, 3, 1, 2, Bytes("B"))));
                    operations.Add(ManagedFragments(TreeGroup2, AlphaBlobElement, 15, ManagedPartial(RootObject, 3, 2, 3, Bytes("C"))));
                    operations.Add(ManagedFragments(BetaGroup2, BlobElement, 9, ManagedPartial(RootObject, 3, 0, 3, Bytes("BAD"))));
                    break;
                case "older-fragment-than-inline-rejected":
                    operations.Add(ManagedGroup(AlphaGroup, 20, RawRecord(RootObject, Bytes("INLINE"))));
                    operations.Add(ManagedFragments(BetaGroup, FragmentA, 19, ManagedPartial(RootObject, 3, 0, 3, Bytes("BAD"))));
                    break;
                case "equal-fragment-to-inline-accepted":
                    operations.Add(ManagedGroup(AlphaGroup, 20, RawRecord(RootObject, Bytes("INLINE"))));
                    operations.Add(ManagedFragments(BetaGroup, FragmentA, 20, ManagedPartial(RootObject, 1, 0, 1, Bytes("A"))));
                    operations.Add(ManagedGroup(TreeGroup2, 20, RawRecord(RootObject, Bytes("TIED"))));
                    break;
                case "duplicate-start-equal-sequence-last-wins":
                    operations.Add(ManagedFragments(AlphaGroup, FragmentA, 10, ManagedPartial(RootObject, 5, 0, 5, Bytes("FIRST"))));
                    operations.Add(ManagedFragments(BetaGroup, FragmentB, 10, ManagedPartial(RootObject, 4, 0, 4, Bytes("LAST"))));
                    break;
                case "duplicate-start-last-accepted-not-highest-sequence":
                    operations.Add(ManagedFragments(AlphaGroup, FragmentA, 10, ManagedPartial(RootObject, 1, 0, 1, Bytes("A"))));
                    operations.Add(ManagedFragments(BetaGroup, FragmentB, 20, ManagedPartial(RootObject, 4, 0, 4, Bytes("HIGH"))));
                    operations.Add(ManagedFragments(TreeGroup2, AlphaBlobElement, 15, ManagedPartial(RootObject, 5, 0, 5, Bytes("LATER"))));
                    break;
                case "overlapping-fragments-append-complete-bytes":
                    operations.Add(ManagedFragments(AlphaGroup, FragmentA, 10, ManagedPartial(RootObject, 4, 0, 4, Bytes("ABCD"))));
                    operations.Add(ManagedFragments(BetaGroup, FragmentB, 10, ManagedPartial(RootObject, 4, 2, 4, Bytes("xy"))));
                    break;
                case "cumulative-output-byte-gap-check":
                    operations.Add(ManagedFragments(AlphaGroup, FragmentA, 10, ManagedPartial(RootObject, 8, 0, 4, Bytes("ABCD"))));
                    operations.Add(ManagedFragments(BetaGroup, FragmentB, 10, ManagedPartial(RootObject, 8, 2, 6, Bytes("efgh"))));
                    operations.Add(ManagedFragments(TreeGroup2, AlphaBlobElement, 10, ManagedPartial(RootObject, 8, 7, 8, Bytes("I"))));
                    break;
                case "newer-inline-keeps-existing-fragments":
                    operations.Add(ManagedFragments(AlphaGroup, FragmentA, 10, ManagedPartial(RootObject, 1, 0, 1, Bytes("A"))));
                    operations.Add(ManagedGroup(BetaGroup, 20, RawRecord(RootObject, Bytes("INLINE"))));
                    break;
                case "inline-raises-fragment-rejection-threshold":
                    operations.Add(ManagedFragments(AlphaGroup, FragmentA, 10, ManagedPartial(RootObject, 2, 0, 1, Bytes("A"))));
                    operations.Add(ManagedGroup(BetaGroup, 20, RawRecord(RootObject, Bytes("INLINE"))));
                    operations.Add(ManagedFragments(TreeGroup2, FragmentB, 15, ManagedPartial(RootObject, 4, 1, 4, Bytes("BAD"))));
                    operations.Add(ManagedFragments(BetaGroup2, AlphaBlobElement, 20, ManagedPartial(RootObject, 2, 1, 2, Bytes("B"))));
                    break;
                case "declared-fragment-size-and-end-ignored":
                    operations.Add(ManagedFragments(AlphaGroup, FragmentA, 10, ManagedPartial(RootObject, 0, 0, UInt64.MaxValue, Bytes("A"))));
                    operations.Add(ManagedFragments(BetaGroup, FragmentB, 10, ManagedPartial(RootObject, UInt64.MaxValue, 1, 0, Bytes("B"))));
                    break;
                case "odb-does-not-override-inline":
                    operations.Add(ManagedGroup(AlphaGroup, 10, RawRecord(RootObject, Bytes("INLINE"))));
                    operations.Add(ManagedGroup(BetaGroup, 100, BlobRecord(RootObject, BetaBlob)));
                    operations.Add(DataBlob(BlobElement, 12, BlobPiece(BetaBlob, 4, 0, Bytes("BLOB"))));
                    break;
                case "fragment-odb-does-not-override-partials":
                    operations.Add(ManagedFragments(AlphaGroup, FragmentA, 10, ManagedPartial(RootObject, 4, 0, 4, Bytes("PART"))));
                    operations.Add(ManagedFragments(BetaGroup, FragmentB, 100, ManagedOdb(RootObject, BetaBlob)));
                    operations.Add(DataBlob(BlobElement, 12, BlobPiece(BetaBlob, 4, 0, Bytes("BLOB"))));
                    break;
                case "declaration-does-not-override-inline":
                    operations.Add(ManagedGroup(AlphaGroup, 10, RawRecord(RootObject, Bytes("INLINE"))));
                    operations.Add(ManagedFragments(BetaGroup, FragmentA, 100, Declaration(7, RootObject, 0)));
                    break;
                case "direct-odb-serial4":
                    operations.Add(ManagedGroup(AlphaGroup, 20, RawRecord(RootObject, Bytes("INLINE"))));
                    operations.Add(DataBlob(BlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 19), 3, 0, 3, Bytes("BAD")), 100));
                    operations.Add(DataBlob(AlphaBlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 20), 2, 0, 2, Bytes("S4")), 1));
                    break;
                case "direct-odb-serial6":
                    operations.Add(ManagedGroup(AlphaGroup, (1UL << 30) + 20, RawRecord(RootObject, Bytes("INLINE"))));
                    operations.Add(DataBlob(BlobElement, 12, ManagedBlobPiece(RootObject, Serial6(2, (1UL << 30) + 19), 3, 0, 3, Bytes("BAD")), 100));
                    operations.Add(DataBlob(AlphaBlobElement, 12, ManagedBlobPiece(RootObject, Serial6(2, (1UL << 30) + 20), 2, 0, 2, Bytes("S6")), 1));
                    break;
                case "direct-odb-serial-full":
                    operations.Add(ManagedGroup(AlphaGroup, (1UL << 40) + 20, RawRecord(RootObject, Bytes("INLINE"))));
                    operations.Add(DataBlob(BlobElement, 12, ManagedBlobPiece(RootObject, SerialFull(DictionaryGuid, (1UL << 40) + 19), 3, 0, 3, Bytes("BAD")), 100));
                    operations.Add(DataBlob(AlphaBlobElement, 12, ManagedBlobPiece(RootObject, SerialFull(DictionaryGuid, (1UL << 40) + 20), 2, 0, 2, Bytes("SF")), 1));
                    break;
                case "direct-odb-serial-threshold-not-host-sequence":
                    operations.Add(DataBlob(BlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 10), 3, 0, 1, Bytes("A")), 100));
                    operations.Add(DataBlob(AlphaBlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 9), 3, 0, 3, Bytes("BAD")), 200));
                    operations.Add(DataBlob(FragmentA, 12, ManagedBlobPiece(RootObject, Serial4(2, 20), 3, 1, 2, Bytes("B")), 1));
                    operations.Add(DataBlob(FragmentB, 12, ManagedBlobPiece(RootObject, Serial4(2, 15), 3, 2, 3, Bytes("C")), 999));
                    break;
                case "direct-odb-duplicate-start-last-accepted-wins":
                    operations.Add(DataBlob(BlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 10), 1, 0, 1, Bytes("A"))));
                    operations.Add(DataBlob(AlphaBlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 20), 4, 0, 4, Bytes("HIGH"))));
                    operations.Add(DataBlob(FragmentA, 12, ManagedBlobPiece(RootObject, Serial4(2, 15), 4, 0, 4, Bytes("LAST"))));
                    break;
                case "direct-odb-overlap-appends-complete-bytes":
                    operations.Add(DataBlob(BlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 10), 4, 0, 4, Bytes("ABCD"))));
                    operations.Add(DataBlob(AlphaBlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 10), 4, 2, 4, Bytes("xy"))));
                    break;
                case "direct-odb-cumulative-byte-gap-check":
                    operations.Add(DataBlob(BlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 10), 8, 0, 4, Bytes("ABCD"))));
                    operations.Add(DataBlob(AlphaBlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 10), 8, 2, 6, Bytes("efgh"))));
                    operations.Add(DataBlob(FragmentA, 12, ManagedBlobPiece(RootObject, Serial4(2, 10), 8, 7, 8, Bytes("I"))));
                    break;
                case "direct-odb-inline-preserves-fragment-preference":
                    operations.Add(DataBlob(BlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 10), 2, 0, 1, Bytes("A"))));
                    operations.Add(ManagedGroup(AlphaGroup, 20, RawRecord(RootObject, Bytes("INLINE"))));
                    operations.Add(DataBlob(AlphaBlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 15), 4, 1, 4, Bytes("BAD")), 999));
                    operations.Add(DataBlob(FragmentA, 12, ManagedBlobPiece(RootObject, Serial4(2, 20), 2, 1, 2, Bytes("B")), 1));
                    break;
                case "node-precedes-direct-odb-leaf":
                    operations.Add(ManagedGroup(AlphaGroup, 10, NodeRecord(RootObject, new[] { AlphaRaw, BetaRaw }, new byte[0])));
                    operations.Add(DataBlob(BlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 20), 6, 0, 6, Bytes("SHADOW"))));
                    break;
                case "direct-odb-declared-size-and-end-ignored":
                    operations.Add(DataBlob(BlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 10), 0, 0, UInt64.MaxValue, Bytes("A"))));
                    operations.Add(DataBlob(AlphaBlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 10), UInt64.MaxValue, 1, 0, Bytes("B"))));
                    break;
                case "malformed-odb-footer-preserves-prior-valid-piece":
                    operations.Add(DataBlob(BlobElement, 12, Join(
                        ManagedBlobPiece(RootObject, Serial4(2, 10), 4, 0, 4, Bytes("GOOD")),
                        ManagedBlobPieceWithoutFooter(RootObject, Serial4(2, 20), Bytes("BAD")))));
                    break;
                case "malformed-odb-footer-does-not-register-piece":
                    operations.Add(DataBlob(BlobElement, 12, ManagedBlobPieceWithoutFooter(RootObject, Serial4(2, 20), Bytes("BAD"))));
                    break;
                case "null-odb-serial-does-not-register-piece":
                    operations.Add(DataBlob(BlobElement, 12, ManagedBlobPiece(RootObject, new byte[] { 0 }, 3, 0, 3, Bytes("BAD"))));
                    break;
                case "direct-odb-gap-is-rejected":
                    operations.Add(DataBlob(BlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 10), 5, 0, 2, Bytes("AB"))));
                    operations.Add(DataBlob(AlphaBlobElement, 12, ManagedBlobPiece(RootObject, Serial4(2, 10), 5, 3, 5, Bytes("CD"))));
                    break;
                case "fragment-gap-is-rejected":
                    operations.Add(ManagedFragments(AlphaGroup, FragmentA, 10, ManagedPartial(RootObject, 5, 0, 2, Bytes("AB"))));
                    operations.Add(ManagedFragments(BetaGroup, FragmentB, 10, ManagedPartial(RootObject, 5, 3, 5, Bytes("CD"))));
                    break;
                case "repeated-leaf-is-rejected":
                    operations.Add(ManagedGroup(AlphaGroup, 10, NodeRecord(RootObject, new[] { AlphaRaw, AlphaRaw }, new byte[0])));
                    break;
                default: throw new ArgumentException("Unknown generic-reader fixture scenario.", "name");
            }
            var rows = new List<StoredChunk>();
            for (int index = 0; index < operations.Count; index++)
                rows.Add(new StoredChunk { Partition = 0, Type = 10, BSN = 1000 - index * 7,
                    StreamId = index + 2, Content = HostBlob(false, new[] { operations[index] }) });
            rows.Add(new StoredChunk { Partition = 0, Type = 11, BSN = 2000, StreamId = 1, Content = HostBlob(true, new[] {
                IndexBlob(new[] { CellEntry(CellElement), RevisionEntry(Id(RevisionId), Id(RevisionElement)) }),
                DataBlob(CellElement, 3, Node(1, Id(RevisionId))),
                DataBlob(RevisionElement, 4, Revision(RevisionId, null, RootObject, TreeGroup, AlphaGroup, BetaGroup)),
                ManagedGroup(TreeGroup, 1, RawRecord(AlphaRaw, Bytes("A")), RawRecord(BetaRaw, Bytes("B"))) }) });
            return rows;
        }
        private static byte[] Bytes(string text) { return System.Text.Encoding.ASCII.GetBytes(text); }
        private static byte[] ManagedGroup(uint group, ulong sequence, params byte[][] states)
        { return DataBlob(group, 5, GroupContent(states), sequence); }
        private static byte[] ManagedFragments(uint group, uint fragment, ulong sequence, params byte[][] states)
        { return Blob(5, Id(group), DataElementBlob(Id(fragment), 11, FragmentContent(Id(group), states)), sequence); }
        private static byte[] ManagedPartial(uint id, ulong size, ulong start, ulong end, byte[] data, params uint[] references)
        {
            var fields = new List<byte[]> { Id(id), Compact(1), Compact(0), Compact((ulong)references.Length) };
            foreach (uint reference in references) fields.Add(Id(reference));
            fields.Add(Compact(0)); fields.Add(Compact(size)); fields.Add(Compact(start)); fields.Add(Compact(end)); fields.Add(Binary(data));
            return Node(6, Join(fields.ToArray()));
        }
        private static byte[] ManagedOdb(uint id, uint blob)
        { return Node(4, Join(Id(id), Compact(1), Compact(0), Compact(0), Compact(0), Id(blob))); }
        private static byte[] ManagedBlobPiece(uint fragmentOf, byte[] serial, ulong size, ulong start, ulong end, byte[] data)
        { return Node(1, Join(Id(fragmentOf), serial, Compact(size), Compact(start), Compact(end), Binary(data))); }
        private static byte[] ManagedBlobPieceWithoutFooter(uint fragmentOf, byte[] serial, byte[] data)
        {
            byte[] piece = ManagedBlobPiece(fragmentOf, serial, (ulong)data.Length, 0, (ulong)data.Length, data);
            piece[piece.Length - 1] = E(1)[0];
            return piece;
        }

        // Host blob: header item, optional GUID table, contained blobs, footer.
        private static byte[] HostBlob(bool table, byte[][] blobs, bool slot5Table = false, bool repeatedSlot = false)
        {
            byte[] header = Item(Join(Compact(0), Compact(0), Compact(0), Compact(1), Compact(0)));
            byte[][] entries = repeatedSlot ? new[] { Entry(2, OtherGuid), Entry(2, DictionaryGuid) } : new[] { Entry(2, DictionaryGuid) };
            byte[] guids = table ? GuidTableElement(4, Join(new Guid(DictionaryGuid).ToByteArray(), Compact(0)), 1, entries)
                : slot5Table ? GuidTableElement(4, Join(new Guid(DictionaryGuid).ToByteArray(), Compact(0)), 1, new[] { Entry(5, DictionaryGuid) }) : new byte[0];
            return Join(E(1), header, guids, E(2), Join(blobs), F, F);
        }
        private static byte[] Entry(ulong index, string guid) { return Node(1, Join(Compact(index), new Guid(guid).ToByteArray())); }
        private static byte[] GuidTableElement(int outer, byte[] outerItem, int table, byte[][] entries)
        {
            return Join(E(outer), Item(outerItem), E(table), Join(entries), F, F);
        }
        // Contained blob: identity, sequence number, type, change frequency, tag and data.
        private static byte[] Blob(int kind, byte[] id, byte[] data, ulong sequence = 1)
        {
            return Node(kind, Join(id, Compact(sequence), Compact(0), Compact(0), Compact(kind == 5 ? 1UL : 0UL), Binary(data), Compact(0), Compact(0)));
        }
        private static byte[] DataBlob(uint id, ulong type, byte[] content, ulong sequence = 1, string ownGuid = null)
        {
            return Blob(5, Id(id), DataElementBlob(Id(id), type, content, ownGuid), sequence);
        }
        // Data element blob: signature, header item, then GUID tables, identity and contents.
        private static byte[] DataElementBlob(byte[] id, ulong type, byte[] content, string ownGuid = null, byte[] serial = null, int tables = 1)
        {
            byte[] table = ownGuid == null ? Join(E(3), F) : Join(E(3), Entry(2, ownGuid), F);
            var declared = new List<byte[]>();
            for (int index = 0; index < tables; index++) declared.Add(table);
            return Join(new byte[] { 0xBA, 0x32, 0x4E, 0x2E }, E(1), Item(new byte[] { 1 }), E(3), Join(declared.ToArray()),
                Node(1, Join(id, Compact(type), serial ?? Serial4(2, 1))), E(2), content, F, F, F);
        }
        private static byte[] IndexBlob(byte[][] entries)
        {
            byte[] content = Join(E(1), Item(new byte[] { 0 }), E(2), Item(BitConverter.GetBytes((uint)entries.Length)), Join(entries), F, F);
            return Node(3, Join(GuidId(IndexGuid, 1), Compact(1), Compact(1), Compact(0), Compact(0), Binary(content), Compact(0), Compact(0)));
        }
        private static byte[] CellEntry(uint target)
        {
            return Node(5, Join(GuidId(PrimaryRoot, 1), GuidId(CellB, 1), Binary(Pointer(Id(target), 2)), Compact(1)));
        }
        private static byte[] RevisionEntry(byte[] revision, byte[] target)
        {
            return Node(4, Join(revision, Binary(Pointer(target, 3)), Compact(1)));
        }
        private static byte[] Pointer(byte[] target, ulong kind) { return Node(1, Join(target, Compact(1), Compact(kind))); }

        // A revision manifest, optionally naming the object of the primary file root.
        private static byte[] Revision(uint id, uint? basis, uint? root, params uint[] groups)
        {
            var references = new List<byte[]>();
            foreach (uint group in groups) references.Add(Node(5, Id(group)));
            return Join(
                Node(1, Join(Id(id), basis.HasValue ? Id(basis.Value) : new byte[] { 0 })),
                root.HasValue ? Join(E(2), Node(3, Join(GuidId(PrimaryRoot, 2), Id(root.Value))), F) : new byte[0],
                E(4), Item(Compact(0)), Join(references.ToArray()), F);
        }
        private static byte[] GroupContent(params byte[][] objects) { return Join(E(1), Join(objects), F); }
        private static byte[] FragmentContent(byte[] group, params byte[][] objects)
        {
            return Join(E(1), Item(Join(group, Serial4(2, 1))), E(2), Join(objects), F, F);
        }
        private static byte[] RawRecord(uint id, byte[] data) { return Node(2, Join(Id(id), Compact(1), Compact(0), Compact(0), Binary(data), Compact(0))); }
        private static byte[] BlobRecord(uint id, uint blob) { return Node(3, Join(Id(id), Compact(1), Compact(0), Compact(0), Id(blob), Compact(0))); }
        private static byte[] NodeRecord(uint id, uint[] references, byte[] payload, ulong changeFrequency = 0)
        {
            var fields = new List<byte[]> { Id(id), Compact(1), Compact((ulong)references.Length) };
            foreach (uint reference in references) fields.Add(Id(reference));
            fields.Add(Compact(0)); fields.Add(Binary(payload)); fields.Add(Compact(changeFrequency));
            return Node(2, Join(fields.ToArray()));
        }
        private static byte[] FragmentState(uint id, byte[] data, params uint[] references)
        {
            var fields = new List<byte[]> { Id(id), Compact(1), Compact(0), Compact((ulong)references.Length) };
            foreach (uint reference in references) fields.Add(Id(reference));
            fields.Add(Compact(0)); fields.Add(Binary(data));
            return Node(3, Join(fields.ToArray()));
        }
        private static byte[] Declaration(int kind, uint id, ulong size) { return Node(kind, Join(Id(id), Compact(1), Compact(0), Compact(0), Compact(0), Compact(size))); }
        private static byte[] Partial(uint id, ulong size, ulong start, byte[] data)
        {
            return Node(6, Join(Id(id), Compact(1), Compact(0), Compact(0), Compact(0), Compact(size), Compact(start), Compact((ulong)data.Length), Binary(data)));
        }
        private static byte[] BlobPiece(uint blob, ulong size, ulong start, byte[] data)
        {
            return Node(1, Join(Id(blob), Serial4(2, 1), Compact(size), Compact(start), Compact(start + (ulong)data.Length), Binary(data)));
        }

        private static byte[] LeafPayload(byte[] data)
        {
            byte[] hash = new byte[20];
            // Independent bit-level implementation of the published circular XOR hash.
            for (int index = 0; index < data.Length; index++)
                for (int bit = 0; bit < 8; bit++)
                    if ((data[index] & (1 << bit)) != 0)
                    {
                        int target = (int)(((long)index * 11 + bit) % 160);
                        hash[target / 8] ^= (byte)(1 << (target % 8));
                    }
            using (MemoryStream output = new MemoryStream())
            using (BinaryWriter writer = new BinaryWriter(output))
            {
                writer.Write(new byte[] { 0xFC, 0, 8, 0x13, 0x11 });
                writer.Write(new byte[8]);
                writer.Write((ushort)0x1110);
                writer.Write((ulong)data.Length);
                writer.Write(new byte[] { 0x78, 0x2B, 0x29 });
                writer.Write(hash);
                writer.Write((byte)0x7D);
                return output.ToArray();
            }
        }
        private static byte[] UnhashedLeaf(byte[] data)
        {
            using (MemoryStream output = new MemoryStream())
            using (BinaryWriter writer = new BinaryWriter(output))
            {
                writer.Write(new byte[] { 0xFC, 0, 8, 0x13, 0x11 });
                writer.Write(new byte[8]);
                writer.Write((ushort)0x1110);
                writer.Write((ulong)data.Length);
                writer.Write((byte)0x7D);
                return output.ToArray();
            }
        }
        private static byte[] ParentPayload(long size)
        {
            using (MemoryStream output = new MemoryStream())
            using (BinaryWriter writer = new BinaryWriter(output))
            {
                writer.Write(new byte[] { 4, 1, 8, 3, 0, 0x10, 0x11 });
                writer.Write((ulong)size);
                writer.Write((byte)0x81);
                return output.ToArray();
            }
        }

        // Structure: element header (1-byte xx01 or 2-byte xx11), item header
        // (2-byte x100 or 4-byte x010) with a length, and a zero footer.
        private static readonly byte[] F = { 0 };
        private static byte[] E(int id)
        {
            if (id < 64) return new[] { (byte)((id << 2) | 1) };
            return BitConverter.GetBytes((ushort)((id << 2) | 3));
        }
        private static byte[] Item(byte[] payload)
        {
            if (payload.LongLength > 0x1FFFFE) throw new ArgumentException("Synthetic item is too large.");
            byte[] header = payload.Length <= 2047 ? BitConverter.GetBytes((ushort)((payload.Length << 5) | 4)) : BitConverter.GetBytes(((uint)payload.Length << 11) | 2U);
            return Join(header, payload);
        }
        private static byte[] Node(int id, byte[] item) { return Join(E(id), Item(item), F); }

        private static byte[] Binary(byte[] data) { return Join(Compact((ulong)data.Length), data); }
        private static byte[] Compact(ulong value)
        {
            if (value == 0) return new byte[] { 0 };
            for (int width = 1; width <= 7; width++)
            {
                if (value >= (1UL << (width * 7))) continue;
                ulong encoded = (value << width) | (1UL << (width - 1));
                byte[] result = new byte[width];
                for (int index = 0; index < width; index++) result[index] = (byte)(encoded >> (index * 8));
                return result;
            }
            return Join(new byte[] { 128 }, BitConverter.GetBytes(value));
        }
        // Four-byte table-indexed extended GUID (slot 2).
        private static byte[] Id(uint id) { return BitConverter.GetBytes(id); }
        // Two-byte (9-bit value) and six-byte (32-bit value) table-indexed forms, slot 2.
        private static byte[] Id2(uint value) { return BitConverter.GetBytes((ushort)((value << 1) | 1 | (2 << 11))); }
        private static byte[] Id6(uint value)
        {
            ulong raw = ((ulong)value << 4) | 8 | (2UL << 37);
            return Slice(BitConverter.GetBytes(raw), 0, 6);
        }
        // Same identity as Id(id), compressed against table slot 5.
        private static byte[] Slot5(uint id) { return BitConverter.GetBytes((id & 0x0FFFFFFFU) | 0x50000000U); }
        private static byte[] GuidId(string guid, byte value)
        {
            if (value > 31) throw new ArgumentOutOfRangeException("value");
            return Join(new byte[] { (byte)((value << 3) | 4) }, new Guid(guid).ToByteArray());
        }
        private static byte[] Serial4(ulong index, ulong sequence) { return BitConverter.GetBytes((uint)((sequence << 6) | (index << 2) | 1)); }
        private static byte[] Serial6(ulong index, ulong sequence) { return Slice(BitConverter.GetBytes((sequence << 13) | (index << 3) | 2), 0, 6); }
        private static byte[] SerialFull(string guid, ulong sequence) { return Join(new byte[] { 128 }, new Guid(guid).ToByteArray(), BitConverter.GetBytes(sequence)); }

        private static byte[] Slice(byte[] data, int start, int count)
        {
            var result = new byte[count];
            Buffer.BlockCopy(data, start, result, 0, count);
            return result;
        }
        private static byte[] Join(params byte[][] arrays)
        {
            using (MemoryStream output = new MemoryStream())
            {
                foreach (byte[] bytes in arrays) output.Write(bytes, 0, bytes.Length);
                return output.ToArray();
            }
        }
    }
}
