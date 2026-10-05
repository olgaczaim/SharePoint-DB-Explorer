using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SharePointExplorer.Tests
{
    // Expected bytes for the generic reconstruction rules described in the
    // README, written independently of the reader. Inputs are serialized
    // host/data-element rows, not injected states.
    internal static class GenericReaderRuleChecks
    {
        internal static void Run()
        {
            var expected = new Dictionary<string, string>
            {
                { "node-precedes-newer-inline", "AB" },
                { "node-precedes-inline-read-first", "AB" },
                { "node-tie-keeps-first", "BA" },
                { "newer-node-replaces-older", "AB" },
                { "inline-tie-keeps-first", "FIRST" },
                { "newer-inline-replaces-older", "NEW" },
                { "partial-with-references-is-leaf", "PART" },
                { "fragments-across-groups", "AB" },
                // Fragment sequences 10, 20, 15 are all accepted; 9 is rejected.
                { "fragment-threshold-does-not-advance", "ABC" },
                { "older-fragment-than-inline-rejected", "INLINE" },
                { "equal-fragment-to-inline-accepted", "A" },
                { "duplicate-start-equal-sequence-last-wins", "LAST" },
                { "duplicate-start-last-accepted-not-highest-sequence", "LATER" },
                { "overlapping-fragments-append-complete-bytes", "ABCDxy" },
                // The overlapping second piece contributes four full output bytes,
                // so start 7 is accepted despite a gap in source offset coverage.
                { "cumulative-output-byte-gap-check", "ABCDefghI" },
                { "newer-inline-keeps-existing-fragments", "A" },
                { "inline-raises-fragment-rejection-threshold", "AB" },
                { "declared-fragment-size-and-end-ignored", "AB" },
                { "odb-does-not-override-inline", "INLINE" },
                { "fragment-odb-does-not-override-partials", "PART" },
                { "declaration-does-not-override-inline", "INLINE" },
                // Object-data BLOB fragment elements also update leaves at
                // FragmentOf using SerialNumberOf.Sequential, rather than BSN
                // or the containing host's sequence. Indexed and full serials
                // below use different widths so truncating a serial is visible.
                { "direct-odb-serial4", "S4" },
                { "direct-odb-serial6", "S6" },
                { "direct-odb-serial-full", "SF" },
                { "direct-odb-serial-threshold-not-host-sequence", "ABC" },
                { "direct-odb-duplicate-start-last-accepted-wins", "LAST" },
                { "direct-odb-overlap-appends-complete-bytes", "ABCDxy" },
                { "direct-odb-cumulative-byte-gap-check", "ABCDefghI" },
                { "direct-odb-inline-preserves-fragment-preference", "AB" },
                { "node-precedes-direct-odb-leaf", "AB" },
                { "direct-odb-declared-size-and-end-ignored", "AB" },
                // Mutation follows a complete inner element: earlier valid
                // fragments survive, and the malformed higher-sequence piece
                // cannot replace their bytes even when its item parsed fully.
                { "malformed-odb-footer-preserves-prior-valid-piece", "GOOD" }
            };
            foreach (var pair in expected)
            {
                using (var output = new MemoryStream())
                {
                    StoredDocumentReader.Write(StorageFixture.GenericReaderScenario(pair.Key), output);
                    string actual = Encoding.ASCII.GetString(output.ToArray());
                    if (actual != pair.Value)
                        throw new Exception("Generic reader rule failed for '" + pair.Key + "': expected '" + pair.Value + "', recovered '" + actual + "'.");
                }
            }
            foreach (string scenario in new[] { "fragment-gap-is-rejected", "repeated-leaf-is-rejected",
                "direct-odb-gap-is-rejected", "malformed-odb-footer-does-not-register-piece",
                "null-odb-serial-does-not-register-piece" })
            {
                using (var output = new MemoryStream())
                {
                    try { StoredDocumentReader.Write(StorageFixture.GenericReaderScenario(scenario), output); }
                    catch (InvalidDataException)
                    {
                        if (output.Length != 0) throw new Exception("A rejected generic-reader scenario published bytes: " + scenario);
                        continue;
                    }
                    catch (NotSupportedException)
                    {
                        if (output.Length != 0) throw new Exception("A rejected generic-reader scenario published bytes: " + scenario);
                        continue;
                    }
                    throw new Exception("Generic reader rules accepted '" + scenario + "'.");
                }
            }
            Console.WriteLine("PASS generic reader rules: independent node/leaf states, sequence ties, fragment thresholds, duplicate offsets, full overlaps, cumulative byte bounds, fragment preference, ignored ODB declarations, direct ODB serial4/6/full leaves, footer mutation boundaries and consumed-leaf rejection");
        }
    }
}
