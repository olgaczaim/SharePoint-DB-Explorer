#nullable enable
using System;
using System.Collections.Generic;
using System.IO;

namespace SharePointExplorer.Tests
{
    internal static class StorageDecoderChecks
    {
        public static void Run()
        {
            TestNames();
            TestStorageTree();
            TestBlobGroupAliases();
            TestPersistedRepresentations();
            TestDamageLog();
            TestSchemaRouting();
            GenericReaderRuleChecks.Run();
            Console.WriteLine("PASS safe filenames, ordered storage graph, stored bytes without size/hash checks, canonical GUID aliases, fragments, data BLOBs, last-root and newest-copy selection, ignored element GUID tables, lenient layouts, logged damaged-row skipping and schema-flag routing");
        }

        private static void TestPersistedRepresentations()
        {
            byte[] alpha=System.Text.Encoding.ASCII.GetBytes("alpha-content");
            byte[] beta=System.Text.Encoding.ASCII.GetBytes("beta-data");
            byte[] newer=(byte[])beta.Clone(); Array.Reverse(newer);
            var expected=new Dictionary<string,string>
            {
                { "fragmented-tree","beta-dataalpha-content" }, { "partial-alpha","beta-dataalpha-content" },
                // Resident object-data BLOB recovery is a fallback when no inline or partial leaf exists.
                { "blob-beta","beta-dataalpha-content" },
                { "element-guid-table","beta-dataalpha-content" }, { "secondary-table","beta-dataalpha-content" }, { "non-aggregatable","beta-dataalpha-content" },
                { "compact-forms","beta-dataalpha-content" }, { "unhashed-leaves","beta-dataalpha-content" }, { "lenient-layout","beta-dataalpha-content" },
                // One object stored twice, or two copies at one sequence number: the copy read first wins.
                { "duplicate-object","beta-dataalpha-content" }, { "tied-conflicting-copy","beta-dataalpha-content" },
                // The rest of a damaged row is skipped; beta's group was read before the damage.
                { "damaged-after-beta","beta-dataalpha-content" },
                // The copy with the higher sequence number wins, whatever its revision.
                { "inherited",System.Text.Encoding.ASCII.GetString(newer)+"alpha-content" },
                { "inherited-cycle",System.Text.Encoding.ASCII.GetString(newer)+"alpha-content" },
                { "inherited-missing-base",System.Text.Encoding.ASCII.GetString(newer)+"alpha-content" },
                { "stale-higher-sequence","beta-dataalpha-content" },
                // The newer altered copy is written as stored; no stored hash is compared.
                { "newer-damaged-copy","beta-data`lpha-content" },
                // Fragment assembly appends every piece's full bytes, including overlaps.
                { "partial-overlap-conflict","beta-dataalpha-cContent" },
                // The root named by the revision read last wins.
                { "last-root-wins","alpha-contentbeta-data" }
            };
            foreach(var pair in expected) MustRead(StorageFixture.Scenario(pair.Key,alpha,beta),pair.Value,"the persisted '"+pair.Key+"' layout");
            MustReject(StorageFixture.Scenario("partial-gap",alpha,beta),"a fragmented object with a missing byte range");
            MustReject(StorageFixture.Scenario("blob-missing",alpha,beta),"an object whose data BLOB is absent");
            MustReject(StorageFixture.Scenario("undeclared-slot",alpha,beta),"an identifier compressed against an undeclared GUID table slot");
            MustReject(StorageFixture.Scenario("damaged-before-beta",alpha,beta),"an object whose group follows damage in its row");
        }

        // A skipped damaged chunk is logged with the document being recovered;
        // an intact store logs nothing.
        private static void TestDamageLog()
        {
            byte[] alpha=System.Text.Encoding.ASCII.GetBytes("alpha-content");
            byte[] beta=System.Text.Encoding.ASCII.GetBytes("beta-data");
            string folder=RecoveryLog.Folder;
            try
            {
                RecoveryLog.Folder=Path.GetFullPath(Path.Combine(".scratch","damage-log",Guid.NewGuid().ToString("N")));
                var document=new Node { Kind=NodeKind.File,SiteId=Guid.NewGuid(),Id=Guid.NewGuid(),UiVersion=512,Level=1,Path="Docs/damaged.bin" };
                using(RecoveryLog.ForDocument(document)) MustRead(StorageFixture.Scenario("plain",alpha,beta),"beta-dataalpha-content","an intact store");
                if(File.Exists(RecoveryLog.CurrentFile)) throw new Exception("An intact store wrote a damage log entry.");
                using(RecoveryLog.ForDocument(document)) MustRead(StorageFixture.Scenario("damaged-after-beta",alpha,beta),"beta-dataalpha-content","a store with a damaged row");
                string log=File.ReadAllText(RecoveryLog.CurrentFile);
                if(!log.Contains(" WARN [site "+document.SiteId.ToString("D")+", document "+document.Id.ToString("D")+", version 512, level 1, Docs/damaged.bin] Stored chunk (stream 3, BSN 2) is damaged: ") ||
                    !log.Contains("The rest of this chunk was skipped."))
                    throw new Exception("A skipped damaged chunk was not logged with its document: "+log);
            }
            finally { RecoveryLog.Folder=folder; }
        }

        private static void TestSchemaRouting()
        {
            DocumentDecoderRegistry registry=DocumentDecoderRegistry.CreateDefault();
            var routes=new Dictionary<string,byte[]>
            {
                { "Resident plain stream", new byte[] {0,1,64,65} },
                { "SharePoint generic document tree", new byte[] {2,66} },
                { "OneNote server section", new byte[] {18,82} },
                { "OneNote server notebook index", new byte[] {34,98} }
            };
            foreach(var route in routes)
                foreach(byte schema in route.Value)
                    if(!registry.Supports(schema) || registry.Resolve(schema).Name!=route.Key) throw new Exception("Stream schema "+schema+" was not routed to "+route.Key+".");
            foreach(byte schema in new byte[] {3,4,7,8,12,17,33,48,50,67,77,80,96,114,128,130,192,194,255})
            {
                if(registry.Supports(schema)) throw new Exception("Undefined stream schema "+schema+" was accepted.");
                try { registry.Resolve(schema); throw new Exception("Undefined stream schema "+schema+" resolved a decoder."); }
                catch(NotSupportedException) { }
            }
        }

        private static void TestNames()
        {
            string[] inputs = { "report.docx", "CON", "NUL.tar.gz", "COM1.test.txt", "LPT\u00b2.csv", "../../escape.txt", "a:b?.pdf", "...", "trail. " };
            string[] expected = { "report.docx", "_CON", "_NUL.tar.gz", "_COM1.test.txt", "_LPT\u00b2.csv", ".._.._escape.txt", "a_b_.pdf", "document", "trail" };
            for (int index=0; index<inputs.Length; index++)
            {
                string actual=DocumentExporter.SafeFileName(inputs[index]);
                if (actual != expected[index]) throw new Exception("Filename " + inputs[index] + " became " + actual);
            }
        }

        private static void TestStorageTree()
        {
            byte[] alpha=System.Text.Encoding.ASCII.GetBytes("alpha");
            byte[] beta=System.Text.Encoding.ASCII.GetBytes("beta");
            List<StoredChunk> chunks=StorageFixture.Create(alpha,beta);
            MustRead(chunks,"betaalpha","tree order rather than physical chunk order");
            // Altered stored bytes are written as stored; no leaf hash is compared.
            byte[] valid=chunks[0].Content;
            chunks[0].Content=(byte[])valid.Clone();
            chunks[0].Content[FindBytes(chunks[0].Content,alpha)]^=1;
            MustRead(chunks,"beta`lpha","altered content under a stored hash");
            chunks[0].Content=valid;
            var missing=new List<StoredChunk>(chunks); missing.RemoveAt(1);
            MustReject(missing,"missing referenced object");
            // A conflicting copy at the same sequence number is read later and loses.
            var conflict=new List<StoredChunk>(chunks);
            byte[] conflicting=(byte[])valid.Clone();
            conflicting[FindBytes(conflicting,alpha)]^=1;
            conflict.Add(new StoredChunk { Partition=0, Type=10, BSN=4, StreamId=2, Content=conflicting });
            MustRead(conflict,"betaalpha","a later conflicting copy at the same sequence number");
            // The tree's length is not compared with the document's declared size.
            if(((IReconstructedDocumentDecoder)DocumentDecoderRegistry.CreateDefault().Resolve(66)).Reconstruct(chunks,10).Length!=9)
                throw new Exception("The generic decoder adjusted its output to the declared document size.");
            // Damage after the last contained blob of a row loses nothing.
            byte[] malformedFooter=(byte[])valid.Clone();
            malformedFooter[malformedFooter.Length-1]=3;
            chunks[0].Content=malformedFooter;
            MustRead(chunks,"betaalpha","a row whose closing footer is damaged");
            chunks[0].Content=valid;
            // A further main row is ignored; only the first is read.
            byte[] badIndex=(byte[])chunks[2].Content.Clone();
            int rootGuid=FindBytes(badIndex,new Guid("fb737e23-2247-4ec8-8d39-bda7f137215e").ToByteArray());
            badIndex[rootGuid+17]=0;
            var duplicateIndex=new List<StoredChunk>(chunks);
            duplicateIndex.Add(new StoredChunk { Partition=0,Type=11,BSN=5,StreamId=1,Content=badIndex });
            MustRead(duplicateIndex,"betaalpha","a second main row");
            byte[] largeAlpha=new byte[5000], largeBeta=new byte[2000];
            for(int index=0;index<largeAlpha.Length;index++) largeAlpha[index]=(byte)(index*13);
            for(int index=0;index<largeBeta.Length;index++) largeBeta[index]=(byte)(index*7);
            using(var output=new MemoryStream())
            {
                StoredDocumentReader.Write(StorageFixture.Create(largeAlpha,largeBeta),output);
                byte[] actual=output.ToArray();
                for(int index=0;index<actual.Length;index++)
                    if(actual[index]!=(index<largeBeta.Length ? largeBeta[index] : largeAlpha[index-largeBeta.Length]))
                        throw new Exception("The 32-bit persisted field variant reconstructed incorrectly.");
                if(actual.Length!=7000) throw new Exception("The 32-bit persisted field variant has the wrong size.");
            }
        }

        private static void TestBlobGroupAliases()
        {
            byte[] alpha=System.Text.Encoding.ASCII.GetBytes("alpha");
            byte[] beta=System.Text.Encoding.ASCII.GetBytes("beta");
            MustRead(StorageFixture.CreateAliasedBlobGroup(alpha,beta),"betaalpha","a canonical outer blob-group identity with a compact inner alias");
            // A fragment's group identity is not compared with its blob; its objects still load.
            MustRead(StorageFixture.CreateAliasedBlobGroup(alpha,beta,wrongGuid:true),"betaalpha","a blob-group GUID different from its fragment's group");
            MustRead(StorageFixture.CreateAliasedBlobGroup(alpha,beta,wrongValue:true),"betaalpha","a blob-group value different from its fragment's group");
        }
        private static void MustRead(List<StoredChunk> chunks,string expected,string reason)
        {
            using(var output=new MemoryStream())
            {
                StoredDocumentReader.Write(chunks,output);
                if(System.Text.Encoding.ASCII.GetString(output.ToArray())!=expected) throw new Exception("The reader recovered "+reason+" incorrectly.");
            }
        }
        private static void MustReject(List<StoredChunk> chunks,string reason)
        {
            using(var output=new MemoryStream())
            {
                try { StoredDocumentReader.Write(chunks,output); }
                catch(InvalidDataException) { if(output.Length!=0) throw new Exception("Unverified bytes were written."); return; }
                catch(NotSupportedException) { if(output.Length!=0) throw new Exception("Unverified bytes were written."); return; }
                throw new Exception("The reader accepted " + reason);
            }
        }

        private static int FindBytes(byte[] haystack,byte[] needle)
        {
            for(int offset=0;offset<=haystack.Length-needle.Length;offset++)
            {
                bool equal=true;
                for(int index=0;index<needle.Length;index++)
                    if(haystack[offset+index]!=needle[index]) { equal=false; break; }
                if(equal) return offset;
            }
            throw new Exception("Synthetic raw payload was not found.");
        }

    }
}

