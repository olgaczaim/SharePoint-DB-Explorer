using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using SharePointExplorer.Desktop;

namespace SharePointExplorer.Tests
{
    internal static class ModernTestRunner
    {
        public static async Task<int> Main(string[] args)
        {
            try
            {
                // Checks log damaged storage to the scratch area, not the user's log folder.
                RecoveryLog.Folder=Path.GetFullPath(Path.Combine(".scratch","logs"));
                StorageDecoderChecks.Run();
                ContentSchemaChecks.Run();
                NativeBlobRecoveryChecks.Run();
                ResidentCompressionChecks.Run();
                OneNoteRecoveryChecks.Run();
                DeletedRecoveryChecks.Run();
                AttachmentCatalogChecks.Run();
                MigrationCatalogChecks.Run();
                MigrationPackageChecks.Run();
                RecoveryEngineChecks.Run();
                DesktopExportChecks.Run();
                LibraryExportChecks.Run();
                ZipExportChecks.Run();
                BulkExportChecks.Run();
                VersionRecoveryChecks.Run();
                ControllerVersionExportChecks.Run(Path.GetFullPath(Path.Combine(".scratch","controller-version-checks",Guid.NewGuid().ToString("N"))));
                CatalogChecks.Run();
                SqlConnectionChecks.Run();
                await ExplorerViewModelChecks.RunAsync();
                await ExtendedExportViewModelChecks.RunAsync();
                bool integration=args.Contains("--integration"),migrationIntegration=args.Contains("--migration-integration");
                if(args.Contains("--package-history-integration"))MigrationPackageSqlChecks.Run(new SqlConnectionOptions { Server = "SQL", Database = "WSS_Content" });
                if (integration) CheckDatabase();
                else if(migrationIntegration)
                {
                    var options=new SqlConnectionOptions { Server = "SQL", Database = "WSS_Content" };
                    DeletedRecoveryChecks.RunSql(options);
                    AttachmentCatalogChecks.RunSql();
                    MigrationSqlChecks.Run(options);
                    MigrationPackageSqlChecks.Run(options);
                }
                if (args.Contains("--native")) await NativeUiChecks.RunAsync(integration || migrationIntegration);
                Console.WriteLine("All modern recovery and interaction checks passed.");
                return 0;
            }
            catch (Exception error) { Console.Error.WriteLine("FAIL " + error); return 1; }
        }

        private static void CheckDatabase()
        {
            var options = new SqlConnectionOptions { Server = "SQL", Database = "WSS_Content" };
            List<string> databases = SqlDatabaseDiscovery.GetAccessibleDatabases(options);
            if (!databases.Contains(options.Database)) throw new Exception("Test database was absent from discovery.");
            using (var controller = ExplorerController.Connect(options))
            {
                List<Node> roots = controller.GetRootSites();
                if (roots.Count == 0) throw new Exception("Modern SQL provider returned no sites.");
                Guid siteId = new Guid("cbd6e0be-6ee8-4b9d-9b04-d9031327831b");
                Node jpeg = new Node { Kind = NodeKind.File, SiteId = siteId, Id = new Guid("44da1cf3-e59f-49bb-b618-bc69c5f5ddcc") };
                Node text = new Node { Kind = NodeKind.File, SiteId = siteId, Id = new Guid("6afefc32-a65e-4c67-82b6-8145cec99fbd") };
                string output = Path.GetFullPath(Path.Combine(".scratch", "modern-integration", Guid.NewGuid().ToString("N")));
                DesktopExportSummary summary = controller.ExportFiles(new[] { jpeg, text }, output, CancellationToken.None, null);
                if (summary.Success != 2 || summary.Failed != 0 || summary.Skipped != 0) throw new Exception("Modern SQL file recovery failed.");
                if (summary.Entries[0].Bytes != 10886 || summary.Entries[0].Sha256 != "cb5357c9daea4fe0e71d5296eb41e7df526e768e3cf236000e6a4d5bf43b8e0a") throw new Exception("Modern provider current Word checksum changed.");
                if (summary.Entries[1].Bytes != 1000) throw new Exception("Modern text export size changed.");
                if(!File.Exists(summary.ReportPath) || Path.GetDirectoryName(summary.ReportPath)!=output ||
                    summary.Entries.Any(entry=>!File.Exists(entry.Path) || Path.GetDirectoryName(entry.Path)!=output))
                    throw new Exception("Checked current documents were not saved directly beside their report in the selected folder.");
                Console.WriteLine("PASS modern Microsoft.Data.SqlClient Windows authentication, discovery and current Word/text recovery: " + output);
                CheckZip(controller,new[] {jpeg,text},summary);
                CheckLibrary(options,controller,siteId,text.Id,summary.Entries[1].Sha256);
                CheckDefaultForms(options,controller,summary.Entries[0]);
                VersionSqlChecks.Run(options,controller);
                ContentSchemaChecks.RunSql(options);
                OneNoteRecoveryChecks.RunSql(options,controller);
                DeletedRecoveryChecks.RunSql(options);
                AttachmentCatalogChecks.RunSql();
                MigrationSqlChecks.Run(options);
                MigrationPackageSqlChecks.Run(options);
            }
        }

        private static void CheckZip(ExplorerController controller,Node[] selected,DesktopExportSummary recoveredFiles)
        {
            string output=Path.GetFullPath(Path.Combine(".scratch","modern-zip-integration",Guid.NewGuid().ToString("N")));
            string destination=Path.Combine(output,"selected-files.zip");
            DesktopExportSummary summary=controller.ExportFilesAsZip(selected,destination,CancellationToken.None,null);
            if(summary.Cancelled || summary.Total!=2 || summary.Success!=2 || summary.Failed!=0 || summary.Skipped!=0 ||
                summary.ArchivePath!=destination || summary.ArchiveReportEntry!="export-report.csv" || !String.IsNullOrEmpty(summary.ReportPath))
                throw new Exception("SQL ZIP export did not publish the expected completed selection.");
            using ZipArchive archive=ZipFile.OpenRead(summary.ArchivePath);
            if(archive.Entries.Count!=4 || archive.GetEntry("export-report.csv")==null || archive.GetEntry("summary.txt")==null)
                throw new Exception("SQL ZIP contains missing or unexpected entries.");
            foreach(DesktopExportEntry result in summary.Entries)
            {
                DesktopExportEntry original=recoveredFiles.Entries.Single(entry=>entry.Document.Id==result.Document.Id);
                string expectedPath=original.Document.Id==Guid.Parse("44da1cf3-e59f-49bb-b618-bc69c5f5ddcc")
                    ? "cbd6e0be-6ee8-4b9d-9b04-d9031327831b/Shared Documents/Document.docx"
                    : original.Document.Id==Guid.Parse("6afefc32-a65e-4c67-82b6-8145cec99fbd")
                    ? "cbd6e0be-6ee8-4b9d-9b04-d9031327831b/Test 2/test 2 sub 1/dummy-1000-chars-lorem.txt"
                    : throw new Exception("Unexpected document in the SQL ZIP fixture.");
                ZipArchiveEntry entry=archive.GetEntry(expectedPath) ?? throw new Exception("ZIP did not preserve the document hierarchy.");
                using Stream data=entry.Open();
                string hash=Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(data)).ToLowerInvariant();
                if(result.Path!=expectedPath || result.Status!=RecoveryStatus.Success || entry.Length!=original.Bytes ||
                    result.Bytes!=original.Bytes || hash!=original.Sha256 || result.Sha256!=original.Sha256)
                    throw new Exception("SQL ZIP content disagrees with verified selected-file recovery.");
            }
            using var reportReader=new StreamReader(archive.GetEntry("export-report.csv").Open());
            string report=reportReader.ReadToEnd();
            if(report.Split('\n',StringSplitOptions.RemoveEmptyEntries).Length!=3 ||
                summary.Entries.Any(entry=>!report.Contains(entry.Sha256,StringComparison.Ordinal)))
                throw new Exception("SQL ZIP report did not preserve both checksums and results.");
            using var summaryReader=new StreamReader(archive.GetEntry("summary.txt").Open());
            string state=summaryReader.ReadToEnd();
            if(!state.Contains("Cancelled=False\n",StringComparison.Ordinal) || !state.Contains("Processed=2\n",StringComparison.Ordinal))
                throw new Exception("SQL ZIP batch summary was incomplete.");
            if(Directory.GetFiles(output,"*.partial",SearchOption.AllDirectories).Length!=0 || Directory.GetFiles(output).Length!=1)
                throw new Exception("SQL ZIP export left temporary or unreported output files.");
            Console.WriteLine("PASS SQL selected-files ZIP: current Word and nested text bytes/SHA-256 match verified exports; hierarchy and embedded reports verified: "+summary.ArchivePath);
        }

        private static void CheckDefaultForms(SqlConnectionOptions options,ExplorerController controller,DesktopExportEntry knownCurrent)
        {
            using var session=RecoverySession.OpenSql(options);
            Node library=session.Catalog.GetChildren(new Node {Kind=NodeKind.Site,SiteId=knownCurrent.Document.SiteId,WebId=knownCurrent.Document.WebId})
                .Single(node=>node.Kind==NodeKind.Library && node.ListId==Guid.Parse("31ba3d41-40a5-4e4b-b985-f50cd926a1d9"));
            Node[] inventory=session.Catalog.EnumerateCurrentFiles(library.SiteId).Where(file=>file.WebId==library.WebId && file.ListId==library.ListId).ToArray();
            Node[] exportable=inventory.Where(session.Engine.CanExport).ToArray();
            Node[] forms=inventory.Where(file=>file.HasStream==false && (file.Name=="DispForm.aspx" || file.Name=="EditForm.aspx")).ToArray();
            if(forms.Length!=2)throw new Exception("The document-library template fixture changed.");
            string destination=Path.GetFullPath(Path.Combine(".scratch","modern-quiet-library",Guid.NewGuid().ToString("N")));
            var progress=new List<DesktopExportProgress>();DesktopExportSummary result=controller.ExportLibrary(library,destination,CancellationToken.None,progress.Add);
            if(result.Total!=exportable.Length || result.Success!=exportable.Length || result.Failed!=0 || result.Skipped!=0 ||
                !result.Entries.Select(entry=>entry.Document.Id).ToHashSet().SetEquals(exportable.Select(file=>file.Id)) ||
                progress.Any(item=>item.Document!=null && !exportable.Any(file=>file.Id==item.Document.Id)))
                throw new Exception("Default forms generated export warnings/results or deleted content entered current recovery.");
            string report=File.ReadAllText(result.ReportPath);
            if(report.Contains("DispForm.aspx",StringComparison.OrdinalIgnoreCase) || report.Contains("EditForm.aspx",StringComparison.OrdinalIgnoreCase) ||
                File.ReadAllLines(result.ReportPath).Length!=exportable.Length+1)throw new Exception("Ignored default forms entered the document library export report.");
            foreach(DesktopExportEntry entry in result.Entries)
            {
                using FileStream saved=File.OpenRead(entry.Path);
                if(saved.Length!=entry.Bytes || Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(saved)).ToLowerInvariant()!=entry.Sha256)
                    throw new Exception("Quiet library export changed verified content.");
            }
            Console.WriteLine("PASS SQL "+library.Name+" quiet library export: "+exportable.Length+" current files, "+(inventory.Length-exportable.Length)+" known template files ignored, no skipped outcomes: "+destination);
        }
        private static void CheckLibrary(SqlConnectionOptions options,ExplorerController controller,Guid siteId,Guid textId,string selectedTextHash)
        {
            using var session = RecoverySession.OpenSql(options);
            Node text = session.Catalog.GetFile(siteId,textId);
            Node library = session.Catalog.GetChildren(new Node {Kind=NodeKind.Site,SiteId=siteId,WebId=text.WebId})
                .Single(node => node.Kind==NodeKind.Library && node.ListId==text.ListId);
            // An independent existing whole-source enumeration establishes membership.
            Node[] allFiles = session.Catalog.EnumerateCurrentFiles(siteId)
                .Where(file => file.WebId==library.WebId && file.ListId==library.ListId).ToArray();
            Node[] expected = allFiles.Where(session.Engine.CanExport).ToArray();
            Node[] ignored = allFiles.Where(file=>!session.Engine.CanExport(file)).ToArray();
            if(!ignored.Any(file=>file.HasStream==false && file.Path.EndsWith("/Forms/DispForm.aspx",StringComparison.OrdinalIgnoreCase)) ||
                !ignored.Any(file=>file.HasStream==false && file.Path.EndsWith("/Forms/EditForm.aspx",StringComparison.OrdinalIgnoreCase)))
                throw new Exception("SQL template fixtures were absent from the independent inventory.");
            if (!text.Path.StartsWith(library.Path+"/",StringComparison.OrdinalIgnoreCase) ||
                text.Path.Substring(library.Path.Length+1).IndexOf('/')<0)
                throw new Exception("The SQL library fixture is no longer in a nested folder.");
            string output = Path.GetFullPath(Path.Combine(".scratch","modern-library-integration",Guid.NewGuid().ToString("N")));
            DesktopExportSummary summary = controller.ExportLibrary(library,output,CancellationToken.None,null);
            if (summary.Cancelled || summary.Total!=expected.Length || summary.Entries.Count!=expected.Length ||
                !summary.Entries.Select(entry=>entry.Document.Id).ToHashSet().SetEquals(expected.Select(file=>file.Id)))
                throw new Exception("Whole-library export omitted an exportable file or included a known unsupported template.");
            if (summary.Skipped!=0 || summary.Entries.Any(entry=>ignored.Any(file=>file.Id==entry.Document.Id)))
                throw new Exception("Known unsupported forms produced export warnings or results.");
            if (summary.Entries.Any(entry=>entry.Document.SiteId!=siteId || entry.Document.WebId!=library.WebId || entry.Document.ListId!=library.ListId))
                throw new Exception("SQL library export escaped its scope.");
            DesktopExportEntry recoveredText = summary.Entries.Single(entry=>entry.Document.Id==textId);
            if (recoveredText.Status!=RecoveryStatus.Success || recoveredText.Bytes!=1000 || recoveredText.Sha256!=selectedTextHash)
                throw new Exception("Whole-library export did not preserve the nested text fixture.");
            foreach (DesktopExportEntry entry in summary.Entries.Where(entry=>entry.Status==RecoveryStatus.Success))
            {
                using FileStream file = File.OpenRead(entry.Path);
                string hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(file)).ToLowerInvariant();
                if (file.Length!=entry.Bytes || hash!=entry.Sha256) throw new Exception("A library export's bytes disagree with its report.");
            }
            if (Directory.GetFiles(output,"*.partial",SearchOption.AllDirectories).Length!=0 ||
                File.ReadAllLines(summary.ReportPath).Length!=summary.Entries.Count+1)
                throw new Exception("SQL library export left partial files or an incomplete CSV report.");
            Console.WriteLine("PASS SQL whole-library export "+library.Name+": "+summary.Total+" exportable files, "+summary.Success+" exported, "+
                summary.Skipped+" runtime skips, "+summary.Failed+" failed; "+ignored.Length+" known templates/formats excluded; nested text and successful-file checksums verified: "+output);
        }
    }
}
