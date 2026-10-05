using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using SharePointExplorer.Desktop;

namespace SharePointExplorer.Tests
{
    internal static class VersionSqlChecks
    {
        // Recorded from raw historical SQL stream captures before the version API
        // existed. These are backup consistency baselines, not original-upload hashes.
        private static readonly (Guid Id,int Ui,long Bytes,string Sha)[] samples = new[]
        {
            (Guid.Parse("67960ae0-5983-45f3-8a63-9bc879876397"),1024,13708L,"31f9c8e1869cec7d692f7fffb2e59250c822597c29eb824406ba4130fbde0e4b"),
            (Guid.Parse("67960ae0-5983-45f3-8a63-9bc879876397"),512,15350L,"98d1752d430dc5e3ada0bb56386c4058f617184430a7c95a71b34393b0493ebb"),
            (Guid.Parse("16e91fbe-9e8e-4f47-8c18-a3d2462b3505"),512,18459L,"73f66648d757c9a3d1ad5006e7ae2f721337a6ab82d6aa78e8220bb6db68807c"),
            (Guid.Parse("6a4afa3f-6591-4a7f-87f0-ec867383dfd4"),512,18462L,"6051c13e558156d8ed9b0b43a8c6797a2ed6822453236567d36582e28f764e87"),
            (Guid.Parse("ee576d94-0532-46f5-8473-a7b262010299"),512,36937L,"aa328b8c9f8a73c2b3775953a6fe45ed9ff65fe15bae2019f57125fcd17b165a"),
            (Guid.Parse("44cc6e4e-8257-4da9-9182-fdc8ca18cf39"),512,2118L,"646efef993df087e956b3bbe5ef1f93a678fd6d93c86291e60cf890f4d1037c9")
        };

        public static void Run(SqlConnectionOptions options,ExplorerController controller)
        {
            Guid site=Guid.Parse("cbd6e0be-6ee8-4b9d-9b04-d9031327831b");
            string destination=Path.GetFullPath(Path.Combine(".scratch","modern-version-integration",Guid.NewGuid().ToString("N")));
            using var session=RecoverySession.OpenSql(options);
            var histories=new Dictionary<Guid,List<Node>>();
            foreach(var group in samples.GroupBy(sample=>sample.Id))
            {
                Node current=session.Catalog.GetFile(site,group.Key);
                List<Node> rows=controller.GetFileVersions(current);
                if(rows.Count!=group.Count()+1 || rows[0].HistoryVersion!=0 || rows[0].Id!=current.Id ||
                    !rows.Skip(1).Select(row=>row.UiVersion).SequenceEqual(group.Select(sample=>sample.Ui).OrderByDescending(ui=>ui)))
                    throw new Exception("SQL version browsing did not distinguish ordered current and retained historical states.");
                histories.Add(group.Key,rows);
            }
            foreach(var sample in samples)
            {
                Node version=histories[sample.Id].Single(row=>row.HistoryVersion==sample.Ui);
                if(version.UiVersion!=sample.Ui || version.Size!=sample.Bytes || version.InternalVersion==version.UiVersion)
                    throw new Exception("Historical SQL metadata used the current state or an internal counter as the version.");
                var progress=new List<DesktopExportProgress>();
                DesktopExportSummary summary=controller.ExportVersion(version,destination,CancellationToken.None,progress.Add);
                if(summary.Total!=1 || summary.Success!=1 || summary.Skipped!=0 || summary.Failed!=0 || summary.Cancelled || summary.Entries.Count!=1)
                    throw new Exception("Historical SQL recovery did not export the chosen version.");
                CheckDirectDestination(summary,destination);
                DesktopExportEntry entry=summary.Entries.Single();
                if(entry.Document.SiteId!=site || entry.Document.Id!=sample.Id || entry.Document.HistoryVersion!=sample.Ui ||
                    entry.Document.UiVersion!=sample.Ui || entry.Document.Level!=version.Level || entry.Document.InternalVersion!=version.InternalVersion ||
                    entry.Bytes!=sample.Bytes || entry.Sha256!=sample.Sha || !File.Exists(entry.Path))
                    throw new Exception("Historical export identity, count or hash disagrees with independently captured source samples.");
                string label=sample.Ui/512+"."+sample.Ui%512;
                if(!Path.GetFileName(entry.Path).Contains("(v"+label+")",StringComparison.Ordinal) ||
                    progress.Count!=1 || progress[0].Document.HistoryVersion!=sample.Ui || progress[0].ExportPath!=entry.Path)
                    throw new Exception("Historical filename/progress lost the selected version identity.");
                using(FileStream saved=File.OpenRead(entry.Path))
                    if(saved.Length!=sample.Bytes || Digest(saved)!=sample.Sha)
                        throw new Exception("Historical published file bytes differ from the report and captured source.");
                string[] report=File.ReadAllLines(summary.ReportPath);
                string fields="\""+sample.Ui+"\",\""+sample.Ui+"\",\""+version.Level+"\",\""+version.InternalVersion+"\"";
                if(report.Length!=2 || !report[0].EndsWith("\"UiVersion\",\"HistoryVersion\",\"Level\",\"InternalVersion\"",StringComparison.Ordinal) ||
                    !report[1].EndsWith(fields,StringComparison.Ordinal) || !report[1].Contains(sample.Sha,StringComparison.Ordinal))
                    throw new Exception("Historical CSV report omitted the exact selected state.");
                if(Path.GetExtension(entry.Path)==".xaml")
                {
                    using FileStream markup=File.OpenRead(entry.Path);
                    if(System.Xml.Linq.XDocument.Load(markup).Root==null) throw new Exception("Historical workflow XML was invalid.");
                }
                else
                {
                    using ZipArchive office=ZipFile.OpenRead(entry.Path);
                    if(office.GetEntry("[Content_Types].xml")==null) throw new Exception("Historical Office content was not a complete document.");
                    foreach(ZipArchiveEntry part in office.Entries) using(Stream data=part.Open()) data.CopyTo(Stream.Null);
                    if(sample.Id==samples[0].Id)
                    {
                        string expectedSheet=sample.Ui==512 ? "6c2e2c295a5aeed490d0b06e6a2e56129fbcc8a87604f7d21f9014ab146b8b9a"
                            : "44f9eb3a26fb2f830c314c71996fbaff4b24312676d81790f36f675bcb7f642d";
                        using Stream sheet=office.GetEntry("xl/worksheets/sheet1.xml").Open();
                        if(Digest(sheet)!=expectedSheet) throw new Exception("Excel history did not preserve the version-specific worksheet content.");
                    }
                }
            }
            Node excelCurrent=histories[samples[0].Id].Single(row=>row.HistoryVersion==0);
            if(excelCurrent.UiVersion!=1536 || excelCurrent.Size!=14479)
                throw new Exception("The current Excel version no longer matches the historical fixture.");
            DesktopExportSummary currentExport=controller.ExportVersion(excelCurrent,destination,CancellationToken.None,null);
            if(currentExport.Success!=1 || currentExport.Entries[0].Bytes!=14479 ||
                samples.Where(sample=>sample.Id==excelCurrent.Id).Any(sample=>sample.Sha==currentExport.Entries[0].Sha256))
                throw new Exception("Historical recovery substituted the current Excel file.");
            CheckDirectDestination(currentExport,destination);
            CheckReportedUserVersion(session,controller,site,destination);
            Node missing=controller.GetFileVersions(session.Catalog.GetFile(site,excelCurrent.Id)).Single(row=>row.HistoryVersion==512);
            missing.HistoryVersion=9999;missing.UiVersion=9999;
            DesktopExportSummary absent=controller.ExportVersion(missing,destination,CancellationToken.None,null);
            if(absent.Success!=0 || absent.Entries.Count!=1 || absent.Entries[0].Status!=RecoveryStatus.Unavailable ||
                !String.IsNullOrEmpty(absent.Entries[0].Path) || absent.Entries[0].Bytes!=0)
                throw new Exception("A missing historical version fell back to current content.");
            if(Directory.GetFiles(destination,"*.partial",SearchOption.AllDirectories).Length!=0)
                throw new Exception("Historical recovery left incomplete files.");
            Console.WriteLine("PASS SQL retained history: six Excel/Word/PowerPoint/workflow versions, captured sizes/hashes, direct file/report destinations, distinct Excel worksheets, exact state reports and no current-content fallback: "+destination);
        }
        private static void CheckReportedUserVersion(RecoverySession session,ExplorerController controller,Guid site,string destination)
        {
            Guid fileId=Guid.Parse("44da1cf3-e59f-49bb-b618-bc69c5f5ddcc");
            const long expectedBytes=10877;
            const string expectedHash="c4fd9e425fab2cef769c8f272a7c248ba6e9977a5f6bc9268d7eafc98dd9b362";
            Node current=session.Catalog.GetFile(site,fileId);
            if(current==null)
            {
                Node[] matches=session.Catalog.EnumerateCurrentFiles(null).Where(document=>document.Id==fileId).Take(2).ToArray();
                if(matches.Length>1) throw new Exception("The reported Word fixture ID is ambiguous across site collections.");
                current=matches.SingleOrDefault();
            }
            if(current==null)
            {
                Console.WriteLine("SKIP reported Word version fixture: current document is absent from this content source.");return;
            }
            site=current.SiteId;
            List<Node> versions=controller.GetFileVersions(current);
            CheckReportedCurrentDocument(controller,current,versions,destination);
            Node version=versions.SingleOrDefault(row=>row.HistoryVersion==512 && row.UiVersion==512);
            if(version==null || version.Size!=expectedBytes)
            {
                Console.WriteLine("SKIP reported Word version fixture: recorded version 1.0 metadata is absent or changed.");return;
            }
            var progress=new List<DesktopExportProgress>();
            DesktopExportSummary summary=controller.ExportVersion(version,destination,CancellationToken.None,progress.Add);
            if(summary.Total!=1 || summary.Success!=1 || summary.Failed!=0 || summary.Skipped!=0 || summary.Entries.Count!=1)
                throw new Exception("The reported Word version recovery produced only an audit report or lost the chosen version.");
            CheckDirectDestination(summary,destination);DesktopExportEntry entry=summary.Entries.Single();
            if(entry.Document.Id!=fileId || entry.Document.SiteId!=site || entry.Document.HistoryVersion!=512 || entry.Document.UiVersion!=512 ||
                entry.Document.Level!=version.Level || entry.Document.InternalVersion!=version.InternalVersion || entry.Bytes!=expectedBytes || entry.Sha256!=expectedHash ||
                !Path.GetFileName(entry.Path).Contains("(v1.0)",StringComparison.Ordinal) || progress.Count!=1 || progress[0].ExportPath!=entry.Path)
                throw new Exception("The reported Word version identity, direct filename or digest differs from its captured historical state.");
            using(FileStream saved=File.OpenRead(entry.Path))
                if(saved.Length!=expectedBytes || Digest(saved)!=expectedHash)
                    throw new Exception("The reported Word version file bytes differ from the independently recorded 10877-byte baseline.");
            using(ZipArchive office=ZipFile.OpenRead(entry.Path))
            {
                if(office.GetEntry("[Content_Types].xml")==null || office.GetEntry("word/document.xml")==null)
                    throw new Exception("The reported Word version is not a complete DOCX file.");
                foreach(ZipArchiveEntry part in office.Entries) using(Stream data=part.Open()) data.CopyTo(Stream.Null);
            }
            string[] report=File.ReadAllLines(summary.ReportPath);
            string exactState="\"512\",\"512\",\""+version.Level+"\",\""+version.InternalVersion+"\"";
            if(report.Length!=2 || !report[0].EndsWith("\"UiVersion\",\"HistoryVersion\",\"Level\",\"InternalVersion\"",StringComparison.Ordinal) ||
                !report[1].EndsWith(exactState,StringComparison.Ordinal) || !report[1].Contains(expectedHash,StringComparison.Ordinal))
                throw new Exception("The reported Word version audit does not identify its exact historical state and published hash.");
            Console.WriteLine("PASS reported Word 1.0: direct DOCX and CSV, 10877 bytes, captured SHA-256: "+entry.Path);
        }
        private static void CheckReportedCurrentDocument(ExplorerController controller,Node current,List<Node> versions,string destination)
        {
            const long expectedBytes=10886;
            const string expectedHash="cb5357c9daea4fe0e71d5296eb41e7df526e768e3cf236000e6a4d5bf43b8e0a";
            if(current.HistoryVersion!=0 || current.UiVersion!=1024 || current.Size!=expectedBytes)
            {
                Console.WriteLine("SKIP reported current Word fixture: recorded version 2.0 metadata is absent or changed.");return;
            }
            Node currentVersion=versions.Single(row=>row.HistoryVersion==0);
            if(!SameCurrentState(currentVersion,current))
                throw new Exception("Current version browsing disagrees with the current document selected for recovery.");
            string directory=Path.Combine(destination,"reported-current-document");
            var progress=new List<DesktopExportProgress>();
            DesktopExportSummary selected=controller.ExportFiles(new[]{current},directory,CancellationToken.None,progress.Add);
            if(selected.Total!=1 || selected.Success!=1 || selected.Failed!=0 || selected.Skipped!=0 || selected.Cancelled || selected.Entries.Count!=1)
                throw new Exception("The reported current Word selection produced only an audit report or failed recovery.");
            CheckDirectDestination(selected,directory);DesktopExportEntry selectedEntry=selected.Entries.Single();
            if(!SameCurrentState(selectedEntry.Document,current) || selectedEntry.Document.Name!=current.Name || selectedEntry.Document.Path!=current.Path ||
                selectedEntry.Bytes!=expectedBytes || selectedEntry.Sha256!=expectedHash || Path.GetFileName(selectedEntry.Path)!=current.Name ||
                progress.Count!=1 || !SameCurrentState(progress[0].Document,current) || progress[0].ExportPath!=selectedEntry.Path)
                throw new Exception("The reported selected Word file did not retain current metadata, plain filename, direct destination and captured digest.");
            using(FileStream saved=File.OpenRead(selectedEntry.Path))
                if(saved.Length!=expectedBytes || Digest(saved)!=expectedHash)
                    throw new Exception("Selected current Word bytes differ from the independently recorded 10886-byte baseline.");
            List<string[]> report=ReadCsv(selected.ReportPath);
            string[] expectedHeader={"Source","SiteId","DocumentId","DocumentPath","Status","Bytes","Sha256","ExportPath","Message"};
            if(report.Count!=2 || !report[0].SequenceEqual(expectedHeader) || report[1].Length!=9 || report[1][0]!=controller.SourceName ||
                report[1][1]!=current.SiteId.ToString("D") || report[1][2]!=current.Id.ToString("D") || report[1][3]!=current.Path ||
                report[1][4]!="Success" || report[1][5]!=expectedBytes.ToString(CultureInfo.InvariantCulture) || report[1][6]!=expectedHash ||
                report[1][7]!=selectedEntry.Path || report[1][8]!=String.Empty)
                throw new Exception("The selected current Word nine-column report does not describe its actual direct file and current source metadata.");
            var versionProgress=new List<DesktopExportProgress>();
            DesktopExportSummary explicitVersion=controller.ExportVersion(currentVersion,directory,CancellationToken.None,versionProgress.Add);
            if(explicitVersion.Total!=1 || explicitVersion.Success!=1 || explicitVersion.Failed!=0 || explicitVersion.Skipped!=0 ||
                explicitVersion.Cancelled || explicitVersion.Entries.Count!=1)
                throw new Exception("Explicit current-version Word recovery failed while the selected-file route succeeded.");
            CheckDirectDestination(explicitVersion,directory);DesktopExportEntry versionEntry=explicitVersion.Entries.Single();
            if(!SameCurrentState(versionEntry.Document,currentVersion) || versionEntry.Bytes!=expectedBytes || versionEntry.Sha256!=expectedHash ||
                !Path.GetFileName(versionEntry.Path).Contains("(v2.0)",StringComparison.Ordinal) || versionProgress.Count!=1 ||
                !SameCurrentState(versionProgress[0].Document,currentVersion) || versionProgress[0].ExportPath!=versionEntry.Path ||
                !File.ReadAllBytes(selectedEntry.Path).SequenceEqual(File.ReadAllBytes(versionEntry.Path)))
                throw new Exception("Checked current-file and explicitly selected version 2.0 recovery did not publish identical captured Word bytes.");
            using(FileStream saved=File.OpenRead(versionEntry.Path))
                if(saved.Length!=expectedBytes || Digest(saved)!=expectedHash)
                    throw new Exception("Explicit current-version Word bytes disagree with the selected-file recovery baseline.");
            List<string[]> versionReport=ReadCsv(explicitVersion.ReportPath);
            if(versionReport.Count!=2 || versionReport[0].Length!=13 || versionReport[1].Length!=13 || versionReport[1][6]!=expectedHash ||
                versionReport[1][7]!=versionEntry.Path || versionReport[1][9]!="1024" || versionReport[1][10]!="0" ||
                versionReport[1][11]!=currentVersion.Level.ToString(CultureInfo.InvariantCulture) ||
                versionReport[1][12]!=currentVersion.InternalVersion.ToString(CultureInfo.InvariantCulture))
                throw new Exception("The explicit current Word version audit lost its exact version key or direct published path.");
            foreach(string path in new[]{selectedEntry.Path,versionEntry.Path})
            {
                using ZipArchive office=ZipFile.OpenRead(path);
                if(office.GetEntry("[Content_Types].xml")==null || office.GetEntry("word/document.xml")==null)
                    throw new Exception("The reported current Word recovery is not a complete DOCX document.");
                foreach(ZipArchiveEntry part in office.Entries) using(Stream data=part.Open()) data.CopyTo(Stream.Null);
            }
            Console.WriteLine("PASS reported current Word: selected DOCX and CSV directly in chosen folder; exact version 2.0 has identical 10886-byte captured SHA-256: "+selectedEntry.Path);
        }
        private static bool SameCurrentState(Node actual,Node expected)=>actual.Kind==NodeKind.File && actual.SiteId==expected.SiteId &&
            actual.Id==expected.Id && actual.WebId==expected.WebId && actual.ListId==expected.ListId && actual.HistoryVersion==0 &&
            actual.UiVersion==expected.UiVersion && actual.Level==expected.Level && actual.InternalVersion==expected.InternalVersion &&
            actual.Size==expected.Size && actual.StreamSchema==expected.StreamSchema;
        private static List<string[]> ReadCsv(string path)
        {
            string text=File.ReadAllText(path);var rows=new List<string[]>();var fields=new List<string>();var value=new StringBuilder();bool quoted=false;
            for(int index=0;index<text.Length;index++)
            {
                char character=text[index];
                if(character=='"')
                {
                    if(quoted && index+1<text.Length && text[index+1]=='"') {value.Append('"');index++;}
                    else quoted=!quoted;
                }
                else if(!quoted && character==',') {fields.Add(value.ToString());value.Clear();}
                else if(!quoted && (character=='\r' || character=='\n'))
                {
                    fields.Add(value.ToString());value.Clear();rows.Add(fields.ToArray());fields.Clear();
                    if(character=='\r' && index+1<text.Length && text[index+1]=='\n') index++;
                }
                else value.Append(character);
            }
            if(quoted) throw new Exception("The export audit contains an unterminated CSV field.");
            if(fields.Count>0 || value.Length>0) {fields.Add(value.ToString());rows.Add(fields.ToArray());}
            return rows;
        }
        private static void CheckDirectDestination(DesktopExportSummary summary,string directory)
        {
            string destination=Path.GetFullPath(directory);
            if(summary.Entries.Count!=1 || !String.Equals(summary.Directory,destination,StringComparison.OrdinalIgnoreCase) ||
                !String.Equals(Path.GetDirectoryName(Path.GetFullPath(summary.Entries[0].Path)),destination,StringComparison.OrdinalIgnoreCase) ||
                !String.Equals(Path.GetDirectoryName(Path.GetFullPath(summary.ReportPath)),destination,StringComparison.OrdinalIgnoreCase) ||
                !File.Exists(summary.Entries[0].Path) || !File.Exists(summary.ReportPath))
                throw new Exception("A chosen document version or its report was not published directly in the selected output folder.");
        }
        private static string Digest(Stream data)=>Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();
    }
}