using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace SharePointExplorer
{
    // Verify each document before adding its entry; publish only a closed ZIP.
    public sealed class ValidatedZipArchive : IDisposable
    {
        public const string ReportEntry = "export-report.csv";
        public const string SummaryEntry = "summary.txt";
        private readonly string directory,requestedName,temporaryPath;
        private readonly FileStream output;
        private ZipArchive archive;
        private readonly HashSet<string> files=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> folders=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string,string> directoryNames=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        private bool disposed,closed;
        public bool IsFaulted { get; private set; }

        public ValidatedZipArchive(string archivePath)
        {
            if(String.IsNullOrWhiteSpace(archivePath)) throw new ArgumentException("Choose a ZIP filename.","archivePath");
            archivePath=Path.GetFullPath(archivePath);
            if(!Path.GetExtension(archivePath).Equals(".zip",StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Choose a filename ending in .zip.","archivePath");
            directory=Path.GetDirectoryName(archivePath);
            requestedName=DocumentExporter.SafeFileName(Path.GetFileName(archivePath));
            Directory.CreateDirectory(directory);
            temporaryPath=Path.Combine(directory,"."+Guid.NewGuid().ToString("N")+".partial");
            output=new FileStream(temporaryPath,FileMode.CreateNew,FileAccess.ReadWrite,FileShare.None);
            try { archive=new ZipArchive(output,ZipArchiveMode.Create,true,Encoding.UTF8); }
            catch { output.Dispose(); File.Delete(temporaryPath); throw; }
            files.Add(ReportEntry); files.Add(SummaryEntry);
        }

        public ExportResult Add(PreparedDocument prepared)
        {
            EnsureOpen();
            if(prepared==null) throw new ArgumentNullException("prepared");
            Node document=prepared.Document;
            string spoolPath=Path.Combine(directory,"."+Guid.NewGuid().ToString("N")+".partial");
            bool entryStarted=false;
            try
            {
                using(var spool=new FileStream(spoolPath,FileMode.CreateNew,FileAccess.ReadWrite,FileShare.None))
                {
                    RecoveryResult recovered=prepared.Recover(spool);
                    if(spool.Length!=prepared.OutputSize) throw new InvalidDataException("Reconstructed length does not match the stored file size.");
                    spool.Position=0;
                    string name=DocumentEntryName(document);
                    try
                    {
                        entryStarted=true;
                        ZipArchiveEntry entry=archive.CreateEntry(name,CompressionLevel.Optimal);
                        if(document.Modified.Year>=1980 && document.Modified.Year<=2107)
                            entry.LastWriteTime=new DateTimeOffset(DateTime.SpecifyKind(document.Modified,DateTimeKind.Utc));
                        using(Stream contents=entry.Open()) spool.CopyTo(contents,65536);
                    }
                    catch { IsFaulted=true; throw; }
                    return new ExportResult { Path=name,Bytes=recovered.Bytes,Sha256=recovered.Sha256,Decoder=recovered.Decoder };
                }
            }
            catch { if(entryStarted) IsFaulted=true; throw; }
            finally { try { if(File.Exists(spoolPath)) File.Delete(spoolPath); } catch { IsFaulted=true; throw; } }
        }

        public string Commit(Action<TextWriter> writeReport,string summary)
        {
            EnsureOpen();
            if(writeReport==null) throw new ArgumentNullException("writeReport");
            try
            {
                using(var report=new StreamWriter(archive.CreateEntry(ReportEntry,CompressionLevel.Optimal).Open(),new UTF8Encoding(true)))
                    writeReport(report);
                using(var state=new StreamWriter(archive.CreateEntry(SummaryEntry,CompressionLevel.Optimal).Open(),new UTF8Encoding(false)))
                    state.Write(summary ?? "");
                archive.Dispose(); archive=null; closed=true;
                output.Flush(true); output.Dispose();
                // Retry a filename collision discovered at the atomic move, never overwrite.
                while(true)
                {
                    string finalPath=AvailablePath(directory,requestedName);
                    try { File.Move(temporaryPath,finalPath); return finalPath; }
                    catch(IOException) when(File.Exists(finalPath) || Directory.Exists(finalPath)) { }
                }
            }
            catch { IsFaulted=true; throw; }
        }

        private string DocumentEntryName(Node document)
        {
            string parent=DirectoryName("",document.SiteId.ToString("D"));
            string[] parts=(document.Path ?? document.Name ?? "").Split(new char[] {'/','\\'},StringSplitOptions.RemoveEmptyEntries);
            for(int index=0;index<parts.Length-1;index++) parent=DirectoryName(parent,DocumentExporter.SafeFileName(parts[index]));
            string name=DocumentExporter.SafeFileName(document.Name);
            string candidate=parent+"/"+name;
            string stem=Path.GetFileNameWithoutExtension(name),extension=Path.GetExtension(name);
            for(int suffix=2;files.Contains(candidate) || folders.Contains(candidate);suffix++)
                candidate=parent+"/"+stem+" ("+suffix+")"+extension;
            files.Add(candidate);
            return candidate;
        }
        private string DirectoryName(string parent,string name)
        {
            string key=parent+"/"+name;
            string existing;
            if(directoryNames.TryGetValue(key,out existing)) return existing;
            string candidate=parent.Length==0 ? name : parent+"/"+name;
            for(int suffix=2;files.Contains(candidate) || folders.Contains(candidate);suffix++)
                candidate=(parent.Length==0 ? "" : parent+"/")+name+" ("+suffix+")";
            folders.Add(candidate); directoryNames.Add(key,candidate);
            return candidate;
        }
        private static string AvailablePath(string directory,string name)
        {
            string candidate=Path.Combine(directory,name);
            string stem=Path.GetFileNameWithoutExtension(name),extension=Path.GetExtension(name);
            for(int suffix=2;File.Exists(candidate) || Directory.Exists(candidate);suffix++)
                candidate=Path.Combine(directory,stem+" ("+suffix+")"+extension);
            return candidate;
        }
        private void EnsureOpen()
        {
            if(disposed || closed) throw new ObjectDisposedException("ValidatedZipArchive");
            if(IsFaulted) throw new IOException("The temporary ZIP is no longer writable.");
        }
        public void Dispose()
        {
            if(disposed) return; disposed=true;
            try { if(archive!=null) archive.Dispose(); }
            finally
            {
                try { output.Dispose(); }
                finally { if(File.Exists(temporaryPath)) File.Delete(temporaryPath); }
            }
        }
    }
}
