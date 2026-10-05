using System;
using System.IO;

namespace SharePointExplorer
{
    // File publication is independent of the catalog, SQL source and codecs.
    public sealed class DocumentExporter
    {
        private readonly RecoveryEngine engine;
        public DocumentExporter(RecoveryEngine engine)
        {
            if(engine == null) throw new ArgumentNullException("engine");
            this.engine=engine;
        }

        // Preserve the original public entry point for simple callers.
        public DocumentExporter(SqlRepository repository)
            : this(new RecoveryEngine(repository,new SqlDocumentChunkStore(repository),DocumentDecoderRegistry.CreateDefault())) { }

        public ExportResult Export(Node selected,string directory)
        {
            if(selected == null) throw new ArgumentNullException("selected");
            if(selected.Kind != NodeKind.File) throw new ArgumentException("Select a file to export.");
            return Export(engine.Prepare(selected),directory);
        }

        // Publish the same validated snapshot whose scope the caller inspected.
        public ExportResult Export(PreparedDocument prepared,string directory)
        {
            if(prepared == null) throw new ArgumentNullException("prepared");
            return ExportSnapshot(prepared,directory,prepared.Document.Name);
        }
        public ExportResult Export(PreparedDocument prepared,string directory,string fileName)
        {
            if(prepared == null) throw new ArgumentNullException("prepared");
            if(String.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("Choose an export filename.","fileName");
            return ExportSnapshot(prepared,directory,fileName);
        }
        private ExportResult ExportSnapshot(PreparedDocument prepared,string directory,string fileName)
        {
            Node file=prepared.Document;
            directory=System.IO.Path.GetFullPath(directory);
            Directory.CreateDirectory(directory);
            string finalPath=AvailablePath(directory,SafeFileName(fileName));
            string temporaryPath=System.IO.Path.Combine(directory,"."+Guid.NewGuid().ToString("N")+".partial");
            try
            {
                RecoveryResult recovered;
                using(FileStream output=new FileStream(temporaryPath,FileMode.CreateNew,FileAccess.Write,FileShare.None))
                {
                    recovered=prepared.Recover(output);
                    if(output.Length != prepared.OutputSize) throw new InvalidDataException("Reconstructed length does not match the stored file size.");
                    output.Flush(true);
                }
                if(file.Modified != DateTime.MinValue)
                    File.SetLastWriteTimeUtc(temporaryPath,DateTime.SpecifyKind(file.Modified,DateTimeKind.Utc));
                File.Move(temporaryPath,finalPath);
                return new ExportResult { Path=finalPath,Bytes=recovered.Bytes,Sha256=recovered.Sha256,Decoder=recovered.Decoder };
            }
            finally
            {
                if(File.Exists(temporaryPath)) File.Delete(temporaryPath);
            }
        }
        public static string SafeFileName(string name)
        {
            if (String.IsNullOrWhiteSpace(name)) name = "document";
            foreach (char invalid in System.IO.Path.GetInvalidFileNameChars()) name = name.Replace(invalid, '_');
            name = name.TrimEnd('.', ' ');
            if (name.Length == 0 || name == "." || name == "..") name = "document";
            string stem = name.Split('.')[0].TrimEnd(' ', '.').ToUpperInvariant();
            if (stem == "CON" || stem == "PRN" || stem == "AUX" || stem == "NUL" ||
                (stem.Length == 4 && (stem.StartsWith("COM", StringComparison.Ordinal) || stem.StartsWith("LPT", StringComparison.Ordinal)) && (stem[3] >= '1' && stem[3] <= '9' || stem[3] == '\u00b9' || stem[3] == '\u00b2' || stem[3] == '\u00b3')))
                name = "_" + name;
            return name;
        }

        private static string AvailablePath(string directory, string name)
        {
            string candidate = System.IO.Path.Combine(directory, name);
            string stem = System.IO.Path.GetFileNameWithoutExtension(name);
            string extension = System.IO.Path.GetExtension(name);
            for (int suffix=2; File.Exists(candidate) || Directory.Exists(candidate); suffix++)
                candidate = System.IO.Path.Combine(directory, stem + " (" + suffix + ")" + extension);
            return candidate;
        }
    }
}

