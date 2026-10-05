using System;
using System.Globalization;
using System.IO;
using System.Threading;

namespace SharePointExplorer
{
    // Application log. Each event is one line appended to a daily file under
    // %LOCALAPPDATA%\SharePointExplorer\Logs. Recovery code reports events here
    // that do not stop an export, such as skipped damaged storage. Writing the
    // log never fails a recovery.
    public static class RecoveryLog
    {
        private static readonly object gate = new object();
        private static readonly AsyncLocal<string> document = new AsyncLocal<string>();
        private static string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SharePointExplorer", "Logs");

        public static string Folder
        {
            get { lock (gate) return folder; }
            set
            {
                if (String.IsNullOrWhiteSpace(value)) throw new ArgumentException("A log folder is required.", "value");
                lock (gate) folder = value;
            }
        }
        public static string CurrentFile
        {
            get { return Path.Combine(Folder, "SharePointExplorer-" + DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + ".log"); }
        }

        // Names the document whose recovery is running on this flow; log lines
        // written until the returned scope is disposed carry its identity.
        public static IDisposable ForDocument(Node node)
        {
            string previous = document.Value;
            document.Value = node == null ? null : "site " + node.SiteId.ToString("D") + ", document " + node.Id.ToString("D") +
                ", version " + node.UiVersion.ToString(CultureInfo.InvariantCulture) + ", level " + node.Level.ToString(CultureInfo.InvariantCulture) +
                (String.IsNullOrEmpty(node.Path) ? "" : ", " + node.Path);
            return new Scope(previous);
        }

        public static void Warning(string message) { Write("WARN", message); }

        private static void Write(string level, string message)
        {
            try
            {
                string scope = document.Value;
                string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture) + " " + level + " " +
                    (scope == null ? "" : "[" + scope + "] ") + message + Environment.NewLine;
                lock (gate)
                {
                    Directory.CreateDirectory(folder);
                    File.AppendAllText(Path.Combine(folder, "SharePointExplorer-" + DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + ".log"), line);
                }
            }
            catch (Exception)
            {
                // An unwritable log must not change the recovery outcome.
            }
        }

        private sealed class Scope : IDisposable
        {
            private readonly string previous;
            private bool disposed;
            internal Scope(string previous) { this.previous = previous; }
            public void Dispose()
            {
                if (disposed) return;
                disposed = true;
                document.Value = previous;
            }
        }
    }
}
