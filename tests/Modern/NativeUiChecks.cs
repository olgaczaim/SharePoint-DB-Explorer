#nullable enable
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace SharePointExplorer.Tests
{
    internal static class NativeUiChecks
    {
        public static async Task RunAsync(bool integration)
        {
            if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("Native desktop checks require Windows.");
            string workspace = FindWorkspace();
            string application = Path.Combine(workspace, "bin", "Desktop", "SharePointExplorer.Desktop.exe");
            if (!File.Exists(application))
                throw new FileNotFoundException("Publish the desktop application into bin/Desktop before running --native.", application);

            string scratch = Path.Combine(workspace, ".scratch");
            Directory.CreateDirectory(scratch);
            // The application's diagnostic filenames are shared. Avoid accepting another run's log.
            using var runLock = new FileStream(Path.Combine(scratch, "native-ui-check.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            string logPath = Path.Combine(scratch, integration ? "winui-runtime-check-integration.txt" : "winui-runtime-check.txt");
            var startInfo = new ProcessStartInfo(application)
            {
                WorkingDirectory = workspace,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            startInfo.ArgumentList.Add(integration ? "--ui-integration" : "--ui-check");
            DateTime startedAt = DateTime.UtcNow;
            using Process process = Process.Start(startInfo) ?? throw new InvalidOperationException("The native diagnostic process could not start.");
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(90));
            try
            {
                await process.WaitForExitAsync(timeout.Token);
            }
            catch (OperationCanceledException) when (timeout.IsCancellationRequested)
            {
                // This handle identifies only the diagnostic process created above.
                if (!process.HasExited) process.Kill();
                throw new TimeoutException("Native desktop checks exceeded 90 seconds. See " + logPath);
            }

            if (!File.Exists(logPath))
                throw new InvalidOperationException("The native application produced no diagnostic log. Exit code: " + process.ExitCode);
            string log = File.ReadAllText(logPath);
            Console.Write(log);
            Match start = Regex.Match(log, @"^START (?<time>\S+) mode=(?<mode>fixture|integration)\r?$", RegexOptions.Multiline);
            bool fresh = File.GetLastWriteTimeUtc(logPath) >= startedAt && start.Success &&
                DateTimeOffset.TryParse(start.Groups["time"].Value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTimeOffset logStartedAt) &&
                logStartedAt.UtcDateTime >= startedAt && start.Groups["mode"].Value == (integration ? "integration" : "fixture");
            if (!fresh) throw new InvalidOperationException("The native application did not create a fresh log for this run: " + logPath);

            Match pass = Regex.Match(log, @"^PASS assertions=(?<count>\d+)\r?$", RegexOptions.Multiline);
            bool passed = pass.Success && int.TryParse(pass.Groups["count"].Value, out int assertionCount) && assertionCount > 0 &&
                Regex.IsMatch(log, @"^CLOSE completed\r?$", RegexOptions.Multiline) &&
                Regex.IsMatch(log, @"^END exitCode=0\r?$", RegexOptions.Multiline) &&
                !Regex.IsMatch(log, @"^(FAIL|CLOSE FAILED)\b", RegexOptions.Multiline);
            if (process.ExitCode != 0 || !passed)
                throw new InvalidOperationException("Native desktop checks failed. Exit code: " + process.ExitCode + ". See " + logPath);
            Console.WriteLine("PASS published native desktop process, fresh diagnostics and clean shutdown");
        }

        private static string FindWorkspace()
        {
            foreach (string location in new[] { Environment.CurrentDirectory, AppContext.BaseDirectory })
                for (DirectoryInfo? directory = new DirectoryInfo(location); directory != null; directory = directory.Parent)
                    if (File.Exists(Path.Combine(directory.FullName, "SharePointExplorer.Modern.Tests.csproj"))) return directory.FullName;
            throw new DirectoryNotFoundException("Run native checks from the SharePoint Explorer workspace.");
        }
    }
}
