#nullable enable
using System;
using System.Globalization;

namespace SharePointExplorer.Desktop.ViewModels
{
    // Version choices preserve their exact source identity independently of checks
    // in the current-folder file list.
    public sealed class VersionViewModel
    {
        private static readonly DocumentDecoderRegistry decoders = DocumentDecoderRegistry.CreateDefault();
        private readonly Node version;
        public VersionViewModel(Node version)
        {
            ArgumentNullException.ThrowIfNull(version);
            this.version = ExplorerNode.Copy(version);
        }
        public Node Node => ExplorerNode.Copy(version);
        public string Name => version.Name ?? "";
        public string Path => version.Path ?? "";
        public string VersionText => FormatVersion(version.UiVersion);
        public string StateText => version.HistoryVersion == 0 ? "Current" : "Historical";
        public string SizeText => ExplorerNode.Size(version.Size);
        public string ModifiedText => version.Modified == DateTime.MinValue ? "" : version.Modified.ToString("g", CultureInfo.CurrentCulture);
        public bool CanExport => version.Kind == NodeKind.File && version.HistoryVersion >= 0 && version.HasStream != false && decoders.Supports(version.StreamSchema);
        public string StatusText => CanExport ? "Available" : "Browse only";
        public int HistoryVersion => version.HistoryVersion;
        public int UiVersion => version.UiVersion;
        public byte Level => version.Level;
        public static string FormatVersion(int uiVersion)
            => uiVersion < 0 ? "Unknown" : (uiVersion / 512).ToString(CultureInfo.InvariantCulture) + "." + (uiVersion % 512).ToString(CultureInfo.InvariantCulture);
    }
}