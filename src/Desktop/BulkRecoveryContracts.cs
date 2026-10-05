using System;
using System.Collections.Generic;
using System.Threading;
namespace SharePointExplorer.Desktop
{
    public interface IExplorerBulkExportController
    {
        DesktopExportSummary ExportScopeAsZip(Node scope,string archivePath,CancellationToken cancellationToken,Action<DesktopExportProgress> progress);
        DesktopExportSummary ExportVersions(IEnumerable<Node> documentsOrVersions,string directory,CancellationToken cancellationToken,Action<DesktopExportProgress> progress);
        DesktopExportSummary ExportMigrationPackage(Node list,string directory,bool includeHistory,CancellationToken cancellationToken,Action<DesktopExportProgress> progress);
    }
}
