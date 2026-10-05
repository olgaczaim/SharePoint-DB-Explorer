# Application downloads

The included build is **1.0.0 for Windows x64**. Each version directory contains:

- **Installer bundle:** extract `SharePointDatabaseExplorerInstaller-v<version>-win-x64.zip` and run `setup.exe`.
- **Individual installer files:** `setup.exe` requires the matching `SharePointDatabaseExplorerSetup.msi` beside it, with both original names.
- **Portable application:** extract the entire `SharePointDatabaseExplorer-v<version>-win-x64.zip` and run `SharePointExplorer.Desktop.exe`.
- **Checksums:** compare `Get-FileHash -Algorithm SHA256` with `SHA256SUMS.txt`.

See [the 1.0.0 downloads](v1.0.0/) and [the application README](../README.md#download-and-install) for requirements and usage. The source and project license are available in this repository.

Future builds can be added as new version directories or uploaded to the repository's [GitHub Releases](https://github.com/olgaczaim/SharePoint-DB-Explorer/releases). Maintainers should follow [Create and publish a release](../README.md#create-and-publish-a-release); build output directories remain ignored.