# Contributing to SharePoint DB Explorer

Thank you for your interest in improving SharePoint DB Explorer. Bug reports, fixes and new storage-format support are all welcome.

## Reporting issues

Open an issue and include:

- what you did, what you expected and what happened;
- the app version (from the release folder name or the installer) and your Windows version;
- the SharePoint version of the content database (shown after connecting) and the `StreamSchema` value of an affected file, if you know it;
- any warnings from `%LOCALAPPDATA%\SharePointExplorer\Logs` and the error text from the export report.

**Never attach content databases, exported documents or connection details.** They can contain confidential data. A description of the stored layout, or a synthetic reproduction, is enough.

## Making changes

1. Fork the repository and create a branch from `main`.
2. Follow [Building from source](README.md#build-from-source) to set up the .NET 10 SDK and build.
3. Keep the existing structure: SQL access in `src/Storage`, format decoding in `src/Serialization`, recovery and publication in `src/Core` and `src/Export`, and UI-independent logic in `src/Desktop`. The WinUI layer in `src/WinUI` should stay thin.
4. Match the style of the surrounding code: naming, comment density and formatting.
5. Add or update checks under `tests/`. Storage-format changes need a synthetic fixture in `tests/StorageFixture.cs` that serializes the stored rows, not just in-memory objects.
6. Run the local checks and make sure they all pass:

   ```powershell
   dotnet run --project SharePointExplorer.Modern.Tests.csproj -c Release
   ```

7. Open a pull request that explains the change and how you tested it.

## Ground rules

- **Read-only:** the app must never write to the source database. Only `SELECT` queries are allowed.
- **No guessed data:** if a file's bytes cannot be recovered from the stored structures, fail with a clear error instead of inventing or padding content.
- **Public sources only:** contributions must be your own work, written from public documentation such as Microsoft's Open Specifications and observations of databases you are authorized to examine. Do not contribute code, text or binaries copied or decompiled from other products.
- **Dependencies:** pinned in the `*.packages.lock.json` files; update them deliberately with `dotnet restore --force-evaluate`.

By contributing, you agree that your contributions are licensed under the project's [GPL-3.0 license](LICENSE).
