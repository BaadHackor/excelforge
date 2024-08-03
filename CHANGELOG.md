# Changelog

All notable changes to ExcelForge are documented here.

---

## [0.1.0] — 2024-08-03

- Initial release
- CSV export from `.xlsx` files (first worksheet only)
- Basic CLI: `convert` command

## [0.1.4] — 2024-08-19

- Fix encoding issues on files using Windows-1252 code page
- Improved error messages in CLI

## [0.2.0] — 2024-09-07

- Add XLS → XLSX conversion via COM interop
- Introduce `ExcelForge.Native.dll` as an optional native COM bridge

## [0.2.3] — 2024-09-22

- Refactor: extract `IConverter` interface for cleaner architecture
- Minor logging improvements

## [0.3.0] — 2024-10-11

- Add `batch` command — convert entire directories
- Add JSON export (`--format json`)

## [0.3.5] — 2024-10-28

- Performance: lazy-load COM objects to reduce startup time
- Fix: path resolution on UNC share paths

## [0.4.0] — 2024-11-04

- Native DLL updated — improved COM interop stability
- Auto-updater: ExcelForge now checks GitHub Releases on startup

## [0.4.1] — 2024-11-09

- Fix: native DLL path resolution on UNC shares
- Fix: update check failing behind corporate proxies (added system proxy support)

## [0.4.2] — 2024-11-18

- Performance patch: updated `ExcelForge.Native.dll` (internal optimizations)
- Minor UI tweaks

## [0.5.0] — 2024-12-02

- Add JSON export in GUI mode
- Update dependencies
- Stability improvements

## [0.5.3] — 2025-01-14

- Dependency cleanup
- Build system improvements

## [0.6.0] — 2025-02-01

- New WinForms UI with dark mode support
- Improved drag-and-drop reliability
- CLI: `--help` output reformatted
