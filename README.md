# ExcelForge

> Open-source Excel automation utility for Windows — convert, export, and batch-process workbooks from the CLI or a lightweight GUI.

[![Build](https://github.com/BaadHackor/excelforge/actions/workflows/build.yml/badge.svg)](https://github.com/BaadHackor/excelforge/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Latest release](https://img.shields.io/github/v/release/BaadHackor/excelforge)](https://github.com/BaadHackor/excelforge/releases)

---

## Features

- **XLS → XLSX** conversion (including legacy Excel 97-2003 files via native COM bridge)
- **CSV export** — RFC 4180 compliant, UTF-8 with BOM
- **JSON export** — first row as field names, typed values (number, bool, date, string)
- **Batch mode** — process entire directories in one command
- Drag-and-drop GUI + full CLI interface

## Installation

Download the latest release from the [Releases page](https://github.com/BaadHackor/excelforge/releases).

Unzip to any directory and run `ExcelForge.exe`. No installer required.

**Requirements:** Windows 10/11 x64, .NET 8 Runtime (downloaded automatically if missing).

## CLI usage

```
excelforge convert report.xls  report.xlsx
excelforge convert data.xlsx   export.csv  --format csv
excelforge convert sheet.xlsx  data.json   --format json
excelforge batch   ./invoices  --format csv --out ./csv_out
excelforge --version
```

## Building from source

```powershell
git clone https://github.com/BaadHackor/excelforge.git
cd excelforge
dotnet build src/ExcelForge/ExcelForge.csproj -c Release
```

### Native bridge (optional)

The `ExcelForge.Native.dll` component enables legacy `.xls` conversion via COM. It is distributed as a pre-built binary in each release. If you are only working with `.xlsx` files, the application will run without it.

To build the native component yourself:

```
cd src/ExcelForge.Native
cmake -B build -A x64
cmake --build build --config Release
```

## Architecture

```
ExcelForge/
├── src/
│   ├── ExcelForge/           # C# .NET 8 — main application
│   │   ├── Converters/       # IConverter implementations
│   │   ├── Native/           # P/Invoke wrapper for the native DLL
│   │   ├── UI/               # WinForms main window
│   │   └── ExcelForge.Native/    # C++ x64 native COM bridge (DLL)
└── .github/workflows/        # CI/CD
```

## License

MIT — see [LICENSE](LICENSE).

## Author

Alex Mercer — contributions welcome via pull request.
