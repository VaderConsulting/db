# db

VB6 Kixtart log importer ("Troll") that walks a UNC log share, lists `*.log` files, and loads hostname plus file datetime into a temporary SQL Server table. Points at paths such as `\\cbdxaai\kixlog$\...` and uses SQLOLEDB against the EUC catalog on PERTHXSAC with integrated security. Start fills the file list and runs the extract into `tblDaveTemp`.

**Source last updated:** 2026-08-27 · **Language:** VB6 · **Target:** VB6 Win32 · **Output:** WinForms exe

_Note: original OneDrive LastWriteTime values were wiped to 2026-08-27 by a zip transfer; date above uses best available evidence (headers/copyright where helpful)._

## Solution structure

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| `Project1` (`Troll.vbp`) | VB6 | WinForms exe | Kixtart log share scanner / SQL import |

## How to open

Open the `.vbp` in Visual Basic 6.0 IDE:
- `Troll.vbp`

## Requirements

- Visual Basic 6.0 IDE
- Microsoft ActiveX Data Objects 2.0 Library
- Registered OCX/DLL dependencies referenced by the `.vbp` (may need to be installed separately):
  - `MSADODC.OCX`
  - `Dblist32.ocx`
  - `MSBIND.DLL`

## Attribution and provenance

Working copy from Dave Robinson's OneDrive Historical Dev folder `VB/Old/db`.
Company names in project files: CSC.

## License

MIT © 2026 VaderConsulting for Dave Robinson's code. See `LICENSE`.
