# VS Code VBA Workflow (Excel Runtime)

This repository is set up so VBA source lives as text files you edit in VS Code, while Excel remains the runtime and debugger.

## What this gives you

- Git version control for VBA source
- Real diffs across modules/classes/forms
- Project-wide search and multi-file editing
- Reusable snippets and editor tooling

## Runtime constraints

- Excel is still the runtime.
- The VBA engine is still required.
- Debugging still happens in Excel.

## Folder convention

Use a folder like `vba/` to store exported components:

- `.bas` -> standard modules
- `.cls` -> class modules
- `.frm` (+ `.frx`) -> userforms

Document modules (`ThisWorkbook`, `Sheet1`, etc.) are not exported/imported by this CLI.

You can also keep ribbon XML as source (for example `ribbon/customUI14.xml`) and round-trip it to `.xlam`.

## Prerequisites

1. Windows with desktop Excel installed.
2. In Excel: **Trust Center -> Macro Settings -> Trust access to the VBA project object model** enabled.
3. Python environment for this repo.

Install dependency:

```powershell
uv add pywin32
```

## CLI usage

Entry point: `excel-com`

### Export workbook VBA to files (pull from Excel)

```powershell
uv run excel-com export --workbook .\Workbook.xlsm --out .\vba
```

### Import files into workbook (push to Excel)

```powershell
uv run excel-com import --workbook .\Workbook.xlsm --src .\vba
```

### Import and remove workbook components not present on disk

```powershell
uv run excel-com import --workbook .\Workbook.xlsm --src .\vba --clean
```

### Sync convenience command

```powershell
# pull
uv run excel-com sync --workbook .\Workbook.xlsm --dir .\vba --direction pull

# push
uv run excel-com sync --workbook .\Workbook.xlsm --dir .\vba --direction push --clean
```

### Ribbon XML pull/push for `.xlam`

Extract ribbon XML from an add-in package to edit in VS Code:

```powershell
uv run excel-com ribbon pull --workbook .\MyAddin.xlam --out .\ribbon\customUI14.xml
```

Inject edited ribbon XML back into the add-in (in-place):

```powershell
uv run excel-com ribbon push --workbook .\MyAddin.xlam --xml .\ribbon\customUI14.xml
```

Inject into a new output add-in file:

```powershell
uv run excel-com ribbon push --workbook .\MyAddin.xlam --xml .\ribbon\customUI14.xml --out-workbook .\MyAddin.updated.xlam
```

## Manual import/export (Excel UI)

- Export: VBA editor -> right click module/class/form -> `Export File...`
- Import: VBA editor -> right click project -> `Import File...`

## Optional in-workbook auto-import macro

You can keep a small bootstrap module inside the workbook to import from disk:

```vb
Public Sub ImportAllFromFolder()
    ' Example only: iterate files in a known folder and import each .bas/.cls/.frm
    ' This still runs in Excel and requires Trust Access to VBProject.
End Sub
```

## Git tips

- Commit the `vba/` folder.
- Exclude generated binaries (`.xlsm`/`.xlsb`) if your team prefers source-first workflow.
- If you keep workbook binaries in git, use frequent commits and tags because diffs are limited.
