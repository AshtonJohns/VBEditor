# VS Code VBA Workflow (Excel Default, Word Optional)

This repository is set up so VBA source lives as text files you edit in VS Code. By default, commands target Excel, and you can optionally target Word with `--app word`.

## What this gives you

- Git version control for VBA source
- Real diffs across modules/classes/forms
- Project-wide search and multi-file editing
- Reusable snippets and editor tooling

## Runtime constraints

- Excel is the default runtime.
- Word is supported when you pass `--app word`.
- The VBA engine is still required.
- Debugging happens in the Office host app you use.

## Folder convention

Use a folder like `vba/` to store exported components:

- `.bas` -> standard modules
- `.cls` -> class modules
- `.frm` (+ `.frx`) -> userforms

Document modules (`ThisWorkbook`, `Sheet1`, etc.) are not exported/imported by this CLI.

You can also keep ribbon XML as source (for example `ribbon/customUI14.xml`) and round-trip it to `.xlam` or `.dotm`.

## Prerequisites

1. Windows with desktop Excel and/or Word installed.
2. In Excel or Word: **Trust Center -> Macro Settings -> Trust access to the VBA project object model** enabled.
3. Python environment for this repo.

Install dependency:

```powershell
uv add pywin32
```

## CLI usage

Entry point: `com`

### Export workbook VBA to files (pull from Excel, default)

```powershell
uv run com export --workbook .\Workbook.xlsm --out .\vba
```

### Import files into workbook (push to Excel, default)

```powershell
uv run com import --workbook .\Workbook.xlsm --src .\vba
```

### Import and remove workbook components not present on disk

```powershell
uv run com import --workbook .\Workbook.xlsm --src .\vba --clean
```

### Export/import with Word (optional)

```powershell
# pull from Word
uv run com export --workbook .\Document.docm --out .\vba --app word

# push to Word
uv run com import --workbook .\Document.docm --src .\vba --clean --app word
```

### Sync convenience command

```powershell
# pull
uv run com sync --workbook .\Workbook.xlsm --dir .\vba --direction pull

# push
uv run com sync --workbook .\Workbook.xlsm --dir .\vba --direction push --clean
```

### Ribbon XML pull/push for `.xlam` or `.dotm`

Extract ribbon XML from an add-in package to edit in VS Code:

```powershell
uv run com ribbon pull --workbook .\MyAddin.xlam --out .\ribbon\customUI14.xml
```

Inject edited ribbon XML back into the add-in (in-place):

```powershell
uv run com ribbon push --workbook .\MyAddin.xlam --xml .\ribbon\customUI14.xml
```

Inject into a new output add-in file:

```powershell
uv run com ribbon push --workbook .\MyAddin.xlam --xml .\ribbon\customUI14.xml --out-workbook .\MyAddin.updated.xlam
```

Word template example:

```powershell
# pull from Word template
uv run com ribbon pull --workbook .\MyTemplate.dotm --out .\ribbon\customUI14.xml

# push to Word template
uv run com ribbon push --workbook .\MyTemplate.dotm --xml .\ribbon\customUI14.xml
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
