from __future__ import annotations

import argparse
import contextlib
import sys
from dataclasses import dataclass
from pathlib import Path

from com.ribbon import pull_ribbon_xml, push_ribbon_xml

try:
    from win32com.client import DispatchEx  # type: ignore[import-untyped]
except ImportError:  # pragma: no cover - exercised in runtime on host
    DispatchEx = None

VBEXT_CT_STDMODULE = 1
VBEXT_CT_CLASSMODULE = 2
VBEXT_CT_MSFORM = 3

COMPONENT_EXTENSIONS = {
    VBEXT_CT_STDMODULE: ".bas",
    VBEXT_CT_CLASSMODULE: ".cls",
    VBEXT_CT_MSFORM: ".frm",
}

SUPPORTED_APPS = ("excel", "word")


@dataclass(frozen=True)
class SourceModule:
    path: Path
    extension: str

    @property
    def component_name(self) -> str:
        return self.path.stem


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="com",
        description="Export and import VBA components so you can edit in VS Code.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    export_parser = subparsers.add_parser("export", help="Export VBA modules from a workbook.")
    export_parser.add_argument("--workbook", required=True, type=Path, help="Path to .xlsm/.xlsb workbook")
    export_parser.add_argument("--out", required=True, type=Path, help="Directory for exported .bas/.cls/.frm files")
    export_parser.add_argument(
        "--app",
        choices=SUPPORTED_APPS,
        default="excel",
        help="Office app host (default: excel).",
    )

    import_parser = subparsers.add_parser("import", help="Import VBA modules into a workbook.")
    import_parser.add_argument("--workbook", required=True, type=Path, help="Path to .xlsm/.xlsb workbook")
    import_parser.add_argument("--src", required=True, type=Path, help="Directory containing .bas/.cls/.frm files")
    import_parser.add_argument(
        "--app",
        choices=SUPPORTED_APPS,
        default="excel",
        help="Office app host (default: excel).",
    )
    import_parser.add_argument(
        "--clean",
        action="store_true",
        help="Remove workbook modules/forms/classes not present in --src (document modules are untouched).",
    )

    sync_parser = subparsers.add_parser("sync", help="Convenience wrapper for export/import.")
    sync_parser.add_argument("--workbook", required=True, type=Path, help="Path to .xlsm/.xlsb workbook")
    sync_parser.add_argument("--dir", required=True, type=Path, help="Shared source folder")
    sync_parser.add_argument(
        "--app",
        choices=SUPPORTED_APPS,
        default="excel",
        help="Office app host (default: excel).",
    )
    sync_parser.add_argument(
        "--direction",
        required=True,
        choices=("pull", "push"),
        help="pull=export workbook to folder, push=import folder to workbook",
    )
    sync_parser.add_argument(
        "--clean",
        action="store_true",
        help="When --direction push, remove modules not present on disk.",
    )

    ribbon_parser = subparsers.add_parser("ribbon", help="Read/write custom ribbon XML from workbook package.")
    ribbon_subparsers = ribbon_parser.add_subparsers(dest="ribbon_command", required=True)

    ribbon_pull_parser = ribbon_subparsers.add_parser("pull", help="Extract ribbon XML from workbook package.")
    ribbon_pull_parser.add_argument("--workbook", required=True, type=Path, help="Path to workbook/add-in")
    ribbon_pull_parser.add_argument("--out", required=True, type=Path, help="Path to output XML file")

    ribbon_push_parser = ribbon_subparsers.add_parser("push", help="Inject ribbon XML into workbook package.")
    ribbon_push_parser.add_argument("--workbook", required=True, type=Path, help="Path to workbook/add-in")
    ribbon_push_parser.add_argument("--xml", required=True, type=Path, help="Path to ribbon XML file")
    ribbon_push_parser.add_argument(
        "--out-workbook",
        type=Path,
        default=None,
        help="Optional output workbook path. If omitted, updates --workbook in place.",
    )
    ribbon_push_parser.add_argument(
        "--target",
        choices=("customUI14.xml", "customUI.xml"),
        default=None,
        help="Override target filename under customUI/.",
    )

    return parser


def discover_source_modules(source_dir: Path) -> list[SourceModule]:
    modules: list[SourceModule] = []
    for suffix in (".bas", ".cls", ".frm"):
        for path in sorted(source_dir.glob(f"*{suffix}")):
            modules.append(SourceModule(path=path, extension=suffix))
    return modules


def assert_dispatch_available() -> None:
    if DispatchEx is None:
        raise RuntimeError(
            "pywin32 is required. Install with `uv add pywin32` or `pip install pywin32`."
        )


def ensure_exists(path: Path, kind: str) -> None:
    if not path.exists():
        raise FileNotFoundError(f"{kind} not found: {path}")


def ensure_office_file(path: Path, app: str) -> Path:
    ensure_exists(path, "Workbook")
    supported_extensions = {
        "excel": {".xlsm", ".xlsb", ".xlam", ".xls"},
        "word": {".docm", ".dotm", ".doc"},
    }
    if path.suffix.lower() not in supported_extensions[app]:
        raise ValueError(
            f"Workbook must be a supported {app.title()} file ({'/'.join(sorted(supported_extensions[app]))})."
        )
    return path.resolve()


def open_host_document(workbook_path: Path, app: str):
    assert_dispatch_available()
    if app == "excel":
        host = DispatchEx("Excel.Application")
        host.Visible = False
        host.DisplayAlerts = False
        document = host.Workbooks.Open(str(workbook_path))
        return host, document

    host = DispatchEx("Word.Application")
    host.Visible = False
    host.DisplayAlerts = 0
    document = host.Documents.Open(str(workbook_path))
    return host, document


def iter_exportable_components(vbproject):
    for component in vbproject.VBComponents:
        if int(component.Type) in COMPONENT_EXTENSIONS:
            yield component


def export_modules(workbook_path: Path, output_dir: Path, app: str = "excel") -> int:
    workbook_path = ensure_office_file(workbook_path, app)
    output_dir.mkdir(parents=True, exist_ok=True)

    host = document = None
    exported = 0
    try:
        host, document = open_host_document(workbook_path, app)
        vbproject = document.VBProject
        for component in iter_exportable_components(vbproject):
            extension = COMPONENT_EXTENSIONS[int(component.Type)]
            target = output_dir / f"{component.Name}{extension}"
            component.Export(str(target.absolute()))
            exported += 1
    finally:
        with contextlib.suppress(Exception):
            if document is not None:
                document.Close(SaveChanges=False)
        with contextlib.suppress(Exception):
            if host is not None:
                host.Quit()

    return exported


def import_modules(workbook_path: Path, source_dir: Path, clean: bool = False, app: str = "excel") -> int:
    workbook_path = ensure_office_file(workbook_path, app)
    ensure_exists(source_dir, "Source directory")

    sources = discover_source_modules(source_dir.resolve())
    source_names = {source.component_name for source in sources}

    host = document = None
    imported = 0
    try:
        host, document = open_host_document(workbook_path, app)
        vbproject = document.VBProject

        existing_components = {component.Name: component for component in iter_exportable_components(vbproject)}

        for source in sources:
            existing = existing_components.get(source.component_name)
            if existing is not None:
                vbproject.VBComponents.Remove(existing)
            vbproject.VBComponents.Import(str(source.path))
            imported += 1

        if clean:
            for component in list(iter_exportable_components(vbproject)):
                if component.Name not in source_names:
                    vbproject.VBComponents.Remove(component)

        document.Save()
    finally:
        with contextlib.suppress(Exception):
            if document is not None:
                document.Close(SaveChanges=False)
        with contextlib.suppress(Exception):
            if host is not None:
                host.Quit()

    cleanup_orphaned_frx(source_dir.resolve(), source_names)
    return imported


def cleanup_orphaned_frx(source_dir: Path, source_names: set[str]) -> None:
    for frx_file in source_dir.glob("*.frx"):
        if frx_file.stem not in source_names:
            frx_file.unlink(missing_ok=True)


def export_command(args: argparse.Namespace) -> int:
    exported = export_modules(args.workbook, args.out, app=args.app)
    print(f"Exported {exported} VBA components to {args.out.resolve()}")
    return 0


def import_command(args: argparse.Namespace) -> int:
    imported = import_modules(args.workbook, args.src, clean=args.clean, app=args.app)
    print(f"Imported {imported} VBA components from {args.src.resolve()}")
    return 0


def sync_command(args: argparse.Namespace) -> int:
    if args.direction == "pull":
        exported = export_modules(args.workbook, args.dir, app=args.app)
        print(f"Pulled {exported} VBA components into {args.dir.resolve()}")
        return 0

    imported = import_modules(args.workbook, args.dir, clean=args.clean, app=args.app)
    print(f"Pushed {imported} VBA components from {args.dir.resolve()}")
    return 0


def ribbon_command(args: argparse.Namespace) -> int:
    if args.ribbon_command == "pull":
        output = pull_ribbon_xml(args.workbook, args.out)
        print(f"Extracted ribbon XML to {output}")
        return 0

    output_workbook = push_ribbon_xml(
        workbook_path=args.workbook,
        ribbon_xml_path=args.xml,
        out_workbook_path=args.out_workbook,
        target_name=args.target,
    )
    print(f"Injected ribbon XML into {output_workbook}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    commands = {
        "export": export_command,
        "import": import_command,
        "sync": sync_command,
        "ribbon": ribbon_command,
    }

    try:
        return commands[args.command](args)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
