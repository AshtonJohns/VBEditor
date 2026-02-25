from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path

RIBBON_CANDIDATE_PATHS = (
    Path("customUI/customUI14.xml"),
    Path("customUI/customUI.xml"),
)


def _extract_zip(zip_path: Path, extract_dir: Path) -> None:
    if extract_dir.exists():
        shutil.rmtree(extract_dir, ignore_errors=True)
    extract_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "r") as archive:
        archive.extractall(extract_dir)


def _rebuild_zip_from_dir(source_dir: Path, zip_path: Path) -> None:
    tmp_zip = zip_path.with_suffix(zip_path.suffix + ".tmp")
    if tmp_zip.exists():
        tmp_zip.unlink()

    with zipfile.ZipFile(tmp_zip, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for file_path in source_dir.rglob("*"):
            if file_path.is_file():
                archive.write(file_path, file_path.relative_to(source_dir).as_posix())

    if zip_path.exists():
        zip_path.unlink()
    tmp_zip.replace(zip_path)


def _extract_workbook_package(workbook_path: Path, workspace_root: Path) -> tuple[Path, Path]:
    temp_zip = workspace_root / "workbook.zip"
    shutil.copy2(workbook_path, temp_zip)

    extracted_root = workspace_root / "extracted"
    _extract_zip(temp_zip, extracted_root)
    return temp_zip, extracted_root


def _find_existing_ribbon_path(extracted_root: Path) -> Path | None:
    for relative in RIBBON_CANDIDATE_PATHS:
        candidate = extracted_root / relative
        if candidate.exists():
            return candidate
    return None


def pull_ribbon_xml(workbook_path: Path, output_xml_path: Path) -> Path:
    workbook_path = workbook_path.resolve()
    output_xml_path = output_xml_path.resolve()

    with tempfile.TemporaryDirectory(prefix="excel-com-ribbon-") as temp_dir:
        workspace_root = Path(temp_dir)
        _, extracted_root = _extract_workbook_package(workbook_path, workspace_root)
        ribbon_path = _find_existing_ribbon_path(extracted_root)
        if ribbon_path is None:
            raise FileNotFoundError(
                "No ribbon XML found. Expected customUI/customUI14.xml or customUI/customUI.xml in workbook package."
            )

        output_xml_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ribbon_path, output_xml_path)

    return output_xml_path


def push_ribbon_xml(
    workbook_path: Path,
    ribbon_xml_path: Path,
    out_workbook_path: Path | None = None,
    target_name: str | None = None,
) -> Path:
    workbook_path = workbook_path.resolve()
    ribbon_xml_path = ribbon_xml_path.resolve()

    if not ribbon_xml_path.exists():
        raise FileNotFoundError(f"Ribbon XML not found: {ribbon_xml_path}")

    if target_name is not None and target_name not in {"customUI14.xml", "customUI.xml"}:
        raise ValueError("target_name must be customUI14.xml or customUI.xml")

    destination_workbook = out_workbook_path.resolve() if out_workbook_path is not None else workbook_path

    with tempfile.TemporaryDirectory(prefix="excel-com-ribbon-") as temp_dir:
        workspace_root = Path(temp_dir)
        temp_zip, extracted_root = _extract_workbook_package(workbook_path, workspace_root)

        existing_ribbon = _find_existing_ribbon_path(extracted_root)
        if target_name is None:
            target_relative = existing_ribbon.relative_to(extracted_root) if existing_ribbon else Path("customUI/customUI14.xml")
        else:
            target_relative = Path("customUI") / target_name

        target_xml_path = extracted_root / target_relative
        target_xml_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ribbon_xml_path, target_xml_path)

        _rebuild_zip_from_dir(extracted_root, temp_zip)

        destination_workbook.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(temp_zip, destination_workbook)

    return destination_workbook
