from pathlib import Path

import pytest

from excel_com.cli import (
    SourceModule,
    build_parser,
    cleanup_orphaned_frx,
    discover_source_modules,
)
from excel_com.ribbon import pull_ribbon_xml, push_ribbon_xml


def test_discover_source_modules_only_returns_supported_files(tmp_path: Path) -> None:
    (tmp_path / "Module1.bas").write_text("Attribute VB_Name = \"Module1\"", encoding="utf-8")
    (tmp_path / "Class1.cls").write_text("Attribute VB_Name = \"Class1\"", encoding="utf-8")
    (tmp_path / "Form1.frm").write_text("VERSION 5.00", encoding="utf-8")
    (tmp_path / "notes.txt").write_text("ignored", encoding="utf-8")

    modules = discover_source_modules(tmp_path)

    assert modules == [
        SourceModule(path=tmp_path / "Module1.bas", extension=".bas"),
        SourceModule(path=tmp_path / "Class1.cls", extension=".cls"),
        SourceModule(path=tmp_path / "Form1.frm", extension=".frm"),
    ]


def test_cleanup_orphaned_frx_removes_only_unmatched_files(tmp_path: Path) -> None:
    keep = tmp_path / "MyForm.frx"
    remove = tmp_path / "OldForm.frx"
    keep.write_bytes(b"keep")
    remove.write_bytes(b"remove")

    cleanup_orphaned_frx(tmp_path, {"MyForm"})

    assert keep.exists()
    assert not remove.exists()


def test_parser_supports_sync_push_clean() -> None:
    parser = build_parser()
    args = parser.parse_args(
        ["sync", "--workbook", "Book1.xlsm", "--dir", "vba", "--direction", "push", "--clean"]
    )

    assert args.command == "sync"
    assert args.direction == "push"
    assert args.clean is True
    assert args.app == "excel"


def test_parser_supports_word_host_option() -> None:
    parser = build_parser()
    args = parser.parse_args(
        ["export", "--workbook", "Doc1.docm", "--out", "vba", "--app", "word"]
    )

    assert args.command == "export"
    assert args.app == "word"


def test_parser_requires_command() -> None:
    parser = build_parser()
    with pytest.raises(SystemExit):
        parser.parse_args([])


def test_parser_supports_ribbon_push() -> None:
    parser = build_parser()
    args = parser.parse_args(
        [
            "ribbon",
            "push",
            "--workbook",
            "Addin.xlam",
            "--xml",
            "customUI14.xml",
            "--target",
            "customUI14.xml",
        ]
    )

    assert args.command == "ribbon"
    assert args.ribbon_command == "push"
    assert args.target == "customUI14.xml"


def test_ribbon_pull_and_push_roundtrip(tmp_path: Path) -> None:
    workbook = tmp_path / "Sample.xlam"
    ribbon_xml = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' />"

    import zipfile

    with zipfile.ZipFile(workbook, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("customUI/customUI14.xml", ribbon_xml)
        archive.writestr("[Content_Types].xml", "<Types />")

    out_xml = tmp_path / "ribbon" / "customUI14.xml"
    pulled = pull_ribbon_xml(workbook, out_xml)
    assert pulled == out_xml.resolve()
    assert out_xml.read_text(encoding="utf-8") == ribbon_xml

    updated_xml = tmp_path / "updated.xml"
    updated_xml.write_text("<customUI/>", encoding="utf-8")
    out_workbook = tmp_path / "Updated.xlam"
    push_ribbon_xml(workbook, updated_xml, out_workbook_path=out_workbook, target_name="customUI14.xml")

    with zipfile.ZipFile(out_workbook, "r") as archive:
        assert archive.read("customUI/customUI14.xml").decode("utf-8") == "<customUI/>"
