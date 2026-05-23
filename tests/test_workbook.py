"""Unit tests for WorkbookReader."""

from __future__ import annotations

import pathlib

import openpyxl
import pytest

from oe_nh.workbook import XLS_MAGIC, WorkbookReader


def _write_xlsx(path: pathlib.Path, rows: list[list]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(path)


def test_sniff_xlsx_signature(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.xlsx"
    _write_xlsx(p, [["a", "b"], [1, 2]])
    assert WorkbookReader.sniff_format(p) == "xlsx"


def test_sniff_xls_signature(tmp_path: pathlib.Path) -> None:
    # Hand-craft a file with just the OLE2 magic. We can't open it as a real
    # workbook, but sniff_format only looks at the header bytes.
    p = tmp_path / "x.xls"
    p.write_bytes(XLS_MAGIC + b"\x00" * 16)
    assert WorkbookReader.sniff_format(p) == "xls"


def test_sniff_rejects_unknown(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.bin"
    p.write_bytes(b"NOTAWORKBOOK")
    with pytest.raises(ValueError):
        WorkbookReader.sniff_format(p)


def test_sniff_dispatches_on_content_not_extension(tmp_path: pathlib.Path) -> None:
    """If SoS hands us an xlsx with a .xls extension, we still parse it."""
    p = tmp_path / "mislabeled.xls"
    _write_xlsx(p, [["a"], [1]])
    assert WorkbookReader.sniff_format(p) == "xlsx"


def test_xlsx_read_basic(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.xlsx"
    _write_xlsx(p, [
        ["Town", "Trump", "Biden"],
        ["Albany", 100, 200],
        ["Bartlett", 50, 75],
    ])
    r = WorkbookReader(p)
    assert r.nrows == 3
    assert r.ncols == 3
    assert r.row_values(0) == ["Town", "Trump", "Biden"]
    assert r.cell_value(1, 1) == 100


def test_xlsx_padding_short_rows(tmp_path: pathlib.Path) -> None:
    """Short rows are padded with '' so cell_value never IndexErrors within bounds."""
    p = tmp_path / "x.xlsx"
    _write_xlsx(p, [
        ["A", "B", "C"],
        ["short"],            # one cell only
        ["a", "b", "c"],
    ])
    r = WorkbookReader(p)
    assert r.ncols == 3
    assert r.cell_value(1, 2) == ""


def test_xlsx_none_normalized_to_empty_string(tmp_path: pathlib.Path) -> None:
    p = tmp_path / "x.xlsx"
    _write_xlsx(p, [["a", None, "c"]])
    r = WorkbookReader(p)
    assert r.cell_value(0, 1) == ""
