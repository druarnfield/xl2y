"""Streaming Excel reader.

Two passes per sheet, both bounded by the *true* data extent rather than the
declared dimension (stray formatting routinely inflates the declared range to
XFD1048576):

1. A streaming XML parse for merges, hidden rows/columns, formula cells and
   the real extent (:func:`sheet_structure`).
2. openpyxl read-only value streaming into a trimmed grid (:func:`read_grid`).

Ported from ``reference/excel_loader.py`` (the frozen prototype); the value
side lands in plain Python lists here — polars typing happens in
``extract.py``.
"""

from __future__ import annotations

import logging
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from xl2y.errors import SheetNotFoundError, UnsupportedFormatError

logger = logging.getLogger(__name__)

SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}


def check_format(path: Path) -> None:
    suffix = path.suffix.lower()
    if suffix in SUPPORTED_SUFFIXES:
        return
    if suffix in {".xls", ".xlsb"}:
        raise UnsupportedFormatError(
            f"{path.name}: {suffix} files cannot be read by openpyxl. "
            "Convert to .xlsx first (Excel: Save As; or `soffice --convert-to "
            "xlsx`; or pandas.read_excel with engine='xlrd'/'pyxlsb')."
        )
    raise UnsupportedFormatError(
        f"{path.name}: unrecognised extension {suffix!r}; expected one of "
        f"{sorted(SUPPORTED_SUFFIXES)}."
    )


def _localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def sheet_meta(path: Path) -> dict[str, tuple[str, bool]]:
    """Sheet order/type without loading the workbook.

    Returns an ordered mapping ``{sheet_name: (zip_member_path, is_chartsheet)}``
    parsed from ``xl/workbook.xml`` and its relationships.
    """
    RELS_NS = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    with zipfile.ZipFile(path) as zf:
        rels: dict[str, tuple[str, str]] = {}  # rId -> (target, type)
        root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        for rel in root:
            rels[rel.get("Id")] = (rel.get("Target"), rel.get("Type", ""))
        meta: dict[str, tuple[str, bool]] = {}
        wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
        for el in wb_root.iter():
            if _localname(el.tag) != "sheet":
                continue
            rid = el.get(f"{{{RELS_NS}}}id")
            target, rel_type = rels.get(rid, ("", ""))
            # Targets may be absolute ("/xl/worksheets/sheet1.xml") or
            # relative to xl/ ("worksheets/sheet1.xml").
            member = (
                target.lstrip("/")
                if target.startswith("/")
                else f"xl/{target}"
            )
            is_chart = "chartsheet" in rel_type or "chartsheets/" in member
            meta[el.get("name")] = (member, is_chart)
        return meta


def _require_sheet(meta: dict[str, tuple[str, bool]], name: str) -> None:
    if name not in meta:
        raise SheetNotFoundError(
            f"Sheet {name!r} not found. Available: {list(meta)}"
        )
    if meta[name][1]:
        raise ValueError(
            f"Sheet {name!r} is a chartsheet and contains no cell data."
        )
