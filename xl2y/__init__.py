"""xl2y — Excel in, cleaning pipeline, Parquet out."""

from __future__ import annotations

from xl2y.errors import (
    EmptySheetError,
    SchemaError,
    SheetNotFoundError,
    UnsupportedFormatError,
    Xl2yError,
)

__all__ = [
    "Xl2yError",
    "UnsupportedFormatError",
    "EmptySheetError",
    "SheetNotFoundError",
    "SchemaError",
]
