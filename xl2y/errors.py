"""Exception hierarchy. Everything the library raises is an Xl2yError."""

from __future__ import annotations


class Xl2yError(Exception):
    """Base for all xl2y errors."""


class UnsupportedFormatError(Xl2yError, ValueError):
    """Spreadsheet format openpyxl cannot open (.xls, .xlsb, ...)."""


class EmptySheetError(Xl2yError, ValueError):
    """Sheet contains no data."""


class SheetNotFoundError(Xl2yError, KeyError):
    """Requested sheet does not exist; message lists available sheets."""


class SchemaError(Xl2yError, ValueError):
    """Validation failed. Carries every problem, not just the first."""

    def __init__(self, message: str, problems: list):
        super().__init__(message)
        self.problems = problems
