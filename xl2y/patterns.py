"""Reusable validation patterns and helpers.

Constants are plain regex *strings* (not compiled), so they drop straight
into ``str_(pattern=...)`` and polars ``str.contains``. Checksum-based
identifiers (ABN/ACN/TFN) also expose predicate functions for use with a
column's ``check=`` argument, since a regex alone cannot validate them.

Australian-flavoured by default (this is an AU-built library); ``PHONE_E164``
and ``EMAIL``/``URL``/``UUID``/``IPV4`` are region-neutral.
"""

from __future__ import annotations

import re

# --------------------------------------------------------------------------- #
# Regex string constants
# --------------------------------------------------------------------------- #

EMAIL = r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"
URL = r"https?://[^\s/$.?#][^\s]*"
UUID = (
    r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-"
    r"[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
)
IPV4 = (
    r"(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}"
    r"(?:25[0-5]|2[0-4]\d|1?\d?\d)"
)

PHONE_E164 = r"\+[1-9]\d{7,14}"
# Optional +61 or trunk 0 (parens allowed around the area digit), then the
# remaining significant digits with optional spaces/hyphens between them.
PHONE_AU = r"(?:\+61\s?\(?\d\)?|\(?0\d\)?)(?:[\s-]?\d){7,8}"

POSTCODE_AU = r"\d{4}"
ABN = r"\d{2}\s?\d{3}\s?\d{3}\s?\d{3}"
ACN = r"\d{3}\s?\d{3}\s?\d{3}"
BSB = r"\d{3}-?\d{3}"
DATE_ISO = r"\d{4}-\d{2}-\d{2}"
CURRENCY = r"-?\(?\$?\s?[\d,]+(?:\.\d+)?\)?"


# --------------------------------------------------------------------------- #
# Composable helpers (all return regex strings)
# --------------------------------------------------------------------------- #


def any_of(*patterns: str) -> str:
    """Match any of the given patterns."""
    return "(?:" + "|".join(patterns) + ")"


def exact(literal: str) -> str:
    """Match one exact literal string (regex-escaped)."""
    return re.escape(literal)


def digits(n: int) -> str:
    """Match exactly ``n`` digits."""
    return rf"\d{{{n}}}"


def one_of(*values: str) -> str:
    """Match any one of the given literal values (regex-escaped)."""
    return "(?:" + "|".join(re.escape(v) for v in values) + ")"


# --------------------------------------------------------------------------- #
# Checksum validators (predicates, for use with ColumnType(check=...))
# --------------------------------------------------------------------------- #


def _only_digits(s: str) -> list[int]:
    return [int(c) for c in re.sub(r"\D", "", s)]


def abn_valid(s: str) -> bool:
    """Validate an Australian Business Number by its weighted checksum."""
    d = _only_digits(s)
    if len(d) != 11:
        return False
    weights = [10, 1, 3, 5, 7, 9, 11, 13, 15, 17, 19]
    d[0] -= 1
    return sum(x * w for x, w in zip(d, weights)) % 89 == 0


def acn_valid(s: str) -> bool:
    """Validate an Australian Company Number by its complement check digit."""
    d = _only_digits(s)
    if len(d) != 9:
        return False
    weights = [8, 7, 6, 5, 4, 3, 2, 1]
    total = sum(x * w for x, w in zip(d[:8], weights))
    check = (10 - (total % 10)) % 10
    return check == d[8]


def tfn_valid(s: str) -> bool:
    """Validate an Australian Tax File Number by its weighted checksum."""
    d = _only_digits(s)
    weights_by_len = {9: [1, 4, 3, 7, 5, 8, 6, 9, 10], 8: [1, 4, 3, 7, 5, 8, 6, 9]}
    weights = weights_by_len.get(len(d))
    if weights is None:
        return False
    return sum(x * w for x, w in zip(d, weights)) % 11 == 0
