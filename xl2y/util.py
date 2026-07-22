"""Small shared helpers: identifier snake_casing and name deduplication."""

from __future__ import annotations

import re


def snake_case(name: str) -> str:
    s = re.sub(r"[^\w]+", "_", name.strip())
    s = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", "_", s)  # camelCase boundaries
    s = re.sub(r"_+", "_", s).strip("_").lower() or "col"
    if not s[0].isalpha():  # "123_totals" is awkward in df.filter/iter
        s = f"col_{s}"
    return s


def unique_name(base: str, taken: set[str]) -> str:
    name, n = base, 2
    while name in taken:
        name, n = f"{base}_{n}", n + 1
    return name
