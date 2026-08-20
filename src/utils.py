from __future__ import annotations
from pathlib import Path
import re


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def safe_name(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*]+', '_', str(value))
    value = re.sub(r'\s+', '_', value.strip())
    return value[:180] or 'file'


def col_to_num(col: str | int) -> int:
    if isinstance(col, int):
        return col
    n = 0
    for ch in str(col).strip().upper():
        if 'A' <= ch <= 'Z':
            n = n * 26 + ord(ch) - 64
    if n < 1:
        raise ValueError(f'Invalid Excel column: {col}')
    return n


def num_to_col(n: int) -> str:
    out = ''
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out
