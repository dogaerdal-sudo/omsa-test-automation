from __future__ import annotations

import re
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd

DEFAULT_DATE_HEADERS = [
    'DTA_RIF', 'DATA_RIF', 'REFERENCE_DATE', 'REPORTING_DATE', 'REF_DATE',
    'BUSINESS_DATE', 'VALUATION_DATE', 'NAV_DATE', 'DTA_NAV', 'DATA_NAV',
    'AS_OF_DATE', 'ASOF_DATE', 'ACCOUNTING_DATE', 'POSITION_DATE',
]


def _norm(value: Any) -> str:
    return re.sub(r'[^A-Z0-9]+', '_', str(value or '').strip().upper()).strip('_')


def _preview(path: Path, slot: dict, nrows: int = 300) -> pd.DataFrame | None:
    suffix = path.suffix.lower()
    header_row = max(1, int((slot.get('source') or {}).get('header_row', 1))) - 1
    try:
        if suffix in {'.csv', '.txt'}:
            for encoding in ('utf-8-sig', 'utf-8', 'latin1'):
                try:
                    return pd.read_csv(path, sep=None, engine='python', header=header_row, nrows=nrows, encoding=encoding)
                except Exception:
                    continue
            return None
        if suffix in {'.xlsx', '.xlsm'}:
            return pd.read_excel(path, sheet_name=int((slot.get('source') or {}).get('sheet_index', 1)) - 1,
                                 header=header_row, nrows=nrows, engine='openpyxl')
        if suffix == '.xlsb':
            return pd.read_excel(path, sheet_name=int((slot.get('source') or {}).get('sheet_index', 1)) - 1,
                                 header=header_row, nrows=nrows, engine='pyxlsb')
        if suffix == '.xls':
            return pd.read_excel(path, sheet_name=int((slot.get('source') or {}).get('sheet_index', 1)) - 1,
                                 header=header_row, nrows=nrows, engine='xlrd')
    except Exception:
        return None
    return None


def _parse_date_values(series: pd.Series) -> tuple[str | None, str | None]:
    values = series.dropna()
    if values.empty:
        return None, None

    parsed: list[date] = []
    for value in values.head(500).tolist():
        if value in (None, ''):
            continue
        try:
            # Excel serial-date handling.
            if isinstance(value, (int, float)) and 20000 <= float(value) <= 80000:
                ts = pd.Timestamp('1899-12-30') + pd.to_timedelta(float(value), unit='D')
            elif isinstance(value, (datetime, date, pd.Timestamp)):
                ts = pd.Timestamp(value)
            else:
                ts = pd.to_datetime(str(value).strip(), errors='coerce', dayfirst=True)
            if pd.isna(ts):
                continue
            d = ts.date()
            if 1990 <= d.year <= 2100:
                parsed.append(d)
        except Exception:
            continue

    if not parsed:
        return None, None
    counts = Counter(parsed)
    most_common, count = counts.most_common(1)[0]
    share = count / len(parsed)
    if len(counts) == 1:
        return most_common.isoformat(), 'single date in column'
    if share >= 0.85:
        return most_common.isoformat(), f'dominant date ({share:.0%} of parsed rows)'
    return None, f'ambiguous: {len(counts)} dates in column'


def _date_from_filename(name: str) -> tuple[str | None, str | None]:
    stem = Path(name).stem
    patterns = [
        (r'(?<!\d)(20\d{2})[-_.]?(0[1-9]|1[0-2])[-_.]?([0-2]\d|3[01])(?!\d)', 'ymd'),
        (r'(?<!\d)([0-2]\d|3[01])[-_.](0[1-9]|1[0-2])[-_.](20\d{2})(?!\d)', 'dmy'),
        (r'(?<!\d)([0-2]\d|3[01])(0[1-9]|1[0-2])(20\d{2})(?!\d)', 'dmy'),
    ]
    for pattern, order in patterns:
        match = re.search(pattern, stem)
        if not match:
            continue
        try:
            if order == 'ymd':
                y, m, d = map(int, match.groups())
            else:
                d, m, y = map(int, match.groups())
            parsed = date(y, m, d)
            return parsed.isoformat(), 'filename'
        except Exception:
            continue
    return None, None


def detect_reference_date(path: Path, slot: dict) -> dict:
    """Best-effort reference-date detection for one uploaded run file.

    The detector never changes workbook data. It only inspects a small preview and the filename.
    Per-slot config can later override candidate headers with date_detection.column_headers.
    """
    path = Path(path)
    config = slot.get('date_detection') or {}
    preferred = config.get('column_headers') or config.get('headers') or DEFAULT_DATE_HEADERS
    preferred_norm = [_norm(x) for x in preferred]

    frame = None if config.get('filename_only') else _preview(path, slot)
    if frame is not None and not frame.empty:
        normalized = {_norm(col): col for col in frame.columns}
        for wanted in preferred_norm:
            if wanted not in normalized:
                continue
            col = normalized[wanted]
            detected, detail = _parse_date_values(frame[col])
            if detected:
                return {
                    'date': detected,
                    'display_date': pd.Timestamp(detected).strftime('%d/%m/%Y'),
                    'status': 'detected',
                    'source': f'column {col}',
                    'detail': detail,
                }

    detected, detail = _date_from_filename(path.name)
    if detected:
        return {
            'date': detected,
            'display_date': pd.Timestamp(detected).strftime('%d/%m/%Y'),
            'status': 'detected',
            'source': 'filename',
            'detail': detail,
        }

    return {
        'date': None,
        'display_date': '—',
        'status': 'not_detected',
        'source': '—',
        'detail': 'No configured reference-date column or date pattern was detected.',
    }


def _role(slot: dict) -> str:
    role = str(slot.get('role') or '').strip().upper().replace(' ', '')
    label = str(slot.get('label') or '').upper().replace(' ', '')
    combined = role or label
    if 'T-1' in combined or 'T_1' in combined or 'T1' == combined:
        return 'T-1'
    return 'T'


def validate_reference_dates(component: dict, detections: dict[str, dict]) -> dict:
    """Validate T and T-1 consistency across uploaded files with detectable dates."""
    primary: list[tuple[str, str]] = []
    previous: list[tuple[str, str]] = []
    undetected: list[str] = []
    labels = {slot['id']: slot.get('label', slot['id']) for slot in component.get('inputs', [])}
    slots = {slot['id']: slot for slot in component.get('inputs', [])}

    for sid, result in detections.items():
        d = result.get('date')
        if not d:
            undetected.append(labels.get(sid, sid))
            continue
        if _role(slots.get(sid, {})) == 'T-1':
            previous.append((sid, d))
        else:
            primary.append((sid, d))

    primary_dates = sorted({d for _, d in primary})
    previous_dates = sorted({d for _, d in previous})
    errors: list[str] = []
    warnings: list[str] = []

    if len(primary_dates) > 1:
        errors.append('Reference date mismatch across current-date (T) input files: ' + ', '.join(primary_dates))
    if len(previous_dates) > 1:
        errors.append('T-1 date mismatch across previous-date input files: ' + ', '.join(previous_dates))

    reference_date = primary_dates[0] if len(primary_dates) == 1 else None
    previous_date = previous_dates[0] if len(previous_dates) == 1 else None
    if reference_date and previous_date and previous_date >= reference_date:
        errors.append(f'T-1 date ({previous_date}) must be earlier than Reference Date T ({reference_date}).')

    if undetected:
        warnings.append('Date not detected in: ' + ', '.join(undetected))
    if not reference_date and primary:
        # primary exists but is inconsistent: already an error above
        pass
    elif not reference_date and detections:
        warnings.append('Reference Date could not be determined from the uploaded files yet.')

    requires_reference_date = any(
        action.get('action') == 'set_column_to_reporting_date'
        for action in component.get('post_import_actions', [])
    )
    if requires_reference_date and detections and not reference_date and not errors:
        errors.append('This template requires a Reference Date, but no date could be detected from the uploaded current-date files.')

    return {
        'reference_date': reference_date,
        'previous_date': previous_date,
        'is_valid': not errors,
        'errors': errors,
        'warnings': warnings,
        'detected_count': sum(1 for r in detections.values() if r.get('date')),
        'undetected_count': sum(1 for r in detections.values() if not r.get('date')),
        'requires_reference_date': requires_reference_date,
    }
