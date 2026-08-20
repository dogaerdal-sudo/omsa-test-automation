from __future__ import annotations
import json
from pathlib import Path
from datetime import datetime
from .utils import project_root

HISTORY = project_root() / 'storage' / 'runs' / 'history.jsonl'


def save_run(row: dict) -> None:
    HISTORY.parent.mkdir(parents=True, exist_ok=True)
    payload = dict(row)
    payload.setdefault('generated_at', datetime.now().isoformat(timespec='seconds'))
    with HISTORY.open('a', encoding='utf-8') as f:
        f.write(json.dumps(payload, ensure_ascii=False, default=str) + '\n')


def list_runs(limit: int = 500) -> list[dict]:
    if not HISTORY.exists():
        return []
    rows=[]
    for line in HISTORY.read_text(encoding='utf-8').splitlines():
        try: rows.append(json.loads(line))
        except Exception: pass
    rows.reverse()
    return rows[:limit]
