from __future__ import annotations
import json
import shutil
from datetime import datetime
from pathlib import Path
from .utils import project_root, safe_name
from .config_store import load_component

REGISTRY = project_root() / 'storage' / 'templates' / 'registry.json'


def _load_registry() -> dict:
    if REGISTRY.exists():
        try:
            return json.loads(REGISTRY.read_text(encoding='utf-8'))
        except Exception:
            return {}
    return {}


def _save_registry(data: dict) -> None:
    REGISTRY.parent.mkdir(parents=True, exist_ok=True)
    REGISTRY.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding='utf-8')


def _local_template(component_id: str) -> Path | None:
    comp = load_component(component_id)
    suggested = (comp.get('template') or {}).get('suggested_filename')
    folder = (comp.get('template') or {}).get('folder')
    if not suggested:
        return None
    if folder:
        p = project_root() / folder / suggested
    else:
        p = project_root() / 'templates_local' / {'loader':'loaders','view':'views','control':'controls'}[comp['component_type']] / suggested
    return p if p.exists() else None


def active_template(component_id: str) -> Path | None:
    row = _load_registry().get(component_id.upper())
    if row:
        p = project_root() / row['path'] if not Path(row['path']).is_absolute() else Path(row['path'])
        if p.exists():
            return p
    return _local_template(component_id)


def template_info(component_id: str) -> dict | None:
    row = _load_registry().get(component_id.upper())
    if row:
        p = project_root() / row['path'] if not Path(row['path']).is_absolute() else Path(row['path'])
        if p.exists():
            return {**row, 'path': str(p), 'source': 'registered'}
    p = _local_template(component_id)
    if p:
        return {'filename': p.name, 'active_version': 'local-fixed', 'path': str(p), 'registered_at': None, 'source': 'local-fixed'}
    return None


def register_template(component_id: str, uploaded_path: Path, original_name: str, version: str) -> Path:
    version = safe_name(version or datetime.now().strftime('%Y%m%d_%H%M%S'))
    dest_dir = project_root() / 'storage' / 'templates' / component_id.upper() / version
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / safe_name(original_name)
    shutil.copy2(uploaded_path, dest)
    reg = _load_registry()
    reg[component_id.upper()] = {
        'active_version': version,
        'path': str(dest.relative_to(project_root())),
        'filename': original_name,
        'registered_at': datetime.now().isoformat(timespec='seconds'),
    }
    _save_registry(reg)
    return dest
