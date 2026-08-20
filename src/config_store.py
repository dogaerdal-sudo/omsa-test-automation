from __future__ import annotations
from pathlib import Path
from typing import Any
import yaml
from .utils import project_root

CATEGORIES = {'loader': 'loaders', 'view': 'views', 'control': 'controls'}


def config_dir(component_type: str) -> Path:
    return project_root() / 'config' / CATEGORIES[component_type]


def load_components(component_type: str | None = None) -> list[dict[str, Any]]:
    types = [component_type] if component_type else list(CATEGORIES)
    rows: list[dict[str, Any]] = []
    for ctype in types:
        folder = config_dir(ctype)
        if not folder.exists():
            continue
        for path in sorted(folder.glob('*.yaml')):
            data = yaml.safe_load(path.read_text(encoding='utf-8')) or {}
            data['_config_path'] = str(path)
            rows.append(data)
    return rows


def load_component(component_id: str) -> dict[str, Any]:
    target = str(component_id).upper()
    for item in load_components():
        if str(item.get('component_id', '')).upper() == target:
            return item
    raise KeyError(f'Component not found: {component_id}')
