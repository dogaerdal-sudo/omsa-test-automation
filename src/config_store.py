from __future__ import annotations

from copy import deepcopy
from functools import lru_cache
from pathlib import Path
from typing import Any

import yaml

from .utils import project_root

CATEGORIES = {'loader': 'loaders', 'view': 'views', 'control': 'controls'}


def config_dir(component_type: str) -> Path:
    return project_root() / 'config' / CATEGORIES[component_type]


@lru_cache(maxsize=4)
def _load_components_cached(component_type: str | None = None) -> tuple[dict[str, Any], ...]:
    # Parse component YAML files once per Python process.
    # Public functions return deep copies, so callers cannot mutate the cache.
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

    return tuple(rows)


@lru_cache(maxsize=1)
def _component_index_cached() -> dict[str, dict[str, Any]]:
    return {
        str(item.get('component_id', '')).upper(): item
        for item in _load_components_cached(None)
    }


def load_components(component_type: str | None = None) -> list[dict[str, Any]]:
    return deepcopy(list(_load_components_cached(component_type)))


def load_component(component_id: str) -> dict[str, Any]:
    target = str(component_id).upper()
    item = _component_index_cached().get(target)
    if item is None:
        raise KeyError(f'Component not found: {component_id}')
    return deepcopy(item)


def clear_config_cache() -> None:
    # Call only if YAML files are edited without restarting the app.
    _load_components_cached.cache_clear()
    _component_index_cached.cache_clear()
