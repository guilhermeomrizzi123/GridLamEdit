"""Utilities for managing default and user-defined materials."""

from __future__ import annotations

from typing import Any, Iterable, Sequence

from PySide6.QtCore import QSettings

DEFAULT_MATERIALS = [
    "E1676818   FABRIC,PREPREG,CARBON/EPOXY RESIN-CLASS 1,T830H-6K-PW/3900-2D",
]

_SETTINGS_KEY = "Materials/custom_list"


def _normalize_materials(items: Iterable[str] | None) -> list[str]:
    """Strip whitespace, drop blanks, and de-duplicate while preserving order."""
    normalized: list[str] = []
    seen: set[str] = set()
    for raw in items or []:
        text = str(raw or "").strip()
        if not text:
            continue
        key = text.casefold()
        if key in seen:
            continue
        seen.add(key)
        normalized.append(text)
    return normalized


def _settings(instance: QSettings | None = None) -> QSettings:
    return instance or QSettings("GridLamEdit", "GridLamEdit")


def load_custom_materials(settings: QSettings | None = None) -> list[str]:
    """Return the list of user-registered materials stored in settings."""
    store = _settings(settings)
    raw = store.value(_SETTINGS_KEY, [])
    if isinstance(raw, str):
        raw_items: Sequence[Any] = [raw]
    elif isinstance(raw, (list, tuple)):
        raw_items = list(raw)
    else:
        raw_items = []
    return _normalize_materials(raw_items)


def save_custom_materials(
    materials: Sequence[str] | None, settings: QSettings | None = None
) -> list[str]:
    """Persist a material list and return its normalized form."""
    normalized = _normalize_materials(materials)
    store = _settings(settings)
    store.setValue(_SETTINGS_KEY, normalized)
    return normalized


def add_custom_material(material: str, settings: QSettings | None = None) -> list[str]:
    """Add a material to the custom list, returning the updated list."""
    current = load_custom_materials(settings)
    text = str(material or "").strip()
    if not text:
        return current
    key = text.casefold()
    if key not in {item.casefold() for item in current}:
        current.append(text)
    return save_custom_materials(current, settings)


def available_materials(
    project: Any = None, settings: QSettings | None = None
) -> list[str]:
    """
    Build the ordered material list exposed to the UI.

    Priority: defaults first, then custom entries, then project-derived materials.
    Duplicates are removed case-insensitively.
    """
    from gridlamedit.services.project_query import project_distinct_materials

    base = _normalize_materials(DEFAULT_MATERIALS)
    custom = load_custom_materials(settings)
    project_materials = project_distinct_materials(project)
    extras = sorted(_normalize_materials([*custom, *project_materials]), key=str.casefold)
    return _normalize_materials([*base, *extras])
