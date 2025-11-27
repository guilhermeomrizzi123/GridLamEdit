"""Project-level queries for shared laminate data."""

from __future__ import annotations

from collections import Counter
from collections.abc import Iterable
from typing import Any

from gridlamedit.io.spreadsheet import GridModel, Laminado, normalize_angle


def _iter_laminates(project: Any) -> Iterable[Laminado]:
    """Yield laminates from a GridModel-like project."""
    if project is None:
        return []

    laminates = getattr(project, "laminados", None)
    if isinstance(laminates, dict):
        return laminates.values()
    if isinstance(laminates, Iterable):
        return laminates
    return []


def project_distinct_materials(project: GridModel | Any) -> list[str]:
    """
    Retorna todos os materiais distintos usados em todos os laminados do projeto.
    Ignora vazios/None. Ordena alfabeticamente. Sem duplicatas.
    """
    materials: set[str] = set()
    for laminate in _iter_laminates(project):
        for layer in getattr(laminate, "camadas", []):
            if getattr(layer, "orientacao", None) is None:
                continue
            material = getattr(layer, "material", "")
            text = str(material).strip()
            if text:
                materials.add(text)
    return sorted(materials, key=str.casefold)


def project_distinct_orientations(project: GridModel | Any) -> list[float]:
    """
    Retorna orienta\u00e7\u00f5es distintas usadas em todos os laminados do projeto.
    Ignora vazios/None. Normaliza para float. Sem duplicatas.
    """
    orientations: set[float] = set()
    for laminate in _iter_laminates(project):
        for layer in getattr(laminate, "camadas", []):
            value = getattr(layer, "orientacao", None)
            try:
                normalized = normalize_angle(value)
            except (TypeError, ValueError):
                continue
            orientations.add(normalized)
    return sorted(orientations)


def project_most_used_material(project: GridModel | Any) -> str | None:
    """Retorna o material mais utilizado considerando apenas camadas com orientacao."""
    counts: Counter[str] = Counter()
    for laminate in _iter_laminates(project):
        for layer in getattr(laminate, "camadas", []):
            if getattr(layer, "orientacao", None) is None:
                continue
            material = getattr(layer, "material", "")
            text = str(material).strip()
            if text:
                counts[text] += 1
    most_common = counts.most_common(1)
    return most_common[0][0] if most_common else None
