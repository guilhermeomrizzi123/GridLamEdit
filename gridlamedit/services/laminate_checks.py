"""Domain checks applied to laminates before exporting."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Sequence

from gridlamedit.io.spreadsheet import (
    Camada,
    Laminado,
    PLY_TYPE_OPTIONS,
    normalize_angle,
)

STRUCTURAL_LABEL = PLY_TYPE_OPTIONS[0].lower()
PLY_TYPE_TOKENS = {option.lower() for option in PLY_TYPE_OPTIONS}


@dataclass
class SymmetryResult:
    """Holds laminate names categorized by symmetry."""

    symmetric: List[str] = field(default_factory=list)
    not_symmetric: List[str] = field(default_factory=list)


@dataclass
class DuplicateGroup:
    """Represents laminates that share the same duplicate signature."""

    signature: str
    summary: str
    laminates: List[str]


@dataclass
class ChecksReport:
    """Aggregate for all laminate checks."""

    symmetry: SymmetryResult
    duplicates: List[DuplicateGroup]
    meta: Dict[str, Any] = field(default_factory=dict)


def run_all_checks(laminates: Iterable[Laminado]) -> ChecksReport:
    """
    Execute every available laminate check and return a consolidated report.
    """

    laminate_list = [lam for lam in laminates if isinstance(lam, Laminado)]
    symmetry = check_symmetry(laminate_list)
    duplicates = check_duplicates(laminate_list)
    return ChecksReport(symmetry=symmetry, duplicates=duplicates, meta={})


def check_symmetry(laminates: Sequence[Laminado]) -> SymmetryResult:
    """
    Reuse the existing symmetry definition: only structural plies participate and
    materials/orientations must match when mirrored.
    """

    symmetric: list[str] = []
    not_symmetric: list[str] = []

    for laminado in laminates:
        if _is_laminate_symmetric(laminado):
            symmetric.append(laminado.nome)
        else:
            not_symmetric.append(laminado.nome)

    symmetric.sort()
    not_symmetric.sort()
    return SymmetryResult(symmetric=symmetric, not_symmetric=not_symmetric)


def check_duplicates(laminates: Sequence[Laminado]) -> List[DuplicateGroup]:
    """
    Group laminates that share the same normalized signature (stacking + type + color).
    """

    groups: dict[str, list[str]] = {}
    for laminado in laminates:
        signature = _build_duplicate_signature(laminado)
        name = str(laminado.nome or "").strip()
        if not signature or not name:
            continue
        groups.setdefault(signature, []).append(name)

    duplicate_groups: list[DuplicateGroup] = []
    for signature, names in groups.items():
        unique_names = sorted({n for n in names if n})
        if len(unique_names) < 2:
            continue
        duplicate_groups.append(
            DuplicateGroup(
                signature=signature,
                summary=_summarize_duplicate_signature(signature),
                laminates=unique_names,
            )
        )

    duplicate_groups.sort(key=lambda group: (len(group.laminates) * -1, group.summary.lower()))
    return duplicate_groups


def _is_laminate_symmetric(laminado: Laminado) -> bool:
    structural_layers = [layer for layer in laminado.camadas if _is_structural(layer)]
    count = len(structural_layers)
    if count <= 1:
        return True

    i, j = 0, count - 1
    while i < j:
        top = structural_layers[i]
        bottom = structural_layers[j]
        if not _layers_match(top, bottom):
            return False
        i += 1
        j -= 1
    return True


def _layers_match(top: Camada, bottom: Camada) -> bool:
    return _normalize_material(top.material) == _normalize_material(bottom.material) and (
        _normalize_orientation(top.orientacao) == _normalize_orientation(bottom.orientacao)
    )


def _is_structural(layer: Camada) -> bool:
    ply_type = getattr(layer, "ply_type", "") or ""
    return ply_type.strip().lower() == STRUCTURAL_LABEL


def _normalize_material(value: object) -> str:
    text = str(value or "").strip()
    collapsed = " ".join(text.split())
    return collapsed.upper()


def _normalize_orientation(value: object) -> int | None:
    try:
        return normalize_angle(value)
    except Exception:
        try:
            return normalize_angle(int(value))
        except Exception:
            return None


def _build_duplicate_signature(laminado: Laminado) -> str:
    stacking_part = _stacking_signature(laminado.camadas)
    lam_type = (laminado.tipo or "").strip().lower()
    color = str(getattr(laminado, "color_index", "") or "").strip()
    return f"{stacking_part}|{lam_type}|{color}"


def _summarize_duplicate_signature(signature: str) -> str:
    try:
        _, lam_type, color = signature.rsplit("|", 2)
    except ValueError:
        return signature
    lam_type = lam_type.strip()
    color = color.strip()
    if lam_type and color:
        return f"Tipo: {lam_type.upper()} | Cor: {color}"
    if lam_type:
        return f"Tipo: {lam_type.upper()}"
    if color:
        return f"Cor: {color}"
    return "Stacking duplicado"


def _stacking_signature(layers: Sequence[Camada]) -> str:
    if not layers:
        return "stacking:empty"

    tokens: list[str] = []
    for layer in layers:
        material = _normalize_material(layer.material)
        orientation = _normalize_orientation(layer.orientacao)
        orientation_token = "none" if orientation is None else f"{orientation:+d}"
        ply = getattr(layer, "ply_type", "") or ""
        ply_token = ply.strip().lower()
        if ply_token not in PLY_TYPE_TOKENS:
            ply_token = STRUCTURAL_LABEL
        tokens.append(f"{material}@{orientation_token}@{ply_token}")
    return ";".join(tokens)


__all__ = [
    "SymmetryResult",
    "DuplicateGroup",
    "ChecksReport",
    "run_all_checks",
    "check_symmetry",
    "check_duplicates",
]
