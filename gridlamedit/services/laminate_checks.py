"""Domain checks applied to laminates before exporting."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Sequence, Tuple

import math

from gridlamedit.io.spreadsheet import (
    Camada,
    DEFAULT_PLY_TYPE,
    Laminado,
    PLY_TYPE_OPTIONS,
    normalize_angle,
    normalize_ply_type_label,
    ply_type_signature_token,
)


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
class LaminateSymmetryEvaluation:
    """Detailed symmetry result for a single laminate."""

    structural_rows: List[int]
    centers: List[int]
    is_symmetric: bool
    first_mismatch: Tuple[int, int] | None = None


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
        evaluation = evaluate_symmetry_for_layers(laminado.camadas)
        if evaluation.is_symmetric:
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
    return evaluate_symmetry_for_layers(laminado.camadas).is_symmetric


def _normalize_material(value: object) -> str:
    text = str(value or "").strip()
    collapsed = " ".join(text.split())
    return collapsed.upper()


def _normalize_orientation(value: object) -> float | None:
    try:
        return normalize_angle(value)
    except Exception:
        try:
            return normalize_angle(str(value))
        except Exception:
            return None


def _normalized_orientation_token(value: object) -> float | None:
    """Normalize orientation returning ``None`` for blanks or invalid values."""
    try:
        return normalize_angle(value)
    except Exception:
        try:
            text = str(value).strip()
        except Exception:
            return None
        if not text:
            return None
        try:
            return normalize_angle(text)
        except Exception:
            return None


def _normalized_material_token(value: object) -> str:
    text = str(value or "").strip()
    return " ".join(text.split()).lower()


def _rows_match(
    layers: Sequence[Camada], left_idx: int, right_idx: int
) -> tuple[bool, tuple[int, int] | None]:
    """
    Compare two rows for symmetry (orientation + material).

    Missing rows are treated as empty. Empty only matches empty.
    """
    left_layer = layers[left_idx] if 0 <= left_idx < len(layers) else None
    right_layer = layers[right_idx] if 0 <= right_idx < len(layers) else None

    left_orientation = (
        _normalized_orientation_token(getattr(left_layer, "orientacao", None))
        if left_layer
        else None
    )
    right_orientation = (
        _normalized_orientation_token(getattr(right_layer, "orientacao", None))
        if right_layer
        else None
    )
    if left_orientation is None or right_orientation is None:
        if not (left_orientation is None and right_orientation is None):
            return False, (left_idx, right_idx)
    elif not math.isclose(left_orientation, right_orientation, abs_tol=1e-6):
        return False, (left_idx, right_idx)

    left_material = (
        _normalized_material_token(getattr(left_layer, "material", "") if left_layer else "")
    )
    right_material = (
        _normalized_material_token(getattr(right_layer, "material", "") if right_layer else "")
    )
    if left_material or right_material:
        if left_material != right_material:
            return False, (left_idx, right_idx)

    return True, None


def evaluate_symmetry_for_layers(layers: Sequence[Camada]) -> LaminateSymmetryEvaluation:
    """
    Evaluate laminate symmetry based on valid sequences (ply_type != ``PLY_TYPE_OPTIONS[1]``).
    """
    structural_rows: list[int] = [
        idx
        for idx, camada in enumerate(layers)
        if normalize_ply_type_label(getattr(camada, "ply_type", DEFAULT_PLY_TYPE)) != PLY_TYPE_OPTIONS[1]
    ]
    centers: list[int] = []
    mismatch: tuple[int, int] | None = None
    symmetric = False

    count = len(structural_rows)
    if count == 0:
        return LaminateSymmetryEvaluation(
            structural_rows=structural_rows,
            centers=centers,
            is_symmetric=False,
            first_mismatch=None,
        )

    if count % 2 == 1:
        mid = count // 2
        centers = [structural_rows[mid]]
        symmetric = True
        for offset in range(1, mid + 1):
            left = structural_rows[mid - offset]
            right = structural_rows[mid + offset]
            matches, bad_pair = _rows_match(layers, left, right)
            if not matches:
                symmetric = False
                mismatch = bad_pair
                break
    else:
        mid_left = count // 2 - 1
        mid_right = count // 2
        centers = [structural_rows[mid_left], structural_rows[mid_right]]
        symmetric = True
        for offset in range(0, mid_left + 1):
            left = structural_rows[mid_left - offset]
            right = structural_rows[mid_right + offset]
            matches, bad_pair = _rows_match(layers, left, right)
            if not matches:
                symmetric = False
                mismatch = bad_pair
                break
        if symmetric:
            matches, bad_pair = _rows_match(layers, centers[0], centers[1])
            if not matches:
                symmetric = False
                mismatch = bad_pair

    return LaminateSymmetryEvaluation(
        structural_rows=structural_rows,
        centers=centers,
        is_symmetric=symmetric,
        first_mismatch=mismatch,
    )


def _orientation_token(value: float | None) -> str:
    if value is None:
        return "none"
    number = float(value)
    if math.isclose(number, 0.0, abs_tol=1e-9):
        number = 0.0
    if number.is_integer():
        base = str(int(number))
    else:
        base = f"{number}".rstrip("0").rstrip(".")
    if base in {"", "-0"}:
        base = "0"
    return f"+{base}" if number >= 0 else base


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
        if getattr(layer, "orientacao", None) is None:
            continue
        material = _normalize_material(layer.material)
        orientation = _normalize_orientation(layer.orientacao)
        orientation_token = _orientation_token(orientation)
        ply_token = ply_type_signature_token(getattr(layer, "ply_type", ""))
        tokens.append(f"{material}@{orientation_token}@{ply_token}")
    if not tokens:
        return "stacking:empty"
    return ";".join(tokens)


__all__ = [
    "SymmetryResult",
    "LaminateSymmetryEvaluation",
    "DuplicateGroup",
    "ChecksReport",
    "run_all_checks",
    "check_symmetry",
    "check_duplicates",
    "evaluate_symmetry_for_layers",
]
