"""Reassociate laminates by contour signatures when cell IDs change."""

from __future__ import annotations

from dataclasses import dataclass, field
import logging
from typing import Dict, Iterable, List, Sequence, Tuple

from gridlamedit.io.spreadsheet import GridModel

logger = logging.getLogger(__name__)
MAX_CONTOUR_SIDES = 30


def _normalize_contour_token(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    return " ".join(text.split()).casefold()


def _contour_signature(values: Sequence[str]) -> Tuple[str, ...]:
    normalized = [_normalize_contour_token(value) for value in values]
    if len(normalized) > MAX_CONTOUR_SIDES:
        normalized = normalized[:MAX_CONTOUR_SIDES]
    while normalized and not normalized[-1]:
        normalized.pop()
    return tuple(normalized)


@dataclass
class ReassociationEntry:
    laminate: str
    old_cell: str
    new_cell: str
    contours: Tuple[str, ...]


@dataclass
class ReassociationIssue:
    laminate: str
    old_cell: str
    details: str
    contours: Tuple[str, ...] = ()


@dataclass
class ReassociationReport:
    reassociated: List[ReassociationEntry] = field(default_factory=list)
    conflicts: List[ReassociationIssue] = field(default_factory=list)
    missing_contours: List[ReassociationIssue] = field(default_factory=list)
    not_found: List[ReassociationIssue] = field(default_factory=list)
    unmapped_new_cells: List[str] = field(default_factory=list)


def _build_contour_index(contours: Dict[str, Sequence[str]]) -> Dict[Tuple[str, ...], List[str]]:
    index: Dict[Tuple[str, ...], List[str]] = {}
    for cell_id, values in contours.items():
        signature = _contour_signature(values)
        if not any(signature):
            continue
        index.setdefault(signature, []).append(cell_id)
    return index


def reassociate_laminates_by_contours(
    old_model: GridModel,
    new_model: GridModel,
    *,
    apply: bool = True,
) -> ReassociationReport:
    """Reassociate laminates from old_model to new_model using contour signatures."""
    report = ReassociationReport()
    contour_index = _build_contour_index(getattr(new_model, "cell_contours", {}))
    if apply:
        cell_map: Dict[str, str] = {}
    else:
        cell_map = dict(new_model.cell_to_laminate)

    for old_cell, laminate_name in old_model.cell_to_laminate.items():
        if not laminate_name:
            continue

        old_contours = getattr(old_model, "cell_contours", {}).get(old_cell)
        if not old_contours:
            report.missing_contours.append(
                ReassociationIssue(
                    laminate=laminate_name,
                    old_cell=old_cell,
                    details="Contornos nao encontrados na base atual.",
                )
            )
            continue

        signature = _contour_signature(old_contours)
        if not any(signature):
            report.missing_contours.append(
                ReassociationIssue(
                    laminate=laminate_name,
                    old_cell=old_cell,
                    details="Contornos vazios ou invalidos.",
                    contours=signature,
                )
            )
            continue

        candidates = contour_index.get(signature, [])
        if not candidates:
            report.not_found.append(
                ReassociationIssue(
                    laminate=laminate_name,
                    old_cell=old_cell,
                    details="Nenhuma celula equivalente encontrada na nova planilha.",
                    contours=signature,
                )
            )
            continue
        if len(candidates) > 1:
            report.conflicts.append(
                ReassociationIssue(
                    laminate=laminate_name,
                    old_cell=old_cell,
                    details=f"Contorno ambiguo em {len(candidates)} celulas: {', '.join(candidates)}.",
                    contours=signature,
                )
            )
            continue

        target_cell = candidates[0]
        current = cell_map.get(target_cell)
        if current and current != laminate_name:
            report.conflicts.append(
                ReassociationIssue(
                    laminate=laminate_name,
                    old_cell=old_cell,
                    details=(
                        f"Celula {target_cell} ja associada ao laminado '{current}'."
                    ),
                    contours=signature,
                )
            )
            continue

        if current != laminate_name:
            cell_map[target_cell] = laminate_name
            report.reassociated.append(
                ReassociationEntry(
                    laminate=laminate_name,
                    old_cell=old_cell,
                    new_cell=target_cell,
                    contours=signature,
                )
            )

    if apply:
        new_model.cell_to_laminate = cell_map
        _rebuild_laminate_cells(new_model)
    report.unmapped_new_cells = [
        cell_id
        for cell_id in new_model.celulas_ordenadas
        if cell_id not in cell_map
    ]
    return report


def _rebuild_laminate_cells(model: GridModel) -> None:
    for laminate in model.laminados.values():
        laminate.celulas = []

    for cell_id in model.celulas_ordenadas:
        laminate_name = model.cell_to_laminate.get(cell_id)
        if not laminate_name:
            continue
        laminate = model.laminados.get(laminate_name)
        if laminate is None:
            logger.warning(
                "Laminado '%s' nao encontrado para a celula %s.",
                laminate_name,
                cell_id,
            )
            continue
        laminate.celulas.append(cell_id)
