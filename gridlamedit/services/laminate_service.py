"""Laminate creation helpers."""

from __future__ import annotations

import logging
from collections import OrderedDict
from typing import Callable, Optional, Union

from gridlamedit.io.spreadsheet import (
    GridModel,
    Laminado,
    MIN_COLOR_INDEX,
    MAX_COLOR_INDEX,
    normalize_color_index,
    DEFAULT_COLOR_INDEX,
    count_oriented_layers,
    StackingTableModel,
)

logger = logging.getLogger(__name__)


class LaminateCreationError(ValueError):
    """Raised when laminate creation fails due to invalid input."""


def _normalize_color(value: Union[str, int]) -> int:
    try:
        return normalize_color_index(value, DEFAULT_COLOR_INDEX)
    except Exception:  # pragma: no cover - defensive
        return DEFAULT_COLOR_INDEX


def create_laminate_with_association(
    model: GridModel,
    nome: str,
    cor: Union[str, int],
    tipo: str,
    celula_id: Union[str, int],
    *,
    tag: str = "",
) -> Laminado:
    """
    Create a Laminate with empty stacking, persist it on ``model`` and bind it to ``celula_id``.

    Parameters
    ----------
    model:
        Active ``GridModel`` receiving the new laminate.
    nome:
        Unique laminate name.
    cor:
        Color index (1-150). Falls back to ``DEFAULT_COLOR_INDEX`` on invalid values.
    tipo:
        Laminate type label.
    celula_id:
        Cell identifier that must exist in ``model.celulas_ordenadas``.

    Returns
    -------
    Laminado
        The newly created laminate.
    """

    if model is None:
        raise LaminateCreationError("Modelo inexistente para criar laminado.")

    name = str(nome or "").strip()
    if not name:
        raise LaminateCreationError("O campo Nome não pode ficar vazio.")

    existing = getattr(model, "laminados", None)
    if isinstance(existing, OrderedDict):
        target_map = existing
    else:
        target_map = OrderedDict(model.laminados or {})
        model.laminados = target_map

    if name in target_map:
        raise LaminateCreationError("Já existe um laminado com este nome.")

    lam_type = str(tipo or "").strip()
    if not lam_type:
        raise LaminateCreationError("Selecione um tipo de laminado válido.")

    color_index = _normalize_color(cor)
    color_index = max(MIN_COLOR_INDEX, min(MAX_COLOR_INDEX, color_index))

    cell = str(celula_id or "").strip()
    if not cell:
        raise LaminateCreationError("Selecione uma célula para associar.")

    if model.celulas_ordenadas and cell not in model.celulas_ordenadas:
        raise LaminateCreationError("Selecione uma célula válida para associar.")

    laminado = Laminado(
        nome=name,
        tipo=lam_type,
        color_index=color_index,
        celulas=[cell],
        camadas=[],
        tag=str(tag or "").strip(),
    )

    target_map[name] = laminado
    previous = model.cell_to_laminate.get(cell)
    if previous and previous in target_map:
        old = target_map[previous]
        if cell in old.celulas:
            old.celulas.remove(cell)
    model.cell_to_laminate[cell] = name
    laminado.celulas = [cell]

    try:
        model.mark_dirty(True)
    except AttributeError:  # pragma: no cover - defensive
        logger.debug("GridModel missing mark_dirty; skipping dirty mark.")

    return laminado


def _build_auto_name(base: str, suffix_index: int, tag: str) -> str:
    suffix = f".{suffix_index}" if suffix_index > 0 else ""
    tag_text = str(tag or "").strip()
    tag_suffix = f"({tag_text})" if tag_text else ""
    return f"{base}{suffix}{tag_suffix}"


def auto_name_for_layers(
    model: Optional[GridModel],
    *,
    layer_count: int,
    tag: str = "",
    target: Optional[Laminado] = None,
) -> str:
    """
    Generate an automatic name following the L<N>[.k][(Tag)] pattern.

    The suffix ``.k`` is assigned based on the number of laminates with the same
    ``layer_count`` already present in ``model`` (in insertion order). The Tag
    text is appended in parentheses when provided.
    """

    count = max(0, int(layer_count))
    base = f"L{count}"
    suffix_index = 0
    used_names: set[str] = set()

    if model is not None:
        same_count = [
            lam
            for lam in model.laminados.values()
            if count_oriented_layers(getattr(lam, "camadas", [])) == count
        ]
        if target in same_count:
            suffix_index = same_count.index(target)
        else:
            suffix_index = len(same_count)
        used_names = {name for name, lam in model.laminados.items() if lam is not target}

    candidate = _build_auto_name(base, suffix_index, tag)
    while candidate in used_names:
        suffix_index += 1
        candidate = _build_auto_name(base, suffix_index, tag)
    return candidate


def auto_name_for_laminate(
    model: Optional[GridModel], laminate: Laminado
) -> str:
    """Convenience wrapper around :func:`auto_name_for_layers` for a laminate."""

    return auto_name_for_layers(
        model,
        layer_count=count_oriented_layers(getattr(laminate, "camadas", [])),
        tag=getattr(laminate, "tag", ""),
        target=laminate,
    )


def sync_material_by_sequence(
    model: Optional[GridModel],
    row: int,
    material: str,
    *,
    stacking_model_provider: Optional[
        Callable[[Laminado], Optional[StackingTableModel]]
    ] = None,
) -> list[Laminado]:
    """
    Apply ``material`` to the same ``row`` across all laminates.

    Returns the list of laminates that were changed.
    """

    if model is None or row < 0:
        return []

    updated: list[Laminado] = []
    new_material = str(material or "").strip()
    for laminate in model.laminados.values():
        layers = getattr(laminate, "camadas", [])
        if row >= len(layers):
            continue
        target_layer = layers[row]
        target_value = new_material
        if str(getattr(target_layer, "material", "") or "").strip() == target_value:
            continue

        stacking_model = stacking_model_provider(laminate) if stacking_model_provider else None
        if stacking_model is not None:
            if stacking_model.apply_field_value(row, StackingTableModel.COL_MATERIAL, target_value):
                laminate.camadas = stacking_model.layers()
                updated.append(laminate)
                continue

        target_layer.material = target_value
        updated.append(laminate)
    return updated
