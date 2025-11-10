"""Laminate creation helpers."""

from __future__ import annotations

import logging
from collections import OrderedDict
from typing import Union

from gridlamedit.io.spreadsheet import (
    GridModel,
    Laminado,
    MIN_COLOR_INDEX,
    MAX_COLOR_INDEX,
    normalize_color_index,
    DEFAULT_COLOR_INDEX,
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
