"""Excel export helpers for GridLamEdit."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

from gridlamedit.io.spreadsheet import GridModel, save_grid_spreadsheet

logger = logging.getLogger(__name__)


def export_grid_xlsx(
    model: GridModel,
    path: Path,
    template_info: Optional[dict] = None,
) -> Path:
    """
    Persist the current GridModel into an Excel workbook matching the import layout.

    Parameters
    ----------
    model:
        GridModel instance containing the state to export.
    path:
        Target file path. When no ``.xlsx``/``.xls`` suffix is provided, ``.xlsx`` is used.
    template_info:
        Placeholder for future template metadata (unused for now).

    Returns
    -------
    Path
        The resolved output path written to disk.
    """
    if model is None:
        raise ValueError("Model ausente para exportacao da planilha.")

    output_path = Path(path)
    if output_path.suffix.lower() not in {".xlsx", ".xls"}:
        output_path = output_path.with_suffix(".xlsx")

    if template_info:
        logger.info("Ignorando template_info nao utilizado: keys=%s", list(template_info))

    save_grid_spreadsheet(str(output_path), model)
    return output_path


__all__ = ["export_grid_xlsx"]

