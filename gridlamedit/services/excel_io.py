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

    _ensure_original_available(model)

    save_grid_spreadsheet(str(output_path), model)
    _restore_preserved_columns(model.source_excel_path, output_path)
    return output_path


__all__ = ["export_grid_xlsx"]


def _ensure_original_available(model: GridModel) -> None:
    source_path = getattr(model, "source_excel_path", None)
    if not source_path:
        raise ValueError(
            "Nao foi possivel localizar o arquivo Excel original para preservar as colunas C-F."
        )
    original = Path(source_path)
    if not original.exists():
        raise ValueError(
            f"Arquivo Excel original nao encontrado em '{original}'. "
            "Nao foi possivel preservar as colunas C-F."
        )


def _restore_preserved_columns(
    source_excel: str | Path,
    output_path: Path,
    sheet_name: str = "Planilha1",
    preserved_columns: tuple[int, ...] = (3, 4, 5, 6),
) -> None:
    """
    Copia os valores das colunas preservadas (C-F) do arquivo original para a exportacao.
    """

    from openpyxl import load_workbook

    original_path = Path(source_excel)

    wb_out = load_workbook(output_path)
    if sheet_name not in wb_out.sheetnames:
        logger.warning(
            "Sheet '%s' nao encontrada no arquivo exportado. Pulando preservacao.", sheet_name
        )
        wb_out.save(output_path)
        return
    ws_out = wb_out[sheet_name]
    max_row_out = ws_out.max_row

    original_suffix = original_path.suffix.lower()
    if original_suffix == ".xls":
        preserved_data = _read_preserved_columns_from_xls(
            original_path, sheet_name, preserved_columns, max_row_out
        )
    else:
        preserved_data = _read_preserved_columns_from_xlsx(
            original_path, sheet_name, preserved_columns, max_row_out
        )

    for row_idx, row_values in enumerate(preserved_data, start=1):
        for col_idx, value in zip(preserved_columns, row_values, strict=True):
            ws_out.cell(row=row_idx, column=col_idx, value=value)

    wb_out.save(output_path)


def _read_preserved_columns_from_xlsx(
    original_path: Path,
    sheet_name: str,
    preserved_columns: tuple[int, ...],
    max_rows: int,
) -> list[list[Optional[object]]]:
    from openpyxl import load_workbook

    wb_in = load_workbook(original_path, data_only=False, read_only=False)
    if sheet_name not in wb_in.sheetnames:
        raise ValueError(
            f"Aba '{sheet_name}' nao encontrada no arquivo original '{original_path.name}'."
        )
    ws_in = wb_in[sheet_name]
    max_row = min(max_rows, ws_in.max_row)
    data: list[list[Optional[object]]] = []
    for row_idx in range(1, max_row + 1):
        row_data: list[Optional[object]] = []
        for col_idx in preserved_columns:
            row_data.append(ws_in.cell(row=row_idx, column=col_idx).value)
        data.append(row_data)
    return data


def _read_preserved_columns_from_xls(
    original_path: Path,
    sheet_name: str,
    preserved_columns: tuple[int, ...],
    max_rows: int,
) -> list[list[Optional[object]]]:
    try:
        import xlrd  # type: ignore
    except ImportError as exc:  # pragma: no cover - dependAancia opcional
        raise ValueError(
            "Preservar colunas C-F de arquivos '.xls' requer a dependAancia 'xlrd==1.2.0'."
        ) from exc

    workbook = xlrd.open_workbook(original_path)  # type: ignore[arg-type]
    try:
        sheet = workbook.sheet_by_name(sheet_name)
    except xlrd.biffh.XLRDError as exc:  # type: ignore[attr-defined]
        raise ValueError(
            f"Aba '{sheet_name}' nao encontrada no arquivo original '{original_path.name}'."
        ) from exc

    max_row = min(max_rows, sheet.nrows)
    data: list[list[Optional[object]]] = []
    for row_idx in range(max_row):
        row_data: list[Optional[object]] = []
        for col_idx in preserved_columns:
            if col_idx - 1 < sheet.ncols:
                cell_value = sheet.cell_value(row_idx, col_idx - 1)
            else:
                cell_value = None
            row_data.append(cell_value)
        data.append(row_data)
    return data
