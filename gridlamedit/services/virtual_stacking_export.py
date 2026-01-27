"""Export Virtual Stacking data to a CATIA-compatible Excel template."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, Sequence

from gridlamedit.io.spreadsheet import DEFAULT_ROSETTE_LABEL, normalize_angle

__all__ = ["export_virtual_stacking"]

# Fixed headers taken from the reference template ``Grid Lam Vs Exported_RevC.xls``.
_HEADER_ROW_ZERO = ["Sequence", "Cell"]
_HEADER_ROW_ONE = ["Virtual Sequence", "Sequence", "Material", "Rosette"]


def _normalize_orientation(value: object) -> float | str | int | None:
    """Convert an orientation value to a float when possible; return empty on blanks."""
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        return float(normalize_angle(value))
    except Exception:
        try:
            return float(text)
        except Exception:
            return text


def _safe_rosette(value: object) -> str:
    text = str(value or "").strip()
    return text or DEFAULT_ROSETTE_LABEL


def _build_virtual_stacking_rows(
    layers: Sequence[object],
    cells: Sequence[object],
) -> list[list[object]]:
    total_columns = len(_HEADER_ROW_ONE) + len(cells)

    header_zero = list(_HEADER_ROW_ZERO) + ["#"] * max(0, total_columns - len(_HEADER_ROW_ZERO))
    header_one = list(_HEADER_ROW_ONE) + [
        str(getattr(cell, "cell_id", f"C{idx + 1}")) for idx, cell in enumerate(cells)
    ]

    rows: list[list[object]] = [header_zero, header_one]

    for row_idx, layer in enumerate(layers, start=1):
        sequence_label = getattr(layer, "sequence_label", "") or f"Seq.{row_idx}"
        material = getattr(layer, "material", "") or ""
        rosette = _safe_rosette(getattr(layer, "rosette", ""))

        row: list[object] = [sequence_label, "", material, rosette]

        for cell in cells:
            laminate = getattr(cell, "laminate", None)
            layers_list: Iterable[object] = getattr(laminate, "camadas", []) if laminate else []
            orientation_value: float | str | int | None = None
            if isinstance(layers_list, Sequence):
                if row_idx - 1 < len(layers_list):
                    layer_obj = layers_list[row_idx - 1]
                    orientation_value = _normalize_orientation(
                        getattr(layer_obj, "orientacao", None)
                    )
            if orientation_value in (None, "", "Empty"):
                row.append("")
            else:
                row.append(orientation_value)

        rows.append(row)

    terminator_row = [""] * total_columns
    if terminator_row:
        terminator_row[0] = "##"
    else:
        terminator_row = ["##"]
    rows.append(terminator_row)
    return rows


def _write_rows_xlsx(rows: list[list[object]], output_path: Path, sheet_name: str) -> None:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    for row_idx, row in enumerate(rows, start=1):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    wb.save(output_path)


def export_virtual_stacking(
    layers: Sequence[object],
    cells: Sequence[object],
    path: Path,
    *,
    sheet_name: str = "Planilha1",
) -> Path:
    """
    Persist the current Virtual Stacking view into a spreadsheet that mirrors
    the bundled ``Grid Lam Vs Exported_RevC.xls`` template used by CATIA V5.

    Parameters
    ----------
    layers:
        Ordered collection of layer descriptors. Each item must expose the
        attributes ``sequence_label``, ``material`` and ``rosette``.
    cells:
        Ordered collection of cell descriptors. Each item must expose ``cell_id``
        and ``laminate``; the laminate must have a ``camadas`` list containing
        orientation information at the same index as ``layers``.
    path:
        Target output path. When no ``.xlsx`` suffix is provided, ``.xlsx`` is used.
        Apenas ``.xlsx`` sera escrito.
    sheet_name:
        Name of the worksheet to create (defaults to ``Planilha1``).

    Returns
    -------
    Path
        The resolved path of the written file.
    """
    if not cells:
        raise ValueError("Nenhuma coluna de Virtual Stacking para exportar.")

    output_path = Path(path)
    if output_path.suffix.lower() != ".xlsx":
        output_path = output_path.with_suffix(".xlsx")
    xlsx_path = output_path
    output_path.parent.mkdir(parents=True, exist_ok=True)

    rows = _build_virtual_stacking_rows(layers, cells)

    try:
        _write_rows_xlsx(rows, xlsx_path, sheet_name)
    except Exception as exc:
        raise ValueError(f"Falha ao exportar arquivo .xlsx: {exc}") from exc

    return output_path
