"""Export Virtual Stacking data to a CATIA-compatible Excel template."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, Sequence

import xlwt

from gridlamedit.io.spreadsheet import DEFAULT_ROSETTE_LABEL, normalize_angle

__all__ = ["export_virtual_stacking"]

# Fixed headers taken from the reference template ``Virtual Stacking.xls``.
_HEADER_ROW_ZERO = ["Ply", "SA"]
_HEADER_ROW_ONE = ["Virtual Sequence", "Sequence", "Virtual Ply", "Ply", "Material", "Rosette"]


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


def export_virtual_stacking(
    layers: Sequence[object],
    cells: Sequence[object],
    path: Path,
    *,
    sheet_name: str = "Planilha1",
) -> Path:
    """
    Persist the current Virtual Stacking view into a spreadsheet that mirrors
    the bundled ``Virtual Stacking.xls`` template used by CATIA V5.

    Parameters
    ----------
    layers:
        Ordered collection of layer descriptors. Each item must expose the
        attributes ``sequence_label``, ``ply_label``, ``material`` and ``rosette``.
    cells:
        Ordered collection of cell descriptors. Each item must expose ``cell_id``
        and ``laminate``; the laminate must have a ``camadas`` list containing
        orientation information at the same index as ``layers``.
    path:
        Target output path. The result will be written as ``.xls`` regardless of
        the provided suffix to keep parity with the reference template.
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
    if output_path.suffix.lower() != ".xls":
        output_path = output_path.with_suffix(".xls")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    book = xlwt.Workbook()
    sheet = book.add_sheet(sheet_name)

    total_columns = len(_HEADER_ROW_ONE) + len(cells)

    # Row 0: leading identifiers ("Ply", "SA", then fill with "#" like the template)
    header_zero = list(_HEADER_ROW_ZERO) + ["#"] * max(0, total_columns - len(_HEADER_ROW_ZERO))
    for col, value in enumerate(header_zero):
        sheet.write(0, col, value)

    # Row 1: titles + cell identifiers.
    header_one = list(_HEADER_ROW_ONE) + [str(getattr(cell, "cell_id", f"C{idx+1}")) for idx, cell in enumerate(cells)]
    for col, value in enumerate(header_one):
        sheet.write(1, col, value)

    # Data rows start at row index 2.
    for row_idx, layer in enumerate(layers, start=2):
        sequence_label = getattr(layer, "sequence_label", "") or f"Seq.{row_idx - 1}"
        ply_label = getattr(layer, "ply_label", "") or f"Ply.{row_idx - 1}"
        material = getattr(layer, "material", "") or ""
        rosette = _safe_rosette(getattr(layer, "rosette", ""))

        sheet.write(row_idx, 0, sequence_label)
        sheet.write(row_idx, 1, "")  # Keep the second column empty as in the template
        sheet.write(row_idx, 2, ply_label)
        sheet.write(row_idx, 3, "")  # Keep the fourth column empty as in the template
        sheet.write(row_idx, 4, material)
        sheet.write(row_idx, 5, rosette)

        for cell_offset, cell in enumerate(cells):
            target_col = len(_HEADER_ROW_ONE) + cell_offset
            laminate = getattr(cell, "laminate", None)
            layers_list: Iterable[object] = getattr(laminate, "camadas", []) if laminate else []
            orientation_value: float | str | int | None = None
            if isinstance(layers_list, Sequence):
                if row_idx - 2 < len(layers_list):
                    layer_obj = layers_list[row_idx - 2]
                    orientation_value = _normalize_orientation(getattr(layer_obj, "orientacao", None))
            if orientation_value in (None, "", "Empty"):
                sheet.write(row_idx, target_col, "")
            else:
                sheet.write(row_idx, target_col, orientation_value)

    # Terminator row: mirrors the reference template end marker.
    terminator_row = len(layers) + 2
    sheet.write(terminator_row, 0, "##")

    book.save(str(output_path))
    return output_path
