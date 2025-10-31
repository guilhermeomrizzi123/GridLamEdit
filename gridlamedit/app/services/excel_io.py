"""Excel IO helpers for GridLamEdit."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, Optional

import math
import re

import pandas as pd

from ..models import Cell, GridDoc, Laminate, Layer


def _normalize_angle(value: object) -> float:
    """Convert CATIA-like angle representations to float degrees."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return 0.0
        return float(value)
    text = str(value).strip()
    if not text:
        return 0.0
    cleaned = re.sub(r"[^\d\-\.+]", "", text)
    try:
        return float(cleaned)
    except ValueError:
        print(f"Warning: could not parse angle value '{value}', defaulting to 0.")
        return 0.0


def load_grid_xlsx(path: str) -> GridDoc:
    """Load a GridDoc from an Excel file exported by CATIA Grid Design."""
    file_path = Path(path)
    if not file_path.exists():
        print(f"Warning: file '{file_path}' not found. Returning empty GridDoc.")
        return GridDoc()

    doc = GridDoc()
    try:
        workbook = pd.ExcelFile(file_path)
    except Exception as exc:  # pragma: no cover - defensive
        print(f"Warning: failed to open '{file_path}': {exc}")
        return doc

    # Parse cells sheet.
    cells: list[Cell] = []
    try:
        cells_df = workbook.parse("Cells")
    except ValueError:
        print("Warning: 'Cells' sheet missing; no cell assignments loaded.")
        cells_df = pd.DataFrame(columns=["Cell ID", "Laminate Name"])

    cell_id_column = _find_first_column(cells_df, ["Cell ID", "Cell", "ID"])
    laminate_column = _find_first_column(cells_df, ["Laminate Name", "Laminate", "Stack"])

    if cell_id_column is None:
        print("Warning: could not find a cell ID column; skipping cell data.")
    else:
        for _, row in cells_df.iterrows():
            cell_id = row.get(cell_id_column)
            if pd.isna(cell_id):
                continue
            laminate_name = None
            if laminate_column and not pd.isna(row.get(laminate_column)):
                laminate_name = str(row.get(laminate_column)).strip() or None
            cells.append(Cell(id=str(cell_id).strip(), laminate_name=laminate_name))

    doc.cells = cells

    # Parse laminate sheets.
    for sheet_name in (name for name in workbook.sheet_names if name != "Cells"):
        laminate = _parse_laminate_sheet(workbook, sheet_name)
        if laminate is None:
            continue
        doc.laminates[laminate.name] = laminate

    doc.ensure_associations()
    return doc


def save_grid_xlsx(doc: GridDoc, path: str) -> None:
    """Persist the GridDoc into an Excel file compatible with Grid Design."""
    target_path = Path(path)
    target_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(target_path, engine="openpyxl") as writer:
        _write_cells_sheet(writer, doc.cells)
        _write_laminate_sheets(writer, doc.laminates.values())


def _find_first_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    """Return the first matching column name from candidates."""
    for candidate in candidates:
        if candidate in df.columns:
            return candidate
    lowered = {col.lower(): col for col in df.columns}
    for candidate in candidates:
        key = candidate.lower()
        if key in lowered:
            return lowered[key]
    return None


def _parse_laminate_sheet(workbook: pd.ExcelFile, sheet_name: str) -> Optional[Laminate]:
    """Parse laminate metadata and layers from a sheet."""
    try:
        metadata_df = workbook.parse(sheet_name, nrows=1)
    except ValueError:
        print(f"Warning: sheet '{sheet_name}' is empty; skipping.")
        return None

    if metadata_df.empty:
        print(f"Warning: sheet '{sheet_name}' has no metadata; skipping.")
        return None

    meta_row = metadata_df.iloc[0]
    name = str(meta_row.get("Name") or "").strip()
    if not name:
        name = sheet_name.replace("Laminate_", "").strip() or sheet_name
        print(
            f"Warning: laminate sheet '{sheet_name}' missing name; "
            f"using '{name}'."
        )

    color = str(meta_row.get("Color") or "").strip()
    laminate_type = str(meta_row.get("Type") or "").strip()
    associated_raw = str(meta_row.get("Associated Cells") or "").strip()
    associated_cells = [cell.strip() for cell in associated_raw.split(",") if cell.strip()]

    data_rows_used = metadata_df.shape[0] + 1  # include header
    skiprows = data_rows_used + 1  # blank row separator

    stacking_df = workbook.parse(sheet_name, skiprows=skiprows)
    layers: list[Layer] = []
    symmetry_index: Optional[int] = None

    if not stacking_df.empty:
        stacking_df = stacking_df.rename(
            columns={col: col.strip() for col in stacking_df.columns if isinstance(col, str)}
        )
        index_column = _find_first_column(stacking_df, ["Index", "#"])
        angle_column = _find_first_column(stacking_df, ["Angle", "Angle Deg"])
        active_column = _find_first_column(stacking_df, ["Active"])
        symmetry_column = _find_first_column(stacking_df, ["Symmetry"])

        for _, row in stacking_df.iterrows():
            if row.isna().all():
                continue

            material = str(row.get("Material") or "").strip()
            if not material:
                print(
                    f"Warning: laminate '{name}' has a layer without material; skipping."
                )
                continue

            if index_column and not pd.isna(row.get(index_column)):
                try:
                    layer_index = int(row.get(index_column))
                except (TypeError, ValueError):
                    layer_index = len(layers)
            else:
                layer_index = len(layers)

            angle_value = _normalize_angle(row.get(angle_column)) if angle_column else 0.0
            active_value = True
            if active_column:
                active_raw = str(row.get(active_column)).strip().lower()
                active_value = active_raw in {"yes", "y", "true", "1"}

            if symmetry_column:
                symmetry_raw = str(row.get(symmetry_column)).strip().lower()
                if symmetry_raw in {"yes", "y", "true", "1"}:
                    symmetry_index = len(layers)

            layers.append(
                Layer(
                    index=layer_index,
                    material=material,
                    angle_deg=angle_value,
                    active=active_value,
                )
            )

    # Reindex layers sequentially.
    for idx, layer in enumerate(layers):
        layer.index = idx

    laminate = Laminate(
        name=name,
        color=color,
        type=laminate_type,
        layers=layers,
        associated_cells=associated_cells,
        symmetry_index=symmetry_index,
    )
    return laminate


def _write_cells_sheet(writer: pd.ExcelWriter, cells: Iterable[Cell]) -> None:
    """Write the Cells sheet from the current document."""
    rows = [
        {"Cell ID": cell.id, "Laminate Name": cell.laminate_name or ""}
        for cell in cells
    ]
    cells_df = pd.DataFrame(rows)
    cells_df.to_excel(writer, sheet_name="Cells", index=False)


def _write_laminate_sheets(writer: pd.ExcelWriter, laminates: Iterable[Laminate]) -> None:
    """Write one sheet per laminate definition."""
    used_names: set[str] = set()
    for laminate in laminates:
        sheet_name = _make_unique_sheet_name(f"Laminate_{laminate.name}", used_names)
        used_names.add(sheet_name)

        metadata_df = pd.DataFrame(
            [
                {
                    "Name": laminate.name,
                    "Color": laminate.color,
                    "Type": laminate.type,
                    "Associated Cells": ", ".join(laminate.associated_cells),
                }
            ]
        )

        stack_rows = []
        for idx, layer in enumerate(laminate.layers):
            stack_rows.append(
                {
                    "Index": idx,
                    "Material": layer.material,
                    "Angle": layer.angle_deg,
                    "Active": "Yes" if layer.active else "No",
                    "Symmetry": "Yes" if laminate.symmetry_index == idx else "",
                }
            )

        stack_df = pd.DataFrame(stack_rows)
        metadata_df.to_excel(writer, sheet_name=sheet_name, index=False)
        startrow = metadata_df.shape[0] + 2  # data rows + header + blank
        stack_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=startrow)


def _make_unique_sheet_name(base: str, used: set[str]) -> str:
    """Return a sheet name compatible with Excel and unique in workbook."""
    invalid_pattern = r"[:\\/?*\[\]]"
    safe_base = re.sub(invalid_pattern, "_", base)[:31]
    if safe_base not in used:
        return safe_base
    counter = 1
    while True:
        candidate = f"{safe_base[:28]}_{counter}"
        if candidate not in used:
            return candidate
        counter += 1
