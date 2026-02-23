"""Virtual Stacking dialog and model."""

from __future__ import annotations

import copy
from collections import Counter, OrderedDict
import math
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Optional

from PySide6 import QtCore, QtGui, QtWidgets

# Allow running as a script without installing the package.
if __package__ in (None, ""):
    project_root = Path(__file__).resolve().parents[2]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))

from gridlamedit.io.spreadsheet import (
    Camada,
    DEFAULT_PLY_TYPE,
    DEFAULT_ROSETTE_LABEL,
    GridModel,
    Laminado,
    StackingTableModel,
    WordWrapHeader,
    format_orientation_value,
    orientation_highlight_color,
    ORIENTATION_SYMMETRY_ROLE,
    count_oriented_layers,
    is_structural_ply_label,
    normalize_angle,
    normalize_ply_type_label,
    PLY_TYPE_OPTIONS,
)
from gridlamedit.app.delegates import (
    MaterialComboDelegate,
    OrientationComboDelegate,
    PlyTypeComboDelegate,
)
from gridlamedit.services.laminate_checks import (
    LaminateSymmetryEvaluation,
    evaluate_symmetry_for_layers,
    evaluate_laminate_balance_clt,
)
from gridlamedit.services.laminate_service import auto_name_for_laminate, sync_material_by_sequence
from gridlamedit.core.paths import package_path
from gridlamedit.services.material_registry import (
    available_materials as registry_available_materials,
)
from gridlamedit.services.project_query import (
    project_distinct_orientations,
    project_most_used_material,
)
from gridlamedit.services.excel_io import ensure_layers_have_material
from gridlamedit.services.virtual_stacking_export import export_virtual_stacking


@dataclass
class VirtualStackingLayer:
    """Represent a stacking layer sequence entry."""

    sequence_label: str
    ply_label: str
    ply_type: str = DEFAULT_PLY_TYPE
    material: str = ""
    rosette: str = DEFAULT_ROSETTE_LABEL


@dataclass
class VirtualStackingCell:
    """Represent orientations of a laminado for each grid cell."""

    cell_id: str
    laminate: Laminado


_PREFIX_NUMBER_PATTERN = re.compile(r"^([^.]*?)[.]?(\d+)$")


class _InsertLayerCommand(QtGui.QUndoCommand):
    """Undoable insertion of new layers for a laminate."""

    def __init__(
        self,
        model: StackingTableModel,
        laminate: Laminado,
        positions: list[int],
        default_material: str = "",
        default_orientation: float | None = None,
    ) -> None:
        super().__init__("Insert layer")
        self._model = model
        self._laminate = laminate
        self._positions = positions
        self._default_material = default_material
        self._default_orientation = default_orientation
        self._inserted: list[int] = []

    def redo(self) -> None:
        self._inserted = []
        for pos in sorted(set(self._positions), reverse=True):
            target_pos = max(0, min(pos, len(self._model.layers())))
            self._model.insert_layer(
                target_pos,
                Camada(
                    idx=0,
                    material=self._default_material,
                    orientacao=self._default_orientation,
                    ativo=True,
                    simetria=False,
                    ply_type=DEFAULT_PLY_TYPE,
                ),
            )
            self._inserted.append(target_pos)
        self._laminate.camadas = self._model.layers()

    def undo(self) -> None:
        if not self._inserted:
            return
        for pos in sorted(self._inserted, reverse=True):
            self._model.remove_rows([pos])
        self._laminate.camadas = self._model.layers()


class _RemoveLayerCommand(QtGui.QUndoCommand):
    """Undoable removal of layers for a laminate."""

    def __init__(
        self,
        model: StackingTableModel,
        laminate: Laminado,
        rows: list[int],
    ) -> None:
        super().__init__("Remove layer")
        self._model = model
        self._laminate = laminate
        self._rows = sorted(set(rows))
        self._backup: list[tuple[int, Camada]] = []

    def redo(self) -> None:
        self._backup = []
        layers = self._model.layers()
        for row in self._rows:
            if 0 <= row < len(layers):
                self._backup.append((row, copy.deepcopy(layers[row])))
        if not self._backup:
            return
        self._model.remove_rows([row for row, _ in self._backup])
        self._laminate.camadas = self._model.layers()

    def undo(self) -> None:
        if not self._backup:
            return
        for row, camada in sorted(self._backup, key=lambda item: item[0]):
            self._model.insert_layer(row, copy.deepcopy(camada))
        self._laminate.camadas = self._model.layers()


class _MoveColumnCommand(QtGui.QUndoCommand):
    """Undoable movement of laminate column (reorder cells)."""

    def __init__(
        self,
        cells: list[VirtualStackingCell],
        old_index: int,
        new_index: int,
    ) -> None:
        direction = "left" if new_index < old_index else "right"
        super().__init__(f"Move column {direction}")
        self._cells_ref = cells
        self._old_index = old_index
        self._new_index = new_index

    def redo(self) -> None:
        if 0 <= self._old_index < len(self._cells_ref) and 0 <= self._new_index < len(self._cells_ref):
            cell = self._cells_ref.pop(self._old_index)
            self._cells_ref.insert(self._new_index, cell)

    def undo(self) -> None:
        if 0 <= self._new_index < len(self._cells_ref) and 0 <= self._old_index < len(self._cells_ref):
            cell = self._cells_ref.pop(self._new_index)
            self._cells_ref.insert(self._old_index, cell)


class _ChangeOrientationCommand(QtGui.QUndoCommand):
    """Undoable change of orientation for a specific cell and row."""

    def __init__(
        self,
        laminate: Laminado,
        row: int,
        old_value: object,
        new_value: object,
    ) -> None:
        super().__init__("Change orientation")
        self._laminate = laminate
        self._row = row
        self._old_value = old_value
        self._new_value = new_value

    def redo(self) -> None:
        layers = getattr(self._laminate, "camadas", [])
        if 0 <= self._row < len(layers):
            layers[self._row].orientacao = self._new_value

    def undo(self) -> None:
        layers = getattr(self._laminate, "camadas", [])
        if 0 <= self._row < len(layers):
            layers[self._row].orientacao = self._old_value


class VirtualStackingHeaderView(WordWrapHeader):
    """Extended header view for laminate columns."""
    
    def __init__(self, orientation: QtCore.Qt.Orientation, parent=None):
        super().__init__(orientation, parent)
        self._laminate_column_offset = 5  # Will be set later
    
    def set_laminate_column_offset(self, offset: int) -> None:
        """Set the offset where laminate columns start."""
        self._laminate_column_offset = offset

    def update_height_from_text(self) -> None:
        """Adjust header height so wrapped text stays centered and visible."""
        model = self.model()
        if model is None:
            return
        fm = self.fontMetrics()
        max_height = 0
        for logical in range(model.columnCount()):
            text = model.headerData(logical, self.orientation(), QtCore.Qt.DisplayRole)
            if not text:
                continue
            width = self.sectionSize(logical) if hasattr(self, "sectionSize") else self.defaultSectionSize()
            if width <= 0:
                width = self.defaultSectionSize()
            text_rect = fm.boundingRect(
                QtCore.QRect(0, 0, int(width - 8), 2000),
                QtCore.Qt.TextWordWrap,
                str(text),
            )
            max_height = max(max_height, text_rect.height())

        padding = 16
        desired = max(80, max_height + padding)
        if desired != self.height():
            self.setMinimumHeight(desired)
            self.setMaximumHeight(desired)
            self.updateGeometry()


class VirtualStackingModel(QtCore.QAbstractTableModel):
    """Qt model responsible for exposing Virtual Stacking data."""

    COL_SEQUENCE = 0
    COL_PLY = 1
    COL_PLY_TYPE = 2
    COL_MATERIAL = 3
    COL_ROSETTE = 4
    LAMINATE_COLUMN_OFFSET = 5

    def __init__(
        self,
        parent: Optional[QtCore.QObject] = None,
        change_callback: Optional[Callable[[list[str]], None]] = None,
        stacking_model_provider: Optional[
            Callable[[Laminado], Optional[StackingTableModel]]
        ] = None,
        post_edit_callback: Optional[Callable[[Laminado], None]] = None,
        most_used_material_provider: Optional[Callable[[], Optional[str]]] = None,
        material_sync_handler: Optional[Callable[[int, str], list[Laminado]]] = None,
    ) -> None:
        super().__init__(parent)
        self.layers: list[VirtualStackingLayer] = []
        self.cells: list[VirtualStackingCell] = []
        self.symmetry_row_index: int | None = None
        self.symmetry_rows: set[int] = set()
        self._change_callback = change_callback
        self._stacking_model_provider = stacking_model_provider or (lambda _lam: None)
        self._post_edit_callback = post_edit_callback
        self._most_used_material_provider = most_used_material_provider or (lambda: None)
        self._material_sync_handler = material_sync_handler
        self._red_cells: set[tuple[int, int]] = set()
        self._green_cells: set[tuple[int, int]] = set()
        self._symmetric_cells: set[tuple[int, int]] = set()
        self._unbalanced_columns: set[int] = set()
        self._warning_icon = QtGui.QIcon()
        self._selected_columns: set[int] = set()
        self._column_summaries: list[str] = []

    # Basic structure -------------------------------------------------
    def rowCount(self, parent: QtCore.QModelIndex = QtCore.QModelIndex()) -> int:  # noqa: N802
        if parent.isValid():
            return 0
        return len(self.layers)

    def columnCount(self, parent: QtCore.QModelIndex = QtCore.QModelIndex()) -> int:  # noqa: N802
        if parent.isValid():
            return 0
        return self.LAMINATE_COLUMN_OFFSET + len(self.cells)

    # Header ----------------------------------------------------------
    def headerData(  # noqa: N802
        self,
        section: int,
        orientation: QtCore.Qt.Orientation,
        role: int = QtCore.Qt.DisplayRole,
    ):
        if role == QtCore.Qt.DecorationRole and orientation == QtCore.Qt.Horizontal:
            if section in self._unbalanced_columns:
                return self._warning_icon or self._default_warning_icon()
            return None
        if role == QtCore.Qt.BackgroundRole and orientation == QtCore.Qt.Horizontal:
            if section in self._selected_columns:
                return QtGui.QBrush(QtGui.QColor(200, 230, 255))
            return None

        if role != QtCore.Qt.DisplayRole:
            return None

        if orientation == QtCore.Qt.Horizontal:
            if section == self.COL_SEQUENCE:
                return "Sequence"
            if section == self.COL_PLY:
                return "Ply"
            if section == self.COL_PLY_TYPE:
                return "Symmetry"
            if section == self.COL_MATERIAL:
                return "Material"
            if section == self.COL_ROSETTE:
                return "Rosette"
            cell_index = section - self.LAMINATE_COLUMN_OFFSET
            if 0 <= cell_index < len(self.cells):
                cell = self.cells[cell_index]
                laminate = cell.laminate
                lam_name = (getattr(laminate, "nome", "") or "").strip() or cell.cell_id
                # Display format: C2 | L7 on first line, Total: X on second line
                label = f"{cell.cell_id} | {lam_name}"
                if 0 <= cell_index < len(self._column_summaries):
                    summary = self._column_summaries[cell_index].strip()
                    if summary:
                        # Summary contains "Total: X" - append it on a new line
                        label = f"{label}\n{summary}"
                return label
        elif orientation == QtCore.Qt.Vertical:
            return str(section + 1)
        return None

    def _default_warning_icon(self) -> QtGui.QIcon:
        if self._warning_icon and not self._warning_icon.isNull():
            return self._warning_icon
        icon = QtWidgets.QApplication.style().standardIcon(QtWidgets.QStyle.SP_MessageBoxWarning)
        self._warning_icon = icon
        return icon

    # Data roles ------------------------------------------------------

    def data(  # noqa: N802
        self,
        index: QtCore.QModelIndex,
        role: int = QtCore.Qt.DisplayRole,
    ):
        if not index.isValid():
            return None

        row = index.row()
        column = index.column()

        if row < 0 or row >= len(self.layers):
            return None

        layer = self.layers[row]

        if role in (QtCore.Qt.DisplayRole, QtCore.Qt.EditRole):
            if column == self.COL_SEQUENCE:
                return layer.sequence_label or f"Seq.{row + 1}"
            if column == self.COL_PLY:
                return layer.ply_label or f"Ply.{row + 1}"
            if column == self.COL_PLY_TYPE:
                return layer.ply_type or DEFAULT_PLY_TYPE
            if column == self.COL_MATERIAL:
                return layer.material
            if column == self.COL_ROSETTE:
                return layer.rosette or DEFAULT_ROSETTE_LABEL

            cell_index = column - self.LAMINATE_COLUMN_OFFSET
            if 0 <= cell_index < len(self.cells):
                cell = self.cells[cell_index]
                layers = getattr(cell.laminate, "camadas", [])
                if 0 <= row < len(layers):
                    value = getattr(layers[row], "orientacao", None)
                    if value is None:
                        return "Empty" if role == QtCore.Qt.DisplayRole else ""
                    try:
                        numeric_value = float(value)
                    except Exception:
                        numeric_value = None
                    if role == QtCore.Qt.DisplayRole:
                        return format_orientation_value(value)
                    if numeric_value is not None:
                        return f"{numeric_value:g}"
                    return str(value)
                return ""

        if role == QtCore.Qt.BackgroundRole:
            if self.symmetry_rows and row in self.symmetry_rows:
                return QtGui.QBrush(QtGui.QColor(250, 128, 114))
            if column < self.LAMINATE_COLUMN_OFFSET:
                return QtGui.QBrush(QtGui.QColor(240, 240, 240))
            if (row, column) in self._red_cells:
                return QtGui.QBrush(QtGui.QColor(220, 53, 69))
            if (row, column) in self._green_cells:
                return QtGui.QBrush(QtGui.QColor(40, 167, 69))
            cell_index = column - self.LAMINATE_COLUMN_OFFSET
            if 0 <= cell_index < len(self.cells):
                laminate = self.cells[cell_index].laminate
                layers = getattr(laminate, "camadas", [])
                if 0 <= row < len(layers):
                    color = orientation_highlight_color(
                        getattr(layers[row], "orientacao", None)
                    )
                    if color is not None:
                        return QtGui.QBrush(color)

        if role == QtCore.Qt.ForegroundRole and column >= self.LAMINATE_COLUMN_OFFSET:
            cell_index = column - self.LAMINATE_COLUMN_OFFSET
            if 0 <= cell_index < len(self.cells):
                layers = getattr(self.cells[cell_index].laminate, "camadas", [])
                if 0 <= row < len(layers):
                    orient = getattr(layers[row], "orientacao", None)
                    if orient is None:
                        return QtGui.QBrush(QtGui.QColor(160, 160, 160))
                    try:
                        angle = normalize_angle(orient)
                    except Exception:
                        angle = None
                    if angle is not None and abs(float(angle) - 90.0) <= 1e-9:
                        return QtGui.QBrush(QtGui.QColor(255, 255, 255))

        if role == QtCore.Qt.TextAlignmentRole and column >= self.LAMINATE_COLUMN_OFFSET:
            return QtCore.Qt.AlignCenter

        if role == ORIENTATION_SYMMETRY_ROLE and column >= self.LAMINATE_COLUMN_OFFSET:
            if (row, column) in self._symmetric_cells:
                return True

        return None

    # Editing ---------------------------------------------------------
    def flags(self, index: QtCore.QModelIndex) -> QtCore.Qt.ItemFlags:  # noqa: N802
        if not index.isValid():
            return QtCore.Qt.NoItemFlags

        base = QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled
        if index.column() in (self.COL_PLY_TYPE, self.COL_MATERIAL, self.COL_ROSETTE) or index.column() >= self.LAMINATE_COLUMN_OFFSET:
            return base | QtCore.Qt.ItemIsEditable
        return base

    def setData(  # noqa: N802
        self,
        index: QtCore.QModelIndex,
        value,
        role: int = QtCore.Qt.EditRole,
    ) -> bool:
        if role != QtCore.Qt.EditRole or not index.isValid():
            return False

        row = index.row()
        column = index.column()
        if row < 0 or row >= len(self.layers):
            return False

        if column == self.COL_PLY_TYPE:
            return self._set_ply_type_for_row(row, value)
        if column == self.COL_MATERIAL:
            return self._set_material_for_row(row, str(value))
        if column == self.COL_ROSETTE:
            return self._set_rosette_for_row(row, str(value))
        if column < self.LAMINATE_COLUMN_OFFSET:
            return False

        cell_index = column - self.LAMINATE_COLUMN_OFFSET
        if cell_index < 0 or cell_index >= len(self.cells):
            return False

        cell = self.cells[cell_index]
        
        # Ensure unique laminate before modifying to prevent affecting other cells
        if hasattr(self.parent(), '_ensure_unique_laminate_for_cell'):
            laminate = self.parent()._ensure_unique_laminate_for_cell(cell)
        else:
            laminate = cell.laminate
            
        stacking_model = self._stacking_model_provider(laminate)
        if stacking_model is None:
            return False
        self._ensure_row_exists(stacking_model, row, laminate)

        text = "" if value is None else str(value).strip()
        if text.endswith("?"):
            text = text[:-1].strip()
        if text.lower() in {"", "empty"}:
            normalized_value: object = ""
        else:
            try:
                normalized_value = normalize_angle(text)
            except Exception:
                return False

        target_index = stacking_model.index(row, StackingTableModel.COL_ORIENTATION)
        if not target_index.isValid():
            return False
        # Permit blank values; delegate validation to StackingTableModel.
        success = stacking_model.setData(target_index, normalized_value, QtCore.Qt.EditRole)
        if not success:
            success = stacking_model.apply_field_value(
                row, StackingTableModel.COL_ORIENTATION, normalized_value
            )
        if not success:
            return False
        laminate.camadas = stacking_model.layers()
        self._auto_fill_material_if_missing(laminate, stacking_model, row)
        self.dataChanged.emit(index, index, [QtCore.Qt.DisplayRole, QtCore.Qt.EditRole])
        self._emit_change_callbacks([laminate])
        return True


    def _ensure_row_exists(
        self,
        stacking_model: StackingTableModel,
        row: int,
        laminate: Optional[Laminado] = None,
    ) -> list[Camada]:
        default_material = str(self._most_used_material_provider() or "").strip()
        layers = stacking_model.layers()
        if row >= len(layers):
            missing = row - len(layers) + 1
            for _ in range(missing):
                stacking_model.insert_layer(
                    len(layers),
                    Camada(
                        idx=0,
                        material=default_material,
                        orientacao=None,
                        ativo=True,
                        simetria=False,
                        ply_type=DEFAULT_PLY_TYPE,
                    ),
                )
                layers = stacking_model.layers()
        if laminate is not None:
            laminate.camadas = stacking_model.layers()
            layers = laminate.camadas
        return layers

    def _emit_change_callbacks(self, laminates: Iterable[Laminado]) -> None:
        names: list[str] = []
        seen: set[str] = set()
        for lam in laminates:
            if lam is None:
                continue
            if lam.nome not in seen:
                names.append(lam.nome)
                seen.add(lam.nome)
            if self._post_edit_callback:
                try:
                    self._post_edit_callback(lam)
                except Exception:
                    pass
        if names and self._change_callback:
            try:
                self._change_callback(names)
            except Exception:
                pass

    def _auto_fill_material_if_missing(
        self, laminate: Laminado, stacking_model: StackingTableModel, row: int
    ) -> None:
        layers = stacking_model.layers()
        if not (0 <= row < len(layers)):
            return
        layer = layers[row]
        if layer.orientacao is None or getattr(layer, "material", ""):
            return
        suggestion = self._most_used_material_provider() or ""
        suggestion = str(suggestion or "").strip()
        if not suggestion:
            return
        target_index = stacking_model.index(row, StackingTableModel.COL_MATERIAL)
        if target_index.isValid():
            stacking_model.setData(target_index, suggestion, QtCore.Qt.EditRole)

    def _set_ply_type_for_row(self, row: int, value: object) -> bool:
        new_value = normalize_ply_type_label(value)
        changed: list[Laminado] = []
        for cell in self.cells:
            laminate = cell.laminate
            stacking_model = self._stacking_model_provider(laminate)
            if stacking_model is None:
                continue
            layers = self._ensure_row_exists(stacking_model, row, laminate)
            if row >= len(layers):
                continue
            target_layer = layers[row]
            if normalize_ply_type_label(getattr(target_layer, "ply_type", DEFAULT_PLY_TYPE)) == new_value:
                continue
            idx = stacking_model.index(row, StackingTableModel.COL_PLY_TYPE)
            if idx.isValid() and stacking_model.setData(idx, new_value, QtCore.Qt.EditRole):
                changed.append(laminate)
            else:
                target_layer.ply_type = new_value
                laminate.camadas = stacking_model.layers()
                changed.append(laminate)
        if changed:
            self.layers[row].ply_type = new_value
            model_index = self.index(row, self.COL_PLY_TYPE)
            self.dataChanged.emit(model_index, model_index, [QtCore.Qt.DisplayRole, QtCore.Qt.EditRole])
            self._emit_change_callbacks(changed)
            return True
        return False

    def _set_material_for_row(self, row: int, value: str) -> bool:
        new_material = str(value or "").strip()
        if self._material_sync_handler is not None:
            changed = list(self._material_sync_handler(row, new_material) or [])
        else:
            changed = []
            for cell in self.cells:
                laminate = cell.laminate
                stacking_model = self._stacking_model_provider(laminate)
                if stacking_model is None:
                    continue
                layers = self._ensure_row_exists(stacking_model, row, laminate)
                if row >= len(layers):
                    continue
                idx = stacking_model.index(row, StackingTableModel.COL_MATERIAL)
                if idx.isValid() and stacking_model.setData(idx, new_material, QtCore.Qt.EditRole):
                    changed.append(laminate)
        if changed:
            self.layers[row].material = new_material
            model_index = self.index(row, self.COL_MATERIAL)
            self.dataChanged.emit(model_index, model_index, [QtCore.Qt.DisplayRole, QtCore.Qt.EditRole])
            self._emit_change_callbacks(changed)
            return True
        return False

    def _set_rosette_for_row(self, row: int, value: str) -> bool:
        rosette_value = str(value or "").strip() or DEFAULT_ROSETTE_LABEL
        changed: list[Laminado] = []
        for cell in self.cells:
            laminate = cell.laminate
            stacking_model = self._stacking_model_provider(laminate)
            if stacking_model is None:
                continue
            layers = self._ensure_row_exists(stacking_model, row, laminate)
            if row >= len(layers):
                continue
            target_layer = layers[row]
            current_value = str(getattr(target_layer, "rosette", "") or "")
            if current_value == rosette_value:
                continue
            target_layer.rosette = rosette_value
            laminate.camadas = stacking_model.layers()
            changed.append(laminate)
        if changed:
            self.layers[row].rosette = rosette_value
            model_index = self.index(row, self.COL_ROSETTE)
            self.dataChanged.emit(model_index, model_index, [QtCore.Qt.DisplayRole, QtCore.Qt.EditRole])
            self._emit_change_callbacks(changed)
            return True
        return False


    # Public API ------------------------------------------------------
    def set_virtual_stacking(
        self,
        layers: Iterable[VirtualStackingLayer],
        cells: Iterable[VirtualStackingCell],
        symmetry_row_index: int | set[int] | list[int] | tuple[int, ...] | None = None,
    ) -> None:
        self.beginResetModel()
        self.layers = list(layers)
        self.cells = list(cells)
        if symmetry_row_index is None:
            self.symmetry_rows = set()
            self.symmetry_row_index = None
        elif isinstance(symmetry_row_index, int):
            self.symmetry_rows = {symmetry_row_index}
            self.symmetry_row_index = symmetry_row_index
        else:
            self.symmetry_rows = {idx for idx in symmetry_row_index}
            self.symmetry_row_index = min(self.symmetry_rows) if self.symmetry_rows else None
        self._red_cells.clear()
        self._green_cells.clear()
        self._symmetric_cells.clear()
        self._unbalanced_columns.clear()
        self.endResetModel()

    def set_symmetry_row_index(self, index: int | None) -> None:
        if index == self.symmetry_row_index:
            return
        if index is None:
            self.symmetry_rows = set()
        else:
            self.symmetry_rows = {index}
        self.symmetry_row_index = index
        if self.layers:
            top_left = self.createIndex(0, 0)
            bottom_right = self.createIndex(len(self.layers) - 1, self.columnCount() - 1)
            self.dataChanged.emit(top_left, bottom_right, [QtCore.Qt.BackgroundRole])

    def set_symmetry_rows(self, rows: Iterable[int] | None) -> None:
        new_rows = {row for row in rows} if rows is not None else set()
        if new_rows == self.symmetry_rows:
            return
        self.symmetry_rows = new_rows
        self.symmetry_row_index = min(new_rows) if new_rows else None
        if self.layers:
            top_left = self.createIndex(0, 0)
            bottom_right = self.createIndex(len(self.layers) - 1, self.columnCount() - 1)
            self.dataChanged.emit(top_left, bottom_right, [QtCore.Qt.BackgroundRole])

    def set_unbalanced_columns(self, columns: Iterable[int]) -> None:
        new_columns = {col for col in columns}
        if new_columns == self._unbalanced_columns:
            return
        self._unbalanced_columns = new_columns
        if self.columnCount() > self.LAMINATE_COLUMN_OFFSET:
            first = self.LAMINATE_COLUMN_OFFSET
            last = self.columnCount() - 1
            self.headerDataChanged.emit(QtCore.Qt.Horizontal, first, last)

    def set_selected_columns(self, columns: Iterable[int]) -> None:
        new_columns = {col for col in columns if col >= self.LAMINATE_COLUMN_OFFSET}
        if new_columns == self._selected_columns:
            return
        self._selected_columns = new_columns
        if self.columnCount() > self.LAMINATE_COLUMN_OFFSET:
            first = min(new_columns) if new_columns else self.LAMINATE_COLUMN_OFFSET
            last = max(new_columns) if new_columns else self.columnCount() - 1
            self.headerDataChanged.emit(QtCore.Qt.Horizontal, first, last)

    def clear_highlights(self) -> None:
        if not (self._red_cells or self._green_cells or self._symmetric_cells):
            return
        self._red_cells.clear()
        self._green_cells.clear()
        self._symmetric_cells.clear()
        if self.layers:
            top_left = self.createIndex(0, 0)
            bottom_right = self.createIndex(len(self.layers) - 1, self.columnCount() - 1)
            self.dataChanged.emit(top_left, bottom_right, [QtCore.Qt.BackgroundRole, ORIENTATION_SYMMETRY_ROLE])

    def set_highlights(
        self,
        red_cells: Iterable[tuple[int, int]],
        green_cells: Iterable[tuple[int, int]],
    ) -> None:
        new_red = {(r, c) for r, c in red_cells}
        new_green = {(r, c) for r, c in green_cells}
        if new_red == self._red_cells and new_green == self._green_cells:
            return
        self._red_cells = new_red
        self._green_cells = new_green
        if self.layers:
            top_left = self.createIndex(0, 0)
            bottom_right = self.createIndex(len(self.layers) - 1, self.columnCount() - 1)
            self.dataChanged.emit(top_left, bottom_right, [QtCore.Qt.BackgroundRole])

    def set_column_summaries(self, summaries: list[str]) -> None:
        """Store summaries per laminate column for display in headers."""
        self._column_summaries = summaries
        if self.columnCount() > self.LAMINATE_COLUMN_OFFSET:
            first = self.LAMINATE_COLUMN_OFFSET
            last = self.columnCount() - 1
            self.headerDataChanged.emit(QtCore.Qt.Horizontal, first, last)



class VirtualStackingWindow(QtWidgets.QDialog):
    """Dialog that renders the Virtual Stacking spreadsheet-like view."""

    def _on_sequence_column_clicked(self, index: QtCore.QModelIndex) -> None:
        if not index.isValid() or index.column() != self.model.COL_SEQUENCE:
            self.table.clearSelection()
            return
        self.table.selectRow(index.row())

    def _setup_table_signals(self) -> None:
        self.table.clicked.connect(self._on_sequence_column_clicked)

    def showEvent(self, event):
        super().showEvent(event)
        self._setup_table_signals()

    stacking_changed = QtCore.Signal(list)
    closed = QtCore.Signal()

    def __init__(
        self,
        parent: Optional[QtWidgets.QWidget] = None,
        *,
        undo_stack: Optional[QtGui.QUndoStack] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Virtual Stacking - View Mode: Cells")
        self.setWindowFlags(
            self.windowFlags()
            | QtCore.Qt.WindowMinMaxButtonsHint
            | QtCore.Qt.WindowSystemMenuHint
        )
        self.setWindowFlag(QtCore.Qt.Window, True)
        self._apply_window_icon()
        self.resize(1200, 700)

        self._settings = QtCore.QSettings("GridLamEdit", "GridLamEdit")
        self._layers: list[VirtualStackingLayer] = []
        self._cells: list[VirtualStackingCell] = []
        self._symmetry_row_index: int | None = None
        self._symmetry_rows: set[int] = set()
        self._symmetry_evaluations: dict[int, LaminateSymmetryEvaluation] = {}
        self._project: Optional[GridModel] = None
        self._sorted_cell_ids: list[str] = []
        self._initial_sort_done = False
        self._stacking_models: dict[int, StackingTableModel] = {}
        self._selected_cell_ids: set[str] = set()
        self._restoring_column_order = False
        self._column_order_key = "virtual_stacking/column_order"
        self._warning_banner_icon = QtGui.QIcon()
        self.undo_stack = undo_stack or QtGui.QUndoStack(self)
        self._neighbors_reorder_snapshot: Optional[dict] = None

        self._build_ui()

    def _apply_window_icon(self) -> None:
        icon = QtGui.QIcon(":/icons/stacking_summary.svg")
        if icon.isNull():
            candidate = package_path("resources", "icons", "stacking_summary.svg")
            if candidate.is_file():
                icon = QtGui.QIcon(str(candidate))
        if icon.isNull():
            candidate_png = package_path("resources", "icons", "stacking_summary.png")
            if candidate_png.is_file():
                icon = QtGui.QIcon(str(candidate_png))
        if icon.isNull():
            app_icon = QtWidgets.QApplication.windowIcon()
            if app_icon is not None and not app_icon.isNull():
                icon = app_icon
        if icon.isNull():
            icon = self.style().standardIcon(QtWidgets.QStyle.SP_FileDialogInfoView)
        if not icon.isNull():
            self.setWindowIcon(icon)
            try:
                QtWidgets.QApplication.setWindowIcon(icon)
            except Exception:
                pass


    def _build_ui(self) -> None:
        toolbar_layout = self._build_toolbar()
        prefix_layout = self._build_prefix_toolbar()

        self.model = VirtualStackingModel(
            self,
            change_callback=self._on_model_change,
            stacking_model_provider=self._stacking_model_for,
            post_edit_callback=self._after_laminate_changed,
            most_used_material_provider=self._most_used_material,
            material_sync_handler=self._sync_material_for_row,
        )
        self.table = QtWidgets.QTableView(self)
        self.table.setModel(self.model)
        header = VirtualStackingHeaderView(QtCore.Qt.Horizontal, self.table)
        header.setDefaultAlignment(QtCore.Qt.AlignCenter)
        header.set_laminate_column_offset(self.model.LAMINATE_COLUMN_OFFSET)
        self.table.setHorizontalHeader(header)
        header.setSectionsClickable(True)
        header.setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        header.setSectionsMovable(True)
        header.setStretchLastSection(False)
        # Keep header tall enough for wrapped text and adjust dynamically on resize
        header.setMinimumHeight(max(header.sizeHint().height(), 80))
        header.update_height_from_text()
        header.sectionResized.connect(lambda *_: header.update_height_from_text())
        header.sectionResized.connect(lambda *_: self._resize_summary_columns())
        header.sectionMoved.connect(self._on_header_section_moved)
        header.sectionClicked.connect(self._on_header_section_clicked)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        # Seleção por linha para destacar bordas sem alterar cores
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.setStyleSheet(
            "QTableView::item:selected { background: transparent; border: 2px solid #0078D7; }"
        )
        self.table.setHorizontalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        self.table.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        self.table.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked
            | QtWidgets.QAbstractItemView.SelectedClicked
        )
        self.table.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self._apply_orientation_delegate()
        # Removemos a tabela de resumo flutuante que criava um bloco branco no topo.
        self.summary_table = None
        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        if toolbar_layout is not None:
            main_layout.addLayout(toolbar_layout)
        if prefix_layout is not None:
            main_layout.addLayout(prefix_layout)
        main_layout.setContentsMargins(4, 4, 4, 4)
        main_layout.addWidget(self.table)
        self.setLayout(main_layout)
        self.undo_stack.canUndoChanged.connect(self._update_undo_buttons)
        self.undo_stack.canRedoChanged.connect(self._update_undo_buttons)
        self.undo_stack.indexChanged.connect(lambda _idx: self._on_undo_stack_changed())

    def _build_prefix_toolbar(self) -> Optional[QtWidgets.QHBoxLayout]:
        layout = QtWidgets.QHBoxLayout()
        layout.setContentsMargins(4, 0, 4, 4)
        layout.setSpacing(8)

        seq_button = QtWidgets.QToolButton(self)
        seq_button.setText("Rename Sequence")
        seq_button.setToolTip("Rename the prefix for every sequence")
        seq_button.clicked.connect(self._rename_all_sequences)
        layout.addWidget(seq_button)

        ply_button = QtWidgets.QToolButton(self)
        ply_button.setText("Rename Ply")
        ply_button.setToolTip("Rename the prefix for every ply")
        ply_button.clicked.connect(self._rename_all_ply)
        layout.addWidget(ply_button)

        layout.addStretch()
        return layout

    def _apply_orientation_delegate(self) -> None:
        material_delegate = MaterialComboDelegate(
            self.table,
            items_provider=lambda: registry_available_materials(
                self._project, settings=self._settings
            ),
        )
        self.table.setItemDelegateForColumn(
            VirtualStackingModel.COL_MATERIAL, material_delegate
        )
        ply_type_delegate = PlyTypeComboDelegate(self.table)
        self.table.setItemDelegateForColumn(
            VirtualStackingModel.COL_PLY_TYPE, ply_type_delegate
        )
        delegate = OrientationComboDelegate(
            self.table,
            items_provider=lambda: project_distinct_orientations(self._project),
        )
        for col in range(self.model.LAMINATE_COLUMN_OFFSET, self.model.columnCount()):
            self.table.setItemDelegateForColumn(col, delegate)

    def _build_summary_table(self) -> QtWidgets.QTableView:
        summary = QtWidgets.QTableView(self)
        summary.setModel(QtGui.QStandardItemModel(1, 0, summary))
        summary.verticalHeader().setVisible(False)
        summary.horizontalHeader().setVisible(False)
        summary.horizontalHeader().setSectionsMovable(True)
        summary.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Fixed)
        summary.setWordWrap(True)
        summary.setTextElideMode(QtCore.Qt.ElideNone)
        summary.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        summary.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        summary.setFocusPolicy(QtCore.Qt.NoFocus)
        summary.setHorizontalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        summary.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        summary.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        summary.setStyleSheet("QTableView { gridline-color: transparent; }")
        summary.setFrameShape(QtWidgets.QFrame.NoFrame)
        summary.setShowGrid(False)
        font = summary.font()
        if font.pointSize() > 0:
            font.setPointSize(max(font.pointSize() - 1, 8))
            summary.setFont(font)
        return summary

    def _sync_scrollbars(self) -> None:
        # Rolagem horizontal unificada via QScrollArea; nada a sincronizar manualmente.
        pass

    def _current_header_order(self) -> list[int]:
        header = self.table.horizontalHeader()
        if header is None:
            return []
        return [header.logicalIndex(i) for i in range(header.count())]

    def _apply_header_order_to_summary(self) -> None:
        header = self.table.horizontalHeader()
        summary_header = (
            self.summary_table.horizontalHeader()
            if getattr(self, "summary_table", None)
            else None
        )
        if (
            header is None
            or summary_header is None
            or summary_header.count() != header.count()
        ):
            return
        order = self._current_header_order()
        for visual_pos, logical in enumerate(order):
            current_visual = summary_header.visualIndex(logical)
            if current_visual != visual_pos:
                summary_header.blockSignals(True)
                try:
                    summary_header.moveSection(current_visual, visual_pos)
                finally:
                    summary_header.blockSignals(False)

    def _mirror_summary_section_move(self, old_visual: int, new_visual: int) -> None:
        summary_header = None
        summary = getattr(self, "summary_table", None)
        if summary is not None:
            summary_header = summary.horizontalHeader()
        if summary_header is None:
            return
        summary_header.blockSignals(True)
        try:
            summary_header.moveSection(old_visual, new_visual)
        finally:
            summary_header.blockSignals(False)

    def _persist_column_order(self) -> None:
        order = self._current_header_order()
        try:
            self._settings.setValue(self._column_order_key, order)
        except Exception:
            pass

    def persist_column_order(self) -> None:
        """Public wrapper to persist the current visual order of the columns."""
        self._persist_column_order()

    def _restore_column_order(self) -> None:
        header = self.table.horizontalHeader()
        if header is None or header.count() == 0:
            return
        raw_value = None
        try:
            raw_value = self._settings.value(self._column_order_key, [])
        except Exception:
            raw_value = []
        if raw_value in (None, ""):
            return
        try:
            saved_order = [int(val) for val in raw_value]
        except Exception:
            return
        current_count = header.count()
        valid = [idx for idx in saved_order if 0 <= idx < current_count]
        remaining = [idx for idx in range(current_count) if idx not in valid]
        target_order = valid + remaining
        self._restoring_column_order = True
        try:
            for visual_pos, logical in enumerate(target_order):
                current_visual = header.visualIndex(logical)
                if current_visual != visual_pos:
                    header.moveSection(current_visual, visual_pos)
            self._apply_header_order_to_summary()
        finally:
            self._restoring_column_order = False

    def _on_header_section_moved(self, logical_index: int, old_visual: int, new_visual: int) -> None:  # noqa: ARG002
        self._mirror_summary_section_move(old_visual, new_visual)
        if not self._restoring_column_order:
            self._persist_column_order()

    def _on_header_section_clicked(self, logical_index: int) -> None:
        if logical_index < self.model.LAMINATE_COLUMN_OFFSET:
            self._selected_cell_ids.clear()
            self._apply_column_selection()
            return
        cell_idx = logical_index - self.model.LAMINATE_COLUMN_OFFSET
        if not (0 <= cell_idx < len(self._cells)):
            return
        cell_id = self._cells[cell_idx].cell_id
        modifiers = QtWidgets.QApplication.keyboardModifiers()
        multi_select = bool(modifiers & (QtCore.Qt.ControlModifier | QtCore.Qt.ShiftModifier))
        if multi_select:
            if cell_id in self._selected_cell_ids:
                self._selected_cell_ids.remove(cell_id)
            else:
                self._selected_cell_ids.add(cell_id)
        else:
            self._selected_cell_ids = {cell_id}
        self._apply_column_selection()

    def _selected_column_indexes(self) -> list[int]:
        columns: list[int] = []
        for idx, cell in enumerate(self._cells):
            if cell.cell_id in self._selected_cell_ids:
                columns.append(idx + self.model.LAMINATE_COLUMN_OFFSET)
        return columns

    def _apply_column_selection(self) -> None:
        if not hasattr(self, "table") or self.table.model() is None:
            return
        valid_ids = {cell.cell_id for cell in self._cells}
        self._selected_cell_ids = {cid for cid in self._selected_cell_ids if cid in valid_ids}
        selected_columns = self._selected_column_indexes()
        self.model.set_selected_columns(selected_columns)
        selection_model = self.table.selectionModel()
        if selection_model is None:
            return
        selection_model.blockSignals(True)
        selection_model.clearSelection()
        if selected_columns and self.model.rowCount() > 0:
            flags = (
                QtCore.QItemSelectionModel.Select
                | QtCore.QItemSelectionModel.Columns
            )
            for col in selected_columns:
                top_index = self.model.index(0, col)
                bottom_index = self.model.index(self.model.rowCount() - 1, col)
                selection = QtCore.QItemSelection(top_index, bottom_index)
                selection_model.select(selection, flags)
        selection_model.blockSignals(False)

    def _update_undo_buttons(self) -> None:
        if hasattr(self, "btn_undo"):
            self.btn_undo.setEnabled(self.undo_stack.canUndo())
        if hasattr(self, "btn_redo"):
            self.btn_redo.setEnabled(self.undo_stack.canRedo())

    def _execute_command(self, command: QtGui.QUndoCommand) -> None:
        if self.undo_stack is not None:
            self.undo_stack.push(command)
        else:
            command.redo()
        self._update_undo_buttons()

    def _move_selected_column(self, direction: int) -> None:
        if not self._cells:
            return
        selected_columns = self._selected_column_indexes()
        if not selected_columns:
            QtWidgets.QMessageBox.information(
                self,
                "Move laminate",
                "Select exactly one column to move.",
            )
            return
        if len(selected_columns) > 1:
            QtWidgets.QMessageBox.information(
                self,
                "Move laminate",
                "Select only one column to move.",
            )
            return
        column = selected_columns[0]
        cell_idx = column - self.model.LAMINATE_COLUMN_OFFSET
        if direction < 0 and cell_idx <= 0:
            QtWidgets.QMessageBox.information(
                self,
                "Move laminate",
                "This column is already at the left edge.",
            )
            return
        if direction > 0 and cell_idx >= len(self._cells) - 1:
            QtWidgets.QMessageBox.information(
                self,
                "Move laminate",
                "This column is already at the right edge.",
            )
            return
        new_pos = cell_idx + direction

        # Record state for undo/redo before applying the move
        self._push_virtual_snapshot()

        current_ids = [cell.cell_id for cell in self._cells]
        current_ids[cell_idx], current_ids[new_pos] = current_ids[new_pos], current_ids[cell_idx]
        self._sorted_cell_ids = current_ids
        self._initial_sort_done = True
        if self._project is not None:
            remaining = [cid for cid in self._project.celulas_ordenadas if cid not in current_ids]
            self._project.celulas_ordenadas = current_ids + remaining
        self._rebuild_view()
        self._mark_project_dirty()
        self._apply_column_selection()

    def _on_undo_stack_changed(self) -> None:
        # Sincronizar todos os stacking models com os dados dos laminados
        for lam_id, model in list(self._stacking_models.items()):
            for cell in self._cells:
                if id(cell.laminate) == lam_id:
                    model.update_layers(list(getattr(cell.laminate, "camadas", [])))
                    break
        
        for cell in getattr(self, "_cells", []):
            self._auto_rename_if_enabled(cell.laminate)
        self._rebuild_view()
        self._update_undo_buttons()
        self._mark_project_dirty()


    def _warning_icon_for_banner(self) -> QtGui.QIcon:
        if self._warning_banner_icon and not self._warning_banner_icon.isNull():
            return self._warning_banner_icon
        icon = QtGui.QIcon(":/icons/warning.png")
        if icon.isNull():
            candidate = package_path("resources", "icons", "warning.png")
            if candidate.is_file():
                icon = QtGui.QIcon(str(candidate))
        if icon.isNull():
            icon = self.style().standardIcon(QtWidgets.QStyle.SP_MessageBoxWarning)
        self._warning_banner_icon = icon
        return icon

    def _build_toolbar(self) -> QtWidgets.QHBoxLayout:
        toolbar = QtWidgets.QHBoxLayout()
        toolbar.setSpacing(10)
        toolbar.setContentsMargins(6, 4, 6, 4)

        # Indicador de salvamento.
        self.lbl_auto_saving = QtWidgets.QLabel("Automatic Saving Active", self)
        self.lbl_auto_saving.setSizePolicy(
            QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed
        )
        self.lbl_auto_saving.setStyleSheet("background: transparent; margin: 0; padding: 0;")
        toolbar.addWidget(self.lbl_auto_saving)

        # Aviso de desbalanceamento inline.
        warning_container = QtWidgets.QHBoxLayout()
        warning_container.setSpacing(4)
        warning_container.setContentsMargins(0, 0, 0, 0)
        warning_icon_label = QtWidgets.QLabel(self)
        warning_icon = self._warning_icon_for_banner()
        if warning_icon is not None and not warning_icon.isNull():
            warning_icon_label.setPixmap(warning_icon.pixmap(18, 18))
        else:
            warning_icon_label.setVisible(False)
        warning_container.addWidget(warning_icon_label)
        warning_container.addWidget(QtWidgets.QLabel("Unbalanced Laminate", self))
        warning_widget = QtWidgets.QWidget(self)
        warning_widget.setLayout(warning_container)
        warning_widget.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        warning_widget.setStyleSheet("background: transparent; margin: 0; padding: 0;")
        toolbar.addWidget(warning_widget)

        self.btn_export_virtual = QtWidgets.QToolButton(self)
        self.btn_export_virtual.setText("Export Virtual Stacking")
        self.btn_export_virtual.setToolTip(
            "Exporta a planilha em .xlsx (template Grid Lam Vs Exported_RevC) para importação no CATIA."
        )
        self.btn_export_virtual.setEnabled(True)
        self.btn_export_virtual.clicked.connect(self._export_virtual_stacking)
        toolbar.addWidget(self.btn_export_virtual)

        toolbar.addStretch()

        # Reorganizar por Vizinhança
        self.btn_reorganize_neighbors = QtWidgets.QToolButton(self)
        self.btn_reorganize_neighbors.setText("Reorder by Neighborhood")
        self.btn_reorganize_neighbors.setToolTip(
            "Reorders sequences based on neighborhood and symmetry rules"
        )
        self.btn_reorganize_neighbors.setCheckable(True)
        self.btn_reorganize_neighbors.toggled.connect(self._toggle_reorganizar_por_vizinhanca)
        toolbar.addWidget(self.btn_reorganize_neighbors)

        # Novo botão: analisar simetria
        self.btn_analyze_symmetry = QtWidgets.QToolButton(self)
        self.btn_analyze_symmetry.setText("Analyze Symmetry")
        self.btn_analyze_symmetry.setToolTip(
            "Analyzes symmetry for each column (ignoring center rows) and highlights symmetric center layers in green."
        )
        self.btn_analyze_symmetry.clicked.connect(self._analyze_symmetry)
        toolbar.addWidget(self.btn_analyze_symmetry)

        # Controles de movimentação.
        self.btn_move_left = QtWidgets.QToolButton(self)
        self.btn_move_left.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_ArrowLeft))
        self.btn_move_left.setToolTip("Move the selected column to the left")
        self.btn_move_left.clicked.connect(lambda: self._move_selected_column(-1))
        toolbar.addWidget(self.btn_move_left)

        self.btn_move_right = QtWidgets.QToolButton(self)
        self.btn_move_right.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_ArrowRight))
        self.btn_move_right.setToolTip("Move the selected column to the right")
        self.btn_move_right.clicked.connect(lambda: self._move_selected_column(1))
        toolbar.addWidget(self.btn_move_right)

        return toolbar

    def _default_export_path(self) -> str:
        try:
            last = self._settings.value("virtual_stacking/last_export_path", "")
        except Exception:
            last = ""
        if last:
            try:
                last_path = Path(str(last))
                return str(last_path.parent)
            except Exception:
                return str(last)
        return str(Path.home())

    def _export_virtual_stacking(self) -> None:
        # Verificar se reordenação por vizinhança foi feita e mostrar aviso se não foi
        if hasattr(self, "btn_reorganize_neighbors") and not self.btn_reorganize_neighbors.isChecked():
            result = QtWidgets.QMessageBox.warning(
                self,
                "Aviso: Reordenação por Vizinhança não foi feita",
                "A reordenação por vizinhança não foi realizada.\n\nVocê pode continuar com a exportação, mas tenha em mente que a ordem das células pode não estar otimizada segundo as regras de vizinhança.",
                QtWidgets.QMessageBox.Ok | QtWidgets.QMessageBox.Cancel,
                QtWidgets.QMessageBox.Cancel,
            )
            if result == QtWidgets.QMessageBox.Cancel:
                return
        if self._project is None or not getattr(self._project, "celulas_ordenadas", []):
            QtWidgets.QMessageBox.information(
                self,
                "Export Virtual Stacking",
                "Nenhum dado de Virtual Stacking disponível para exportar.",
            )
            return

        try:
            ensure_layers_have_material(self._project)
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(
                self,
                "Export Virtual Stacking",
                str(exc),
            )
            return

        suggested = self._default_export_path()
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Export Virtual Stacking",
            suggested,
            "Planilhas Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            self._settings.setValue("virtual_stacking/last_export_path", path)
        except Exception:
            pass

        try:
            from types import SimpleNamespace

            export_cells = [
                SimpleNamespace(
                    cell_id=cell_id,
                    laminate=self._laminate_for_cell(self._project, cell_id),
                )
                for cell_id in self._project.celulas_ordenadas
            ]
            if self._sorted_cell_ids:
                order = {cell_id: pos for pos, cell_id in enumerate(self._sorted_cell_ids)}
                export_cells.sort(key=lambda c: order.get(c.cell_id, len(order)))

            output_path = export_virtual_stacking(self._layers, export_cells, Path(path))
        except Exception as exc:
            QtWidgets.QMessageBox.critical(
                self,
                "Export Virtual Stacking",
                f"Falha ao exportar a planilha:\n{exc}",
            )
            return

        try:
            self._settings.setValue("virtual_stacking/last_export_path", str(output_path))
        except Exception:
            pass

        QtWidgets.QMessageBox.information(
            self,
            "Export Virtual Stacking",
            "Planilha exportada para:\n"
            f"{output_path}",
        )

    # Data binding ----------------------------------------------------

    def populate_from_project(self, project: Optional[GridModel]) -> None:
        """Populate labels and table using the provided project snapshot."""
        previous_project = getattr(self, "_project", None)
        self._project = project
        if project is None:
            self._sorted_cell_ids = []
            self._initial_sort_done = False
        elif not self._initial_sort_done or project is not previous_project:
            self._sorted_cell_ids = self._compute_sorted_cell_ids(project)
            self._initial_sort_done = True
        self._rebuild_view()

    def _rebuild_view(self) -> None:
        if self._project is None:
            self._layers = []
            self._cells = []
            self._symmetry_row_index = None
            self._selected_cell_ids.clear()
            self.model.set_virtual_stacking([], [], None)
            self.model.set_selected_columns([])
            self._stacking_models.clear()
            self._update_summary_row()
            return

        layers, cells, symmetry_index, evaluations = self._collect_virtual_data(self._project)
        self._layers = layers
        self._cells = cells
        self._symmetry_evaluations = evaluations
        if isinstance(symmetry_index, set):
            self._symmetry_rows = symmetry_index
            self._symmetry_row_index = min(symmetry_index) if symmetry_index else None
        elif symmetry_index is None:
            self._symmetry_rows = set()
            self._symmetry_row_index = None
        else:
            self._symmetry_rows = {symmetry_index}
            self._symmetry_row_index = symmetry_index
        valid_ids = {cell.cell_id for cell in cells}
        self._selected_cell_ids = {cid for cid in self._selected_cell_ids if cid in valid_ids}
        self._refresh_stacking_models(cells)
        self.model.set_virtual_stacking(layers, cells, self._symmetry_rows or self._symmetry_row_index)
        self._apply_orientation_delegate()
        self._restore_column_order()
        self._resize_columns()
        self._apply_column_selection()
        self._update_summary_row()
        self._check_symmetry()

    def _collect_virtual_data(
        self, project: GridModel
    ) -> tuple[
        list[VirtualStackingLayer],
        list[VirtualStackingCell],
        set[int] | int | None,
        dict[int, LaminateSymmetryEvaluation],
    ]:
        cells: list[VirtualStackingCell] = []
        laminates: list[Laminado] = []
        for cell_id in project.celulas_ordenadas:
            laminate = self._laminate_for_cell(project, cell_id)
            if laminate is None:
                continue
            laminates.append(laminate)
            cells.append(
                VirtualStackingCell(
                    cell_id=cell_id,
                    laminate=laminate,
                )
            )

        # Ordena apenas na abertura inicial; depois mantém a ordem já mostrada.
        if self._sorted_cell_ids:
            order = {cell_id: pos for pos, cell_id in enumerate(self._sorted_cell_ids)}
            cells.sort(key=lambda c: order.get(c.cell_id, len(order)))
        laminates = [cell.laminate for cell in cells]
        evaluations: dict[int, LaminateSymmetryEvaluation] = {}

        max_layers = max((len(lam.camadas) for lam in laminates), default=0)
        layers: list[VirtualStackingLayer] = []
        for idx in range(max_layers):
            material = ""
            sequence_label = ""
            ply_label = ""
            ply_type = DEFAULT_PLY_TYPE
            rosette = ""
            for lam in laminates:
                if idx >= len(lam.camadas):
                    continue
                camada = lam.camadas[idx]
                if not material and getattr(camada, "material", ""):
                    material = camada.material
                if not sequence_label and getattr(camada, "sequence", ""):
                    sequence_label = camada.sequence
                if not ply_label and getattr(camada, "ply_label", ""):
                    ply_label = camada.ply_label
                if getattr(camada, "ply_type", ""):
                    candidate_type = normalize_ply_type_label(camada.ply_type)
                    if candidate_type == PLY_TYPE_OPTIONS[1]:
                        ply_type = candidate_type
                    elif ply_type == DEFAULT_PLY_TYPE:
                        ply_type = candidate_type
                if not rosette and getattr(camada, "rosette", ""):
                    rosette = camada.rosette
                if material and sequence_label and ply_label and rosette and ply_type:
                    break
            if not sequence_label:
                sequence_label = f"Seq.{idx + 1}"
            if not ply_label:
                ply_label = f"Ply.{idx + 1}"
            if not rosette:
                rosette = DEFAULT_ROSETTE_LABEL
            layers.append(
                VirtualStackingLayer(
                    sequence_label=sequence_label,
                    ply_label=ply_label,
                    ply_type=ply_type,
                    material=material,
                    rosette=rosette,
                )
            )

        for lam in laminates:
            evaluation = evaluate_symmetry_for_layers(getattr(lam, "camadas", []))
            evaluations[id(lam)] = evaluation
        symmetry_rows = self._compute_symmetry_axis_from_layers(layers)
        return layers, cells, symmetry_rows, evaluations

    def _refresh_stacking_models(self, cells: list[VirtualStackingCell]) -> None:
        active_ids = {id(cell.laminate) for cell in cells}
        for lam_id in list(self._stacking_models.keys()):
            if lam_id not in active_ids:
                del self._stacking_models[lam_id]
        for cell in cells:
            laminate = cell.laminate
            model = self._stacking_model_for(laminate)
            if model is not None:
                model.update_layers(list(getattr(laminate, "camadas", [])))

    def _stacking_model_for(self, laminate: Laminado) -> Optional[StackingTableModel]:
        key = id(laminate)
        model = self._stacking_models.get(key)
        if model is None:
            model = StackingTableModel(
                camadas=list(getattr(laminate, "camadas", [])),
                change_callback=lambda layers, lam=laminate: self._on_layers_replaced(
                    lam, layers
                ),
                undo_stack=self.undo_stack,
                most_used_material_provider=self._most_used_material,
            )
            self._stacking_models[key] = model
        else:
            model.set_undo_stack(self.undo_stack)
        return model

    def _on_layers_replaced(self, laminate: Laminado, layers: list[Camada]) -> None:
        laminate.camadas = layers

    def _mark_project_dirty(self) -> None:
        if self._project is not None:
            try:
                self._project.mark_dirty(True)
            except Exception:
                pass

    def _extract_label_prefix(self, label: str, default: str) -> str:
        text = str(label or "").strip()
        match = _PREFIX_NUMBER_PATTERN.match(text)
        if match:
            prefix = (match.group(1) or "").strip()
            return prefix or default
        return default

    def _extract_label_number(self, label: str, fallback: int) -> int:
        text = str(label or "").strip()
        match = _PREFIX_NUMBER_PATTERN.match(text)
        if match:
            try:
                return int(match.group(2))
            except ValueError:
                return fallback
        return fallback

    def _current_prefix_from_project(self, attr: str, default: str) -> str:
        if self._project is None:
            return default
        for laminate in getattr(self._project, "laminados", {}).values():
            for layer in getattr(laminate, "camadas", []):
                label = getattr(layer, attr, "")
                if label:
                    return self._extract_label_prefix(label, default)
        return default

    def _current_label_seed_from_project(
        self,
        attr: str,
        default_prefix: str,
        default_separator: str = ".",
        default_start: int = 1,
    ) -> tuple[str, str, int]:
        if self._project is None:
            return default_prefix, default_separator, default_start
        for laminate in getattr(self._project, "laminados", {}).values():
            layers = getattr(laminate, "camadas", [])
            for idx, layer in enumerate(layers):
                label = str(getattr(layer, attr, "") or "").strip()
                if not label:
                    continue
                match = _PREFIX_NUMBER_PATTERN.fullmatch(label)
                if match:
                    prefix = (match.group(1) or default_prefix).strip() or default_prefix
                    separator = "." if "." in label else ""
                    try:
                        number = int(match.group(2))
                    except ValueError:
                        number = idx + 1
                    start_number = max(1, number - idx)
                    return prefix, separator, start_number
        return default_prefix, default_separator, default_start

    def _current_sequence_prefix(self) -> str:
        return self._current_prefix_from_project("sequence", "Seq")

    def _current_ply_prefix(self) -> str:
        return self._current_prefix_from_project("ply_label", "Ply")

    def _most_used_material(self) -> Optional[str]:
        """Retorna o material mais utilizado em todos os laminados carregados."""
        return project_most_used_material(self._project)

    def _sync_material_for_row(self, row: int, material: str) -> list[Laminado]:
        if self._project is None:
            return []
        return sync_material_by_sequence(
            self._project,
            row,
            material,
            stacking_model_provider=self._stacking_model_for,
        )

    def _prompt_prefix_dialog(self, title: str, current_label: str) -> Optional[str]:
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle(title)
        layout = QtWidgets.QVBoxLayout(dialog)
        layout.addWidget(
            QtWidgets.QLabel(f"Current base label: {current_label}", dialog)
        )
        input_box = QtWidgets.QLineEdit(dialog)
        input_box.setText(current_label)
        layout.addWidget(input_box)
        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel,
            parent=dialog,
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        input_box.selectAll()
        input_box.setFocus()
        if dialog.exec() == QtWidgets.QDialog.Accepted:
            return input_box.text().strip()
        return None

    def _apply_global_prefix_change(
        self,
        target_attr: str,
        column: int,
        new_prefix: str,
        *,
        separator: str,
        start_number: int,
    ) -> list[str]:
        if self._project is None:
            return []
        prefix = new_prefix.strip().rstrip(".")
        if not prefix:
            return []
        base_number = max(1, int(start_number))
        changed: list[str] = []
        for laminate in getattr(self._project, "laminados", {}).values():
            model = self._stacking_model_for(laminate)
            if model is None:
                continue
            layers = model.layers()
            updated = False
            for idx, layer in enumerate(layers):
                current_label = getattr(layer, target_attr, "")
                target_label = f"{prefix}{separator}{base_number + idx}"
                if str(current_label or "") == target_label:
                    continue
                if model.apply_field_value(idx, column, target_label):
                    updated = True
                else:
                    setattr(layer, target_attr, target_label)
                    updated = True
            if updated:
                laminate.camadas = model.layers()
                changed.append(laminate.nome)
        if changed:
            self._notify_changes(changed)
        return changed

    def _rename_all_sequences(self) -> None:
        current_prefix, current_sep, current_start = self._current_label_seed_from_project(
            "sequence", "Seq"
        )
        current_label = f"{current_prefix}{current_sep}{current_start}"
        new_label = self._prompt_prefix_dialog("Rename Sequence", current_label)
        if new_label is None:
            return
        cleaned = new_label.strip()
        if not cleaned:
            return
        cleaned_no_dot = cleaned.rstrip(".")
        match = _PREFIX_NUMBER_PATTERN.fullmatch(cleaned_no_dot)
        if match:
            prefix = (match.group(1) or current_prefix).strip() or current_prefix
            separator = "." if "." in cleaned_no_dot else ""
            try:
                start_number = int(match.group(2))
            except ValueError:
                start_number = current_start
        else:
            prefix = cleaned_no_dot
            separator = "." if "." in cleaned_no_dot else ""
            start_number = current_start
        if not prefix:
            return
        if (
            prefix == current_prefix
            and separator == current_sep
            and int(start_number) == int(current_start)
        ):
            return
        self._push_virtual_snapshot()
        self._apply_global_prefix_change(
            "sequence",
            StackingTableModel.COL_SEQUENCE,
            prefix,
            separator=separator,
            start_number=start_number,
        )

    def _rename_all_ply(self) -> None:
        current_prefix, current_sep, current_start = self._current_label_seed_from_project(
            "ply_label", "Ply"
        )
        current_label = f"{current_prefix}{current_sep}{current_start}"
        new_label = self._prompt_prefix_dialog("Rename Ply", current_label)
        if new_label is None:
            return
        cleaned = new_label.strip()
        if not cleaned:
            return
        cleaned_no_dot = cleaned.rstrip(".")
        match = _PREFIX_NUMBER_PATTERN.fullmatch(cleaned_no_dot)
        if match:
            prefix = (match.group(1) or current_prefix).strip() or current_prefix
            separator = "." if "." in cleaned_no_dot else ""
            try:
                start_number = int(match.group(2))
            except ValueError:
                start_number = current_start
        else:
            prefix = cleaned_no_dot
            separator = "." if "." in cleaned_no_dot else ""
            start_number = current_start
        if not prefix:
            return
        if (
            prefix == current_prefix
            and separator == current_sep
            and int(start_number) == int(current_start)
        ):
            return
        self._push_virtual_snapshot()
        self._apply_global_prefix_change(
            "ply_label",
            StackingTableModel.COL_PLY,
            prefix,
            separator=separator,
            start_number=start_number,
        )

    def _auto_rename_if_enabled(self, laminate: Laminado) -> None:
        if self._project is None:
            return
        laminate.auto_rename_enabled = True
        new_name = auto_name_for_laminate(self._project, laminate)
        if not new_name or new_name == laminate.nome:
            return
        self._rename_laminate(laminate, new_name)

    def _rename_laminate(self, laminate: Laminado, new_name: str) -> None:
        if self._project is None:
            return
        old_name = laminate.nome
        if not new_name or new_name == old_name:
            return
        laminados = self._project.laminados
        if new_name in laminados and laminados[new_name] is not laminate:
            return
        updated = OrderedDict()
        for name, lam in laminados.items():
            if lam is laminate:
                updated[new_name] = lam
            else:
                updated[name] = lam
        self._project.laminados = updated
        laminate.nome = new_name
        for cell_id, mapped in list(self._project.cell_to_laminate.items()):
            if mapped == old_name:
                self._project.cell_to_laminate[cell_id] = new_name

    def _ensure_unique_laminate_for_cell(self, cell: VirtualStackingCell) -> Laminado:
        """
        Ensure the cell has a unique laminate object.
        If the laminate is shared with other cells, create a copy for this cell.
        Returns the laminate (either the original if unique, or a new copy).
        """
        if self._project is None:
            return cell.laminate
        
        # Check how many cells in the ENTIRE PROJECT reference this laminate by NAME
        laminate_name = cell.laminate.nome
        cells_using_this_laminate = [
            cell_id for cell_id, lam_name in self._project.cell_to_laminate.items()
            if lam_name == laminate_name
        ]
        
        # Also check laminate.celulas for backward compatibility
        if hasattr(cell.laminate, 'celulas'):
            for c_id in cell.laminate.celulas:
                if c_id not in cells_using_this_laminate:
                    cells_using_this_laminate.append(c_id)
        
        # If only this cell uses the laminate, no need to copy
        if len(cells_using_this_laminate) <= 1:
            return cell.laminate
        
        # Create a deep copy of the laminate
        import copy as copy_module
        new_laminate = copy_module.deepcopy(cell.laminate)
        
        # Generate a unique name for the new laminate
        original_name = cell.laminate.nome
        counter = 1
        base_name = original_name
        # Remove existing number suffix if present
        match = _PREFIX_NUMBER_PATTERN.match(original_name)
        if match:
            base_name = match.group(1)
            counter = int(match.group(2))
        
        # Find a unique name
        while True:
            new_name = f"{base_name}.{counter}"
            if new_name not in self._project.laminados:
                break
            counter += 1
        
        new_laminate.nome = new_name
        
        # Remove this cell from the old laminate's cell list
        old_laminate = cell.laminate
        if hasattr(old_laminate, 'celulas') and cell.cell_id in old_laminate.celulas:
            old_laminate.celulas.remove(cell.cell_id)
        
        # Add this cell to the new laminate's cell list
        if hasattr(new_laminate, 'celulas'):
            if cell.cell_id not in new_laminate.celulas:
                new_laminate.celulas.append(cell.cell_id)
        else:
            new_laminate.celulas = [cell.cell_id]
        
        # Add the new laminate to the project
        self._project.laminados[new_name] = new_laminate
        
        # Update the cell-to-laminate mapping for this cell
        self._project.cell_to_laminate[cell.cell_id] = new_name
        
        # Update the cell's laminate reference
        cell.laminate = new_laminate
        
        # Update any other cells in self._cells that were pointing to the old laminate
        # but should now point to the new one (only this specific cell)
        for other_cell in self._cells:
            if other_cell.cell_id == cell.cell_id and id(other_cell.laminate) == id(old_laminate):
                other_cell.laminate = new_laminate
        
        # Create a new stacking model for the new laminate
        # (the old model will be cleaned up automatically)
        self._stacking_models[id(new_laminate)] = StackingTableModel(
            camadas=list(getattr(new_laminate, "camadas", [])),
            change_callback=lambda layers, lam=new_laminate: self._on_layers_replaced(
                lam, layers
            ),
            undo_stack=self.undo_stack,
            most_used_material_provider=self._most_used_material,
        )
        
        return new_laminate

    def _after_laminate_changed(self, laminate: Laminado) -> None:
        self._auto_rename_if_enabled(laminate)
        self._mark_project_dirty()
        self._update_summary_row()


    def _selected_targets(self) -> dict[Laminado, set[int]]:
        if not hasattr(self, "table") or self.table.selectionModel() is None:
            return {}
        targets: dict[Laminado, set[int]] = {}
        selected_rows: set[int] = set()
        selection = self.table.selectionModel()
        for index in selection.selectedIndexes():
            if not index.isValid():
                continue
            selected_rows.add(index.row())
            if index.column() < self.model.LAMINATE_COLUMN_OFFSET:
                continue
            cell_idx = index.column() - self.model.LAMINATE_COLUMN_OFFSET
            if 0 <= cell_idx < len(self._cells):
                laminate = self._cells[cell_idx].laminate
                targets.setdefault(laminate, set()).add(index.row())
        if not targets and not selected_rows:
            current = self.table.currentIndex()
            if current.isValid():
                selected_rows.add(current.row())
                if current.column() >= self.model.LAMINATE_COLUMN_OFFSET:
                    cell_idx = current.column() - self.model.LAMINATE_COLUMN_OFFSET
                    if 0 <= cell_idx < len(self._cells):
                        laminate = self._cells[cell_idx].laminate
                        targets.setdefault(laminate, set()).add(current.row())
        if not targets and selected_rows:
            for cell in self._cells:
                targets.setdefault(cell.laminate, set()).update(selected_rows)
        return targets

    def _orientation_token(self, value: object) -> Optional[float]:
        if value is None:
            return None
        try:
            return normalize_angle(value)
        except Exception:
            try:
                return normalize_angle(str(value))
            except Exception:
                return None

    def _orientation_bucket(self, value: object) -> Optional[str]:
        token = self._orientation_token(value)
        if token is None:
            return None
        angle = float(token)
        candidates = [
            (0.0, "0"),
            (45.0, "+45"),
            (-45.0, "-45"),
            (90.0, "90"),
            (-90.0, "90"),
            (180.0, "0"),
            (-180.0, "0"),
        ]
        closest = min(candidates, key=lambda item: abs(angle - item[0]))
        if abs(angle - closest[0]) <= 10.0:
            return closest[1]
        return "other"

    def _orientation_counts_for_laminate(
        self, laminate: Optional[Laminado]
    ) -> tuple[Counter[str], int]:
        counts: Counter[str] = Counter()
        if laminate is None:
            return counts, 0
        for camada in getattr(laminate, "camadas", []):
            bucket = self._orientation_bucket(getattr(camada, "orientacao", None))
            if bucket is None:
                continue
            counts[bucket] += 1
        total = sum(counts.values())
        return counts, total

    def _classify_laminate_type(self, counts: Counter[str], total: int) -> tuple[str, float]:
        if total <= 0:
            return "-", 0.0

        pct_zero = counts.get("0", 0) / total
        pct_45 = (counts.get("+45", 0) + counts.get("-45", 0)) / total
        pct_90 = counts.get("90", 0) / total

        threshold = 0.45
        if pct_zero >= threshold and pct_zero >= pct_45 and pct_zero >= pct_90:
            return "Hard", pct_zero
        if pct_45 >= threshold and pct_45 >= pct_zero and pct_45 >= pct_90:
            return "Soft", pct_45

        dominant_pct = max(pct_zero, pct_45, pct_90)
        return "Quasi-isotropic", dominant_pct

    def _row_considered(self, row: int) -> bool:
        if not (0 <= row < len(self._layers)):
            return True
        layer = self._layers[row]
        return normalize_ply_type_label(getattr(layer, "ply_type", DEFAULT_PLY_TYPE)) != PLY_TYPE_OPTIONS[1]

    def _considered_sequence_rows_from_layers(self, layers: list[VirtualStackingLayer]) -> list[int]:
        rows: list[int] = []
        for idx, layer in enumerate(layers):
            ply_type = normalize_ply_type_label(getattr(layer, "ply_type", DEFAULT_PLY_TYPE))
            if ply_type != PLY_TYPE_OPTIONS[1]:
                rows.append(idx)
        return rows

    def _summary_for_laminate(self, laminate: Optional[Laminado]) -> str:
        """Return total layer count plus laminate type for the header."""
        counts, total = self._orientation_counts_for_laminate(laminate)
        laminate_type, dominant_pct = self._classify_laminate_type(counts, total)

        lines = [f"Total: {total}"]
        if total > 0:
            pct_text = f" ({dominant_pct * 100:.0f}%)" if dominant_pct > 0 else ""
            lines.append(f"Tipo: {laminate_type}{pct_text}")
        else:
            lines.append("Tipo: -")

        return "\n".join(lines)

    def _get_orientation_summary_for_laminate(self, laminate: Optional[Laminado]) -> str:
        """Build a compact summary with counts by orientation and laminate type."""
        counts, total = self._orientation_counts_for_laminate(laminate)
        laminate_type, dominant_pct = self._classify_laminate_type(counts, total)

        lines: list[str] = []
        lines.append(f"Total oriented layers: {total}")
        if total > 0:
            pct_text = f" ({dominant_pct * 100:.0f}%)" if dominant_pct > 0 else ""
            lines.append(f"Dominant type: {laminate_type}{pct_text}")
        else:
            lines.append("Dominant type: -")

        lines.append("")
        lines.append("Orientation counts:")
        for key, title in [
            ("0", "0 deg"),
            ("+45", "+45 deg"),
            ("-45", "-45 deg"),
            ("90", "90 deg"),
        ]:
            lines.append(f"  {title}: {counts.get(key, 0)}")
        other = counts.get("other", 0)
        if other:
            lines.append(f"  Outras: {other}")

        return "\n".join(lines)

    def _update_summary_row(self) -> None:
        # Gera resumos e injeta no cabeçalho das colunas de laminados.
        column_count = self.model.columnCount()
        summaries: list[str] = []
        for col in range(column_count - self.model.LAMINATE_COLUMN_OFFSET):
            cell_idx = col
            laminate = self._cells[cell_idx].laminate if 0 <= cell_idx < len(self._cells) else None
            summaries.append(self._summary_for_laminate(laminate))
        self.model.set_column_summaries(summaries)

    def _update_summary_height(self, lines: int, metrics: QtGui.QFontMetrics) -> None:
        if not getattr(self, "summary_table", None):
            return
        line_height = metrics.lineSpacing() + 2
        padding = 10
        target_height = max(48, lines * line_height + padding)
        self.summary_table.setFixedHeight(target_height)
        try:
            if self.summary_table.model() is not None and self.summary_table.model().rowCount() > 0:
                self.summary_table.setRowHeight(0, max(36, target_height - 8))
        except Exception:
            pass

    def _resize_summary_columns(self) -> None:
        header = self.table.horizontalHeader()
        summary_header = (
            self.summary_table.horizontalHeader()
            if getattr(self, "summary_table", None)
            else None
        )
        if header is None or summary_header is None:
            return
        for col in range(self.model.columnCount()):
            summary_header.resizeSection(col, header.sectionSize(col))

    def _resize_columns(self) -> None:
        header = self.table.horizontalHeader()
        if header is None:
            return
        metrics = header.fontMetrics()
        padding = 18
        if self.model.columnCount() >= 1:
            seq_width = max(metrics.horizontalAdvance("Sequence") + padding, 70)
            header.resizeSection(VirtualStackingModel.COL_SEQUENCE, seq_width)
        if self.model.columnCount() >= 2:
            ply_width = max(metrics.horizontalAdvance("Ply") + padding, 55)
            header.resizeSection(VirtualStackingModel.COL_PLY, ply_width)
        if self.model.columnCount() >= 3:
            ply_type_width = max(metrics.horizontalAdvance("Symmetry") + padding, 90)
            header.resizeSection(VirtualStackingModel.COL_PLY_TYPE, ply_type_width)
        if self.model.columnCount() >= 4:
            material_width = max(metrics.horizontalAdvance("Material") + padding * 2, 140)
            header.resizeSection(VirtualStackingModel.COL_MATERIAL, material_width)
        if self.model.columnCount() >= 5:
            rosette_width = max(metrics.horizontalAdvance("Rosette") + padding, 80)
            header.resizeSection(VirtualStackingModel.COL_ROSETTE, rosette_width)
        for col in range(self.model.LAMINATE_COLUMN_OFFSET, self.model.columnCount()):
            if header.sectionSize(col) < 110:
                header.resizeSection(col, 110)
        self._resize_summary_columns()

    def _orientations_match(self, left: Optional[float], right: Optional[float]) -> bool:
        if left is None and right is None:
            return True
        if left is None or right is None:
            return False
        return math.isclose(left, right, abs_tol=1e-6)

    def _update_symmetry_rows_from_union(self, rows: set[int]) -> None:
        new_rows = set(rows)
        if new_rows == self._symmetry_rows:
            return
        self._symmetry_rows = new_rows
        self._symmetry_row_index = min(new_rows) if new_rows else None
        if hasattr(self, "model"):
            try:
                self.model.set_symmetry_rows(new_rows)
            except Exception:
                pass

    def _check_symmetry(self) -> None:
        """
        Check for unbalanced columns using Classical Lamination Theory (CLT) criterion.
        Note: Green borders for symmetric central layers are now controlled
        by the explicit "Analisar Simetria" button, not by this method.
        Red cells are no longer used in symmetry analysis.
        """
        # No more red cells for symmetry analysis
        unbalanced_columns: set[int] = set()
        evaluations: dict[int, LaminateSymmetryEvaluation] = {}

        for col, cell in enumerate(self._cells):
            laminate = cell.laminate
            layers = getattr(laminate, "camadas", [])
            evaluation = evaluate_symmetry_for_layers(layers)
            evaluations[id(laminate)] = evaluation

            # Check for unbalanced columns using CLT criterion
            balance_evaluation = evaluate_laminate_balance_clt(layers)
            if not balance_evaluation.is_balanced:
                unbalanced_columns.add(col + self.model.LAMINATE_COLUMN_OFFSET)

        self._symmetry_evaluations = evaluations
        # No red cells - only unbalanced columns
        self.model.set_highlights(set(), set())
        self.model.set_unbalanced_columns(unbalanced_columns)

    def _is_unbalanced(
        self,
        layers: list[Camada],
        structural_rows: list[int],
        center_rows: list[int],
    ) -> bool:
        """
        Deprecated: Use evaluate_laminate_balance_clt() instead.
        Kept for backward compatibility.
        """
        balance_evaluation = evaluate_laminate_balance_clt(layers)
        return not balance_evaluation.is_balanced

    def _targets_for_insertion(
        self, index: Optional[QtCore.QModelIndex] = None
    ) -> dict[Laminado, set[int]]:
        if index is not None and index.isValid() and index.column() >= self.model.LAMINATE_COLUMN_OFFSET:
            cell_idx = index.column() - self.model.LAMINATE_COLUMN_OFFSET
            if 0 <= cell_idx < len(self._cells):
                laminate = self._cells[cell_idx].laminate
                return {laminate: {index.row()}}
        return self._selected_targets()

    def _stacking_entry_for_index(
        self, index: QtCore.QModelIndex
    ) -> Optional[tuple[Laminado, StackingTableModel, int]]:
        if not index.isValid() or index.column() < self.model.LAMINATE_COLUMN_OFFSET:
            return None
        cell_idx = index.column() - self.model.LAMINATE_COLUMN_OFFSET
        if not (0 <= cell_idx < len(self._cells)):
            return None
        laminate = self._cells[cell_idx].laminate
        model = self._stacking_model_for(laminate)
        if model is None:
            return None
        return laminate, model, index.row()

    def _prompt_custom_orientation(self, parent: QtWidgets.QWidget) -> float | None:
        dialog = QtWidgets.QInputDialog(parent)
        dialog.setInputMode(QtWidgets.QInputDialog.DoubleInput)
        dialog.setWindowTitle("Custom value")
        dialog.setLabelText("Enter the orientation (-100 to 100 degrees):")
        dialog.setDoubleRange(-100.0, 100.0)
        dialog.setDoubleDecimals(1)
        dialog.setDoubleStep(1.0)
        dialog.setDoubleValue(0.0)
        dialog.setTextValue("")
        if dialog.exec() != dialog.Accepted:
            return None
        return dialog.doubleValue()

    def _edit_orientation_at(self, index: QtCore.QModelIndex) -> None:
        entry = self._stacking_entry_for_index(index)
        if entry is None:
            return
        laminate, stacking_model, row = entry
        if not (0 <= row < stacking_model.rowCount()):
            QtWidgets.QMessageBox.information(
                self,
                "Orientation unavailable",
                "The selected row does not exist for this laminate.",
            )
            return
        options = ["Empty", "0", "45", "-45", "90", "Custom value..."]
        selected, ok = QtWidgets.QInputDialog.getItem(
            self,
            "Edit orientation",
            "Select the orientation:",
            options,
            0,
            False,
        )
        if not ok:
            return
        selected = selected.strip()
        if selected == "Custom value...":
            custom_value = self._prompt_custom_orientation(self)
            if custom_value is None:
                return
            payload = f"{custom_value:g}"
        elif selected.lower() == "empty":
            payload = ""
        else:
            payload = selected
        if not self.model.setData(index, payload, QtCore.Qt.EditRole):
            QtWidgets.QMessageBox.warning(
                self,
                "Invalid orientation",
                "Use values between -100 and 100 degrees or leave it blank.",
            )

    def _clear_orientation_at(self, index: QtCore.QModelIndex) -> None:
        if not index.isValid():
            return
        if not self.model.setData(index, "", QtCore.Qt.EditRole):
            QtWidgets.QMessageBox.information(
                self,
                "Clear orientation",
                "Could not clear this cell. Verify that the row exists.",
            )

    def _remove_layer_at(self, index: QtCore.QModelIndex) -> None:
        entry = self._stacking_entry_for_index(index)
        if entry is None:
            return
        laminate, stacking_model, row = entry
        if not (0 <= row < stacking_model.rowCount()):
            QtWidgets.QMessageBox.information(
                self,
                "Remove layer",
                "Select an existing row to remove.",
            )
            return
        
        # Ensure this cell has a unique laminate before removing the layer
        # to prevent affecting other cells that might share the same laminate
        cell_idx = index.column() - self.model.LAMINATE_COLUMN_OFFSET
        if 0 <= cell_idx < len(self._cells):
            current_cell = self._cells[cell_idx]
            laminate = self._ensure_unique_laminate_for_cell(current_cell)
            # Get the updated stacking model for the potentially new laminate
            stacking_model = self._stacking_model_for(laminate)
            if stacking_model is None:
                return
        
        command = _RemoveLayerCommand(stacking_model, laminate, [row])
        self._execute_command(command)
        self._after_laminate_changed(laminate)
        self._notify_changes([laminate.nome])

    def _add_sequence_at_position(self, row: int, insert_below: bool) -> None:
        if row < 0 or not self._cells:
            return
        
        # Create undo/redo command
        before_state = self._capture_virtual_snapshot()
        
        self._push_virtual_snapshot()
        changed: list[str] = []
        insert_at = row + (1 if insert_below else 0)
        for cell in self._cells:
            laminate = self._ensure_unique_laminate_for_cell(cell)
            stacking_model = self._stacking_model_for(laminate)
            if stacking_model is None:
                continue
            try:
                self.model._ensure_row_exists(stacking_model, row, laminate)
            except Exception:
                pass
            target_row = max(0, min(insert_at, stacking_model.rowCount()))
            new_layer = Camada(
                idx=0,
                material="",
                orientacao=None,
                ativo=True,
                simetria=False,
                ply_type=DEFAULT_PLY_TYPE,
                rosette=DEFAULT_ROSETTE_LABEL,
            )
            stacking_model.insert_layer(target_row, new_layer)
            laminate.camadas = stacking_model.layers()
            self._after_laminate_changed(laminate)
            changed.append(laminate.nome)
        if changed:
            self._notify_changes(changed)

    def _add_sequence_above_row(self, row: int) -> None:
        self._add_sequence_at_position(row, insert_below=False)

    def _add_sequence_below_row(self, row: int) -> None:
        self._add_sequence_at_position(row, insert_below=True)

    def _delete_sequence_row(self, row: int) -> None:
        if row < 0 or not self._cells:
            return
        has_targets = False
        for cell in self._cells:
            model = self._stacking_model_for(cell.laminate)
            if model is not None and 0 <= row < model.rowCount():
                has_targets = True
                break
        if not has_targets:
            return
        self._push_virtual_snapshot()
        changed: list[str] = []
        for cell in self._cells:
            laminate = self._ensure_unique_laminate_for_cell(cell)
            stacking_model = self._stacking_model_for(laminate)
            if stacking_model is None:
                continue
            if 0 <= row < stacking_model.rowCount():
                stacking_model.remove_rows([row])
                laminate.camadas = stacking_model.layers()
                self._after_laminate_changed(laminate)
                changed.append(laminate.nome)
        if changed:
            self._notify_changes(changed)

    def _insert_layer_for_cell(self, index: QtCore.QModelIndex, insert_at: int) -> None:
        if not index.isValid() or index.column() < self.model.LAMINATE_COLUMN_OFFSET:
            return
        cell_idx = index.column() - self.model.LAMINATE_COLUMN_OFFSET
        if not (0 <= cell_idx < len(self._cells)):
            return
        current_cell = self._cells[cell_idx]
        laminate = self._ensure_unique_laminate_for_cell(current_cell)
        stacking_model = self._stacking_model_for(laminate)
        if stacking_model is None:
            return
        target_pos = max(0, min(insert_at, stacking_model.rowCount()))
        default_material = self._most_used_material() or ""
        command = _InsertLayerCommand(
            stacking_model,
            laminate,
            [target_pos],
            default_material=default_material,
        )
        self._execute_command(command)
        laminate.camadas = stacking_model.layers()
        try:
            self.model._auto_fill_material_if_missing(laminate, stacking_model, target_pos)
        except Exception:
            pass
        self._after_laminate_changed(laminate)
        self._notify_changes([laminate.nome])

    def _add_layer_above(self, index: QtCore.QModelIndex) -> None:
        self._insert_layer_for_cell(index, index.row())

    def _add_layer_below(self, index: QtCore.QModelIndex) -> None:
        self._insert_layer_for_cell(index, index.row() + 1)

    def _insert_layer(self, index: Optional[QtCore.QModelIndex] = None) -> None:
        # If a specific cell was clicked, handle that single insertion directly.
        if index is not None and index.isValid() and index.column() >= self.model.LAMINATE_COLUMN_OFFSET:
            self._insert_layer_for_cell(index, index.row())
            return

        targets = self._targets_for_insertion(index)
        if not targets:
            QtWidgets.QMessageBox.information(
                self,
                "Insert layer",
                "Select at least one row to insert a new layer.",
            )
            return
        affected: list[str] = []
        touched: list[str] = []
        for laminate, rows in targets.items():
            model = self._stacking_model_for(laminate)
            if model is None:
                continue
            touched.append(laminate.nome)
            requested_rows = sorted(rows)
            if not requested_rows:
                continue
            max_row = max(requested_rows)
            positions: list[int] = list(requested_rows)
            if max_row >= model.rowCount():
                positions.extend(range(model.rowCount(), max_row + 1))
            positions = sorted(set(positions))
            if positions:
                default_material = self._most_used_material() or ""
                command = _InsertLayerCommand(
                    model,
                    laminate,
                    positions,
                    default_material=default_material,
                )
                self._execute_command(command)
                laminate.camadas = model.layers()
                for pos in positions:
                    try:
                        self.model._auto_fill_material_if_missing(laminate, model, pos)
                    except Exception:
                        pass
                self._after_laminate_changed(laminate)
                affected.append(laminate.nome)
        if affected or touched:
            self._notify_changes(affected or touched)
        else:
            # Ensure UI refresh even if nothing was recorded as affected.
            self._rebuild_view()

    def _on_model_change(self, laminate_names: list[str]) -> None:
        self._notify_changes(laminate_names)

    def _notify_changes(self, laminate_names: list[str], *, mark_dirty: bool = True) -> None:
        if mark_dirty and self._project is not None:
            try:
                self._project.mark_dirty(True)
            except Exception:
                pass
        self._rebuild_view()
        self._update_undo_buttons()
        try:
            self.stacking_changed.emit(laminate_names)
        except Exception:
            pass

    def _show_context_menu(self, pos: QtCore.QPoint) -> None:
        index = self.table.indexAt(pos)
        if (
            not index.isValid()
            or index.row() < 0
            or index.row() >= self.model.rowCount()
        ):
            return
        self.table.setCurrentIndex(index)
        menu = QtWidgets.QMenu(self)
        
        # Add sequence actions
        add_sequence_above_action = menu.addAction(
            "Add sequence above the selected sequence"
        )
        add_sequence_below_action = menu.addAction(
            "Add sequence below the selected sequence"
        )
        delete_sequence_action = menu.addAction("Delete selected sequence")

        is_sequence_column = index.column() < self.model.LAMINATE_COLUMN_OFFSET
        if not is_sequence_column:
            # Add layer actions for laminate columns
            menu.addSeparator()
            add_layer_above_action = menu.addAction(
                "Add layer above the selected layer"
            )
            add_layer_below_action = menu.addAction(
                "Add layer below the selected layer"
            )
            menu.addSeparator()
            edit_action = menu.addAction("Edit orientation...")
            clear_action = menu.addAction("Clear orientation")
            remove_action = menu.addAction("Remove layer")
        
        chosen = menu.exec(self.table.viewport().mapToGlobal(pos))
        if chosen == add_sequence_above_action:
            self._add_sequence_above_row(index.row())
        elif chosen == add_sequence_below_action:
            self._add_sequence_below_row(index.row())
        elif chosen == delete_sequence_action:
            self._delete_sequence_row(index.row())
        elif not is_sequence_column:
            if chosen == add_layer_above_action:
                self._add_layer_above(index)
            elif chosen == add_layer_below_action:
                self._add_layer_below(index)
            elif chosen == edit_action:
                self._edit_orientation_at(index)
            elif chosen == clear_action:
                self._clear_orientation_at(index)
            elif chosen == remove_action:
                self._remove_layer_at(index)

    def _move_column_left(self, column: int) -> None:
        """Move a laminate column one position to the left."""
        if column < self.model.LAMINATE_COLUMN_OFFSET:
            return  # Not a laminate column
        
        cell_idx = column - self.model.LAMINATE_COLUMN_OFFSET
        if cell_idx <= 0 or cell_idx >= len(self._cells):
            return  # Already at the left or invalid index
        
        # Record before state
        self._push_virtual_snapshot()
        
        # Swap cells in the list
        old_index = cell_idx
        new_index = cell_idx - 1
        command = _MoveColumnCommand(self._cells, old_index, new_index)
        self.undo_stack.push(command)
        
        # Update the sorted cell IDs order
        if len(self._cells) >= 2 and new_index >= 0:
            self._sorted_cell_ids = [cell.cell_id for cell in self._cells]
            if self._project is not None:
                remaining = [cid for cid in self._project.celulas_ordenadas if cid not in self._sorted_cell_ids]
                self._project.celulas_ordenadas = self._sorted_cell_ids + remaining
        
        # Rebuild view and update UI
        self._rebuild_view()
        self._check_symmetry()
        self._mark_project_dirty()
        self._notify_changes([cell.laminate.nome for cell in self._cells])

    def _move_column_right(self, column: int) -> None:
        """Move a laminate column one position to the right."""
        if column < self.model.LAMINATE_COLUMN_OFFSET:
            return  # Not a laminate column
        
        cell_idx = column - self.model.LAMINATE_COLUMN_OFFSET
        if cell_idx < 0 or cell_idx >= len(self._cells) - 1:
            return  # Already at the right or invalid index
        
        # Record before state
        self._push_virtual_snapshot()
        
        # Swap cells in the list
        old_index = cell_idx
        new_index = cell_idx + 1
        command = _MoveColumnCommand(self._cells, old_index, new_index)
        self.undo_stack.push(command)
        
        # Update the sorted cell IDs order
        if len(self._cells) >= 2 and new_index < len(self._cells):
            self._sorted_cell_ids = [cell.cell_id for cell in self._cells]
            if self._project is not None:
                remaining = [cid for cid in self._project.celulas_ordenadas if cid not in self._sorted_cell_ids]
                self._project.celulas_ordenadas = self._sorted_cell_ids + remaining
        
        # Rebuild view and update UI
        self._rebuild_view()
        self._check_symmetry()
        self._mark_project_dirty()
        self._notify_changes([cell.laminate.nome for cell in self._cells])

    def _laminate_for_cell(
        self, model: GridModel, cell_id: str
    ) -> Optional[Laminado]:
        lam_name = model.cell_to_laminate.get(cell_id)
        if lam_name:
            laminate = model.laminados.get(lam_name)
            if laminate is not None:
                return laminate
        for laminate in model.laminados.values():
            if cell_id in laminate.celulas:
                return laminate
        return None

    def _compute_symmetry_axis_from_layers(self, layers: list[VirtualStackingLayer]) -> set[int]:
        total = len(layers)
        if total <= 0:
            return set()
        if total % 2 == 1:
            return {total // 2}
        lower = total // 2 - 1
        return {lower, lower + 1}

    def _compute_sorted_cell_ids(self, project: Optional[GridModel]) -> list[str]:
        """
        Define a ordem inicial das colunas (decrescente por quantidade de camadas)
        apenas quando a janela e populada. Edicoes posteriores preservam essa ordem.
        """
        if project is None:
            return []
        temp_cells: list[tuple[int, int, str]] = []
        for idx, cell_id in enumerate(project.celulas_ordenadas):
            laminate = self._laminate_for_cell(project, cell_id)
            if laminate is None:
                continue
            temp_cells.append((count_oriented_layers(getattr(laminate, "camadas", [])), idx, cell_id))
        temp_cells.sort(key=lambda entry: (-entry[0], entry[1]))
        return [cell_id for _, _, cell_id in temp_cells]

    def _capture_virtual_snapshot(self) -> dict:
        """Return a deep-copied snapshot of the current virtual stacking state."""
        return {
            "layers": copy.deepcopy(self._layers),
            "cells": copy.deepcopy(self._cells),
            "symmetry_row_index": copy.deepcopy(self._symmetry_rows or self._symmetry_row_index),
            "project": copy.deepcopy(self._project),
        }

    def _restore_virtual_snapshot(self, snapshot: Optional[dict], *, keep_project_ref: bool = False) -> None:
        """Restore a previously captured snapshot (used by undo/redo)."""
        if snapshot is None:
            return
        self._layers = copy.deepcopy(snapshot.get("layers", []))
        self._cells = copy.deepcopy(snapshot.get("cells", []))
        symmetry_index = snapshot.get("symmetry_row_index")
        if isinstance(symmetry_index, set):
            self._symmetry_rows = symmetry_index
            self._symmetry_row_index = min(symmetry_index) if symmetry_index else None
        elif symmetry_index is None:
            self._symmetry_rows = set()
            self._symmetry_row_index = None
        else:
            self._symmetry_rows = {symmetry_index}
            self._symmetry_row_index = symmetry_index
        snapshot_project = snapshot.get("project")
        if keep_project_ref and self._project is not None and snapshot_project is not None:
            restored_project = copy.deepcopy(snapshot_project)
            self._project.__dict__.clear()
            self._project.__dict__.update(restored_project.__dict__)
        else:
            self._project = copy.deepcopy(snapshot_project)
        self._rebuild_view()

    def _push_virtual_snapshot(self) -> None:
        """Record a snapshot on the undo stack so reorg operations can be undone."""
        if self.undo_stack is None:
            return
        window = self
        before_state = self._capture_virtual_snapshot()

        class _VirtualStackingSnapshotCommand(QtGui.QUndoCommand):
            def __init__(self, state: dict) -> None:
                super().__init__("Virtual Stacking Snapshot")
                self.before_state = state
                self.after_state: Optional[dict] = None

            def undo(self) -> None:
                self.after_state = window._capture_virtual_snapshot()
                window._restore_virtual_snapshot(self.before_state)

            def redo(self) -> None:
                if self.after_state is not None:
                    window._restore_virtual_snapshot(self.after_state)

        self.undo_stack.push(_VirtualStackingSnapshotCommand(before_state))
        self._update_undo_buttons()

    # ---------------------------------------------------------------
    # Reorganizar por vizinhança
    def on_reorganizar_por_vizinhanca_clicked(self) -> None:
        """Slot conectado ao botão da toolbar."""
        if hasattr(self, "btn_reorganize_neighbors"):
            self.btn_reorganize_neighbors.setChecked(
                not self.btn_reorganize_neighbors.isChecked()
            )
        else:
            self.reorganizar_por_vizinhanca()

    def _toggle_reorganizar_por_vizinhanca(self, checked: bool) -> None:
        if checked:
            self._neighbors_reorder_snapshot = self._capture_virtual_snapshot()
            applied = self.reorganizar_por_vizinhanca()
            if not applied:
                self._neighbors_reorder_snapshot = None
                blocker = QtCore.QSignalBlocker(self.btn_reorganize_neighbors)
                self.btn_reorganize_neighbors.setChecked(False)
                del blocker
                if hasattr(self, "btn_export_virtual"):
                    self.btn_export_virtual.setEnabled(False)
            else:
                if hasattr(self, "btn_export_virtual"):
                    self.btn_export_virtual.setEnabled(True)
            return

        snapshot = self._neighbors_reorder_snapshot
        self._neighbors_reorder_snapshot = None
        if snapshot is not None:
            self._restore_virtual_snapshot(snapshot, keep_project_ref=True)
            laminate_names = [cell.laminate.nome for cell in self._cells]
            self._notify_changes(laminate_names, mark_dirty=False)

    def _neighbors_adjacency(self) -> dict[str, set[str]]:
        """Converte o mapeamento de vizinhos do projeto em um grafo não direcionado."""
        cell_ids = [cell.cell_id for cell in self._cells]
        adjacency: dict[str, set[str]] = {cid: set() for cid in cell_ids}
        project = self._project
        node_payload = list(getattr(project, "cell_neighbor_nodes", []) or []) if project is not None else []
        if node_payload:
            for entry in node_payload:
                src_cell = str(entry.get("cell", "")).strip()
                if not src_cell or src_cell not in adjacency:
                    continue
                neighbors = entry.get("neighbors", {}) or {}
                if not isinstance(neighbors, dict):
                    continue
                for data in neighbors.values():
                    if isinstance(data, dict):
                        dst_cell = str(data.get("cell", "") or "").strip()
                    else:
                        dst_cell = str(data or "").strip()
                    if not dst_cell or dst_cell not in adjacency:
                        continue
                    adjacency[src_cell].add(dst_cell)
                    adjacency.setdefault(dst_cell, set()).add(src_cell)
        else:
            mapping = getattr(project, "cell_neighbors", {}) if project is not None else {}
            for src, neighbors in mapping.items():
                if src not in adjacency or not isinstance(neighbors, dict):
                    continue
                for raw in (neighbors or {}).values():
                    values: list[str] = []
                    if isinstance(raw, (list, tuple, set)):
                        values = [str(v) for v in raw if v]
                    elif raw:
                        values = [str(raw)]
                    for dst in values:
                        if dst and dst in adjacency:
                            adjacency[src].add(dst)
                            adjacency.setdefault(dst, set()).add(src)
        return adjacency

    def _normalize_orientation_value(self, value: object) -> object:
        """Normaliza orientação para float quando possível, mantendo o valor bruto em caso de falha."""
        try:
            return normalize_angle(value)
        except Exception:
            try:
                return normalize_angle(str(value))
            except Exception:
                return value

    def _layer_for_cell(self, cell_layers: dict[str, list[Camada]], cell_id: str, row_idx: int) -> Camada:
        layers = cell_layers.get(cell_id, [])
        if 0 <= row_idx < len(layers):
            return copy.deepcopy(layers[row_idx])
        return Camada(
            idx=0,
            material="",
            orientacao=None,
            ativo=True,
            simetria=False,
            ply_type=DEFAULT_PLY_TYPE,
            ply_label="",
            sequence="",
            rosette=DEFAULT_ROSETTE_LABEL,
        )

    def _build_row_groups(
        self,
        row_idx: int,
        adjacency: dict[str, set[str]],
        cell_layers: dict[str, list[Camada]],
        cell_order: list[str],
        order_index: dict[str, int],
        passive_cells: set[str],
    ) -> list[dict[str, object]]:
        """Agrupa células com orientação preenchida por conectividade e orientação.

        Células sem vizinhos definidos (passivas) são ignoradas aqui para que
        permaneçam no laminado original, evitando que sejam tratadas como
        componentes isolados na reorganização.
        """
        oriented: dict[str, object] = {}
        for cid in cell_order:
            if cid in passive_cells:
                continue
            layer = self._layer_for_cell(cell_layers, cid, row_idx)
            if getattr(layer, "orientacao", None) is None:
                continue
            oriented[cid] = self._normalize_orientation_value(getattr(layer, "orientacao", None))

        visited: set[str] = set()
        groups: list[tuple[int, dict[str, object]]] = []
        for cid in cell_order:
            if cid in visited or cid not in oriented:
                continue
            component: set[str] = set()
            stack = [cid]
            visited.add(cid)
            while stack:
                current = stack.pop()
                component.add(current)
                for neighbor in adjacency.get(current, ()):  # Only consider oriented neighbors
                    if neighbor in oriented and neighbor not in visited:
                        visited.add(neighbor)
                        stack.append(neighbor)
            orientation_groups: dict[object, set[str]] = {}
            for cell_id in component:
                ori = oriented[cell_id]
                orientation_groups.setdefault(ori, set()).add(cell_id)
            for ori, cells in orientation_groups.items():
                anchor = min(order_index.get(c, len(order_index)) for c in cells)
                groups.append((anchor, {"cells": set(cells), "orientation": ori}))
        groups.sort(key=lambda entry: entry[0])
        return [item for _, item in groups]

    def _build_row_snapshot(
        self,
        row_idx: int,
        group: Optional[dict[str, object]],
        cell_layers: dict[str, list[Camada]],
        cell_order: list[str],
        passive_cells: set[str],
        *,
        use_original_when_empty: bool = False,
        include_passive: bool = False,
    ) -> dict[str, Camada]:
        """Gera uma linha virtual com base em um grupo específico ou mantém original."""
        snapshot: dict[str, Camada] = {}
        for cid in cell_order:
            base_layer = self._layer_for_cell(cell_layers, cid, row_idx)
            layer_copy = copy.deepcopy(base_layer)
            if group is None:
                if use_original_when_empty or (include_passive and cid in passive_cells):
                    snapshot[cid] = layer_copy
                else:
                    layer_copy.orientacao = None
                    snapshot[cid] = layer_copy
                continue
            if cid in group.get("cells", set()):
                layer_copy.orientacao = group.get("orientation")
            elif include_passive and cid in passive_cells:
                # Mantém a orientação original apenas na primeira linha gerada
                # para esse índice, evitando replicar camadas para células passivas.
                layer_copy.orientacao = getattr(layer_copy, "orientacao", None)
            else:
                layer_copy.orientacao = None
            snapshot[cid] = layer_copy
        return snapshot

    def _rows_for_side(
        self,
        row_idx: int,
        groups: list[dict[str, object]],
        slot_count: int,
        cell_layers: dict[str, list[Camada]],
        cell_order: list[str],
        passive_cells: set[str],
        *,
        preserve_when_empty: bool,
    ) -> list[dict[str, Camada]]:
        rows: list[dict[str, Camada]] = []
        for pos in range(slot_count):
            group = groups[pos] if pos < len(groups) else None
            keep_original = preserve_when_empty and group is None and not groups and pos == 0
            rows.append(
                self._build_row_snapshot(
                    row_idx,
                    group,
                    cell_layers,
                    cell_order,
                    passive_cells,
                    use_original_when_empty=keep_original,
                    include_passive=pos == 0,
                )
            )
        return rows

    def _apply_virtual_rows_to_laminates(self, rows: list[dict[str, Camada]]) -> None:
        if not rows:
            return
        new_layers_by_cell: dict[str, list[Camada]] = {cell.cell_id: [] for cell in self._cells}
        for idx, row in enumerate(rows):
            for cid, layer in row.items():
                layer.idx = idx
                new_layers_by_cell.setdefault(cid, []).append(layer)
        for cell in self._cells:
            layers = new_layers_by_cell.get(cell.cell_id, [])
            self._sync_sequence_and_ply_labels(layers)
            cell.laminate.camadas = layers
            model = self._stacking_model_for(cell.laminate)
            if model is not None:
                model.update_layers(copy.deepcopy(layers))
            self._after_laminate_changed(cell.laminate)

    def _sync_sequence_and_ply_labels(self, layers: list[Camada]) -> None:
        if not layers:
            return
        seq_prefix, seq_sep, seq_start = self._label_prefix_and_separator(
            layers, "sequence", "Seq"
        )
        ply_prefix, ply_sep, ply_start = self._label_prefix_and_separator(
            layers, "ply_label", "Ply"
        )
        for idx, layer in enumerate(layers):
            layer.sequence = f"{seq_prefix}{seq_sep}{seq_start + idx}"
            layer.ply_label = f"{ply_prefix}{ply_sep}{ply_start + idx}"

    def _label_prefix_and_separator(
        self,
        layers: list[Camada],
        attr: str,
        default_prefix: str,
    ) -> tuple[str, str, int]:
        for idx, layer in enumerate(layers):
            text = str(getattr(layer, attr, "") or "").strip()
            match = _PREFIX_NUMBER_PATTERN.fullmatch(text)
            if match:
                prefix = (match.group(1) or default_prefix).strip() or default_prefix
                separator = "." if "." in text else ""
                try:
                    number = int(match.group(2))
                except ValueError:
                    number = idx + 1
                start_number = max(1, number - idx)
                return prefix, separator, start_number
        return default_prefix, ".", 1

    def _split_center_row_groups(
        self,
        row: dict[str, Camada],
        adjacency: dict[str, set[str]],
        cell_order: list[str],
        order_index: dict[str, int],
        passive_cells: set[str],
    ) -> list[dict[str, Camada]]:
        oriented: dict[str, object] = {}
        for cid in cell_order:
            layer = row.get(cid)
            if layer is None or cid in passive_cells:
                continue
            orientation = getattr(layer, "orientacao", None)
            if orientation is None:
                continue
            oriented[cid] = self._normalize_orientation_value(orientation)

        if not oriented:
            return [copy.deepcopy(row)]

        visited: set[str] = set()
        groups: list[tuple[int, set[str]]] = []
        for cid in cell_order:
            if cid in visited or cid not in oriented:
                continue
            component: set[str] = set()
            stack = [cid]
            visited.add(cid)
            while stack:
                current = stack.pop()
                component.add(current)
                for neighbor in adjacency.get(current, set()):
                    if neighbor in oriented and neighbor not in visited:
                        visited.add(neighbor)
                        stack.append(neighbor)

            orientation_groups: dict[object, set[str]] = {}
            for cell_id in component:
                ori = oriented[cell_id]
                orientation_groups.setdefault(ori, set()).add(cell_id)
            for cells in orientation_groups.values():
                anchor = min(order_index.get(c, len(order_index)) for c in cells)
                groups.append((anchor, cells))

        groups.sort(key=lambda entry: entry[0])

        if len(groups) <= 1:
            normalized_row: dict[str, Camada] = {}
            for cid in cell_order:
                base = row.get(cid)
                if base is None:
                    continue
                layer_copy = copy.deepcopy(base)
                if cid in oriented:
                    layer_copy.orientacao = oriented[cid]
                elif cid in passive_cells:
                    layer_copy.orientacao = getattr(base, "orientacao", None)
                else:
                    layer_copy.orientacao = None
                normalized_row[cid] = layer_copy
            return [normalized_row]

        split_rows: list[dict[str, Camada]] = []
        for idx_group, (_, cells) in enumerate(groups):
            new_row: dict[str, Camada] = {}
            include_passive = idx_group == 0
            for cid in cell_order:
                base = row.get(cid)
                if base is None:
                    continue
                layer_copy = copy.deepcopy(base)
                if cid in cells:
                    layer_copy.orientacao = oriented[cid]
                elif include_passive and cid in passive_cells:
                    layer_copy.orientacao = getattr(base, "orientacao", None)
                else:
                    layer_copy.orientacao = None
                new_row[cid] = layer_copy
            split_rows.append(new_row)

        return split_rows

    def _process_center_sequences(
        self,
        rows: list[dict[str, Camada]],
        adjacency: dict[str, set[str]],
        cell_order: list[str],
        passive_cells: set[str],
    ) -> list[dict[str, Camada]]:
        if not rows:
            return rows

        order_index = {cid: idx for idx, cid in enumerate(cell_order)}

        if len(rows) % 2 == 0:
            center_indices = [len(rows) // 2 - 1, len(rows) // 2]
        else:
            center_indices = [len(rows) // 2]

        result = list(rows)
        offset = 0
        for center in center_indices:
            idx = center + offset
            if idx < 0 or idx >= len(result):
                continue
            split_rows = self._split_center_row_groups(
                result[idx], adjacency, cell_order, order_index, passive_cells
            )
            if len(split_rows) == 1:
                result[idx] = split_rows[0]
                continue
            result.pop(idx)
            for insert_pos, new_row in enumerate(split_rows):
                result.insert(idx + insert_pos, new_row)
            offset += len(split_rows) - 1

        return result

    def reorganizar_por_vizinhanca(self) -> bool:
        """Aplica as regras de agrupamento por vizinhança e simetria."""
        if self._project is None or not self._cells:
            return False

        adjacency = self._neighbors_adjacency()
        has_neighbors = any(neighbors for neighbors in adjacency.values())
        if not has_neighbors:
            QtWidgets.QMessageBox.information(
                self,
                "Reorder By Neighborhood",
                "No neighbor cells are registered. Define cell neighbors before reorganizing by neighborhood.",
            )
            return False

        passive_cells: set[str] = {cid for cid, neighbors in adjacency.items() if not neighbors}

        # Habilitar desfazer
        self._push_virtual_snapshot()

        # Garantir laminado único por célula para evitar efeitos colaterais
        for cell in self._cells:
            self._ensure_unique_laminate_for_cell(cell)

        cell_order = [cell.cell_id for cell in self._cells]
        order_index = {cid: idx for idx, cid in enumerate(cell_order)}
        def _run_reorder_pass(*, disable_center_symmetry: bool, progress_label: str) -> bool:
            cell_layers: dict[str, list[Camada]] = {
                cell.cell_id: copy.deepcopy(getattr(cell.laminate, "camadas", []))
                for cell in self._cells
            }

            total_rows = max((len(layers) for layers in cell_layers.values()), default=0)
            if total_rows == 0:
                return False

            loop_iterations = total_rows // 2
            progress_steps = max(loop_iterations + 3, 1)
            progress = QtWidgets.QProgressDialog(
                progress_label,
                None,
                0,
                progress_steps,
                self,
            )
            progress.setWindowTitle("Reordenação")
            progress.setWindowModality(QtCore.Qt.WindowModal)
            progress.setMinimumDuration(0)
            progress.setAutoClose(True)
            progress.setAutoReset(True)
            progress.setValue(0)
            QtWidgets.QApplication.processEvents()

            if disable_center_symmetry:
                if total_rows % 2 == 0:
                    center_indices = [total_rows // 2 - 1, total_rows // 2]
                else:
                    center_indices = [total_rows // 2]
                center_start = center_indices[0]
                center_end = center_indices[-1]
                top_limit = center_start - 1
                bottom_limit = center_end + 1
            else:
                center_indices = []
                top_limit = None
                bottom_limit = None

            top = 0
            bottom = total_rows - 1
            front_rows: list[dict[str, Camada]] = []
            back_rows: list[list[dict[str, Camada]]] = []

            current_step = 0
            while top < bottom:
                if disable_center_symmetry and (top > top_limit or bottom < bottom_limit):
                    break
                top_groups = self._build_row_groups(
                    top, adjacency, cell_layers, cell_order, order_index, passive_cells
                )
                bottom_groups = self._build_row_groups(
                    bottom, adjacency, cell_layers, cell_order, order_index, passive_cells
                )
                slot_count = max(len(top_groups), len(bottom_groups), 1)

                top_rows = self._rows_for_side(
                    top,
                    top_groups,
                    slot_count,
                    cell_layers,
                    cell_order,
                    passive_cells,
                    preserve_when_empty=not bool(top_groups),
                )
                bottom_rows = self._rows_for_side(
                    bottom,
                    bottom_groups,
                    slot_count,
                    cell_layers,
                    cell_order,
                    passive_cells,
                    preserve_when_empty=not bool(bottom_groups),
                )

                front_rows.extend(top_rows)
                back_rows.append(bottom_rows)
                top += 1
                bottom -= 1

                current_step += 1
                progress.setValue(current_step)
                QtWidgets.QApplication.processEvents()

            center_rows: list[dict[str, Camada]] = []
            if disable_center_symmetry:
                for center_idx in center_indices:
                    center_groups = self._build_row_groups(
                        center_idx,
                        adjacency,
                        cell_layers,
                        cell_order,
                        order_index,
                        passive_cells,
                    )
                    slot_count = max(len(center_groups), 1)
                    center_rows.extend(
                        self._rows_for_side(
                            center_idx,
                            center_groups,
                            slot_count,
                            cell_layers,
                            cell_order,
                            passive_cells,
                            preserve_when_empty=not bool(center_groups),
                        )
                    )
            else:
                if top == bottom:
                    # Sequência central permanece como está para preservar o eixo
                    center_rows.append(
                        self._build_row_snapshot(
                            top,
                            None,
                            cell_layers,
                            cell_order,
                            passive_cells,
                            use_original_when_empty=True,
                            include_passive=True,
                        )
                    )

            final_rows: list[dict[str, Camada]] = []
            final_rows.extend(front_rows)
            final_rows.extend(center_rows)
            for bottom_rows in reversed(back_rows):
                for row in reversed(bottom_rows):
                    final_rows.append(row)

            current_step += 1
            progress.setValue(current_step)
            QtWidgets.QApplication.processEvents()

            # Após aplicar as regras existentes, tratar as sequências centrais para separar orientações e blocos desconectados
            final_rows = self._process_center_sequences(final_rows, adjacency, cell_order, passive_cells)

            current_step += 1
            progress.setValue(current_step)
            QtWidgets.QApplication.processEvents()

            self._apply_virtual_rows_to_laminates(final_rows)
            self._notify_changes([cell.laminate.nome for cell in self._cells])

            current_step += 1
            progress.setValue(current_step)
            QtWidgets.QApplication.processEvents()
            return True

        applied = _run_reorder_pass(
            disable_center_symmetry=False,
            progress_label="Reordenação em andamento...",
        )
        if not applied:
            return False

        _run_reorder_pass(
            disable_center_symmetry=True,
            progress_label="Reordenação em andamento (2/2)...",
        )

        QtWidgets.QMessageBox.information(
            self,
            "Reorder By Neighborhood",
            "Reordenação concluída.",
        )
        return True

    # ---------------------------------------------------------------
    # Nova funcionalidade: analisar simetria
    def _analyze_symmetry(self) -> None:
        """
        Analisa a simetria de cada coluna (laminado).
        
        Nova lógica:
        - Utiliza a função evaluate_symmetry_for_layers() para determinar o eixo de simetria
        - Marca as células/camadas que pertencem ao eixo de simetria com borda verde
        - Independentemente da posição no laminado (não precisa estar no centro estrutural)
        """
        if self._project is None or not self._cells:
            return

        # Limpar todas as bordas verdes de simetria existentes
        self.model._symmetric_cells.clear()

        # Para cada coluna (cada célula/laminado)
        for cell_idx, cell in enumerate(self._cells):
            laminate = cell.laminate
            layers = getattr(laminate, "camadas", [])
            
            if not layers:
                continue
            
            # Usar a função de simetria CLT para obter a avaliação
            evaluation = evaluate_symmetry_for_layers(layers)
            
            # Se é simétrico, marcar as camadas de simetria com borda verde
            if evaluation.is_symmetric and evaluation.centers:
                column_idx = cell_idx + self.model.LAMINATE_COLUMN_OFFSET
                
                # Marcar todas as camadas que representam o eixo de simetria
                # com borda verde, independentemente da sua posição no laminado
                for center_row in evaluation.centers:
                    self.model._symmetric_cells.add((center_row, column_idx))
        
        # Atualizar a visualização para mostrar as bordas verdes
        if self.model.layers:
            top_left = self.model.createIndex(0, self.model.LAMINATE_COLUMN_OFFSET)
            bottom_right = self.model.createIndex(
                len(self.model.layers) - 1,
                self.model.columnCount() - 1
            )
            self.model.dataChanged.emit(top_left, bottom_right, [ORIENTATION_SYMMETRY_ROLE])

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:  # noqa: N802
        if hasattr(self, "btn_reorganize_neighbors"):
            if self.btn_reorganize_neighbors.isChecked() or self._neighbors_reorder_snapshot is not None:
                self._toggle_reorganizar_por_vizinhanca(False)
                blocker = QtCore.QSignalBlocker(self.btn_reorganize_neighbors)
                self.btn_reorganize_neighbors.setChecked(False)
                del blocker
        super().closeEvent(event)
        try:
            self.persist_column_order()
        except Exception:
            pass
        try:
            self.closed.emit()
        except Exception:
            pass
