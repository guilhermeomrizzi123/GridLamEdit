"""Virtual Stacking dialog and model."""

from __future__ import annotations

import copy
from collections import Counter, OrderedDict
import math
import re
from dataclasses import dataclass
from typing import Callable, Iterable, Optional

from PySide6 import QtCore, QtGui, QtWidgets

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
from gridlamedit.services.laminate_service import auto_name_for_laminate
from gridlamedit.core.paths import package_path
from gridlamedit.services.project_query import (
    project_distinct_materials,
    project_distinct_orientations,
    project_most_used_material,
)


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
        default_orientation: float = 0.0,
    ) -> None:
        super().__init__("Inserir camada")
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
        super().__init__("Remover camada")
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
    ) -> None:
        super().__init__(parent)
        self.layers: list[VirtualStackingLayer] = []
        self.cells: list[VirtualStackingCell] = []
        self.symmetry_row_index: int | None = None
        self._change_callback = change_callback
        self._stacking_model_provider = stacking_model_provider or (lambda _lam: None)
        self._post_edit_callback = post_edit_callback
        self._most_used_material_provider = most_used_material_provider or (lambda: None)
        self._red_cells: set[tuple[int, int]] = set()
        self._green_cells: set[tuple[int, int]] = set()
        self._unbalanced_columns: set[int] = set()
        self._warning_icon = QtGui.QIcon()

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

        if role != QtCore.Qt.DisplayRole:
            return None

        if orientation == QtCore.Qt.Horizontal:
            if section == self.COL_SEQUENCE:
                return "Sequence"
            if section == self.COL_PLY:
                return "Ply"
            if section == self.COL_PLY_TYPE:
                return "Simetria"
            if section == self.COL_MATERIAL:
                return "Material"
            if section == self.COL_ROSETTE:
                return "Rosette"
            cell_index = section - self.LAMINATE_COLUMN_OFFSET
            if 0 <= cell_index < len(self.cells):
                cell = self.cells[cell_index]
                laminate = cell.laminate
                lam_name = (getattr(laminate, "nome", "") or "").strip() or cell.cell_id
                tag_text = (getattr(laminate, "tag", "") or "").strip()
                display_name = lam_name or cell.cell_id
                if tag_text:
                    tag_suffix = f"({tag_text})"
                    if tag_suffix not in display_name:
                        display_name = f"{display_name}{tag_suffix}"
                label = f"#\n{cell.cell_id} | {display_name}"
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
                        if (
                            self.symmetry_row_index is not None
                            and row == self.symmetry_row_index
                        ):
                            return QtGui.QBrush(color.lighter(110))
                        return QtGui.QBrush(color)
            if (
                self.symmetry_row_index is not None
                and row == self.symmetry_row_index
            ):
                return QtGui.QBrush(QtGui.QColor(220, 235, 255))

        if role == QtCore.Qt.ForegroundRole and column >= self.LAMINATE_COLUMN_OFFSET:
            cell_index = column - self.LAMINATE_COLUMN_OFFSET
            if 0 <= cell_index < len(self.cells):
                layers = getattr(self.cells[cell_index].laminate, "camadas", [])
                if 0 <= row < len(layers) and getattr(layers[row], "orientacao", None) is None:
                    return QtGui.QBrush(QtGui.QColor(160, 160, 160))

        if role == QtCore.Qt.TextAlignmentRole and column >= self.LAMINATE_COLUMN_OFFSET:
            return QtCore.Qt.AlignCenter

        if role == ORIENTATION_SYMMETRY_ROLE and column >= self.LAMINATE_COLUMN_OFFSET:
            if (row, column) in self._green_cells:
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
        if not stacking_model.setData(target_index, normalized_value, QtCore.Qt.EditRole):
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
                        orientacao=0.0,
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
            if target_layer.orientacao is None:
                if target_layer.material:
                    idx = stacking_model.index(row, StackingTableModel.COL_MATERIAL)
                    if idx.isValid() and stacking_model.setData(idx, "", QtCore.Qt.EditRole):
                        changed.append(laminate)
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
        symmetry_row_index: int | None = None,
    ) -> None:
        self.beginResetModel()
        self.layers = list(layers)
        self.cells = list(cells)
        self.symmetry_row_index = symmetry_row_index
        self._red_cells.clear()
        self._green_cells.clear()
        self._unbalanced_columns.clear()
        self.endResetModel()

    def set_symmetry_row_index(self, index: int | None) -> None:
        if index == self.symmetry_row_index:
            return
        self.symmetry_row_index = index
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

    def clear_highlights(self) -> None:
        if not (self._red_cells or self._green_cells):
            return
        self._red_cells.clear()
        self._green_cells.clear()
        if self.layers:
            top_left = self.createIndex(0, 0)
            bottom_right = self.createIndex(len(self.layers) - 1, self.columnCount() - 1)
            self.dataChanged.emit(top_left, bottom_right, [QtCore.Qt.BackgroundRole])

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


class VirtualStackingWindow(QtWidgets.QDialog):
    """Dialog that renders the Virtual Stacking spreadsheet-like view."""

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

        self._layers: list[VirtualStackingLayer] = []
        self._cells: list[VirtualStackingCell] = []
        self._symmetry_row_index: int | None = None
        self._project: Optional[GridModel] = None
        self._sorted_cell_ids: list[str] = []
        self._initial_sort_done = False
        self._stacking_models: dict[int, StackingTableModel] = {}
        self.undo_stack = undo_stack or QtGui.QUndoStack(self)

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
        )
        self.table = QtWidgets.QTableView(self)
        self.table.setModel(self.model)
        header = WordWrapHeader(QtCore.Qt.Horizontal, self.table)
        header.setDefaultAlignment(QtCore.Qt.AlignCenter)
        self.table.setHorizontalHeader(header)
        header.setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        header.setStretchLastSection(False)
        header.sectionResized.connect(lambda *_: self._resize_summary_columns())
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectItems)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.table.setHorizontalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        self.table.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        self.table.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked
            | QtWidgets.QAbstractItemView.SelectedClicked
        )
        self.table.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self._apply_orientation_delegate()
        self.summary_table = self._build_summary_table()
        self._sync_scrollbars()

        main_layout = QtWidgets.QVBoxLayout(self)
        if toolbar_layout is not None:
            main_layout.addLayout(toolbar_layout)
        if prefix_layout is not None:
            main_layout.addLayout(prefix_layout)
        main_layout.addWidget(self.table)
        if self.summary_table is not None:
            main_layout.addWidget(self.summary_table)
        self.setLayout(main_layout)
        self.undo_stack.canUndoChanged.connect(self._update_undo_buttons)
        self.undo_stack.canRedoChanged.connect(self._update_undo_buttons)
        self.undo_stack.indexChanged.connect(lambda _idx: self._on_undo_stack_changed())

    def _build_prefix_toolbar(self) -> Optional[QtWidgets.QHBoxLayout]:
        layout = QtWidgets.QHBoxLayout()
        layout.setContentsMargins(4, 0, 4, 4)
        layout.setSpacing(8)

        seq_button = QtWidgets.QToolButton(self)
        seq_button.setText("Renomear Sequence")
        seq_button.setToolTip("Renomear prefixo de todas as sequencias")
        seq_button.clicked.connect(self._rename_all_sequences)
        layout.addWidget(seq_button)

        ply_button = QtWidgets.QToolButton(self)
        ply_button.setText("Renomear Ply")
        ply_button.setToolTip("Renomear prefixo de todos os plies")
        ply_button.clicked.connect(self._rename_all_ply)
        layout.addWidget(ply_button)

        layout.addStretch()
        return layout

    def _apply_orientation_delegate(self) -> None:
        material_delegate = MaterialComboDelegate(
            self.table,
            items_provider=lambda: project_distinct_materials(self._project),
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
        summary.setWordWrap(True)
        summary.setTextElideMode(QtCore.Qt.ElideNone)
        summary.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        summary.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        summary.setFocusPolicy(QtCore.Qt.NoFocus)
        summary.setHorizontalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        summary.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        summary.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        summary.setStyleSheet("QTableView { gridline-color: transparent; }")
        font = summary.font()
        if font.pointSize() > 0:
            font.setPointSize(max(font.pointSize() - 1, 8))
            summary.setFont(font)
        return summary

    def _sync_scrollbars(self) -> None:
        table_scroll = self.table.horizontalScrollBar()
        summary_scroll = self.summary_table.horizontalScrollBar()
        if table_scroll is None or summary_scroll is None:
            return

        def mirror(value: int, target: QtWidgets.QScrollBar) -> None:
            target.blockSignals(True)
            target.setValue(value)
            target.blockSignals(False)

        table_scroll.valueChanged.connect(lambda v: mirror(v, summary_scroll))
        summary_scroll.valueChanged.connect(lambda v: mirror(v, table_scroll))

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

    def _on_undo_stack_changed(self) -> None:
        for cell in getattr(self, "_cells", []):
            self._auto_rename_if_enabled(cell.laminate)
        self._rebuild_view()
        self._update_undo_buttons()
        self._mark_project_dirty()


    def _build_toolbar(self) -> QtWidgets.QHBoxLayout:
        layout = QtWidgets.QHBoxLayout()
        layout.setSpacing(8)

        layout.addStretch()

        self.lbl_auto_saving = QtWidgets.QLabel("Automatic Saving", self)
        layout.addWidget(self.lbl_auto_saving)

        self.btn_undo = QtWidgets.QToolButton(self)
        self.btn_undo.setText("Desfazer")
        self.btn_undo.setToolTip("Desfazer alteracao (Ctrl+Z)")
        self.btn_undo.clicked.connect(self.undo_stack.undo)
        self.btn_undo.setEnabled(self.undo_stack.canUndo())
        layout.addWidget(self.btn_undo)

        self.btn_redo = QtWidgets.QToolButton(self)
        self.btn_redo.setText("Refazer")
        self.btn_redo.setToolTip("Refazer alteracao (Ctrl+Y)")
        self.btn_redo.clicked.connect(self.undo_stack.redo)
        self.btn_redo.setEnabled(self.undo_stack.canRedo())
        layout.addWidget(self.btn_redo)

        return layout

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
            self.model.set_virtual_stacking([], [], None)
            self._stacking_models.clear()
            self._update_summary_row()
            return

        layers, cells, symmetry_index = self._collect_virtual_data(self._project)
        self._layers = layers
        self._cells = cells
        self._symmetry_row_index = symmetry_index
        self._refresh_stacking_models(cells)
        self.model.set_virtual_stacking(layers, cells, symmetry_index)
        self._apply_orientation_delegate()
        self._resize_columns()
        self._update_summary_row()
        self._check_symmetry()

    def _collect_virtual_data(
        self, project: GridModel
    ) -> tuple[list[VirtualStackingLayer], list[VirtualStackingCell], int | None]:
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

        symmetry_index = self._detect_symmetry_row(laminates, max_layers)
        return layers, cells, symmetry_index

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

    def _current_sequence_prefix(self) -> str:
        return self._current_prefix_from_project("sequence", "Seq")

    def _current_ply_prefix(self) -> str:
        return self._current_prefix_from_project("ply_label", "Ply")

    def _most_used_material(self) -> Optional[str]:
        """Retorna o material mais utilizado em todos os laminados carregados."""
        return project_most_used_material(self._project)

    def _prompt_prefix_dialog(self, title: str, current_prefix: str) -> Optional[str]:
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle(title)
        layout = QtWidgets.QVBoxLayout(dialog)
        layout.addWidget(QtWidgets.QLabel(f"Prefixo atual: {current_prefix}", dialog))
        input_box = QtWidgets.QLineEdit(dialog)
        input_box.setText(current_prefix)
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
    ) -> list[str]:
        if self._project is None:
            return []
        prefix = new_prefix.strip().rstrip(".")
        if not prefix:
            return []
        changed: list[str] = []
        for laminate in getattr(self._project, "laminados", {}).values():
            model = self._stacking_model_for(laminate)
            if model is None:
                continue
            layers = model.layers()
            updated = False
            for idx, layer in enumerate(layers):
                current_label = getattr(layer, target_attr, "")
                separator = "." if "." in str(current_label) else ""
                target_label = f"{prefix}{separator}{idx + 1}"
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
        current_prefix = self._current_sequence_prefix()
        new_prefix = self._prompt_prefix_dialog("Renomear Sequence", current_prefix)
        if new_prefix is None:
            return
        cleaned = new_prefix.strip().rstrip(".")
        if not cleaned or cleaned == current_prefix.rstrip("."):
            return
        self._apply_global_prefix_change(
            "sequence",
            StackingTableModel.COL_SEQUENCE,
            cleaned,
        )

    def _rename_all_ply(self) -> None:
        current_prefix = self._current_ply_prefix()
        new_prefix = self._prompt_prefix_dialog("Renomear Ply", current_prefix)
        if new_prefix is None:
            return
        cleaned = new_prefix.strip().rstrip(".")
        if not cleaned or cleaned == current_prefix.rstrip("."):
            return
        self._apply_global_prefix_change(
            "ply_label",
            StackingTableModel.COL_PLY,
            cleaned,
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

    def _row_considered(self, row: int) -> bool:
        if not (0 <= row < len(self._layers)):
            return True
        layer = self._layers[row]
        return is_structural_ply_label(getattr(layer, "ply_type", DEFAULT_PLY_TYPE))

    def _summary_for_laminate(self, laminate: Optional[Laminado]) -> str:
        if laminate is None:
            return ""
        layers = getattr(laminate, "camadas", [])
        orientations = Counter()
        oriented_count = 0
        for camada in layers:
            row_idx = getattr(camada, "idx", None)
            if row_idx is not None and not self._row_considered(row_idx):
                continue
            if not is_structural_ply_label(getattr(camada, "ply_type", DEFAULT_PLY_TYPE)):
                continue
            token = self._orientation_token(getattr(camada, "orientacao", None))
            if token is None:
                continue
            oriented_count += 1
            label = format_orientation_value(token)
            orientations[label] += 1
        parts = [f"Total: {oriented_count}"]
        if orientations:
            ordered = sorted(
                orientations.items(),
                key=lambda pair: self._orientation_token(pair[0]) or 0.0,
            )
            ori_parts = [f"{label} [{count}]" for label, count in ordered]
            parts.extend(ori_parts)
        return "\n".join(parts)

    def _update_summary_row(self) -> None:
        summary_model = (
            self.summary_table.model()
            if hasattr(self, "summary_table")
            else None
        )
        if not isinstance(summary_model, QtGui.QStandardItemModel):
            return
        column_count = self.model.columnCount()
        summary_model.setColumnCount(column_count)
        metrics = (
            self.summary_table.fontMetrics()
            if hasattr(self, "summary_table")
            else self.fontMetrics()
        )
        max_lines = 1
        for col in range(column_count):
            if col < self.model.LAMINATE_COLUMN_OFFSET:
                text = ""
            else:
                cell_idx = col - self.model.LAMINATE_COLUMN_OFFSET
                laminate = self._cells[cell_idx].laminate if 0 <= cell_idx < len(self._cells) else None
                text = self._summary_for_laminate(laminate)
            item = QtGui.QStandardItem(text)
            if col < self.model.LAMINATE_COLUMN_OFFSET:
                alignment = QtCore.Qt.AlignCenter
            else:
                alignment = QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter
            item.setTextAlignment(alignment)
            summary_model.setItem(0, col, item)
            max_lines = max(max_lines, max(1, text.count("\n") + 1))
        self._update_summary_height(max_lines, metrics)
        self._resize_summary_columns()

    def _update_summary_height(self, lines: int, metrics: QtGui.QFontMetrics) -> None:
        if not hasattr(self, "summary_table"):
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
            if hasattr(self, "summary_table")
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
        if self.model.columnCount() >= 1:
            header.resizeSection(VirtualStackingModel.COL_SEQUENCE, 150)
        if self.model.columnCount() >= 2:
            header.resizeSection(VirtualStackingModel.COL_PLY, 90)
        if self.model.columnCount() >= 3:
            header.resizeSection(VirtualStackingModel.COL_PLY_TYPE, 150)
        if self.model.columnCount() >= 4:
            header.resizeSection(VirtualStackingModel.COL_MATERIAL, 160)
        if self.model.columnCount() >= 5:
            header.resizeSection(VirtualStackingModel.COL_ROSETTE, 140)
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

    def _check_symmetry(self) -> None:
        red_cells: set[tuple[int, int]] = set()
        green_cells: set[tuple[int, int]] = set()
        unbalanced_columns: set[int] = set()
        for col, cell in enumerate(self._cells):
            laminate = cell.laminate
            layers = getattr(laminate, "camadas", [])
            structural_rows = [
                idx
                for idx, camada in enumerate(layers)
                if self._row_considered(idx)
                and camada.orientacao is not None
                and is_structural_ply_label(getattr(camada, "ply_type", DEFAULT_PLY_TYPE))
            ]
            count_struct = len(structural_rows)
            if count_struct == 0:
                continue
            center_rows: list[int] = []
            is_symmetrical = False
            if count_struct == 1:
                center_rows = [structural_rows[0]]
                green_cells.add((structural_rows[0], col + self.model.LAMINATE_COLUMN_OFFSET))
                is_symmetrical = True
            else:
                i, j = 0, count_struct - 1
                broken = False
                while i < j:
                    r_top = structural_rows[i]
                    r_bot = structural_rows[j]
                    camada_top = layers[r_top]
                    camada_bot = layers[r_bot]
                    mat_top = (getattr(camada_top, "material", "") or "").strip().lower()
                    mat_bot = (getattr(camada_bot, "material", "") or "").strip().lower()
                    ori_top = self._orientation_token(getattr(camada_top, "orientacao", None))
                    ori_bot = self._orientation_token(getattr(camada_bot, "orientacao", None))
                    if not (mat_top == mat_bot and self._orientations_match(ori_top, ori_bot)):
                        red_cells.update(
                            {
                                (r_top, col + self.model.LAMINATE_COLUMN_OFFSET),
                                (r_bot, col + self.model.LAMINATE_COLUMN_OFFSET),
                            }
                        )
                        broken = True
                        break
                    i += 1
                    j -= 1
                if not broken:
                    is_symmetrical = True
                    if count_struct % 2 == 1:
                        center = structural_rows[count_struct // 2]
                        center_rows = [center]
                        green_cells.add((center, col + self.model.LAMINATE_COLUMN_OFFSET))
                    else:
                        center_rows = [
                            structural_rows[count_struct // 2 - 1],
                            structural_rows[count_struct // 2],
                        ]
                        green_cells.update(
                            {
                                (
                                    structural_rows[count_struct // 2 - 1],
                                    col + self.model.LAMINATE_COLUMN_OFFSET,
                                ),
                                (
                                    structural_rows[count_struct // 2],
                                    col + self.model.LAMINATE_COLUMN_OFFSET,
                                ),
                            }
                        )
            if is_symmetrical and center_rows and self._is_unbalanced(layers, structural_rows, center_rows):
                unbalanced_columns.add(col + self.model.LAMINATE_COLUMN_OFFSET)
        self.model.set_highlights(red_cells, green_cells)
        self.model.set_unbalanced_columns(unbalanced_columns)

    def _is_unbalanced(
        self,
        layers: list[Camada],
        structural_rows: list[int],
        center_rows: list[int],
    ) -> bool:
        if not structural_rows or not center_rows:
            return False
        center_min = min(center_rows)
        center_set = set(center_rows)
        pos45 = 0
        neg45 = 0
        for row in structural_rows:
            if row in center_set:
                continue
            if row > center_min:
                # Only consider layers above the symmetry plane.
                continue
            orientation = self._orientation_token(getattr(layers[row], "orientacao", None))
            if orientation is None:
                continue
            if math.isclose(orientation, 45.0, abs_tol=1e-6):
                pos45 += 1
            elif math.isclose(orientation, -45.0, abs_tol=1e-6):
                neg45 += 1
        return pos45 != neg45

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
        dialog.setWindowTitle("Outro valor")
        dialog.setLabelText("Informe a orientacao (-100 a 100 graus):")
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
                "Orientacao inexistente",
                "A linha selecionada nao existe para este laminado.",
            )
            return
        options = ["Empty", "0", "45", "-45", "90", "Outro valor..."]
        selected, ok = QtWidgets.QInputDialog.getItem(
            self,
            "Editar orientacao",
            "Selecione a orientacao:",
            options,
            0,
            False,
        )
        if not ok:
            return
        selected = selected.strip()
        if selected == "Outro valor...":
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
                "Orientacao invalida",
                "Use valores entre -100 e 100 graus ou deixe em branco.",
            )

    def _clear_orientation_at(self, index: QtCore.QModelIndex) -> None:
        if not index.isValid():
            return
        if not self.model.setData(index, "", QtCore.Qt.EditRole):
            QtWidgets.QMessageBox.information(
                self,
                "Limpar orientacao",
                "Nao foi possivel limpar esta celula. Verifique se a linha existe.",
            )

    def _remove_layer_at(self, index: QtCore.QModelIndex) -> None:
        entry = self._stacking_entry_for_index(index)
        if entry is None:
            return
        laminate, stacking_model, row = entry
        if not (0 <= row < stacking_model.rowCount()):
            QtWidgets.QMessageBox.information(
                self,
                "Remover camada",
                "Selecione uma linha existente para remover.",
            )
            return
        command = _RemoveLayerCommand(stacking_model, laminate, [row])
        self._execute_command(command)
        self._after_laminate_changed(laminate)
        self._notify_changes([laminate.nome])

    def _insert_layer(self, index: Optional[QtCore.QModelIndex] = None) -> None:
        # If a specific cell was clicked, handle that single insertion directly.
        if index is not None and index.isValid() and index.column() >= self.model.LAMINATE_COLUMN_OFFSET:
            entry = self._stacking_entry_for_index(index)
            if entry is None:
                return
            laminate, stacking_model, row = entry
            insert_at = max(0, min(row, stacking_model.rowCount()))
            default_material = self._most_used_material() or ""
            command = _InsertLayerCommand(
                stacking_model,
                laminate,
                [insert_at],
                default_material=default_material,
            )
            self._execute_command(command)
            laminate.camadas = stacking_model.layers()
            try:
                self.model._auto_fill_material_if_missing(laminate, stacking_model, insert_at)
            except Exception:
                pass
            self._after_laminate_changed(laminate)
            self._notify_changes([laminate.nome])
            return

        targets = self._targets_for_insertion(index)
        if not targets:
            QtWidgets.QMessageBox.information(
                self,
                "Inserir camada",
                "Selecione pelo menos uma linha para inserir uma nova camada.",
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

    def _notify_changes(self, laminate_names: list[str]) -> None:
        if self._project is not None:
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
        if not index.isValid() or index.column() < self.model.LAMINATE_COLUMN_OFFSET:
            return
        self.table.setCurrentIndex(index)
        menu = QtWidgets.QMenu(self)
        above_action = menu.addAction("Adicionar camada")
        menu.addSeparator()
        edit_action = menu.addAction("Editar orientacao...")
        clear_action = menu.addAction("Limpar orientacao")
        remove_action = menu.addAction("Remover camada")
        chosen = menu.exec(self.table.viewport().mapToGlobal(pos))
        if chosen == above_action:
            self._insert_layer(index=index)
        elif chosen == edit_action:
            self._edit_orientation_at(index)
        elif chosen == clear_action:
            self._clear_orientation_at(index)
        elif chosen == remove_action:
            self._remove_layer_at(index)

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

    def _detect_symmetry_row(
        self, laminates: Iterable[Laminado], layer_count: int
    ) -> int | None:
        for idx in range(layer_count):
            for lam in laminates:
                if (
                    idx < len(lam.camadas)
                    and getattr(lam.camadas[idx], "orientacao", None) is not None
                    and getattr(lam.camadas[idx], "simetria", False)
                ):
                    return idx
        return None

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

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:  # noqa: N802
        super().closeEvent(event)
        try:
            self.closed.emit()
        except Exception:
            pass
