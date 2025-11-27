"""Virtual Stacking dialog and model."""

from __future__ import annotations

import copy
from collections import Counter, OrderedDict
import math
from dataclasses import dataclass
from typing import Callable, Iterable, Optional

from PySide6 import QtCore, QtGui, QtWidgets

from gridlamedit.io.spreadsheet import (
    Camada,
    DEFAULT_PLY_TYPE,
    GridModel,
    Laminado,
    StackingTableModel,
    WordWrapHeader,
    format_orientation_value,
    is_structural_ply_label,
    normalize_angle,
)
from gridlamedit.app.delegates import OrientationComboDelegate
from gridlamedit.services.laminate_service import auto_name_for_laminate
from gridlamedit.services.project_query import project_distinct_orientations


@dataclass
class VirtualStackingLayer:
    """Represent a stacking layer sequence entry."""

    sequence_label: str
    material: str


@dataclass
class VirtualStackingCell:
    """Represent orientations of a laminado for each grid cell."""

    cell_id: str
    laminate: Laminado


class _InsertLayerCommand(QtGui.QUndoCommand):
    """Undoable insertion of new layers for a laminate."""

    def __init__(
        self,
        model: StackingTableModel,
        laminate: Laminado,
        positions: list[int],
    ) -> None:
        super().__init__("Inserir camada")
        self._model = model
        self._laminate = laminate
        self._positions = positions
        self._inserted: list[int] = []

    def redo(self) -> None:
        self._inserted = []
        for pos in sorted(set(self._positions), reverse=True):
            target_pos = max(0, min(pos, len(self._model.layers())))
            self._model.insert_layer(
                target_pos,
                Camada(
                    idx=0,
                    material="",
                    orientacao=None,
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

    def __init__(
        self,
        parent: Optional[QtCore.QObject] = None,
        change_callback: Optional[Callable[[list[str]], None]] = None,
        stacking_model_provider: Optional[
            Callable[[Laminado], Optional[StackingTableModel]]
        ] = None,
        post_edit_callback: Optional[Callable[[Laminado], None]] = None,
    ) -> None:
        super().__init__(parent)
        self.layers: list[VirtualStackingLayer] = []
        self.cells: list[VirtualStackingCell] = []
        self.symmetry_row_index: int | None = None
        self._change_callback = change_callback
        self._stacking_model_provider = stacking_model_provider or (lambda _lam: None)
        self._post_edit_callback = post_edit_callback
        self._red_cells: set[tuple[int, int]] = set()
        self._green_cells: set[tuple[int, int]] = set()

    # Basic structure -------------------------------------------------
    def rowCount(self, parent: QtCore.QModelIndex = QtCore.QModelIndex()) -> int:  # noqa: N802
        if parent.isValid():
            return 0
        return len(self.layers)

    def columnCount(self, parent: QtCore.QModelIndex = QtCore.QModelIndex()) -> int:  # noqa: N802
        if parent.isValid():
            return 0
        return 2 + len(self.cells)

    # Header ----------------------------------------------------------
    def headerData(  # noqa: N802
        self,
        section: int,
        orientation: QtCore.Qt.Orientation,
        role: int = QtCore.Qt.DisplayRole,
    ):
        if role != QtCore.Qt.DisplayRole:
            return None

        if orientation == QtCore.Qt.Horizontal:
            if section == 0:
                return "Sequence\nVirtual Sequence"
            if section == 1:
                return "#\nMaterial"
            cell_index = section - 2
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
            if column == 0:
                return layer.sequence_label or f"Seq.{row + 1}"
            if column == 1:
                return layer.material

            cell_index = column - 2
            if 0 <= cell_index < len(self.cells):
                cell = self.cells[cell_index]
                layers = getattr(cell.laminate, "camadas", [])
                if 0 <= row < len(layers):
                    value = getattr(layers[row], "orientacao", None)
                    if value is None:
                        return ""
                    if role == QtCore.Qt.DisplayRole:
                        return format_orientation_value(value)
                    return f"{value:g}"
                return ""

        if role == QtCore.Qt.BackgroundRole:
            if column in (0, 1):
                return QtGui.QBrush(QtGui.QColor(240, 240, 240))
            if (row, column) in self._red_cells:
                return QtGui.QBrush(QtGui.QColor(220, 53, 69))
            if (row, column) in self._green_cells:
                return QtGui.QBrush(QtGui.QColor(40, 167, 69))
            if (
                self.symmetry_row_index is not None
                and row == self.symmetry_row_index
            ):
                return QtGui.QBrush(QtGui.QColor(220, 235, 255))

        if role == QtCore.Qt.TextAlignmentRole and column >= 2:
            return QtCore.Qt.AlignCenter

        return None

    # Editing ---------------------------------------------------------
    def flags(self, index: QtCore.QModelIndex) -> QtCore.Qt.ItemFlags:  # noqa: N802
        if not index.isValid():
            return QtCore.Qt.NoItemFlags

        base = QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled
        if index.column() >= 2:
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
        if row < 0 or row >= len(self.layers) or column < 2:
            return False

        cell_index = column - 2
        if cell_index < 0 or cell_index >= len(self.cells):
            return False

        cell = self.cells[cell_index]
        laminate = cell.laminate
        stacking_model = self._stacking_model_provider(laminate)
        if stacking_model is None:
            return False
        layers = getattr(laminate, "camadas", [])
        if row >= len(layers):
            # If the laminate has fewer layers, create placeholder rows up to the target.
            missing = row - len(layers) + 1
            for _ in range(missing):
                stacking_model.insert_layer(len(layers), Camada(
                    idx=0,
                    material="",
                    orientacao=None,
                    ativo=True,
                    simetria=False,
                    ply_type=DEFAULT_PLY_TYPE,
                ))
                layers = stacking_model.layers()

        text = str(value).strip()
        if text.endswith("?"):
            text = text[:-1].strip()

        target_index = stacking_model.index(row, StackingTableModel.COL_ORIENTATION)
        if not target_index.isValid():
            return False
        # Permit blank values; delegate validation to StackingTableModel.
        if not stacking_model.setData(target_index, text, QtCore.Qt.EditRole):
            return False
        laminate.camadas = stacking_model.layers()
        self.dataChanged.emit(index, index, [QtCore.Qt.DisplayRole, QtCore.Qt.EditRole])
        if self._change_callback:
            try:
                self._change_callback([laminate.nome])
            except Exception:
                pass
        if self._post_edit_callback:
            try:
                self._post_edit_callback(laminate)
            except Exception:
                pass
        return True


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
        self.endResetModel()

    def set_symmetry_row_index(self, index: int | None) -> None:
        if index == self.symmetry_row_index:
            return
        self.symmetry_row_index = index
        if self.layers:
            top_left = self.createIndex(0, 0)
            bottom_right = self.createIndex(len(self.layers) - 1, self.columnCount() - 1)
            self.dataChanged.emit(top_left, bottom_right, [QtCore.Qt.BackgroundRole])

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
        self.resize(1200, 700)

        self._layers: list[VirtualStackingLayer] = []
        self._cells: list[VirtualStackingCell] = []
        self._symmetry_row_index: int | None = None
        self._project: Optional[GridModel] = None
        self._stacking_models: dict[int, StackingTableModel] = {}
        self.undo_stack = undo_stack or QtGui.QUndoStack(self)

        self._build_ui()


    def _build_ui(self) -> None:
        toolbar_layout = self._build_toolbar()

        self.model = VirtualStackingModel(
            self,
            change_callback=self._on_model_change,
            stacking_model_provider=self._stacking_model_for,
            post_edit_callback=self._after_laminate_changed,
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
        main_layout.addWidget(self.table)
        if self.summary_table is not None:
            main_layout.addWidget(self.summary_table)
        self.setLayout(main_layout)
        self.undo_stack.canUndoChanged.connect(self._update_undo_buttons)
        self.undo_stack.canRedoChanged.connect(self._update_undo_buttons)
        self.undo_stack.indexChanged.connect(lambda _idx: self._on_undo_stack_changed())

    def _apply_orientation_delegate(self) -> None:
        delegate = OrientationComboDelegate(
            self.table,
            items_provider=lambda: project_distinct_orientations(self._project),
        )
        for col in range(2, self.model.columnCount()):
            self.table.setItemDelegateForColumn(col, delegate)

    def _build_summary_table(self) -> QtWidgets.QTableView:
        summary = QtWidgets.QTableView(self)
        summary.setModel(QtGui.QStandardItemModel(1, 0, summary))
        summary.verticalHeader().setVisible(False)
        summary.horizontalHeader().setVisible(False)
        summary.setWordWrap(True)
        summary.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        summary.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        summary.setFocusPolicy(QtCore.Qt.NoFocus)
        summary.setHorizontalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        summary.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        summary.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        summary.setStyleSheet("QTableView { gridline-color: transparent; }")
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
        self._project = project
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
            cells.append(VirtualStackingCell(cell_id=cell_id, laminate=laminate))

        max_layers = max((len(lam.camadas) for lam in laminates), default=0)
        layers: list[VirtualStackingLayer] = []
        for idx in range(max_layers):
            material = ""
            sequence_label = ""
            for lam in laminates:
                if idx >= len(lam.camadas):
                    continue
                camada = lam.camadas[idx]
                if not material and getattr(camada, "material", ""):
                    material = camada.material
                if not sequence_label and getattr(camada, "sequence", ""):
                    sequence_label = camada.sequence
                if material and sequence_label:
                    break
            if not sequence_label:
                sequence_label = f"Seq.{idx + 1}"
            layers.append(
                VirtualStackingLayer(
                    sequence_label=sequence_label,
                    material=material,
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

    def _auto_rename_if_enabled(self, laminate: Laminado) -> None:
        if (
            self._project is None
            or not getattr(laminate, "auto_rename_enabled", False)
        ):
            return
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
            if index.column() < 2:
                continue
            cell_idx = index.column() - 2
            if 0 <= cell_idx < len(self._cells):
                laminate = self._cells[cell_idx].laminate
                targets.setdefault(laminate, set()).add(index.row())
        if not targets and not selected_rows:
            current = self.table.currentIndex()
            if current.isValid():
                selected_rows.add(current.row())
                if current.column() >= 2:
                    cell_idx = current.column() - 2
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

    def _summary_for_laminate(self, laminate: Optional[Laminado]) -> str:
        if laminate is None:
            return ""
        layers = getattr(laminate, "camadas", [])
        orientations = Counter()
        for camada in layers:
            token = self._orientation_token(getattr(camada, "orientacao", None))
            if token is None:
                continue
            label = format_orientation_value(token)
            orientations[label] += 1
        parts = [f"Total: {len(layers)}"]
        if orientations:
            ordered = sorted(
                orientations.items(),
                key=lambda pair: self._orientation_token(pair[0]) or 0.0,
            )
            ori_parts = [f"{label}: {count}" for label, count in ordered]
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
        metrics = self.fontMetrics()
        max_lines = 1
        for col in range(column_count):
            if col < 2:
                text = ""
            else:
                cell_idx = col - 2
                laminate = self._cells[cell_idx].laminate if 0 <= cell_idx < len(self._cells) else None
                text = self._summary_for_laminate(laminate)
            item = QtGui.QStandardItem(text)
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            if col in (0, 1):
                item.setBackground(QtGui.QBrush(QtGui.QColor(240, 240, 240)))
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
            header.resizeSection(0, 170)
        if self.model.columnCount() >= 2:
            header.resizeSection(1, 140)
        for col in range(2, self.model.columnCount()):
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
        for col, cell in enumerate(self._cells):
            laminate = cell.laminate
            layers = getattr(laminate, "camadas", [])
            structural_rows = [
                idx for idx, camada in enumerate(layers) if is_structural_ply_label(getattr(camada, "ply_type", DEFAULT_PLY_TYPE))
            ]
            count_struct = len(structural_rows)
            if count_struct == 0:
                continue
            if count_struct == 1:
                green_cells.add((structural_rows[0], col + 2))
                continue
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
                    red_cells.update({(r_top, col + 2), (r_bot, col + 2)})
                    broken = True
                    break
                i += 1
                j -= 1
            if broken:
                continue
            if count_struct % 2 == 1:
                green_cells.add((structural_rows[count_struct // 2], col + 2))
            else:
                green_cells.update(
                    {
                        (structural_rows[count_struct // 2 - 1], col + 2),
                        (structural_rows[count_struct // 2], col + 2),
                    }
                )
        self.model.set_highlights(red_cells, green_cells)

    def _targets_for_insertion(
        self, index: Optional[QtCore.QModelIndex] = None
    ) -> dict[Laminado, set[int]]:
        if index is not None and index.isValid() and index.column() >= 2:
            cell_idx = index.column() - 2
            if 0 <= cell_idx < len(self._cells):
                laminate = self._cells[cell_idx].laminate
                return {laminate: {index.row()}}
        return self._selected_targets()

    def _stacking_entry_for_index(
        self, index: QtCore.QModelIndex
    ) -> Optional[tuple[Laminado, StackingTableModel, int]]:
        if not index.isValid() or index.column() < 2:
            return None
        cell_idx = index.column() - 2
        if not (0 <= cell_idx < len(self._cells)):
            return None
        laminate = self._cells[cell_idx].laminate
        model = self._stacking_model_for(laminate)
        if model is None:
            return None
        return laminate, model, index.row()

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
        current_index = stacking_model.index(row, StackingTableModel.COL_ORIENTATION)
        current_layers = stacking_model.layers()
        current_value = None
        if 0 <= row < len(current_layers):
            current_value = getattr(current_layers[row], "orientacao", None)
        current_text = "" if current_value is None else f"{current_value:g}"

        dialog = QtWidgets.QInputDialog(self)
        dialog.setInputMode(QtWidgets.QInputDialog.TextInput)
        dialog.setWindowTitle("Editar orientacao")
        dialog.setLabelText("Defina a orientacao (-100 a 100 graus ou vazio):")
        dialog.setTextValue(str(current_text))
        editor = dialog.findChild(QtWidgets.QLineEdit)
        if editor is not None:
            editor.setValidator(None)  # Validation via normalize_angle below.
            editor.setPlaceholderText("Ex.: -45, 0, 12.5 ou deixe em branco")
        if dialog.exec() != QtWidgets.QDialog.Accepted:
            return
        new_value = dialog.textValue().strip()
        normalized_value = ""
        if new_value:
            try:
                normalized_value = f"{normalize_angle(new_value):g}"
            except ValueError:
                QtWidgets.QMessageBox.warning(
                    self,
                    "Orientacao invalida",
                    "Use valores entre -100 e 100 graus ou deixe em branco.",
                )
                return
        if not normalized_value and current_value is None:
            return
        if normalized_value and current_value is not None:
            try:
                if math.isclose(current_value, float(normalized_value), rel_tol=0.0, abs_tol=1e-9):
                    return
            except Exception:
                pass
        if not self.model.setData(index, normalized_value, QtCore.Qt.EditRole):
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
        if index is not None and index.isValid() and index.column() >= 2:
            entry = self._stacking_entry_for_index(index)
            if entry is None:
                return
            laminate, stacking_model, row = entry
            insert_at = max(0, min(row, stacking_model.rowCount()))
            command = _InsertLayerCommand(stacking_model, laminate, [insert_at])
            self._execute_command(command)
            laminate.camadas = stacking_model.layers()
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
                command = _InsertLayerCommand(model, laminate, positions)
                self._execute_command(command)
                laminate.camadas = model.layers()
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
        if not index.isValid() or index.column() < 2:
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
                if idx < len(lam.camadas) and getattr(lam.camadas[idx], "simetria", False):
                    return idx
        return None
