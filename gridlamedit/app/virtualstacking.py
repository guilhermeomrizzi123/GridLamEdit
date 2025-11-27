"""Virtual Stacking dialog and model."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
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
    normalize_angle,
)
from gridlamedit.app.dialogs.bulk_orientation_dialog import BulkOrientationDialog
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


class VirtualStackingModel(QtCore.QAbstractTableModel):
    """Qt model responsible for exposing Virtual Stacking data."""

    def __init__(
        self,
        parent: Optional[QtCore.QObject] = None,
        change_callback: Optional[Callable[[list[str]], None]] = None,
    ) -> None:
        super().__init__(parent)
        self.layers: list[VirtualStackingLayer] = []
        self.cells: list[VirtualStackingCell] = []
        self.symmetry_row_index: int | None = None
        self._change_callback = change_callback

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
                lam_name = getattr(laminate, "nome", "") or cell.cell_id
                tag_text = (getattr(laminate, "tag", "") or "").strip()
                label = f"#\n{cell.cell_id} | {lam_name}"
                if tag_text:
                    label = f"{label} ({tag_text})"
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
        layers = getattr(laminate, "camadas", [])
        if row >= len(layers):
            return False

        text = str(value).strip()
        if text.endswith("?"):
            text = text[:-1].strip()

        if not text:
            angle: Optional[float] = None
        else:
            try:
                angle = normalize_angle(text.replace(",", "."))
            except ValueError:
                return False

        model = StackingTableModel(camadas=layers)
        if not model.apply_field_value(row, StackingTableModel.COL_ORIENTATION, angle):
            return False
        laminate.camadas = model.layers()
        self.dataChanged.emit(index, index, [QtCore.Qt.DisplayRole, QtCore.Qt.EditRole])
        if self._change_callback:
            try:
                self._change_callback([laminate.nome])
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
        self.endResetModel()

    def set_symmetry_row_index(self, index: int | None) -> None:
        if index == self.symmetry_row_index:
            return
        self.symmetry_row_index = index
        if self.layers:
            top_left = self.createIndex(0, 0)
            bottom_right = self.createIndex(len(self.layers) - 1, self.columnCount() - 1)
            self.dataChanged.emit(top_left, bottom_right, [QtCore.Qt.BackgroundRole])


class VirtualStackingWindow(QtWidgets.QDialog):
    """Dialog that renders the Virtual Stacking spreadsheet-like view."""

    stacking_changed = QtCore.Signal(list)

    def __init__(self, parent: Optional[QtWidgets.QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Virtual Stacking - View Mode: Cells")
        self.resize(1200, 700)

        self._layers: list[VirtualStackingLayer] = []
        self._cells: list[VirtualStackingCell] = []
        self._symmetry_row_index: int | None = None
        self._project: Optional[GridModel] = None

        self._build_ui()


    def _build_ui(self) -> None:
        self.grid_label = QtWidgets.QLabel("Grid: [indefinido]", self)
        self.sheet_label = QtWidgets.QLabel("Folha: [nao definida]", self)
        self.range_label = QtWidgets.QLabel("Intervalo de celulas: [nao definido]", self)

        top_layout = QtWidgets.QHBoxLayout()
        top_layout.addWidget(self.grid_label)
        top_layout.addWidget(self.sheet_label)
        top_layout.addWidget(self.range_label)
        top_layout.addStretch()

        toolbar_layout = self._build_toolbar()

        self.model = VirtualStackingModel(self, change_callback=self._on_model_change)
        self.table = QtWidgets.QTableView(self)
        self.table.setModel(self.model)
        header = WordWrapHeader(QtCore.Qt.Horizontal, self.table)
        header.setDefaultAlignment(QtCore.Qt.AlignCenter)
        self.table.setHorizontalHeader(header)
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectItems)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked
            | QtWidgets.QAbstractItemView.SelectedClicked
        )

        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.addLayout(top_layout)
        if toolbar_layout is not None:
            main_layout.addLayout(toolbar_layout)
        main_layout.addWidget(self.table)
        self.setLayout(main_layout)


    def _build_toolbar(self) -> QtWidgets.QHBoxLayout:
        layout = QtWidgets.QHBoxLayout()
        layout.setSpacing(8)

        self.btn_change_orientation = QtWidgets.QToolButton(self)
        self.btn_change_orientation.setText("Trocar orientacao")
        self.btn_change_orientation.setToolTip("Trocar orientacao das celulas selecionadas")
        self.btn_change_orientation.clicked.connect(self._change_orientation)
        layout.addWidget(self.btn_change_orientation)

        self.btn_insert_above = QtWidgets.QToolButton(self)
        self.btn_insert_above.setText("Inserir acima")
        self.btn_insert_above.setToolTip("Inserir linha acima da selecao")
        self.btn_insert_above.clicked.connect(lambda: self._insert_layer(below=False))
        layout.addWidget(self.btn_insert_above)

        self.btn_insert_below = QtWidgets.QToolButton(self)
        self.btn_insert_below.setText("Inserir abaixo")
        self.btn_insert_below.setToolTip("Inserir linha abaixo da selecao")
        self.btn_insert_below.clicked.connect(lambda: self._insert_layer(below=True))
        layout.addWidget(self.btn_insert_below)

        self.btn_clear_layer = QtWidgets.QToolButton(self)
        self.btn_clear_layer.setText("Limpar camada")
        self.btn_clear_layer.setToolTip("Limpar material e orientacao das camadas selecionadas")
        self.btn_clear_layer.clicked.connect(self._clear_selected_layers)
        layout.addWidget(self.btn_clear_layer)

        layout.addStretch()

        self.btn_toggle_maximize = QtWidgets.QToolButton(self)
        self.btn_toggle_maximize.setText("Maximizar")
        self.btn_toggle_maximize.setToolTip("Alternar visualizacao maximizada")
        self.btn_toggle_maximize.setCheckable(True)
        self.btn_toggle_maximize.clicked.connect(self._toggle_maximize)
        layout.addWidget(self.btn_toggle_maximize)

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
            self.grid_label.setText("Grid: [indefinido]")
            self.sheet_label.setText("Folha: [nao definida]")
            self.range_label.setText("Intervalo de celulas: [indefinido]")
            return

        layers, cells, symmetry_index = self._collect_virtual_data(self._project)
        self._layers = layers
        self._cells = cells
        self._symmetry_row_index = symmetry_index
        self.model.set_virtual_stacking(layers, cells, symmetry_index)
        self._update_labels(cells)

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

    def _update_labels(self, cells: list[VirtualStackingCell]) -> None:
        grid_name = self._project.source_excel_path if self._project is not None else None
        display_name = grid_name or "[indefinido]"
        if display_name not in ("[indefinido]", ""):
            display_name = Path(display_name).name
        self.grid_label.setText(f"Grid: {display_name or '[indefinido]'}")
        self.sheet_label.setText("Folha: [nao definida]")
        if cells:
            self.range_label.setText(
                f"Intervalo de celulas: {cells[0].cell_id} ... {cells[-1].cell_id}"
            )
        else:
            self.range_label.setText("Intervalo de celulas: [vazio]")


    def _selected_targets(self) -> dict[Laminado, set[int]]:
        if not hasattr(self, "table") or self.table.selectionModel() is None:
            return {}
        targets: dict[Laminado, set[int]] = {}
        for index in self.table.selectionModel().selectedIndexes():
            if not index.isValid() or index.column() < 2:
                continue
            cell_idx = index.column() - 2
            if cell_idx < 0 or cell_idx >= len(self._cells):
                continue
            laminate = self._cells[cell_idx].laminate
            targets.setdefault(laminate, set()).add(index.row())
        return targets

    def _change_orientation(self) -> None:
        targets = self._selected_targets()
        if not targets:
            QtWidgets.QMessageBox.information(
                self,
                "Trocar orientacao",
                "Selecione pelo menos uma celula para alterar a orientacao.",
            )
            return
        project_orientations = project_distinct_orientations(self._project)
        dialog = BulkOrientationDialog(
            parent=self,
            available_orientations=project_orientations,
        )
        if dialog.exec() != QtWidgets.QDialog.Accepted:
            return
        new_orientation = dialog.new_orientation
        if new_orientation is None:
            return
        affected: list[str] = []
        for laminate, rows in targets.items():
            model = StackingTableModel(camadas=list(laminate.camadas))
            changed = False
            for row in sorted(rows):
                if row >= model.rowCount():
                    continue
                if model.apply_field_value(
                    row, StackingTableModel.COL_ORIENTATION, new_orientation
                ):
                    changed = True
            if changed:
                laminate.camadas = model.layers()
                affected.append(laminate.nome)
        if affected:
            self._notify_changes(affected)

    def _insert_layer(self, *, below: bool) -> None:
        targets = self._selected_targets()
        if not targets:
            QtWidgets.QMessageBox.information(
                self,
                "Inserir camada",
                "Selecione pelo menos uma linha para inserir uma nova camada.",
            )
            return
        affected: list[str] = []
        for laminate, rows in targets.items():
            model = StackingTableModel(camadas=list(laminate.camadas))
            changed = False
            ordered = sorted(rows, reverse=below)
            for row in ordered:
                insert_at = row + 1 if below else row
                insert_at = min(max(insert_at, 0), model.rowCount())
                model.insert_layer(
                    insert_at,
                    Camada(
                        idx=0,
                        material="",
                        orientacao=0,
                        ativo=True,
                        simetria=False,
                        ply_type=DEFAULT_PLY_TYPE,
                    ),
                )
                changed = True
            if changed:
                laminate.camadas = model.layers()
                affected.append(laminate.nome)
        if affected:
            self._notify_changes(affected)

    def _clear_selected_layers(self) -> None:
        targets = self._selected_targets()
        if not targets:
            QtWidgets.QMessageBox.information(
                self,
                "Limpar camada",
                "Selecione pelo menos uma celula para limpar.",
            )
            return
        affected: list[str] = []
        for laminate, rows in targets.items():
            model = StackingTableModel(camadas=list(laminate.camadas))
            changed = False
            for row in sorted(rows):
                if row >= model.rowCount():
                    continue
                if model.apply_field_value(row, StackingTableModel.COL_MATERIAL, ""):
                    changed = True
                if model.apply_field_value(row, StackingTableModel.COL_ORIENTATION, None):
                    changed = True
            if changed:
                laminate.camadas = model.layers()
                affected.append(laminate.nome)
        if affected:
            self._notify_changes(affected)

    def _toggle_maximize(self) -> None:
        if self.isMaximized():
            self.showNormal()
            self.btn_toggle_maximize.setText("Maximizar")
            self.btn_toggle_maximize.setChecked(False)
        else:
            self.showMaximized()
            self.btn_toggle_maximize.setText("Restaurar")
            self.btn_toggle_maximize.setChecked(True)

    def _on_model_change(self, laminate_names: list[str]) -> None:
        self._notify_changes(laminate_names)

    def _notify_changes(self, laminate_names: list[str]) -> None:
        if self._project is not None:
            try:
                self._project.mark_dirty(True)
            except Exception:
                pass
        self._rebuild_view()
        try:
            self.stacking_changed.emit(laminate_names)
        except Exception:
            pass

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
