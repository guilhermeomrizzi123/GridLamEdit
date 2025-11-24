"""Virtual Stacking dialog and model."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional

from PySide6 import QtCore, QtGui, QtWidgets

from gridlamedit.io.spreadsheet import GridModel, Laminado, WordWrapHeader


@dataclass
class VirtualStackingLayer:
    """Represent a stacking layer sequence entry."""

    sequence_label: str
    material: str


@dataclass
class VirtualStackingCell:
    """Represent orientations of a laminado for each grid cell."""

    cell_id: str
    laminate_name: str
    orientations: list[Optional[float]] = field(default_factory=list)


class VirtualStackingModel(QtCore.QAbstractTableModel):
    """Qt model responsible for exposing Virtual Stacking data."""

    def __init__(self, parent: Optional[QtCore.QObject] = None) -> None:
        super().__init__(parent)
        self.layers: list[VirtualStackingLayer] = []
        self.cells: list[VirtualStackingCell] = []
        self.symmetry_row_index: int | None = None

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
                return f"#\n{cell.cell_id} | {cell.laminate_name}"
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
                if row < len(cell.orientations):
                    value = cell.orientations[row]
                    if value is None:
                        return ""
                    if role == QtCore.Qt.DisplayRole:
                        return f"{value:g}°"
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
        if row >= len(cell.orientations):
            return False

        text = str(value).strip()
        if text.endswith("°"):
            text = text[:-1].strip()

        if not text:
            angle: Optional[float] = None
        else:
            try:
                angle = float(text.replace(",", "."))
            except ValueError:
                return False

            if angle < -100 or angle > 100:
                return False

        cell.orientations[row] = angle
        self.dataChanged.emit(index, index, [QtCore.Qt.DisplayRole, QtCore.Qt.EditRole])
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
        layer_count = len(self.layers)
        normalized_cells: list[VirtualStackingCell] = []
        for cell in cells:
            orientations = list(cell.orientations)
            if len(orientations) < layer_count:
                orientations.extend([None] * (layer_count - len(orientations)))
            elif len(orientations) > layer_count:
                orientations = orientations[:layer_count]
            normalized_cells.append(
                VirtualStackingCell(
                    cell_id=cell.cell_id,
                    laminate_name=cell.laminate_name,
                    orientations=orientations,
                )
            )
        self.cells = normalized_cells
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

    def __init__(self, parent: Optional[QtWidgets.QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Virtual Stacking – View Mode: Cells")
        self.resize(1200, 700)

        self._layers: list[VirtualStackingLayer] = []
        self._cells: list[VirtualStackingCell] = []
        self._symmetry_row_index: int | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        self.grid_label = QtWidgets.QLabel("Grid: [não definido]", self)
        self.sheet_label = QtWidgets.QLabel("Folha: [não definida]", self)
        self.range_label = QtWidgets.QLabel("Intervalo de células: [não definido]", self)

        top_layout = QtWidgets.QHBoxLayout()
        top_layout.addWidget(self.grid_label)
        top_layout.addWidget(self.sheet_label)
        top_layout.addWidget(self.range_label)
        top_layout.addStretch()

        self.model = VirtualStackingModel(self)
        self.table = QtWidgets.QTableView(self)
        self.table.setModel(self.model)
        header = WordWrapHeader(QtCore.Qt.Horizontal, self.table)
        header.setDefaultAlignment(QtCore.Qt.AlignCenter)
        self.table.setHorizontalHeader(header)
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectItems)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked
            | QtWidgets.QAbstractItemView.SelectedClicked
        )

        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.addLayout(top_layout)
        main_layout.addWidget(self.table)
        self.setLayout(main_layout)

    # Data binding ----------------------------------------------------
    def populate_from_project(self, project: Optional[GridModel]) -> None:
        """Populate labels and table using the provided project snapshot."""
        if project is None:
            self._layers = []
            self._cells = []
            self._symmetry_row_index = None
            self.model.set_virtual_stacking([], [], None)
            self.grid_label.setText("Grid: [indefinido]")
            self.sheet_label.setText("Folha: [não definida]")
            self.range_label.setText("Intervalo de células: [não definido]")
            return

        cells: list[VirtualStackingCell] = []
        laminates: list[Laminado] = []
        for cell_id in project.celulas_ordenadas:
            laminate = self._laminate_for_cell(project, cell_id)
            if laminate is None:
                continue
            laminates.append(laminate)
            orientations = [layer.orientacao for layer in laminate.camadas]
            cells.append(
                VirtualStackingCell(
                    cell_id=cell_id,
                    laminate_name=laminate.nome or cell_id,
                    orientations=orientations,
                )
            )

        max_layers = max((len(lam.camadas) for lam in laminates), default=0)
        layers: list[VirtualStackingLayer] = []
        for idx in range(max_layers):
            material = ""
            sequence_label = ""
            for lam in laminates:
                if idx >= len(lam.camadas):
                    continue
                camada = lam.camadas[idx]
                if not material and camada.material:
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

        self._layers = layers
        self._cells = cells
        self._symmetry_row_index = symmetry_index

        self.model.set_virtual_stacking(layers, cells, symmetry_index)

        grid_name = project.source_excel_path or "[indefinido]"
        if grid_name not in ("[indefinido]", ""):
            grid_name = Path(grid_name).name

        self.grid_label.setText(f"Grid: {grid_name or '[indefinido]'}")
        self.sheet_label.setText("Folha: [não definida]")
        if cells:
            self.range_label.setText(
                f"Intervalo de células: {cells[0].cell_id} ... {cells[-1].cell_id}"
            )
        else:
            self.range_label.setText("Intervalo de células: [vazio]")

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
