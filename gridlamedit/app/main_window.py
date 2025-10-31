"""Main application window for GridLamEdit."""

from __future__ import annotations

from typing import Iterable

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAction,
    QCheckBox,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMainWindow,
    QSplitter,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolBar,
    QVBoxLayout,
    QWidget,
)


class MainWindow(QMainWindow):
    """Primary window scaffolding the GridLamEdit interface."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("GridLamEdit")
        self.resize(1200, 800)

        self._setup_toolbar()
        self._setup_central_widget()
        self._setup_status_bar()

    def _setup_toolbar(self) -> None:
        toolbar = QToolBar("Main Toolbar", self)
        toolbar.setMovable(False)

        for text in [
            "Carregar Planilha",
            "Novo Laminado",
            "Importar Laminado",
            "Salvar",
            "Desfazer",
            "Verificar Simetria",
        ]:
            action = QAction(text, self)
            action.setStatusTip(f"TODO: implementar ação '{text}'.")
            action.triggered.connect(self._show_todo_message)  # type: ignore[arg-type]
            toolbar.addAction(action)

        self.addToolBar(toolbar)

    def _setup_central_widget(self) -> None:
        central = QWidget(self)
        outer_layout = QVBoxLayout(central)
        splitter = QSplitter(Qt.Horizontal, central)
        splitter.addWidget(self._build_cells_panel())
        splitter.addWidget(self._build_laminate_panel())
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        outer_layout.addWidget(splitter)
        self.setCentralWidget(central)

    def _setup_status_bar(self) -> None:
        status_bar = QStatusBar(self)
        status_bar.showMessage("Pronto")
        self.setStatusBar(status_bar)

    def _build_cells_panel(self) -> QWidget:
        panel = QWidget(self)
        layout = QVBoxLayout(panel)
        title = QLabel("Células", panel)
        title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        self.cells_list = QListWidget(panel)
        self.cells_list.addItems(["C1", "C2", "C3", "C4"])
        self.cells_list.setSelectionMode(QListWidget.SingleSelection)

        layout.addWidget(title)
        layout.addWidget(self.cells_list)
        return panel

    def _build_laminate_panel(self) -> QWidget:
        panel = QWidget(self)
        layout = QVBoxLayout(panel)

        layout.addLayout(self._build_laminate_form())
        layout.addWidget(self._build_associated_cells_view())
        layout.addWidget(self._build_layers_table())
        return panel

    def _build_laminate_form(self) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.setSpacing(12)

        name_layout, self.laminate_name_combo = self._combo_with_label(
            "Nome", ["LAM-1", "LAM-2"]
        )
        layout.addLayout(name_layout)

        color_layout, self.laminate_color_combo = self._combo_with_label(
            "Cor", ["#FFFFFF", "#FF0000", "#00FF00"]
        )
        layout.addLayout(color_layout)

        type_layout, self.laminate_type_combo = self._combo_with_label(
            "Tipo", ["Core", "Skin", "Custom"]
        )
        layout.addLayout(type_layout)

        layout.addStretch()
        return layout

    def _combo_with_label(
        self, label_text: str, items: Iterable[str]
    ) -> tuple[QHBoxLayout, QComboBox]:
        layout = QHBoxLayout()
        label = QLabel(f"{label_text}:", self)
        combo = QComboBox(self)
        combo.addItems(list(items))
        combo.setEditable(True)
        layout.addWidget(label)
        layout.addWidget(combo)
        return layout, combo

    def _build_associated_cells_view(self) -> QWidget:
        container = QWidget(self)
        layout = QVBoxLayout(container)
        label = QLabel("Células associadas com esse laminado", container)
        self.associated_cells = QTextEdit(container)
        self.associated_cells.setReadOnly(True)
        self.associated_cells.setPlaceholderText("C1, C2, C3")

        layout.addWidget(label)
        layout.addWidget(self.associated_cells)
        return container

    def _build_layers_table(self) -> QWidget:
        container = QWidget(self)
        layout = QVBoxLayout(container)
        label = QLabel("Stacking", container)
        self.layers_table = QTableWidget(container)
        self.layers_table.setColumnCount(5)
        self.layers_table.setHorizontalHeaderLabels(
            ["#", "Material", "Ângulo", "Ativo", "Simetria"]
        )
        self.layers_table.setRowCount(4)
        sample_data = [
            ("0", "Carbon", "0", "Sim", False),
            ("1", "Glass", "45", "Sim", True),
            ("2", "Kevlar", "-45", "Não", False),
            ("3", "Foam", "90", "Sim", False),
        ]

        for row, data in enumerate(sample_data):
            for col, value in enumerate(data[:-1]):
                item = QTableWidgetItem(value)
                if col == 0:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                self.layers_table.setItem(row, col, item)

            check_box = QCheckBox(self.layers_table)
            check_box.setChecked(data[-1])
            check_box.setTristate(False)
            self.layers_table.setCellWidget(row, 4, check_box)

        self.layers_table.resizeColumnsToContents()
        self.layers_table.setAlternatingRowColors(True)
        self.layers_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        layout.addWidget(label)
        layout.addWidget(self.layers_table)
        return container

    def _show_todo_message(self, checked: bool = False) -> None:  # noqa: ARG002
        """Placeholder slot for unimplemented actions."""
        if self.statusBar():
            self.statusBar().showMessage("TODO: implementar ação.", 2000)
