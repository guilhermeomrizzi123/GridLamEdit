"""Main application window for GridLamEdit."""

from __future__ import annotations

from typing import Iterable, List

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMainWindow,
    QPushButton,
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
        toolbar.setToolButtonStyle(Qt.ToolButtonTextOnly)
        toolbar.setContentsMargins(8, 4, 8, 4)

        actions: List[str] = [
            "Carregar Planilha",
            "Novo Laminado",
            "Importar Laminado",
            "Salvar",
            "Desfazer",
            "Verificar Simetria",
        ]
        for index, text in enumerate(actions):
            action = QAction(text, self)
            action.setStatusTip(f"TODO: implementar ação '{text}'.")
            action.triggered.connect(self._show_todo_message)  # type: ignore[arg-type]
            toolbar.addAction(action)
            if index != len(actions) - 1:
                toolbar.addSeparator()

        self.addToolBar(toolbar)

    def _setup_central_widget(self) -> None:
        central = QWidget(self)
        outer_layout = QVBoxLayout(central)
        outer_layout.setContentsMargins(0, 0, 0, 0)

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
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

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
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(16)

        header = QLabel("Laminado Associado a Célula", panel)
        header_font: QFont = header.font()
        header_font.setBold(True)
        header_font.setPointSize(header_font.pointSize() + 1)
        header.setFont(header_font)

        layout.addWidget(header)
        layout.addLayout(self._build_laminate_form())
        layout.addWidget(self._build_associated_cells_view())

        stacking_label = QLabel("Stacking", panel)
        stacking_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        layout.addWidget(stacking_label)
        layout.addWidget(self._build_layers_section())
        return panel

    def _build_laminate_form(self) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.setSpacing(12)
        layout.setContentsMargins(0, 0, 0, 0)

        layout.addLayout(self._combo_with_label("Nome:", ["LAM-1", "LAM-2"], "name"))
        layout.addLayout(
            self._combo_with_label("Cor:", ["#FFFFFF", "#FF0000", "#00FF00"], "color")
        )
        layout.addLayout(
            self._combo_with_label("Tipo:", ["Core", "Skin", "Custom"], "type")
        )
        layout.addStretch()
        return layout

    def _combo_with_label(
        self, label_text: str, items: Iterable[str], attr_prefix: str
    ) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.setSpacing(6)
        layout.setContentsMargins(0, 0, 0, 0)

        label = QLabel(label_text, self)
        combo = QComboBox(self)
        combo.addItems(list(items))
        combo.setEditable(True)
        combo.setMinimumWidth(180)

        layout.addWidget(label)
        layout.addWidget(combo)
        setattr(self, f"laminate_{attr_prefix}_combo", combo)
        return layout

    def _build_associated_cells_view(self) -> QWidget:
        container = QWidget(self)
        layout = QVBoxLayout(container)
        layout.setSpacing(6)
        layout.setContentsMargins(0, 0, 0, 0)

        label = QLabel("Células associadas com esse laminado", container)
        self.associated_cells = QTextEdit(container)
        self.associated_cells.setReadOnly(True)
        self.associated_cells.setPlaceholderText("C3, C5")
        self.associated_cells.setMaximumHeight(80)
        self.associated_cells.setStyleSheet("background-color: #ffffff;")

        layout.addWidget(label)
        layout.addWidget(self.associated_cells)
        return container

    def _build_layers_section(self) -> QWidget:
        container = QWidget(self)
        layout = QHBoxLayout(container)
        layout.setSpacing(12)
        layout.setContentsMargins(0, 0, 0, 0)

        self.layers_table = self._create_layers_table(container)
        layout.addWidget(self.layers_table, stretch=1)
        layout.addLayout(self._create_layers_buttons())
        return container

    def _create_layers_table(self, parent: QWidget) -> QTableWidget:
        table = QTableWidget(parent)
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels(
            ["#", "Material", "Orientação", "Ativo", "Simetria"]
        )
        table.setRowCount(4)
        sample_data = [
            ("0", "Carbon", "0°", "Sim", False),
            ("1", "Glass", "45°", "Sim", True),
            ("2", "Kevlar", "-45°", "Não", False),
            ("3", "Foam", "90°", "Sim", False),
        ]

        for row, data in enumerate(sample_data):
            for col, value in enumerate(data[:-1]):
                item = QTableWidgetItem(value)
                if col == 0:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                table.setItem(row, col, item)

            check_box = QCheckBox(parent)
            check_box.setChecked(data[-1])
            check_box.setTristate(False)
            table.setCellWidget(row, 4, check_box)

        table.setAlternatingRowColors(True)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setVerticalScrollMode(QAbstractItemView.ScrollPerItem)
        table.verticalHeader().setVisible(False)

        header = table.horizontalHeader()
        header.setDefaultSectionSize(140)
        header.setMinimumSectionSize(80)
        table.setColumnWidth(0, 60)
        table.setColumnWidth(1, 180)
        table.setColumnWidth(2, 140)
        table.setColumnWidth(3, 110)
        table.setColumnWidth(4, 110)

        table.horizontalHeaderItem(0).setTextAlignment(Qt.AlignCenter)
        table.horizontalHeaderItem(4).setTextAlignment(Qt.AlignCenter)
        return table

    def _create_layers_buttons(self) -> QVBoxLayout:
        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(0, 0, 0, 0)

        buttons_info = [
            "➕ Adicionar Camada",
            "📄 Duplicar Camada",
            "⬆️ Mover Acima",
            "⬇️ Mover Abaixo",
            "🧱 Alterar Material",
            "🔄 Alterar Orientação",
            "✅ Verificar Simetria",
        ]

        self.layer_buttons: list[QPushButton] = []
        for text in buttons_info:
            button = QPushButton(text, self)
            button.setFixedWidth(200)
            button.clicked.connect(self._show_todo_message)  # type: ignore[arg-type]
            self.layer_buttons.append(button)
            layout.addWidget(button)

        layout.addStretch()
        return layout

    def _show_todo_message(self, checked: bool = False) -> None:  # noqa: ARG002
        """Placeholder slot for unimplemented actions."""
        if self.statusBar():
            self.statusBar().showMessage("TODO: implementar ação.", 2000)
