"""Main application window for GridLamEdit."""

from __future__ import annotations

import logging
from enum import Enum, auto
from pathlib import Path
from typing import Iterable, List, Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QColor, QCloseEvent, QFont, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QColorDialog,
    QFileDialog,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QMainWindow,
    QPushButton,
    QSplitter,
    QStatusBar,
    QStackedWidget,
    QTableView,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from gridlamedit.io.spreadsheet import (
    Camada,
    GridModel,
    Laminado,
    bind_cells_to_ui,
    bind_model_to_ui,
    load_grid_spreadsheet,
    normalize_angle,
)

logger = logging.getLogger(__name__)


class UiState(Enum):
    VIEW = auto()
    CREATING = auto()


class MainWindow(QMainWindow):
    """Primary window scaffolding the GridLamEdit interface."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("GridLamEdit")
        self.resize(1200, 800)

        self._grid_model: Optional[GridModel] = None

        self.ui_state = UiState.VIEW
        self._setup_toolbar()
        self._setup_central_widget()
        self._setup_status_bar()

    def _setup_toolbar(self) -> None:
        toolbar = QToolBar("Main Toolbar", self)
        toolbar.setMovable(False)
        toolbar.setToolButtonStyle(Qt.ToolButtonTextOnly)
        toolbar.setContentsMargins(8, 4, 8, 4)

        actions: List[tuple[str, callable]] = [
            ("Carregar Planilha", self._load_spreadsheet),
            ("Novo Laminado", self._enter_creating_mode),
            ("Importar Laminado", self._show_todo_message),
            ("Salvar", self._show_todo_message),
            ("Desfazer", self._show_todo_message),
            ("Verificar Simetria", self._show_todo_message),
        ]
        for index, (text, handler) in enumerate(actions):
            action = QAction(text, self)
            action.setStatusTip(f"TODO: implementar ação '{text}'.")
            action.triggered.connect(handler)  # type: ignore[arg-type]
            toolbar.addAction(action)
            if index != len(actions) - 1:
                toolbar.addSeparator()

        self.addToolBar(toolbar)

    def _setup_central_widget(self) -> None:
        self.view_editor = self._build_editor_view()
        self.view_new_laminate = self._build_new_laminate_view()

        self.central_stack = QStackedWidget(self)
        self.central_stack.addWidget(self.view_editor)
        self.central_stack.addWidget(self.view_new_laminate)
        self.central_stack.setCurrentWidget(self.view_editor)

        self.setCentralWidget(self.central_stack)

    def _build_editor_view(self) -> QWidget:
        editor = QWidget(self)
        outer_layout = QVBoxLayout(editor)
        outer_layout.setContentsMargins(0, 0, 0, 0)

        splitter = QSplitter(Qt.Horizontal, editor)
        splitter.addWidget(self._build_cells_panel())
        splitter.addWidget(self._build_laminate_panel())
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        outer_layout.addWidget(splitter)
        return editor

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
        self.cells_list.setSelectionMode(QListWidget.SingleSelection)
        self.lstCelulas = self.cells_list

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

    def _create_layers_table(self, parent: QWidget) -> QTableView:
        table = QTableView(parent)
        table.setAlternatingRowColors(True)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setVerticalScrollMode(QAbstractItemView.ScrollPerItem)
        table.verticalHeader().setVisible(False)
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

    def _build_new_laminate_view(self) -> QWidget:
        view = QWidget(self)
        layout = QVBoxLayout(view)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(16)

        form_row = QHBoxLayout()
        form_row.setSpacing(12)

        name_label = QLabel("Name:", view)
        self.new_laminate_name_edit = QLineEdit(view)
        self.new_laminate_name_edit.setPlaceholderText("Ex.: Web-RIB-26")
        form_row.addWidget(name_label)
        form_row.addWidget(self.new_laminate_name_edit, stretch=1)

        color_label = QLabel("ColorIdx:", view)
        color_layout = QHBoxLayout()
        color_layout.setSpacing(6)
        self.new_laminate_color_display = QLineEdit("#FFFFFF", view)
        self.new_laminate_color_display.setReadOnly(True)
        self.new_laminate_color_display.setMaximumWidth(120)
        self.new_laminate_color_button = QPushButton("Selecionar Cor", view)
        self.new_laminate_color_button.clicked.connect(self._select_new_laminate_color)
        color_layout.addWidget(self.new_laminate_color_display)
        color_layout.addWidget(self.new_laminate_color_button)

        form_row.addWidget(color_label)
        form_row.addLayout(color_layout)

        type_label = QLabel("Type:", view)
        self.new_laminate_type_combo = QComboBox(view)
        self.new_laminate_type_combo.addItems(["SS", "Core", "Skin", "RIB", "Other"])
        form_row.addWidget(type_label)
        form_row.addWidget(self.new_laminate_type_combo)

        form_row.addStretch()
        layout.addLayout(form_row)

        stacking_label = QLabel("Stacking do novo laminado", view)
        layout.addWidget(stacking_label)

        self.new_laminate_stacking_table = QTableWidget(view)
        self.new_laminate_stacking_table.setColumnCount(4)
        self.new_laminate_stacking_table.setHorizontalHeaderLabels(
            ["Material", "Orientação", "Ativo", "Simetria"]
        )
        self.new_laminate_stacking_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch
        )
        self.new_laminate_stacking_table.verticalHeader().setVisible(False)
        self.new_laminate_stacking_table.setSelectionBehavior(
            QTableWidget.SelectRows
        )
        self.new_laminate_stacking_table.setSelectionMode(
            QTableWidget.SingleSelection
        )
        layout.addWidget(self.new_laminate_stacking_table, stretch=1)

        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(8)
        self.new_laminate_add_layer_btn = QPushButton("Adicionar Camada", view)
        self.new_laminate_add_layer_btn.clicked.connect(self._new_laminate_add_layer)
        buttons_layout.addWidget(self.new_laminate_add_layer_btn)

        self.new_laminate_remove_layer_btn = QPushButton(
            "Remover Selecionada", view
        )
        self.new_laminate_remove_layer_btn.clicked.connect(
            self._new_laminate_remove_layer
        )
        buttons_layout.addWidget(self.new_laminate_remove_layer_btn)

        self.new_laminate_move_up_btn = QPushButton("Mover Acima", view)
        self.new_laminate_move_up_btn.clicked.connect(
            lambda: self._new_laminate_move_layer(-1)
        )
        buttons_layout.addWidget(self.new_laminate_move_up_btn)

        self.new_laminate_move_down_btn = QPushButton("Mover Abaixo", view)
        self.new_laminate_move_down_btn.clicked.connect(
            lambda: self._new_laminate_move_layer(1)
        )
        buttons_layout.addWidget(self.new_laminate_move_down_btn)

        buttons_layout.addStretch()
        layout.addLayout(buttons_layout)

        action_layout = QHBoxLayout()
        action_layout.addStretch()

        self.new_laminate_save_btn = QPushButton("Salvar", view)
        self.new_laminate_save_btn.setDefault(True)
        self.new_laminate_save_btn.clicked.connect(self._save_new_laminate)
        action_layout.addWidget(self.new_laminate_save_btn)

        self.new_laminate_cancel_btn = QPushButton("Cancelar", view)
        self.new_laminate_cancel_btn.clicked.connect(self._cancel_new_laminate)
        action_layout.addWidget(self.new_laminate_cancel_btn)

        layout.addLayout(action_layout)

        self._new_laminate_shortcuts = [
            QShortcut(QKeySequence("Ctrl+Return"), view),
            QShortcut(QKeySequence("Ctrl+Enter"), view),
        ]
        for shortcut in self._new_laminate_shortcuts:
            shortcut.activated.connect(self._save_new_laminate)
        self._new_laminate_cancel_shortcut = QShortcut(QKeySequence("Esc"), view)
        self._new_laminate_cancel_shortcut.activated.connect(
            self._cancel_new_laminate
        )

        self._reset_new_laminate_form()
        return view

    def _reset_new_laminate_form(self) -> None:
        self.new_laminate_name_edit.clear()
        self._update_color_preview("#FFFFFF")
        self.new_laminate_type_combo.setCurrentIndex(0)

        table = self.new_laminate_stacking_table
        table.setRowCount(0)
        self._new_laminate_add_layer()
        table.setCurrentCell(0, 0)

    def _update_color_preview(self, hex_color: str) -> None:
        hex_color = hex_color.upper()
        if not hex_color.startswith("#"):
            hex_color = f"#{hex_color}"
        self.new_laminate_color_display.setText(hex_color)
        self.new_laminate_color_display.setStyleSheet(
            f"background-color: {hex_color};"
        )

    def _select_new_laminate_color(self) -> None:
        initial = QColor(self.new_laminate_color_display.text())
        if not initial.isValid():
            initial = QColor("#FFFFFF")
        color = QColorDialog.getColor(initial, self, "Selecione a cor do laminado")
        if color.isValid():
            self._update_color_preview(color.name().upper())

    def _new_laminate_add_layer(self) -> None:
        table = self.new_laminate_stacking_table
        row = table.rowCount()
        table.insertRow(row)
        self._apply_layer_row(table, row, ("", "0", True, False))
        table.setCurrentCell(row, 0)

    def _new_laminate_remove_layer(self) -> None:
        table = self.new_laminate_stacking_table
        if table.rowCount() == 0:
            return
        current = table.currentRow()
        if current < 0:
            current = table.rowCount() - 1
        table.removeRow(current)
        if table.rowCount() == 0:
            self._new_laminate_add_layer()

    def _new_laminate_move_layer(self, direction: int) -> None:
        table = self.new_laminate_stacking_table
        current = table.currentRow()
        if current < 0:
            return
        target = current + direction
        if not 0 <= target < table.rowCount():
            return
        current_data = self._collect_layer_row(table, current)
        target_data = self._collect_layer_row(table, target)
        self._apply_layer_row(table, current, target_data)
        self._apply_layer_row(table, target, current_data)
        table.setCurrentCell(target, 0)

    def _collect_layer_row(
        self, table: QTableWidget, row: int
    ) -> tuple[str, str, bool, bool]:
        material = self._text(table.item(row, 0))
        orientation = self._text(table.item(row, 1))
        active = self._checkbox_value(table, row, 2)
        symmetry = self._checkbox_value(table, row, 3)
        return material, orientation, active, symmetry

    def _apply_layer_row(
        self,
        table: QTableWidget,
        row: int,
        data: tuple[str, str, bool, bool],
    ) -> None:
        material, orientation, active, symmetry = data
        table.setItem(row, 0, QTableWidgetItem(str(material)))
        table.setItem(row, 1, QTableWidgetItem(str(orientation)))

        active_checkbox = QCheckBox(table)
        active_checkbox.setChecked(active)
        table.setCellWidget(row, 2, self._wrap_checkbox(active_checkbox))

        symmetry_checkbox = QCheckBox(table)
        symmetry_checkbox.setChecked(symmetry)
        table.setCellWidget(row, 3, self._wrap_checkbox(symmetry_checkbox))

    def _enter_creating_mode(self, checked: bool = False) -> None:  # noqa: ARG002
        if self.ui_state == UiState.CREATING:
            return
        if self._grid_model is None:
            self._grid_model = GridModel()
        self.ui_state = UiState.CREATING
        self._reset_new_laminate_form()
        if hasattr(self, "cells_list"):
            self.cells_list.setEnabled(False)
        self.central_stack.setCurrentWidget(self.view_new_laminate)
        self.new_laminate_name_edit.setFocus()

    def _exit_creating_mode(self) -> None:
        self.ui_state = UiState.VIEW
        if hasattr(self, "cells_list"):
            self.cells_list.setEnabled(True)
        self.central_stack.setCurrentWidget(self.view_editor)

    def _cancel_new_laminate(self) -> None:
        self._exit_creating_mode()

    def _save_new_laminate(self) -> None:
        if self._grid_model is None:
            self._grid_model = GridModel()
        name = self.new_laminate_name_edit.text().strip()
        if not name:
            QMessageBox.warning(
                self, "Campos obrigatórios", "Informe o Name do laminado."
            )
            return
        if name in self._grid_model.laminados:
            QMessageBox.warning(
                self,
                "Nome duplicado",
                f"Já existe um laminado chamado '{name}'. Escolha outro nome.",
            )
            return

        color_hex = self.new_laminate_color_display.text().strip() or "#FFFFFF"
        tipo = self.new_laminate_type_combo.currentText()

        table = self.new_laminate_stacking_table
        camadas: list[Camada] = []
        for row in range(table.rowCount()):
            material = self._text(table.item(row, 0))
            orientation_text = self._text(table.item(row, 1)) or "0"
            if not material:
                continue
            try:
                orientacao = normalize_angle(orientation_text)
            except ValueError as exc:
                QMessageBox.warning(
                    self,
                    "Orientação inválida",
                    f"Linha {row + 1}: {exc}",
                )
                return
            ativo = self._checkbox_value(table, row, 2)
            simetria = self._checkbox_value(table, row, 3)
            camadas.append(
                Camada(
                    idx=len(camadas),
                    material=material,
                    orientacao=orientacao,
                    ativo=ativo,
                    simetria=simetria,
                )
            )

        if not camadas:
            QMessageBox.warning(
                self,
                "Stacking obrigatório",
                "Adicione ao menos uma camada ao laminado.",
            )
            return

        laminado = Laminado(
            nome=name,
            cor_hex=color_hex,
            tipo=tipo,
            celulas=[],
            camadas=camadas,
        )

        self._grid_model.laminados[name] = laminado
        self._refresh_after_new_laminate(name)
        self._exit_creating_mode()

    def _refresh_after_new_laminate(self, laminate_name: str) -> None:
        if self._grid_model is None:
            return
        bind_model_to_ui(self._grid_model, self)
        bind_cells_to_ui(self._grid_model, self)
        if hasattr(self, "laminate_name_combo"):
            idx = self.laminate_name_combo.findText(laminate_name)
            if idx >= 0:
                self.laminate_name_combo.setCurrentIndex(idx)
        binding = getattr(self, "_grid_binding", None)
        if binding is not None and hasattr(binding, "_apply_laminate"):
            try:
                binding._apply_laminate(laminate_name)  # type: ignore[attr-defined]
            except Exception as exc:  # pragma: no cover - defensive
                logger.warning("Não foi possível aplicar novo laminado: %s", exc)

    def _text(self, item: Optional[QTableWidgetItem]) -> str:
        return item.text().strip() if item is not None else ""

    def _checkbox_value(
        self, table: QTableWidget, row: int, column: int
    ) -> bool:
        widget = table.cellWidget(row, column)
        checkbox = widget
        if isinstance(widget, QWidget) and widget.layout() is not None:
            layout = widget.layout()
            if layout.count():
                checkbox = layout.itemAt(0).widget()
        if isinstance(checkbox, QCheckBox):
            return checkbox.isChecked()
        return False

    def _wrap_checkbox(self, checkbox: QCheckBox) -> QWidget:
        container = QWidget(self.new_laminate_stacking_table)
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignCenter)
        layout.addWidget(checkbox)
        return container

    def _show_todo_message(self, checked: bool = False) -> None:  # noqa: ARG002
        """Placeholder slot for unimplemented actions."""
        if self.statusBar():
            self.statusBar().showMessage("TODO: implementar ação.", 2000)

    def _load_spreadsheet(self, checked: bool = False) -> None:  # noqa: ARG002
        """Abre um arquivo Excel e popula a interface."""
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Carregar planilha do Grid Design",
            "",
            "Planilhas Excel (*.xlsx *.xls);;Todos os arquivos (*)",
        )
        if not path:
            return

        try:
            model = load_grid_spreadsheet(path)
        except ValueError as exc:
            logger.error("Falha ao carregar planilha: %s", exc)
            QMessageBox.critical(self, "Erro ao importar planilha", str(exc))
            if self.statusBar():
                self.statusBar().showMessage("Falha ao carregar planilha.", 4000)
            return

        if self.ui_state == UiState.CREATING:
            self._exit_creating_mode()

        bind_model_to_ui(model, self)
        bind_cells_to_ui(model, self)
        self._grid_model = model
        if self.statusBar():
            self.statusBar().showMessage(
                f"Planilha carregada: {Path(path).name}", 5000
            )
