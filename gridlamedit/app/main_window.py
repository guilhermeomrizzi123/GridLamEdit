"""Main application window for GridLamEdit."""

from __future__ import annotations

import logging
from enum import Enum, auto
from pathlib import Path
from typing import Iterable, List, Optional

from PySide6.QtCore import Qt, QSize, QTimer, QEvent, QObject
from PySide6.QtGui import (
    QAction,
    QCloseEvent,
    QIcon,
    QFont,
    QGuiApplication,
    QKeySequence,
    QShortcut,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QMainWindow,
    QPushButton,
    QStyle,
    QSizePolicy,
    QSplitter,
    QStatusBar,
    QStackedWidget,
    QToolButton,
    QTableView,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from gridlamedit.app.delegates import (
    CenteredCheckBoxDelegate,
    MaterialComboDelegate,
    OrientationComboDelegate,
    PlyTypeComboDelegate,
)
from gridlamedit.core.project_manager import ProjectManager
from gridlamedit.io.spreadsheet import (
    Camada,
    DEFAULT_COLOR_INDEX,
    DEFAULT_PLY_TYPE,
    GridModel,
    Laminado,
    StackingTableModel,
    WordWrapHeader,
    bind_cells_to_ui,
    bind_model_to_ui,
    load_grid_spreadsheet,
    normalize_angle,
)
from gridlamedit.services.excel_io import export_grid_xlsx

logger = logging.getLogger(__name__)

ICONS_DIR = Path(__file__).resolve().parent / "icons"


class UiState(Enum):
    VIEW = auto()
    CREATING = auto()


class MainWindow(QMainWindow):
    """Primary window scaffolding the GridLamEdit interface."""

    def __init__(self) -> None:
        super().__init__()
        self.base_title = "GridLamEdit"
        self.project_manager = ProjectManager(self._on_project_dirty_changed)
        self.setWindowTitle(self.base_title)
        self._apply_initial_geometry()

        self._grid_model: Optional[GridModel] = None
        self._stacking_checkbox_delegate = None
        self._stacking_ply_delegate = None
        self._stacking_material_delegate = None
        self._stacking_orientation_delegate = None
        self._selection_column_index = StackingTableModel.COL_SELECT
        self._stacking_header_band: Optional[QWidget] = None
        self._band_labels: list[QLabel] = []
        self._header_band_mapping: list[int] = [
            StackingTableModel.COL_NUMBER,
            StackingTableModel.COL_SELECT,
            StackingTableModel.COL_PLY_TYPE,
            StackingTableModel.COL_MATERIAL,
            StackingTableModel.COL_ORIENTATION,
        ]
        self._band_frame_margin = 0
        self._header_band_scroll_connected = False

        self.ui_state = UiState.VIEW
        self._create_actions()
        self._setup_menu_bar()
        self._setup_central_widget()
        self._setup_status_bar()
        self._update_save_actions_enabled()

    def _apply_initial_geometry(self) -> None:
        screen = QGuiApplication.primaryScreen()
        if screen is None:
            self.resize(1200, 800)
            return
        geometry = screen.availableGeometry()
        width = int(geometry.width() * 0.9)
        height = int(geometry.height() * 0.9)
        self.resize(width, height)
        self.move(
            geometry.x() + (geometry.width() - width) // 2,
            geometry.y() + (geometry.height() - height) // 2,
        )

    def _create_actions(self) -> None:
        action_specs: List[
            tuple[str, str, callable, str, Optional[QKeySequence]]
        ] = [
            (
                "open_project_action",
                "Abrir Projeto",
                self._on_open_project,
                "Abrir arquivo de projeto GridLam.",
                QKeySequence.Open,
            ),
            (
                "load_spreadsheet_action",
                "Carregar Planilha",
                self._load_spreadsheet,
                "Importar planilha do Grid Design.",
                None,
            ),
            (
                "new_laminate_action",
                "Novo Laminado",
                self._enter_creating_mode,
                "Cadastrar um novo laminado.",
                None,
            ),
            (
                "save_action",
                "Salvar",
                self._on_save_triggered,
                "Salvar alteracoes no projeto atual.",
                QKeySequence.Save,
            ),
            (
                "save_as_action",
                "Salvar Como",
                self._on_save_as_triggered,
                "Salvar o projeto em um novo arquivo.",
                QKeySequence.SaveAs,
            ),
            (
                "export_excel_action",
                "Exportar Planilha",
                self._on_export_excel,
                "Exportar planilha Excel com as alteracoes atuais.",
                QKeySequence("Ctrl+E"),
            ),
        ]

        for attr_name, text, handler, tip, shortcut in action_specs:
            action = QAction(text, self)
            action.setStatusTip(tip)
            if shortcut is not None:
                action.setShortcut(shortcut)
            action.triggered.connect(handler)  # type: ignore[arg-type]
            setattr(self, attr_name, action)
        self._update_save_actions_enabled()

    def _setup_menu_bar(self) -> None:
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("Arquivo")
        file_menu.addAction(self.open_project_action)
        file_menu.addAction(self.load_spreadsheet_action)
        file_menu.addAction(self.new_laminate_action)
        file_menu.addSeparator()
        file_menu.addAction(self.save_action)
        file_menu.addAction(self.save_as_action)
        file_menu.addAction(self.export_excel_action)

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
        splitter.setStretchFactor(1, 3)

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

        title = QLabel("Celulas", panel)
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

        header = QLabel("Laminado Associado a Celula", panel)
        header_font: QFont = header.font()
        header_font.setBold(True)
        header_font.setPointSize(header_font.pointSize() + 1)
        header.setFont(header_font)

        layout.addWidget(header)
        layout.addLayout(self._build_laminate_form())
        layout.addWidget(
            self._build_associated_cells_view(), alignment=Qt.AlignLeft
        )

        stacking_label = QLabel("Stacking", panel)
        stacking_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        layout.addWidget(stacking_label)
        layout.addWidget(self._build_layers_section(), stretch=1)
        return panel

    def _build_laminate_form(self) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.setSpacing(12)
        layout.setContentsMargins(0, 0, 0, 0)

        layout.addLayout(self._combo_with_label("Nome:", ["LAM-1", "LAM-2"], "name"))
        layout.addLayout(
            self._combo_with_label(
                "Cor:", (str(i) for i in range(1, 151)), "color", editable=False
            )
        )
        layout.addLayout(
            self._combo_with_label("Tipo:", ["Core", "Skin", "Custom"], "type")
        )
        layout.addStretch()
        return layout

    def _combo_with_label(
        self,
        label_text: str,
        items: Iterable[str],
        attr_prefix: str,
        *,
        editable: bool = True,
    ) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.setSpacing(6)
        layout.setContentsMargins(0, 0, 0, 0)

        label = QLabel(label_text, self)
        combo = QComboBox(self)
        items_list = [str(item) for item in items]
        combo.addItems(items_list)
        combo.setEditable(editable)
        if not editable:
            combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setMinimumWidth(180)
        if items_list:
            combo.setCurrentIndex(0)

        layout.addWidget(label)
        layout.addWidget(combo)
        setattr(self, f"laminate_{attr_prefix}_combo", combo)
        return layout

    def _build_associated_cells_view(self) -> QWidget:
        container = QWidget(self)
        layout = QVBoxLayout(container)
        layout.setSpacing(6)
        layout.setContentsMargins(0, 0, 0, 0)

        label = QLabel("Celulas associadas com esse laminado", container)
        label.setWordWrap(True)
        label.setMaximumWidth(220)
        label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
        self.associated_cells = QTextEdit(container)
        self.associated_cells.setReadOnly(True)
        self.associated_cells.setPlaceholderText("C3, C5")
        self.associated_cells.setMaximumWidth(220)
        self.associated_cells.setFixedHeight(40)
        self.associated_cells.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
        self.associated_cells.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.associated_cells.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.associated_cells.setStyleSheet("background-color: #ffffff;")

        layout.addWidget(label)
        layout.addWidget(self.associated_cells)
        container.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
        return container

    def _build_layers_section(self) -> QWidget:
        container = QWidget(self)
        layout = QHBoxLayout(container)
        layout.setSpacing(12)
        layout.setContentsMargins(0, 0, 0, 0)

        table_layout = QVBoxLayout()
        table_layout.setSpacing(6)
        table_layout.setContentsMargins(0, 0, 0, 0)

        self.layers_table = self._create_layers_table(container)
        header_widget = self._create_layers_header_widget(container)
        table_layout.addWidget(header_widget)
        table_layout.addWidget(self.layers_table, stretch=1)

        self.layers_count_label = QLabel(
            "Quantidade Total de Camadas: 0", container
        )
        self.layers_count_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        table_layout.addWidget(self.layers_count_label)

        layout.addLayout(table_layout, stretch=1)
        layout.addLayout(self._create_layers_buttons())
        container.setMinimumHeight(0)
        header = self.layers_table.horizontalHeader()
        header.sectionResized.connect(self._sync_header_band)
        header.sectionMoved.connect(self._sync_header_band)
        header.geometriesChanged.connect(self._sync_header_band)
        if not self._header_band_scroll_connected:
            self.layers_table.horizontalScrollBar().valueChanged.connect(
                self._sync_header_band
            )
            self._header_band_scroll_connected = True
        self.layers_table.viewport().installEventFilter(self)
        header.installEventFilter(self)
        QTimer.singleShot(0, self._sync_header_band)
        return container

    def _create_layers_table(self, parent: QWidget) -> QTableView:
        table = QTableView(parent)
        table.setAlternatingRowColors(True)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setVerticalScrollMode(QAbstractItemView.ScrollPerItem)
        table.verticalHeader().setVisible(False)
        table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        return table

    def _create_layers_header_widget(self, parent: QWidget) -> QWidget:
        band = QWidget(parent)
        band.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        band.setMinimumHeight(28)
        band.setMaximumHeight(28)
        band.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        band.setContentsMargins(0, 0, 0, 0)
        band.setLayout(None)

        table = getattr(self, "layers_table", None)
        if isinstance(table, QTableView):
            self._band_frame_margin = table.frameWidth()
        else:
            self._band_frame_margin = 0

        titles = ["#", "Selection", "Ply Type", "Material", "Orientação"]
        self._band_labels = []
        for title in titles:
            label = QLabel(title, band)
            label.setAlignment(Qt.AlignCenter)
            label.setStyleSheet("font-weight: 600;")
            label.setFixedHeight(28)
            label.setMinimumWidth(40)
            label.setAttribute(Qt.WA_TransparentForMouseEvents, True)
            label.setGeometry(0, 0, 0, band.height())
            label.hide()
            self._band_labels.append(label)
        band.installEventFilter(self)
        self._stacking_header_band = band
        return band

    def _create_layers_buttons(self) -> QVBoxLayout:
        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(0, 0, 0, 0)

        self.layer_buttons: list[QToolButton] = []

        def make_button(
            icon_name: str | None,
            tooltip: str,
            slot,
            accessible_name: str,
            fallback_icon: QStyle.StandardPixmap | None = None,
            *,
            text: str | None = None,
            tool_button_style: Qt.ToolButtonStyle | None = None,
            fixed_width: Optional[int] = 42,
        ) -> QToolButton:
            button = QToolButton(self)
            button.setText(text or "")
            if tool_button_style is None:
                tool_button_style = (
                    Qt.ToolButtonIconOnly if icon_name else Qt.ToolButtonTextOnly
                )
            button.setToolButtonStyle(tool_button_style)
            if fixed_width is not None:
                button.setFixedWidth(fixed_width)
            button.setAutoRaise(True)
            button.setToolTip(tooltip)
            button.setAccessibleName(accessible_name)
            icon = QIcon()
            if icon_name:
                button.setIconSize(QSize(24, 24))
                if icon_name.startswith(":/"):
                    icon = QIcon(icon_name)
                else:
                    icon_path = ICONS_DIR / icon_name
                    if icon_path.is_file():
                        icon = QIcon(str(icon_path))
                    else:
                        logger.debug("Icon file not found at %s", icon_path)
            if icon.isNull() and fallback_icon is not None:
                icon = self.style().standardIcon(fallback_icon)
            if not icon.isNull():
                button.setIcon(icon)
            button.clicked.connect(slot)
            self.layer_buttons.append(button)
            layout.addWidget(button)
            return button

        self.add_layer_button = make_button(
            "add-layer.svg",
            "Adicionar camada",
            self._on_add_layer_clicked,
            "Adicionar camada",
            QStyle.SP_FileDialogNewFolder,
        )

        self.duplicate_layer_button = make_button(
            "copy-layer.svg",
            "Duplicar camada",
            self._show_todo_message,  # type: ignore[arg-type]
            "Duplicar camada",
            QStyle.SP_FileDialogDetailedView,
        )

        self.move_up_button = make_button(
            "arrow-up.svg",
            "Mover camada para cima",
            self._on_move_up_clicked,
            "Mover camada para cima",
            QStyle.SP_ArrowUp,
        )

        self.move_down_button = make_button(
            "arrow-down.svg",
            "Mover camada para baixo",
            self._on_move_down_clicked,
            "Mover camada para baixo",
            QStyle.SP_ArrowDown,
        )

        self.symmetry_button = make_button(
            "symmetry.svg",
            "Verificar simetria",
            self.check_symmetry,
            "Verificar simetria",
            QStyle.SP_BrowserReload,
        )

        self.delete_layers_button = make_button(
            "trash.svg",
            "Excluir camadas selecionadas",
            self._on_delete_layers_clicked,
            "Excluir camadas selecionadas",
            QStyle.SP_TrashIcon,
        )

        self.select_all_layers_button = make_button(
            ":/icons/select_all.svg",
            "Selecionar todos",
            self._on_select_all_layers_clicked,
            "Selecionar todos",
            QStyle.SP_DialogYesButton,
        )

        self.clear_selection_button = make_button(
            ":/icons/clear_selection.svg",
            "Limpar seleção",
            self._on_clear_selection_clicked,
            "Limpar seleção",
            QStyle.SP_DialogResetButton,
        )

        layout.addStretch()
        return layout

    def _configure_stacking_table(self, binding) -> None:
        view = getattr(self, "layers_table", None)
        if not isinstance(view, QTableView):
            return
        model = binding.stacking_model

        view.setModel(None)

        header = view.horizontalHeader()
        if not isinstance(header, WordWrapHeader):
            header = WordWrapHeader(Qt.Horizontal, view)
            header.setDefaultAlignment(Qt.AlignCenter)
            view.setHorizontalHeader(header)
        else:
            header.set_checkbox_section(None)
            header.setDefaultAlignment(Qt.AlignCenter)

        view.setModel(model)

        if isinstance(model, StackingTableModel):
            self._connect_header_band_model_signals(model)

        self._install_stacking_delegates(view, binding)
        self._apply_stacking_column_setup(view)
        binding.set_header_view(header)
        self._refresh_stacking_header(view, model, header)

    def _install_stacking_delegates(self, view: QTableView, binding) -> None:
        self._stacking_checkbox_delegate = CenteredCheckBoxDelegate(view)
        self._stacking_ply_delegate = PlyTypeComboDelegate(view)
        self._stacking_material_delegate = MaterialComboDelegate(
            view, items_provider=binding.material_options
        )
        self._stacking_orientation_delegate = OrientationComboDelegate(
            view, items_provider=binding.orientation_options
        )
        view.setItemDelegateForColumn(
            StackingTableModel.COL_SELECT, self._stacking_checkbox_delegate
        )
        view.setItemDelegateForColumn(
            StackingTableModel.COL_PLY_TYPE, self._stacking_ply_delegate
        )
        view.setItemDelegateForColumn(
            StackingTableModel.COL_MATERIAL, self._stacking_material_delegate
        )
        view.setItemDelegateForColumn(
            StackingTableModel.COL_ORIENTATION, self._stacking_orientation_delegate
        )

    def _apply_stacking_column_setup(self, view: QTableView) -> None:
        view.setSortingEnabled(False)
        view.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.SelectedClicked
            | QAbstractItemView.EditKeyPressed
        )
        view.setSelectionBehavior(QAbstractItemView.SelectItems)
        view.setSelectionMode(QAbstractItemView.SingleSelection)

        header = view.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setSectionResizeMode(StackingTableModel.COL_NUMBER, QHeaderView.Fixed)
        header.setSectionResizeMode(StackingTableModel.COL_SELECT, QHeaderView.Fixed)
        header.setSectionResizeMode(StackingTableModel.COL_PLY_TYPE, QHeaderView.Fixed)
        header.setSectionResizeMode(StackingTableModel.COL_MATERIAL, QHeaderView.Stretch)
        header.setSectionResizeMode(
            StackingTableModel.COL_ORIENTATION, QHeaderView.Stretch
        )
        header.setMinimumSectionSize(60)
        header.setFixedHeight(max(header.height(), header.sizeHint().height()))

        view.setColumnWidth(StackingTableModel.COL_NUMBER, 60)
        view.setColumnWidth(StackingTableModel.COL_SELECT, 120)
        view.setColumnWidth(StackingTableModel.COL_PLY_TYPE, 160)
        view.verticalHeader().setVisible(False)
        self._sync_header_band()

    def _sync_header_band(self, *args) -> None:  # noqa: ARG002
        table = getattr(self, "layers_table", None)
        band = self._stacking_header_band
        if not isinstance(table, QTableView) or band is None:
            return
        if not self._band_labels:
            return
        header = table.horizontalHeader()
        column_count = header.count()
        if column_count == 0:
            for label in self._band_labels:
                label.hide()
            band.update()
            return

        self._band_frame_margin = table.frameWidth()
        x_offset = -table.horizontalScrollBar().value() + self._band_frame_margin
        band_height = max(1, band.height())

        for label, column in zip(self._band_labels, self._header_band_mapping):
            if column >= column_count or header.isSectionHidden(column):
                label.hide()
                continue
            section_pos = header.sectionViewportPosition(column)
            section_width = header.sectionSize(column)
            if section_width <= 0:
                label.hide()
                continue
            label.setGeometry(
                int(section_pos + x_offset),
                0,
                int(section_width),
                band_height,
            )
            label.show()
        band.update()

    def _refresh_stacking_header(
        self, view: QTableView, model: StackingTableModel, header: WordWrapHeader
    ) -> None:
        column_count = model.columnCount()
        if column_count > 0:
            model.headerDataChanged.emit(Qt.Horizontal, 0, column_count - 1)
        model.layoutChanged.emit()
        header.updateGeometry()
        header.viewport().update()
        view.viewport().update()

        self._sync_header_band()

        def _post_update() -> None:
            header.updateGeometry()
            header.viewport().update()
            view.viewport().update()
            self._sync_header_band()

        QTimer.singleShot(0, _post_update)

    def _connect_header_band_model_signals(
        self, model: StackingTableModel
    ) -> None:
        try:
            model.modelReset.disconnect(self._sync_header_band)
        except (TypeError, RuntimeError):
            pass
        model.modelReset.connect(self._sync_header_band)
        try:
            model.layoutChanged.disconnect(self._sync_header_band)
        except (TypeError, RuntimeError):
            pass
        model.layoutChanged.connect(self._sync_header_band)

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        if event.type() == QEvent.Resize:
            table = getattr(self, "layers_table", None)
            header = table.horizontalHeader() if isinstance(table, QTableView) else None
            viewport = table.viewport() if isinstance(table, QTableView) else None
            if watched in {
                self._stacking_header_band,
                header,
                viewport,
            }:
                QTimer.singleShot(0, self._sync_header_band)
        return super().eventFilter(watched, event)

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
        self.new_laminate_color_combo = QComboBox(view)
        self.new_laminate_color_combo.addItems([str(i) for i in range(1, 151)])
        default_idx = self.new_laminate_color_combo.findText(str(DEFAULT_COLOR_INDEX))
        if default_idx >= 0:
            self.new_laminate_color_combo.setCurrentIndex(default_idx)
        color_layout.addWidget(self.new_laminate_color_combo)

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
            ["Material", "Orientacao", "Ativo", "Simetria"]
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
        if hasattr(self, "new_laminate_color_combo"):
            default_idx = self.new_laminate_color_combo.findText(
                str(DEFAULT_COLOR_INDEX)
            )
            self.new_laminate_color_combo.setCurrentIndex(
                default_idx if default_idx >= 0 else 0
            )
        self.new_laminate_type_combo.setCurrentIndex(0)

        table = self.new_laminate_stacking_table
        table.setRowCount(0)
        self._new_laminate_add_layer()
        table.setCurrentCell(0, 0)

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
                self, "Campos obrigatorios", "Informe o Name do laminado."
            )
            return
        if name in self._grid_model.laminados:
            QMessageBox.warning(
                self,
                "Nome duplicado",
                f"Ja existe um laminado chamado '{name}'. Escolha outro nome.",
            )
            return

        try:
            color_index = int(self.new_laminate_color_combo.currentText())
        except (ValueError, AttributeError):
            color_index = DEFAULT_COLOR_INDEX
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
                    "Orientacao invalida",
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
                    ply_type=DEFAULT_PLY_TYPE,
                )
            )

        if not camadas:
            QMessageBox.warning(
                self,
                "Stacking obrigatorio",
                "Adicione ao menos uma camada ao laminado.",
            )
            return

        laminado = Laminado(
            nome=name,
            tipo=tipo,
            color_index=color_index,
            celulas=[],
            camadas=camadas,
        )

        self._grid_model.laminados[name] = laminado
        self._refresh_after_new_laminate(name)
        self._exit_creating_mode()
        self._mark_dirty()

    def _refresh_after_new_laminate(self, laminate_name: str) -> None:
        if self._grid_model is None:
            return
        bind_model_to_ui(self._grid_model, self)
        binding = getattr(self, "_grid_binding", None)
        if binding is not None:
            self._configure_stacking_table(binding)
        bind_cells_to_ui(self._grid_model, self)
        if hasattr(self, "laminate_name_combo"):
            idx = self.laminate_name_combo.findText(laminate_name)
            if idx >= 0:
                self.laminate_name_combo.setCurrentIndex(idx)
        if binding is not None and hasattr(binding, "_apply_laminate"):
            try:
                binding._apply_laminate(laminate_name)  # type: ignore[attr-defined]
            except Exception as exc:  # pragma: no cover - defensive
                logger.warning("Nao foi possivel aplicar novo laminado: %s", exc)
        self._show_model_warnings()
        self._update_save_actions_enabled()

    def _show_model_warnings(self) -> bool:
        if self._grid_model is None or not self._grid_model.compat_warnings:
            return False
        status_bar = self.statusBar()
        message = " | ".join(self._grid_model.compat_warnings)
        if status_bar:
            status_bar.showMessage(message, 7000)
        else:
            logger.warning("Avisos de compatibilidade: %s", message)
        self._grid_model.compat_warnings.clear()
        return True

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
            self.statusBar().showMessage("TODO: implementar acao.", 2000)

    def check_symmetry(self) -> None:
        """Verifica se o laminado atual e simetrico considerando apenas Structural Ply."""
        _, model = self._get_stacking_view_and_model()
        if model is None:
            self._info("Tabela de camadas indisponivel para verificar a simetria.")
            return
        if hasattr(model, "clear_all_highlights"):
            model.clear_all_highlights()

        try:
            colmap = self._column_map_by_header(
                model, ["#", "Ply Type", "Material", "Orientacao", "Orientation"]
            )
        except ValueError as exc:
            logger.warning("Falha ao mapear colunas da tabela: %s", exc)
            self._info(
                "Nao foi possivel verificar a simetria porque as colunas esperadas nao estao disponiveis."
            )
            return

        missing = [name for name in ("#", "Ply Type", "Material") if name not in colmap]
        if "Orientacao" not in colmap and "Orientation" not in colmap:
            missing.append("Orientacao/Orientation")
        if missing:
            self._info(
                "Nao foi possivel verificar a simetria. Colunas ausentes: "
                + ", ".join(missing)
                + "."
            )
            return

        col_num = colmap["#"]
        col_ply = colmap["Ply Type"]
        col_mat = colmap["Material"]
        if "Orientacao" in colmap:
            col_ori = colmap["Orientacao"]
        else:
            col_ori = colmap["Orientation"]

        row_count = model.rowCount()

        def is_structural(row: int) -> bool:
            value = self._data_str(model, row, col_ply)
            return value.lower() == "structural ply"

        structural_rows = [r for r in range(row_count) if is_structural(r)]
        count_struct = len(structural_rows)
        if count_struct == 0:
            self._info("Laminado simetrico (0 ou 1 camada estrutural).")
            return
        if count_struct == 1:
            if hasattr(model, "add_green_rows"):
                model.add_green_rows(structural_rows)
            self._scroll_to_rows(structural_rows)
            self._info("Laminado simetrico (0 ou 1 camada estrutural).")
            return

        i, j = 0, count_struct - 1
        while i < j:
            r_top = structural_rows[i]
            r_bot = structural_rows[j]

            mat_top = self._data_str(model, r_top, col_mat)
            mat_bot = self._data_str(model, r_bot, col_mat)
            ori_top = self._normalize_orientation(self._data_str(model, r_top, col_ori))
            ori_bot = self._normalize_orientation(self._data_str(model, r_bot, col_ori))

            if not (self._eq(mat_top, mat_bot) and self._eq(ori_top, ori_bot)):
                layer_num = self._data_str(model, r_top, col_num) or str(r_top + 1)
                pair_rows = [r_top, r_bot]
                if hasattr(model, "add_red_rows"):
                    model.add_red_rows(pair_rows)
                self._scroll_to_rows(pair_rows)
                self._warn_asymmetry(layer_num, mat_top, ori_top, mat_bot, ori_bot)
                return

            i += 1
            j -= 1

        if count_struct % 2 == 1:
            center_rows = [structural_rows[count_struct // 2]]
        else:
            center_rows = [
                structural_rows[count_struct // 2 - 1],
                structural_rows[count_struct // 2],
            ]
        if hasattr(model, "add_green_rows"):
            model.add_green_rows(center_rows)
        self._scroll_to_rows(center_rows)
        self._info(
            f"Laminado simetrico considerando apenas Structural Ply ({count_struct} camadas estruturais)."
        )

    def _scroll_to_rows(self, rows: Iterable[int]) -> None:
        rows_list = sorted({r for r in rows if isinstance(r, int) and r >= 0})
        if not rows_list:
            return
        view, model = self._get_stacking_view_and_model()
        if view is None or model is None:
            return
        first = rows_list[0]
        index = model.index(first, 0)
        if not index.isValid():
            return
        view.scrollTo(index, QAbstractItemView.PositionAtCenter)
        view.selectRow(first)

    def _column_map_by_header(
        self, model, wanted_names: list[str]
    ) -> dict[str, int]:
        if not hasattr(model, "columnCount") or not callable(model.columnCount):
            raise ValueError("Modelo invalido para leitura de colunas.")
        column_count = model.columnCount()
        headers: dict[str, int] = {}
        for column in range(column_count):
            header = model.headerData(column, Qt.Horizontal, Qt.DisplayRole)
            if header is None:
                continue
            text = str(header).strip()
            if not text:
                continue
            headers[text] = column
            headers[text.lower()] = column

        mapping: dict[str, int] = {}
        for name in wanted_names:
            candidates = [name, name.lower()]
            if name.lower() == "orientacao":
                candidates.extend(["Orientation", "orientation"])
            if name.lower() == "orientation":
                candidates.extend(["Orientacao", "orientacao"])
            for candidate in candidates:
                if candidate in headers:
                    mapping[name] = headers[candidate]
                    break
        return mapping

    def _data_str(self, model, row: int, column: int) -> str:
        index = model.index(row, column)
        if not index.isValid():
            return ""
        value = model.data(index, Qt.DisplayRole)
        if value is None:
            return ""
        return str(value).strip()

    def _normalize_orientation(self, raw: str) -> str:
        text = (raw or "").strip()
        if not text:
            return ""
        cleaned = (
            text.replace("\N{DEGREE SIGN}", "")
            .replace("\u00ba", "")
            .replace("deg", "")
            .replace("DEG", "")
            .strip()
        )
        try:
            angle = normalize_angle(cleaned)
        except Exception:
            filtered = "".join(ch for ch in cleaned if ch.isdigit() or ch in "+-")
            if not filtered:
                return text
            if filtered[0] not in "+-":
                filtered = f"+{filtered}"
            return filtered
        if angle > 0:
            return f"+{angle}"
        if angle < 0:
            return f"{angle}"
        return "0"

    def _eq(self, left: str, right: str) -> bool:
        return (left or "").strip().lower() == (right or "").strip().lower()

    def _info(self, message: str) -> None:
        QMessageBox.information(self, "Verificar simetria", message)

    def _warn_asymmetry(
        self,
        layer_num: str,
        mat_top: str,
        ori_top: str,
        mat_bot: str,
        ori_bot: str,
    ) -> None:
        message = (
            f"Quebra de simetria a partir da camada # {layer_num}.\n"
            f"Topo:   Material={mat_top or '-'}, Orientacao={ori_top or '-'}\n"
            f"Base:   Material={mat_bot or '-'}, Orientacao={ori_bot or '-'}"
        )
        QMessageBox.warning(self, "Verificar simetria", message)

    def _on_add_layer_clicked(self) -> None:
        binding = getattr(self, "_grid_binding", None)
        if binding is None:
            return
        checked_rows = binding.checked_rows()
        if len(checked_rows) > 1:
            QMessageBox.warning(
                self,
                "Adicionar camada",
                "Apenas uma camada deve estar selecionada para adicionar uma nova abaixo.",
            )
            return
        target_row = checked_rows[0] if checked_rows else None
        if not binding.add_layer(target_row):
            QMessageBox.information(
                self,
                "Adicionar camada",
                "Nenhum laminado ativo para receber a nova camada.",
            )
            return
        if self.statusBar():
            self.statusBar().showMessage("Camada adicionada.", 3000)
        self._update_save_actions_enabled()

    def _on_delete_layers_clicked(self) -> None:
        binding = getattr(self, "_grid_binding", None)
        if binding is None:
            return
        selected = binding.checked_rows()
        count = len(selected)
        if count == 0:
            QMessageBox.information(
                self,
                "Excluir camadas",
                "Selecione pelo menos uma camada para excluir.",
            )
            return
        reply = QMessageBox.question(
            self,
            "Confirmar exclusao",
            f"{count} camada(s) serao excluidas. Deseja continuar?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        removed = binding.delete_checked_layers()
        if removed == 0:
            return
        if self.statusBar():
            self.statusBar().showMessage(
                f"{removed} camada(s) excluidas.", 3000
            )
        self._update_save_actions_enabled()

    def _get_stacking_view_and_model(
        self,
    ) -> tuple[Optional[QTableView], Optional[StackingTableModel]]:
        view = getattr(self, "layers_table", None)
        if not isinstance(view, QTableView):
            view = None
        model: Optional[StackingTableModel] = None
        binding = getattr(self, "_grid_binding", None)
        if binding is not None:
            candidate = getattr(binding, "stacking_model", None)
            if isinstance(candidate, StackingTableModel):
                model = candidate
        if model is None and view is not None:
            table_model = view.model()
            if isinstance(table_model, StackingTableModel):
                model = table_model
        return view, model

    def _select_all_rows(self) -> None:
        view, model = self._get_stacking_view_and_model()
        if model is None or model.rowCount() == 0:
            return
        model.set_all_checked(True)
        if view is not None:
            view.viewport().update()

    def _clear_all_selections(self) -> None:
        view, model = self._get_stacking_view_and_model()
        if model is None:
            return
        if not model.any_checked():
            QMessageBox.information(
                self,
                "Aviso",
                "Nenhum item está selecionado.",
            )
            return
        model.set_all_checked(False)
        if view is not None:
            view.viewport().update()

    def _on_select_all_layers_clicked(self) -> None:
        self._select_all_rows()

    def _on_clear_selection_clicked(self) -> None:
        self._clear_all_selections()

    def _on_move_up_clicked(self) -> None:
        binding = getattr(self, "_grid_binding", None)
        if binding is None:
            return
        success, reason = binding.move_selected_layer(-1)
        if success:
            if self.statusBar():
                self.statusBar().showMessage("Camada movida para cima.", 3000)
            self._update_save_actions_enabled()
            return
        self._handle_move_error(reason, "acima")

    def _on_move_down_clicked(self) -> None:
        binding = getattr(self, "_grid_binding", None)
        if binding is None:
            return
        success, reason = binding.move_selected_layer(1)
        if success:
            if self.statusBar():
                self.statusBar().showMessage("Camada movida para baixo.", 3000)
            self._update_save_actions_enabled()
            return
        self._handle_move_error(reason, "abaixo")

    def _handle_move_error(self, reason: str, direction_label: str) -> None:
        if reason == "none":
            QMessageBox.information(
                self,
                "Nenhuma camada selecionada",
                "Selecione uma camada para mover.",
            )
        elif reason == "multi":
            QMessageBox.warning(
                self,
                "Selecao invalida",
                "Apenas uma camada deve estar selecionada para mover.",
            )
        elif reason == "edge":
            QMessageBox.information(
                self,
                "Movimento invalido",
                f"A camada ja esta na posicao limite {direction_label}.",
            )
        else:
            QMessageBox.information(
                self,
                "Movimento invalido",
                "Nao foi possivel mover a camada selecionada.",
            )


    def _load_spreadsheet(self, checked: bool = False) -> None:  # noqa: ARG002
        """Open an Excel file and populate the UI."""
        if not self._confirm_discard_changes():
            return
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
            QMessageBox.critical(self, "Erro", str(exc))
            if self.statusBar():
                self.statusBar().showMessage("Falha ao carregar planilha.", 4000)
            return

        model.source_excel_path = path
        model.dirty = False
        self._grid_model = model
        self.project_manager.current_path = None

        if self.ui_state == UiState.CREATING:
            self._exit_creating_mode()

        bind_model_to_ui(self._grid_model, self)
        binding = getattr(self, "_grid_binding", None)
        if binding is not None:
            self._configure_stacking_table(binding)
        bind_cells_to_ui(self._grid_model, self)
        self._apply_ui_state(self.project_manager.get_ui_state())
        self.project_manager.capture_from_model(
            self._grid_model, self._collect_ui_state()
        )
        self.project_manager.mark_dirty(True)
        self._update_save_actions_enabled()
        self._update_window_title()

        warnings_shown = self._show_model_warnings()
        if self.statusBar() and not warnings_shown:
            self.statusBar().showMessage(
                f"Planilha carregada: {Path(path).name}", 5000
            )

    def _on_open_project(self, checked: bool = False) -> None:  # noqa: ARG002
        if not self._confirm_discard_changes():
            return
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Abrir projeto GridLam",
            str(self.project_manager.current_path or ""),
            "Projetos GridLam (*.gridlam);;Todos os arquivos (*)",
        )
        if not path:
            return
        try:
            self.project_manager.load(Path(path))
            model = self.project_manager.build_model()
        except ValueError as exc:
            QMessageBox.critical(self, "Erro", str(exc))
            return

        self._grid_model = model
        if self.ui_state == UiState.CREATING:
            self._exit_creating_mode()

        bind_model_to_ui(self._grid_model, self)
        binding = getattr(self, "_grid_binding", None)
        if binding is not None:
            self._configure_stacking_table(binding)
        bind_cells_to_ui(self._grid_model, self)
        self._apply_ui_state(self.project_manager.get_ui_state())
        self.project_manager.capture_from_model(
            self._grid_model, self._collect_ui_state()
        )
        self.project_manager.mark_dirty(False)
        self._update_save_actions_enabled()
        self._update_window_title()

        warnings_shown = self._show_model_warnings()
        if self.statusBar() and not warnings_shown:
            self.statusBar().showMessage(
                f"Projeto carregado: {Path(path).name}", 4000
            )

    def _collect_ui_state(self) -> dict:
        state: dict = {}
        if getattr(self, "lstCelulas", None) and self.lstCelulas.currentItem():
            item = self.lstCelulas.currentItem()
            cell_id = item.data(Qt.UserRole) if item is not None else None
            if not cell_id and item is not None:
                cell_id = item.text().split("|")[0].strip()
            if cell_id:
                state["selected_cell"] = str(cell_id)
        if getattr(self, "laminate_name_combo", None):
            state["selected_laminate"] = self.laminate_name_combo.currentText()
        return state

    def _apply_ui_state(self, state: dict) -> None:
        if not state:
            return
        cell_id = state.get("selected_cell")
        list_widget = getattr(self, "lstCelulas", None)
        if not isinstance(list_widget, QListWidget):
            list_widget = getattr(self, "cells_list", None)
        if cell_id and isinstance(list_widget, QListWidget):
            for idx in range(list_widget.count()):
                item = list_widget.item(idx)
                item_cell = item.data(Qt.UserRole)
                if not item_cell and item is not None:
                    item_cell = item.text().split("|")[0].strip()
                if str(item_cell) == str(cell_id):
                    list_widget.setCurrentItem(item)
                    break

        laminate_name = state.get("selected_laminate")
        if (
            laminate_name
            and self._grid_model is not None
            and laminate_name in self._grid_model.laminados
        ):
            combo = getattr(self, "laminate_name_combo", None)
            if isinstance(combo, QComboBox):
                if combo.currentText() != laminate_name:
                    combo.blockSignals(True)
                    index = combo.findText(laminate_name)
                    if index >= 0:
                        combo.setCurrentIndex(index)
                    else:
                        combo.setEditText(laminate_name)
                    combo.blockSignals(False)
            binding = getattr(self, "_grid_binding", None)
            if binding is not None and hasattr(binding, "_apply_laminate"):
                try:
                    binding._apply_laminate(laminate_name)  # type: ignore[attr-defined]
                except Exception as exc:  # pragma: no cover - defensive
                    logger.debug("Falha ao aplicar estado de laminado: %s", exc)

    def _snapshot_from_model(self) -> None:
        if self._grid_model is None:
            return
        ui_state = self._collect_ui_state()
        self.project_manager.capture_from_model(self._grid_model, ui_state)

    def _perform_save(self, path: Optional[str]) -> bool:
        if self._grid_model is None:
            return False
        self._snapshot_from_model()
        try:
            self.project_manager.save(Path(path) if path else None)
        except ValueError as exc:
            QMessageBox.critical(self, "Erro", str(exc))
            return False
        if self.statusBar():
            target = self.project_manager.current_path
            if target is not None:
                self.statusBar().showMessage(
                    f"Projeto salvo: {target.name}", 4000
                )
        return True

    def _on_save_triggered(self, checked: bool = False) -> bool:  # noqa: ARG002
        if self._grid_model is None:
            QMessageBox.information(self, "Salvar", "Nao ha projeto carregado.")
            return False
        if self.project_manager.current_path is None:
            return self._on_save_as_triggered()
        if self._perform_save(None):
            self.project_manager.mark_dirty(False)
            self._update_window_title()
            self._update_save_actions_enabled()
            return True
        return False

    def _on_save_as_triggered(self, checked: bool = False) -> bool:  # noqa: ARG002
        if self._grid_model is None:
            QMessageBox.information(self, "Salvar", "Nao ha projeto carregado.")
            return False
        initial_path = (
            str(self.project_manager.current_path)
            if self.project_manager.current_path
            else str(Path.cwd() / "projeto.gridlam")
        )
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar projeto",
            initial_path,
            "Projetos GridLam (*.gridlam);;Todos os arquivos (*)",
        )
        if not path:
            return False
        if not path.lower().endswith(".gridlam"):
            path = f"{path}.gridlam"
        if self._perform_save(path):
            self.project_manager.current_path = Path(path)
            self.project_manager.mark_dirty(False)
            self._update_window_title()
            self._update_save_actions_enabled()
            return True
        return False

    def _on_export_excel(self, checked: bool = False) -> bool:  # noqa: ARG002
        if self._grid_model is None or not self._grid_model.laminados:
            QMessageBox.information(
                self,
                "Exportar planilha",
                "Carregue uma planilha ou projeto antes de exportar.",
            )
            return False

        source_path = self._grid_model.source_excel_path
        if source_path:
            base_path = Path(source_path)
            suggested = base_path.with_name(f"{base_path.stem}_editado.xlsx")
        else:
            suggested = Path.cwd() / "grid_export.xlsx"

        path_str, _ = QFileDialog.getSaveFileName(
            self,
            "Exportar planilha do Grid Design",
            str(suggested),
            "Planilhas Excel (*.xlsx *.xls);;Todos os arquivos (*)",
        )
        if not path_str:
            return False

        target_path = Path(path_str)
        try:
            final_path = export_grid_xlsx(self._grid_model, target_path)
        except ValueError as exc:
            QMessageBox.critical(self, "Erro ao exportar", str(exc))
            return False
        except Exception as exc:  # pragma: no cover - defensivo
            logger.error("Falha ao exportar planilha: %s", exc, exc_info=True)
            QMessageBox.critical(
                self,
                "Erro ao exportar",
                f"Falha ao exportar a planilha: {exc}",
            )
            return False

        if self.statusBar():
            self.statusBar().showMessage(
                f"Planilha exportada: {final_path.name}", 5000
            )
        return True

    def _on_project_dirty_changed(self, is_dirty: bool) -> None:
        if self._grid_model is not None:
            self._grid_model.dirty = is_dirty
        self._update_window_title()
        self._update_save_actions_enabled()

    def _mark_dirty(self) -> None:
        if self._grid_model is None:
            return
        self._grid_model.dirty = True
        self.project_manager.mark_dirty(True)

    def _update_save_actions_enabled(self) -> None:
        has_model = self._grid_model is not None
        data_ready = bool(self._grid_model and self._grid_model.laminados)
        if getattr(self, "save_action", None) is not None:
            self.save_action.setEnabled(has_model and self.project_manager.is_dirty)
        if getattr(self, "save_as_action", None) is not None:
            self.save_as_action.setEnabled(has_model)
        if getattr(self, "export_excel_action", None) is not None:
            self.export_excel_action.setEnabled(data_ready)

    def _update_window_title(self) -> None:
        title = self.base_title
        if self.project_manager.current_path:
            title = f"{title} - {self.project_manager.current_path.name}"
        if self.project_manager.is_dirty:
            title = f"{title} *"
        self.setWindowTitle(title)

    def _confirm_discard_changes(self) -> bool:
        if not self.project_manager.is_dirty:
            return True
        response = QMessageBox.question(
            self,
            "Alteracoes pendentes",
            "Voce possui alteracoes nao salvas. Deseja salvar antes de continuar?",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
        )
        if response == QMessageBox.Yes:
            return self._on_save_triggered()
        if response == QMessageBox.No:
            return True
        return False

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._confirm_discard_changes():
            event.accept()
        else:
            event.ignore()
