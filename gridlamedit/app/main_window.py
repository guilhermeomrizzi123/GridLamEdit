"""Main application window for GridLamEdit."""

from __future__ import annotations

import logging
import copy
import os
import re
import secrets
import math
import unicodedata
from collections import Counter, OrderedDict
from dataclasses import dataclass
from enum import Enum, auto
from pathlib import Path
from typing import Iterable, List, Optional

from PySide6.QtCore import (
    Qt,
    QSize,
    QTimer,
    QEvent,
    QObject,
    QByteArray,
    QSettings,
    QThread,
    QUrl,
    Signal,
    Slot,
    QSortFilterProxyModel,
)
from PySide6.QtGui import (
    QDesktopServices,
    QAction,
    QCloseEvent,
    QIcon,
    QFont,
    QColor,
    QGuiApplication,
    QKeySequence,
    QShortcut,
    QTextOption,
    QUndoCommand,
    QUndoStack,
    QStandardItemModel,
    QStandardItem,
)
from PySide6.QtWidgets import (
    QAbstractButton,
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QGridLayout,
    QHeaderView,
    QHBoxLayout,
    QFrame,
    QLabel,
    QLineEdit,
    QInputDialog,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QMainWindow,
    QPushButton,
    QStyle,
    QSizePolicy,
    QSpacerItem,
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
from gridlamedit.app.dialogs.associated_cells_dialog import AssociatedCellsDialog
from gridlamedit.app.dialogs.bulk_material_dialog import BulkMaterialDialog
from gridlamedit.app.dialogs.bulk_orientation_dialog import BulkOrientationDialog
from gridlamedit.app.dialogs.duplicate_laminate_dialog import DuplicateLaminateDialog
from gridlamedit.app.dialogs.name_laminate_dialog import NameLaminateDialog
from gridlamedit.app.dialogs.new_laminate_paste_dialog import NewLaminatePasteDialog
from gridlamedit.app.dialogs.stacking_summary_dialog import StackingSummaryDialog
from gridlamedit.app.virtualstacking import VirtualStackingWindow
from gridlamedit.app.cell_neighbors import CellNeighborsWindow
from gridlamedit.core.project_manager import ProjectManager
from gridlamedit.core.paths import package_path
from gridlamedit.io.spreadsheet import (
    Camada,
    DEFAULT_COLOR_INDEX,
    DEFAULT_PLY_TYPE,
    is_structural_ply_label,
    GridModel,
    Laminado,
    MIN_COLOR_INDEX,
    MAX_COLOR_INDEX,
    StackingTableModel,
    WordWrapHeader,
    bind_cells_to_ui,
    bind_model_to_ui,
    _format_cell_label,
    format_orientation_value,
    load_grid_spreadsheet,
    normalize_angle,
    normalize_ply_type_label,
    PLY_TYPE_OPTIONS,
    orientation_highlight_color,
    count_oriented_layers,
    NO_LAMINATE_COMBO_OPTION,
)
from gridlamedit.services.excel_io import export_grid_xlsx, ensure_layers_have_material
from gridlamedit.services.laminate_batch_import import (
    BatchLaminateInput,
    create_blank_batch_template,
    parse_batch_template,
)
from gridlamedit.services.material_registry import (
    add_custom_material,
    available_materials as registry_available_materials,
)
from gridlamedit.services.project_query import (
    project_distinct_orientations,
    project_most_used_material,
)
from gridlamedit.services.laminate_checks import (
    ChecksReport,
    evaluate_symmetry_for_layers,
    run_all_checks,
)
from gridlamedit.services.laminate_service import (
    auto_name_for_laminate,
    auto_name_for_layers,
    sync_material_by_sequence,
)
from gridlamedit.ui.dialogs.duplicate_removal_dialog import DuplicateRemovalDialog
from gridlamedit.ui.dialogs.new_laminate_dialog import NewLaminateDialog
from gridlamedit.ui.dialogs.verification_report_dialog import VerificationReportDialog

logger = logging.getLogger(__name__)

ICONS_DIR = package_path("app", "icons")
RESOURCES_ICONS_DIR = package_path("resources", "icons")

COL_NUM = StackingTableModel.COL_NUMBER
COL_SEQUENCE = StackingTableModel.COL_SEQUENCE
COL_PLY = StackingTableModel.COL_PLY
COL_SELECTION = StackingTableModel.COL_SELECT
COL_PLY_TYPE = StackingTableModel.COL_PLY_TYPE
COL_MATERIAL = StackingTableModel.COL_MATERIAL
COL_ORIENTATION = StackingTableModel.COL_ORIENTATION


class LaminateFilterProxy(QSortFilterProxyModel):
    """Case-insensitive filter that matches raw or normalized text and keeps the sentinel."""

    def __init__(
        self, parent: QObject | None = None, *, sentinel: str = NO_LAMINATE_COMBO_OPTION
    ) -> None:
        super().__init__(parent)
        self._filter_text: str = ""
        self._filter_norm: str = ""
        self._sentinel: str = sentinel
        self.setFilterCaseSensitivity(Qt.CaseInsensitive)

    @staticmethod
    def _normalize(text: str) -> str:
        base = unicodedata.normalize("NFKD", text)
        stripped = "".join(ch for ch in base if not unicodedata.combining(ch))
        return re.sub(r"[^0-9A-Za-z]+", "", stripped).lower()

    def set_filter_text(self, text: str) -> None:
        self._filter_text = text.strip()
        self._filter_norm = self._normalize(self._filter_text) if self._filter_text else ""
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row: int, source_parent) -> bool:  # type: ignore[override]
        if not self._filter_text:
            return True
        index = self.sourceModel().index(source_row, self.filterKeyColumn(), source_parent)
        raw_text = str(index.data() or "")
        if raw_text == self._sentinel:
            return True

        # Raw substring match (case-insensitive)
        if self._filter_text.lower() in raw_text.lower():
            return True

        # Normalized match handles accents and punctuation removal (e.g., "4109" in "L30(4109)")
        if self._filter_norm:
            norm_text = self._normalize(raw_text)
            if self._filter_norm in norm_text:
                return True

        return False


class _LaminateChecksWorker(QObject):
    """Background worker responsible for executing laminate verifications."""

    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, laminates: list[Laminado]) -> None:
        super().__init__()
        self._laminates = laminates

    @Slot()
    def run(self) -> None:
        try:
            report = run_all_checks(self._laminates)
        except Exception as exc:  # pragma: no cover - defensive
            logger.error("Falha ao executar verificacoes de laminados: %s", exc, exc_info=True)
            self.failed.emit(str(exc))
        else:
            self.finished.emit(report)


def _normalize_orientation_for_summary(value: object) -> Optional[float]:
    if value is None:
        return None
    try:
        return normalize_angle(value)
    except Exception:
        try:
            return normalize_angle(str(value))
        except Exception:
            return None


def _build_auto_name_from_layers(
    layers: Iterable[Camada],
    *,
    model: Optional[GridModel] = None,
    tag: str = "",
    target: Optional[Laminado] = None,
) -> str:
    """Return automatic name using the shared auto-name helper."""
    layer_count = count_oriented_layers(layers)
    return auto_name_for_layers(
        model,
        layer_count=layer_count,
        tag=tag,
        target=target,
    )


def _load_icon_from_resources(
    resource_path: str, fallback_filename: str
) -> QIcon:
    icon = QIcon(resource_path)
    if icon.isNull():
        fallback_path = RESOURCES_ICONS_DIR / fallback_filename
        if fallback_path.is_file():
            icon = QIcon(str(fallback_path))
    return icon


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
        self._associated_cells_dialog: Optional[
            AssociatedCellsDialog
        ] = None
        self._new_laminate_dialog: Optional[NewLaminateDialog] = None
        self._virtual_stacking_window: Optional[VirtualStackingWindow] = None
        self._cell_neighbors_window: Optional[CellNeighborsWindow] = None
        self._new_laminate_button_icon: Optional[QIcon] = None
        self._new_laminate_icon_warning_emitted = False
        self._setting_new_laminate_name = False
        self._current_associated_cells: list[str] = []
        self._selection_column_index = StackingTableModel.COL_SELECT
        self._stacking_header_band: Optional[QWidget] = None
        self._band_labels: list[QLabel] = []
        self._auto_rename_guard = False
        self._header_band_mapping: list[int] = [
            StackingTableModel.COL_NUMBER,
            StackingTableModel.COL_SELECT,
            StackingTableModel.COL_SEQUENCE,
            StackingTableModel.COL_PLY,
            StackingTableModel.COL_PLY_TYPE,
            StackingTableModel.COL_MATERIAL,
            StackingTableModel.COL_ORIENTATION,
        ]
        self._band_frame_margin = 0
        self._header_band_scroll_connected = False
        self._stacking_summary_model: Optional[StackingTableModel] = None
        self._material_sync_guard = False
        self._material_sync_models: set[int] = set()
        self._export_checks_thread: Optional[QThread] = None
        self._export_checks_worker: Optional[_LaminateChecksWorker] = None
        self._last_checks_report: Optional[ChecksReport] = None
        self.undo_stack = QUndoStack(self)
        self._undo_shortcuts: list[QShortcut] = []
        self._settings = QSettings("GridLamEdit", "GridLamEdit")
        self.stacking_summary_dialog = StackingSummaryDialog(self)

        self.ui_state = UiState.VIEW
        self._create_actions()
        self._setup_menu_bar()
        self._setup_central_widget()
        self._setup_status_bar()
        self._setup_undo_shortcuts()
        self._update_undo_buttons_state()
        self._update_save_actions_enabled()
        self._restore_stacking_summary_dialog_state()
        self._center_on_screen()

    def _file_dialog_options(self, *, force_qt_dialog: bool = False) -> QFileDialog.Options:
        """
        Return QFileDialog options.

        By default we keep the native Windows dialog (preferred layout).
        If freezes reappear, set env GRIDLAMEDIT_FORCE_QT_DIALOGS=1 to force the Qt dialog.
        """
        options = QFileDialog.Options()
        if force_qt_dialog or os.getenv("GRIDLAMEDIT_FORCE_QT_DIALOGS") == "1":
            options |= QFileDialog.Option.DontUseNativeDialog
        return options

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

    def _center_on_screen(self) -> None:
        handle = self.windowHandle()
        screen = handle.screen() if handle is not None else None
        if screen is None:
            screen = QGuiApplication.primaryScreen()
        if screen is None:
            return
        geo = self.frameGeometry()
        if not geo.isValid():
            geo = self.geometry()
        geo.moveCenter(screen.availableGeometry().center())
        self.move(geo.topLeft())

    def _create_actions(self) -> None:
        action_specs: List[
            tuple[str, str, callable, str, Optional[QKeySequence]]
        ] = [
            (
                "open_project_action",
                "Open Project",
                self._on_open_project,
                "Open a GridLam project file.",
                QKeySequence.Open,
            ),
            (
                "load_spreadsheet_action",
                "Load Grid Spreadsheet",
                self._load_spreadsheet,
                "Import a Grid Design spreadsheet.",
                None,
            ),
            (
                "batch_import_action",
                "Import Laminates in Batch",
                self._on_batch_import_laminates,
                "Fill and import laminates using the Excel template.",
                None,
            ),
            (
                "save_action",
                "Save",
                self._on_save_triggered,
                "Save changes to the current project.",
                QKeySequence.Save,
            ),
            (
                "save_as_action",
                "Save As",
                self._on_save_as_triggered,
                "Save the project to a new file.",
                QKeySequence.SaveAs,
            ),
            (
                "export_excel_action",
                "Export Grid Spreadsheet",
                self._on_export_excel,
                "Export an Excel spreadsheet with the current changes.",
                QKeySequence("Ctrl+E"),
            ),
            (
                "register_material_action",
                "Register Material...",
                self._on_register_material,
                "Register a new standard material.",
                None,
            ),
            (
                "virtual_stacking_action",
                "Open Virtual Stacking",
                self.open_virtual_stacking,
                "Open the Virtual Stacking interface.",
                QKeySequence("Ctrl+Shift+V"),
            ),
            (
                "cell_neighbors_action",
                "Define Cell Neighbors",
                self.open_cell_neighbors,
                "Open the interface to define cell neighbors.",
                QKeySequence("Ctrl+Shift+N"),
            ),
            (
                "exit_action",
                "Close",
                self.close,
                "Close the application.",
                QKeySequence.Quit,
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
        file_menu.addAction(self.batch_import_action)
        file_menu.addSeparator()
        file_menu.addAction(self.save_action)
        file_menu.addAction(self.save_as_action)
        file_menu.addAction(self.export_excel_action)
        file_menu.addAction(self.register_material_action)
        file_menu.addSeparator()
        file_menu.addAction(self.exit_action)

        virtual_menu = menu_bar.addMenu("Virtual Stacking")
        virtual_menu.addAction(self.virtual_stacking_action)
        virtual_menu.addAction(self.cell_neighbors_action)

    def open_virtual_stacking(self) -> None:
        """Open the Virtual Stacking dialog populated with the current project."""
        if self._virtual_stacking_window is None:
            self._virtual_stacking_window = VirtualStackingWindow(
                self, undo_stack=self.undo_stack
            )
            try:
                self._virtual_stacking_window.stacking_changed.connect(
                    self._on_virtual_stacking_changed
                )
            except Exception:
                logger.debug("Nao foi possivel conectar sinal do Virtual Stacking.", exc_info=True)
            try:
                self._virtual_stacking_window.closed.connect(
                    self._on_virtual_stacking_closed
                )
            except Exception:
                logger.debug("Nao foi possivel conectar fechamento do Virtual Stacking.", exc_info=True)

        project = self._grid_model
        self._virtual_stacking_window.populate_from_project(project)
        self._virtual_stacking_window.show()
        self._virtual_stacking_window.raise_()
        self._virtual_stacking_window.activateWindow()
        self.hide()

    def open_cell_neighbors(self) -> None:
        """Open the Cell Neighbors dialog populated with the current project."""
        if self._cell_neighbors_window is None:
            self._cell_neighbors_window = CellNeighborsWindow(self)
        self._cell_neighbors_window.populate_from_project(self._grid_model, self.project_manager)
        self._cell_neighbors_window.show()
        self._cell_neighbors_window.raise_()
        self._cell_neighbors_window.activateWindow()

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
        layout.setSpacing(12)

        # Header row with title and laminate selector
        header_layout = QHBoxLayout()
        header_layout.setSpacing(12)
        
        header = QLabel("Laminado Associado a Celula", panel)
        header_font: QFont = header.font()
        header_font.setBold(True)
        header_font.setPointSize(header_font.pointSize() + 1)
        header.setFont(header_font)
        header_layout.addWidget(header)

        # Laminate selector (Trocar Laminado)
        selector_label = QLabel("Trocar Laminado:", panel)
        self.laminate_name_combo = QComboBox(panel)
        self.laminate_name_combo.setMinimumWidth(220)
        self.laminate_name_combo.setEditable(True)
        self.laminate_name_combo.setInsertPolicy(QComboBox.NoInsert)
        self._init_laminate_search_combo()
        # Connect signal to handle laminate change
        self.laminate_name_combo.activated.connect(self._on_laminate_combo_changed)
        
        header_layout.addWidget(selector_label)
        header_layout.addWidget(self.laminate_name_combo)
        header_layout.addStretch()

        layout.addLayout(header_layout)
        layout.addLayout(self._build_laminate_form())
        associated_view = self._build_associated_cells_view()
        layout.addWidget(associated_view, alignment=Qt.AlignLeft)

        stacking_label = QLabel("Stacking", panel)
        stacking_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        layout.addWidget(stacking_label)
        layers_section = self._build_layers_section()
        layout.addWidget(layers_section, stretch=2)
        layout.setStretchFactor(layers_section, 2)

        return panel

    def _build_laminate_form(self) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.setSpacing(12)
        layout.setContentsMargins(0, 0, 0, 0)

        # Removed "Nome" combo from here as it is now in the header

        layout.addLayout(
            self._combo_with_label(
                "Cor:", (str(i) for i in range(1, 151)), "color", editable=False
            )
        )
        type_layout = self._combo_with_label(
            "Tipo:", ["Core", "Skin", "Custom"], "type"
        )
        layout.addLayout(type_layout)
        tag_layout = QHBoxLayout()
        tag_layout.setSpacing(6)
        tag_layout.setContentsMargins(0, 0, 0, 0)
        tag_label = QLabel("Tag:", self)
        self.laminate_tag_edit = QLineEdit(self)
        self.laminate_tag_edit.setPlaceholderText("Opcional")
        self.laminate_tag_edit.setMinimumWidth(140)
        self.laminate_tag_edit.editingFinished.connect(self._on_tag_changed)
        tag_layout.addWidget(tag_label)
        tag_layout.addWidget(self.laminate_tag_edit)
        layout.addLayout(tag_layout)
        self._attach_new_laminate_button(type_layout)
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

    def _init_laminate_search_combo(self) -> None:
        """Attach filtering behavior and search placeholder to the laminate combo."""
        combo = getattr(self, "laminate_name_combo", None)
        if not isinstance(combo, QComboBox):
            return

        source_model = QStandardItemModel(combo)
        proxy_model = LaminateFilterProxy(combo, sentinel=NO_LAMINATE_COMBO_OPTION)
        proxy_model.setSourceModel(source_model)
        proxy_model.setFilterKeyColumn(0)

        combo.setModel(proxy_model)
        combo.setModelColumn(0)

        line_edit = combo.lineEdit()
        if line_edit is not None:
            line_edit.setPlaceholderText("Buscar laminado...")
            line_edit.textChanged.connect(self._filter_laminate_combo)
            line_edit.returnPressed.connect(self._select_first_visible_laminate)

        self._laminate_source_model = source_model
        self._laminate_filter_model = proxy_model
        self._reset_laminate_filter(clear_text=True)

    def _filter_laminate_combo(self, text: str) -> None:
        proxy = getattr(self, "_laminate_filter_model", None)
        if isinstance(proxy, LaminateFilterProxy):
            proxy.set_filter_text(text)
        combo = getattr(self, "laminate_name_combo", None)
        if isinstance(combo, QComboBox) and text.strip():
            view = combo.view()
            if view is not None and not view.isVisible():
                combo.showPopup()

    def _reset_laminate_filter(self, *, clear_text: bool = False) -> None:
        proxy = getattr(self, "_laminate_filter_model", None)
        combo = getattr(self, "laminate_name_combo", None)
        if isinstance(proxy, LaminateFilterProxy):
            proxy.set_filter_text("")
        if clear_text and isinstance(combo, QComboBox):
            line_edit = combo.lineEdit()
            if line_edit is not None:
                line_edit.blockSignals(True)
                line_edit.clear()
                line_edit.blockSignals(False)

    def _clear_laminate_combo_display(self) -> None:
        """Show an empty search field while keeping the sentinel option in the list."""
        combo = getattr(self, "laminate_name_combo", None)
        if not isinstance(combo, QComboBox):
            return
        self._reset_laminate_filter(clear_text=True)
        combo.blockSignals(True)
        combo.setCurrentIndex(-1)  # No pre-selected item; dropdown still shows all entries
        combo.blockSignals(False)

    def _select_first_visible_laminate(self) -> None:
        combo = getattr(self, "laminate_name_combo", None)
        proxy = getattr(self, "_laminate_filter_model", None)
        if not isinstance(combo, QComboBox) or not isinstance(proxy, LaminateFilterProxy):
            return
        for row in range(proxy.rowCount()):
            idx = proxy.index(row, 0)
            text = str(idx.data() or "")
            if text and text != NO_LAMINATE_COMBO_OPTION:
                combo.blockSignals(True)
                combo.setCurrentIndex(row)
                combo.blockSignals(False)
                self._on_laminate_combo_changed(row)
                break

    def _set_laminate_combo_selection(self, name: Optional[str]) -> None:
        combo = getattr(self, "laminate_name_combo", None)
        proxy = getattr(self, "_laminate_filter_model", None)
        source = getattr(self, "_laminate_source_model", None)
        target = name or NO_LAMINATE_COMBO_OPTION

        if not isinstance(combo, QComboBox):
            return
        if isinstance(source, QStandardItemModel) and isinstance(proxy, LaminateFilterProxy):
            self._reset_laminate_filter(clear_text=True)
            match_source_idx = None
            for row in range(source.rowCount()):
                idx = source.index(row, 0)
                if str(idx.data() or "") == target:
                    match_source_idx = idx
                    break
            if match_source_idx is None:
                return
            proxy_idx = proxy.mapFromSource(match_source_idx)
            if proxy_idx.isValid():
                combo.blockSignals(True)
                combo.setCurrentIndex(proxy_idx.row())
                combo.blockSignals(False)
            return

        idx = combo.findText(target)
        if idx >= 0:
            combo.blockSignals(True)
            combo.setCurrentIndex(idx)
            combo.blockSignals(False)

    # Automatic rename helpers --------------------------------------------------

    def _color_token(self, value: object) -> str:
        if isinstance(value, int):
            return str(value)
        text = str(value or "").strip()
        if not text:
            return str(DEFAULT_COLOR_INDEX)
        return text.upper()

    def _set_color_combo_value(self, laminate: Optional[Laminado]) -> None:
        combo = getattr(self, "laminate_color_combo", None)
        if not isinstance(combo, QComboBox) or laminate is None:
            return
        token = self._color_token(getattr(laminate, "color_index", DEFAULT_COLOR_INDEX))
        combo.blockSignals(True)
        idx = combo.findText(token)
        if idx < 0:
            combo.addItem(token)
            idx = combo.findText(token)
        combo.setCurrentIndex(idx if idx >= 0 else 0)
        combo.blockSignals(False)

    def _ensure_unique_laminate_color(self, laminate: Laminado) -> None:
        if self._grid_model is None:
            return
        used = {
            self._color_token(other.color_index)
            for other in self._grid_model.laminados.values()
            if other is not laminate
        }
        current = self._color_token(laminate.color_index)
        if current and current not in used:
            return
        for idx in range(MIN_COLOR_INDEX, MAX_COLOR_INDEX + 1):
            token = str(idx)
            if token not in used:
                laminate.color_index = idx
                self._set_color_combo_value(laminate)
                return
        for _ in range(32):
            token = f"#{secrets.token_hex(3).upper()}"
            if token not in used:
                laminate.color_index = token
                self._set_color_combo_value(laminate)
                return
        logger.warning(
            "Nao foi possivel atribuir uma cor unica ao laminado '%s'.",
            laminate.nome,
        )

    def _rename_laminate(self, laminate: Laminado, new_name: str) -> None:
        if self._grid_model is None:
            return
        old_name = laminate.nome
        if not new_name or new_name == old_name:
            return
        laminados = self._grid_model.laminados
        if (
            new_name in laminados
            and laminados[new_name] is not laminate
        ):
            return
        updated = OrderedDict()
        for name, lam in laminados.items():
            if name == old_name:
                updated[new_name] = lam
            else:
                updated[name] = lam
        self._grid_model.laminados = updated
        laminate.nome = new_name
        for cell_id, mapped in list(self._grid_model.cell_to_laminate.items()):
            if mapped == old_name:
                self._grid_model.cell_to_laminate[cell_id] = new_name
        binding = getattr(self, "_grid_binding", None)
        if binding is not None:
            if getattr(binding, "_current_laminate", None) == old_name:
                binding._current_laminate = new_name  # type: ignore[attr-defined]
            if hasattr(binding, "_build_cell_index"):
                binding._laminates_by_cell = binding._build_cell_index()  # type: ignore[attr-defined]
            for cell_id, mapped in self._grid_model.cell_to_laminate.items():
                if mapped == new_name:
                    binding._refresh_cell_item_label(cell_id)  # type: ignore[attr-defined]
        self._refresh_cells_list_labels()
        combo = getattr(self, "laminate_name_combo", None)
        if isinstance(combo, QComboBox):
            self._reset_laminate_filter(clear_text=True)
            combo.blockSignals(True)
            idx = combo.findText(old_name)
            if idx >= 0:
                combo.setItemText(idx, new_name)
                combo.setCurrentIndex(idx)
            combo.blockSignals(False)
        self._update_window_title()

    def _refresh_cells_list_labels(self) -> None:
        """Atualiza os r├│tulos das c├®lulas para refletir nomes atuais dos laminados."""
        if self._grid_model is None:
            return
        list_widget = getattr(self, "lstCelulas", None)
        if not isinstance(list_widget, QListWidget):
            list_widget = getattr(self, "cells_list", None)
        if not isinstance(list_widget, QListWidget):
            return
        list_widget.blockSignals(True)
        for idx in range(list_widget.count()):
            item = list_widget.item(idx)
            cell_id = item.data(Qt.UserRole)
            if cell_id:
                item.setText(_format_cell_label(self._grid_model, cell_id))
        list_widget.blockSignals(False)

    def _apply_auto_rename_if_needed(
        self, laminate: Optional[Laminado], *, force: bool = False
    ) -> None:
        if (
            laminate is None
            or self._grid_model is None
        ):
            return
        laminate.auto_rename_enabled = True
        if self._auto_rename_guard:
            return
        self._auto_rename_guard = True
        try:
            auto_name = auto_name_for_laminate(self._grid_model, laminate)
            changed = False
            if force or auto_name != laminate.nome:
                before = laminate.nome
                self._rename_laminate(laminate, auto_name)
                if laminate.nome != before:
                    changed = True
            before_color = self._color_token(laminate.color_index)
            self._ensure_unique_laminate_color(laminate)
            if self._color_token(laminate.color_index) != before_color:
                changed = True
            self._set_color_combo_value(laminate)
            self._update_auto_rename_controls(laminate)
            if changed:
                self._mark_dirty()
        finally:
            self._auto_rename_guard = False

    def _update_auto_rename_controls(
        self, laminate: Optional[Laminado]
    ) -> None:
        checkbox = getattr(self, "auto_rename_checkbox", None)
        color_combo = getattr(self, "laminate_color_combo", None)
        name_combo = getattr(self, "laminate_name_combo", None)
        enabled = True
        if isinstance(checkbox, QCheckBox):
            checkbox.setChecked(True)
            checkbox.setEnabled(False)
            checkbox.hide()
        if isinstance(name_combo, QComboBox):
            name_combo.setEditable(True)
            line_edit = name_combo.lineEdit()
            if line_edit is not None:
                line_edit.setReadOnly(False)
        if isinstance(color_combo, QComboBox):
            color_combo.setEnabled(not enabled)

    def _on_auto_rename_checkbox_toggled(self, checked: bool) -> None:
        laminate = self._current_laminate_instance()
        if laminate is None:
            return
        laminate.auto_rename_enabled = True
        self._apply_auto_rename_if_needed(laminate, force=True)
        self._mark_dirty()

    def _on_manual_name_edited(self) -> None:
        laminate = self._current_laminate_instance()
        combo = getattr(self, "laminate_name_combo", None)
        if laminate is None or not isinstance(combo, QComboBox):
            return
        line_edit = combo.lineEdit()
        if line_edit is None:
            return
        line_edit.setText(laminate.nome)

    def _on_tag_changed(self) -> None:
        laminate = self._current_laminate_instance()
        edit = getattr(self, "laminate_tag_edit", None)
        if laminate is None or not isinstance(edit, QLineEdit):
            return
        new_tag = edit.text().strip()
        if getattr(laminate, "tag", "") == new_tag:
            return
        laminate.tag = new_tag
        self._mark_dirty()
        self._apply_auto_rename_if_needed(laminate, force=True)

    def _on_binding_laminate_changed(
        self, laminate_name: Optional[str]
    ) -> None:
        laminate: Optional[Laminado] = None
        if self._grid_model is not None and laminate_name:
            laminate = self._grid_model.laminados.get(laminate_name)
        if laminate is None:
            laminate = self._current_laminate_instance()
        self._update_auto_rename_controls(laminate)
        self._set_color_combo_value(laminate)
        self._apply_auto_rename_if_needed(laminate)
        if laminate is not None:
            self.check_symmetry()

    def _on_binding_layers_modified(
        self, laminate_name: Optional[str]
    ) -> None:
        if self._grid_model is None:
            return
        if laminate_name:
            laminate = self._grid_model.laminados.get(laminate_name)
        else:
            laminate = self._current_laminate_instance()
        self._apply_auto_rename_if_needed(laminate)
        self.check_symmetry()
        self._refresh_virtual_stacking_view()
        self.check_symmetry()

    def _on_virtual_stacking_changed(self, laminate_names: list[str]) -> None:
        if self._grid_model is None:
            return
        binding, model, current_laminate = self._stacking_binding_context()
        if model is not None and current_laminate is not None:
            if current_laminate.nome in laminate_names:
                model.update_layers(current_laminate.camadas)
        for name in laminate_names:
            laminate = self._grid_model.laminados.get(name)
            if laminate is not None:
                self._apply_auto_rename_if_needed(laminate, force=True)
        self._refresh_cells_list_labels()
        self.update_stacking_summary_ui()
        self._mark_dirty()

    def _on_virtual_stacking_closed(self) -> None:
        if self._grid_model is not None:
            try:
                laminate_names = list(self._grid_model.laminados.keys())
            except Exception:
                laminate_names = []
            if laminate_names:
                self._on_virtual_stacking_changed(laminate_names)
        self.show()
        self.raise_()
        self.activateWindow()

    def _sync_all_auto_renamed_laminates(self) -> None:
        if self._grid_model is None:
            return
        for laminate in self._grid_model.laminados.values():
            self._apply_auto_rename_if_needed(laminate, force=True)

    def _on_new_laminate_auto_rename_toggled(self, checked: bool) -> None:
        name_edit = getattr(self, "new_laminate_name_edit", None)
        color_combo = getattr(self, "new_laminate_color_combo", None)
        if isinstance(name_edit, QLineEdit):
            name_edit.setReadOnly(True)
        if isinstance(color_combo, QComboBox):
            color_combo.setEnabled(False)
        self._update_new_laminate_auto_name()

    def _collect_new_laminate_layers_for_auto_name(self) -> list[Camada]:
        layers: list[Camada] = []
        table = getattr(self, "new_laminate_stacking_table", None)
        if not isinstance(table, QTableWidget):
            return layers
        for row in range(table.rowCount()):
            orientation_item = table.item(row, 1)
            orientation_text = orientation_item.text() if orientation_item else ""
            orientacao = None
            if orientation_text.strip():
                try:
                    orientacao = normalize_angle(orientation_text)
                except ValueError:
                    orientacao = None
            material_item = table.item(row, 0)
            material = material_item.text() if material_item else ""
            layers.append(
                Camada(
                    idx=row,
                    material=material,
                    orientacao=orientacao,
                    ativo=True,
                    simetria=False,
                )
            )
        return layers

    def _update_new_laminate_auto_name(self) -> None:
        checkbox = getattr(self, "new_laminate_auto_rename_checkbox", None)
        name_edit = getattr(self, "new_laminate_name_edit", None)
        if (
            not isinstance(checkbox, QCheckBox)
            or not checkbox.isChecked()
            or not isinstance(name_edit, QLineEdit)
        ):
            return
        if self._setting_new_laminate_name:
            return
        layers = self._collect_new_laminate_layers_for_auto_name()
        tag_text = ""
        tag_edit = getattr(self, "new_laminate_tag_edit", None)
        if hasattr(tag_edit, "text"):
            tag_text = tag_edit.text()
        auto_name = _build_auto_name_from_layers(
            layers,
            model=self._grid_model,
            tag=tag_text,
            target=None,
        )
        self._setting_new_laminate_name = True
        try:
            name_edit.setText(auto_name)
        finally:
            self._setting_new_laminate_name = False

    def _attach_new_laminate_button(self, container: QHBoxLayout) -> None:
        button = QPushButton("Novo Laminado", self)
        button.setObjectName("btnNovoLaminado")
        button.setAccessibleName("Novo Laminado")
        button.setToolTip("Criar um novo laminado")
        button.setIcon(self._load_new_laminate_button_icon())
        button.setIconSize(QSize(20, 20))
        button.setMinimumSize(28, 28)
        button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        button.clicked.connect(self._open_new_laminate_dialog)
        container.addWidget(button)
        self.btn_new_laminate = button

    def _load_new_laminate_button_icon(self) -> QIcon:
        if self._new_laminate_button_icon is not None:
            return self._new_laminate_button_icon
        candidates = [
            package_path("assets", "icons", "Criar_novo_laminado_ControlV.png"),
            RESOURCES_ICONS_DIR / "Criar_novo_laminado_ControlV.png",
            RESOURCES_ICONS_DIR / "Criar_novo_laminado_ControlV.jpg",
        ]
        for candidate in candidates:
            if candidate.is_file():
                icon = QIcon(str(candidate))
                if not icon.isNull():
                    self._new_laminate_button_icon = icon
                    return icon
        if not self._new_laminate_icon_warning_emitted:
            logger.warning(
                "Icone Criar_novo_laminado_ControlV.* nao encontrado; usando icone padrao."
            )
            self._new_laminate_icon_warning_emitted = True
        fallback = self.style().standardIcon(QStyle.SP_FileDialogNewFolder)
        self._new_laminate_button_icon = fallback
        return fallback

    def _laminate_color_options(self) -> list[str]:
        return [str(idx) for idx in range(MIN_COLOR_INDEX, MAX_COLOR_INDEX + 1)]

    def _laminate_type_options(self) -> list[str]:
        defaults = ["SS", "Core", "Skin", "Custom"]
        if not self._grid_model or not self._grid_model.laminados:
            return defaults
        ordered: list[str] = []
        for item in defaults:
            if item not in ordered:
                ordered.append(item)
        for laminado in self._grid_model.laminados.values():
            tipo = (laminado.tipo or "").strip()
            if tipo and tipo not in ordered:
                ordered.append(tipo)
        return ordered

    def available_materials(self) -> list[str]:
        """Return materials merging defaults, user entries, and project data."""
        return registry_available_materials(self._grid_model, settings=self._settings)

    def _available_cells(self) -> list[str]:
        if self._grid_model is None:
            return []
        return list(self._grid_model.celulas_ordenadas or [])

    def _get_new_laminate_dialog(self) -> NewLaminateDialog:
        if self._new_laminate_dialog is None:
            model = self._grid_model or GridModel()
            self._new_laminate_dialog = NewLaminateDialog(
                model,
                color_options=self._laminate_color_options(),
                type_options=self._laminate_type_options(),
                cell_options=self._available_cells(),
                parent=self,
            )
        return self._new_laminate_dialog

    def _open_new_laminate_dialog(self) -> None:
        if self._grid_model is None:
            self._grid_model = GridModel()
        cells = self._available_cells()
        if not cells:
            QMessageBox.warning(
                self,
                "C├®lulas indispon├¡veis",
                "Nenhuma c├®lula foi carregada para associar ao novo laminado.",
            )
            return
        dialog = self._get_new_laminate_dialog()
        dialog.set_grid_model(self._grid_model)
        dialog.refresh_options(
            color_options=self._laminate_color_options(),
            type_options=self._laminate_type_options(),
            cell_options=cells,
        )
        dialog.reset_fields()
        if dialog.exec() == QDialog.Accepted and dialog.created_laminate is not None:
            self._refresh_after_new_laminate(dialog.created_laminate.nome)
            self._mark_dirty()

    def _build_associated_cells_view(self) -> QWidget:
        container = QWidget(self)
        layout = QHBoxLayout(container)
        layout.setSpacing(12)
        layout.setContentsMargins(0, 0, 0, 0)

        self.btn_associated_cells = QPushButton(
            "Celulas associadas com esse laminado", container
        )
        self.btn_associated_cells.setObjectName("btn_associated_cells")
        self.btn_associated_cells.setToolTip(
            "Abrir lista de celulas associadas ao laminado atual"
        )
        self.btn_associated_cells.clicked.connect(
            self._open_associated_cells_dialog
        )

        layout.addWidget(self.btn_associated_cells, alignment=Qt.AlignLeft)
        layout.addStretch()

        self.layers_count_label = QLabel("Quantidade Total de Camadas: 0", container)
        self.layers_count_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.layers_count_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        layout.addWidget(self.layers_count_label, alignment=Qt.AlignRight)

        container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        return container

    def _build_layers_section(self) -> QWidget:
        container = QWidget(self)
        outer_layout = QVBoxLayout(container)
        outer_layout.setSpacing(12)
        outer_layout.setContentsMargins(0, 0, 0, 0)

        # Main content row: table + buttons
        content_row = QHBoxLayout()
        content_row.setSpacing(12)
        content_row.setContentsMargins(0, 0, 0, 0)

        table_layout = QVBoxLayout()
        table_layout.setSpacing(6)
        table_layout.setContentsMargins(0, 0, 0, 0)

        self.layers_table = self._create_layers_table(container)
        header_widget = self._create_layers_header_widget(container)
        table_layout.addWidget(header_widget)
        table_layout.addWidget(self.layers_table, stretch=1)
        table_layout.setStretch(0, 0)
        table_layout.setStretch(1, 1)

        content_row.addLayout(table_layout, stretch=1)
        content_row.addLayout(self._create_layers_buttons())
        outer_layout.addLayout(content_row)

        # No footer here; footer is placed at the laminate panel level

        container.setMinimumHeight(0)
        container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
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
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        table.setVerticalScrollMode(QAbstractItemView.ScrollPerItem)
        table.setViewportMargins(0, 0, 0, 8)  # leave breathing room for the last row
        table.verticalHeader().setVisible(False)
        table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        table.setMinimumHeight(440)
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

        titles = [
            "#",
            "Selection",
            "Sequence",
            "Ply",
            "Simetria",
            "Material",
            "Orienta├º├úo",
        ]
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
                    if icon.isNull():
                        fallback_name = icon_name.rsplit("/", 1)[-1]
                        icon = _load_icon_from_resources(
                            icon_name, fallback_name
                        )
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

        self.btn_icon_new_laminate_from_paste = make_button(
            "",
            "Criar novo laminado por colagem (Ctrl+V)",
            self.on_new_laminate_from_paste,
            "Criar novo laminado por colagem (Ctrl+V)",
            tool_button_style=Qt.ToolButtonIconOnly,
        )
        self.btn_icon_new_laminate_from_paste.setObjectName(
            "btn_icon_new_laminate_from_paste"
        )
        self.btn_icon_new_laminate_from_paste.setIcon(
            _load_icon_from_resources(
                ":/icons/Criar_novo_laminado_ControlV.jpg",
                "Criar_novo_laminado_ControlV.jpg",
            )
        )
        self.btn_icon_new_laminate_from_paste.setToolTip(
            "Criar novo laminado por colagem (Ctrl+V)"
        )
        self.btn_icon_new_laminate_from_paste.setIconSize(QSize(24, 24))
        self.btn_new_laminate_from_paste = self.btn_icon_new_laminate_from_paste

        self.btn_duplicate_laminate = make_button(
            "",
            "Duplicar laminado existente",
            self._open_duplicate_laminate_dialog,
            "Duplicar laminado existente",
            tool_button_style=Qt.ToolButtonIconOnly,
            fixed_width=42,
        )
        duplicate_icon = _load_icon_from_resources(
            ":/icons/duplicar.png", "duplicar.png"
        )
        if duplicate_icon.isNull():
            duplicate_icon = self.style().standardIcon(
                QStyle.SP_FileDialogDetailedView
            )
        self.btn_duplicate_laminate.setIcon(duplicate_icon)
        self.btn_duplicate_laminate.setIconSize(QSize(20, 20))
        self.btn_duplicate_laminate.setMinimumSize(QSize(28, 28))
        self.btn_duplicate_laminate.setObjectName("btn_duplicate_laminate")

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
        self.btn_renumber_sequence = make_button(
            ":/icons/renumber_sequence.svg",
            "Renumerar sequ\u00eancia (Seq.1, Seq.2...)",
            self.on_renumber_sequences,
            "Renumerar sequ\u00eancia das camadas",
            QStyle.SP_BrowserReload,
            tool_button_style=Qt.ToolButtonIconOnly,
        )
        self.btn_bulk_change_material = QToolButton(self)
        self.btn_bulk_change_material.setObjectName("btn_bulk_change_material")
        self.btn_bulk_change_material.setToolButtonStyle(Qt.ToolButtonIconOnly)
        self.btn_bulk_change_material.setAutoRaise(True)
        self.btn_bulk_change_material.setFixedWidth(42)
        self.btn_bulk_change_material.setIcon(
            _load_icon_from_resources(
                ":/icons/Trocar_Materiais.jpg", "Trocar_Materiais.jpg"
            )
        )
        self.btn_bulk_change_material.setIconSize(QSize(24, 24))
        self.btn_bulk_change_material.setToolTip(
            "Trocar material das camadas selecionadas"
        )
        self.btn_bulk_change_material.setAccessibleName(
            "Trocar material das camadas selecionadas"
        )
        self.btn_bulk_change_material.clicked.connect(
            self.on_bulk_change_material
        )
        self.layer_buttons.append(self.btn_bulk_change_material)
        layout.addWidget(self.btn_bulk_change_material)

        self.btn_bulk_change_orientation = QToolButton(self)
        self.btn_bulk_change_orientation.setObjectName(
            "btn_bulk_change_orientation"
        )
        self.btn_bulk_change_orientation.setToolButtonStyle(
            Qt.ToolButtonIconOnly
        )
        self.btn_bulk_change_orientation.setAutoRaise(True)
        self.btn_bulk_change_orientation.setFixedWidth(42)
        self.btn_bulk_change_orientation.setIcon(
            _load_icon_from_resources(
                ":/icons/Trocar_Orientacao.jpg", "Trocar_Orientacao.jpg"
            )
        )
        self.btn_bulk_change_orientation.setIconSize(QSize(24, 24))
        self.btn_bulk_change_orientation.setToolTip(
            "Trocar orienta├º├úo das camadas selecionadas"
        )
        self.btn_bulk_change_orientation.setAccessibleName(
            "Trocar orienta├º├úo das camadas selecionadas"
        )
        self.btn_bulk_change_orientation.clicked.connect(
            self.on_bulk_change_orientation
        )
        self.layer_buttons.append(self.btn_bulk_change_orientation)
        layout.addWidget(self.btn_bulk_change_orientation)

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

        self.delete_layers_button = make_button(
            "trash.svg",
            "Excluir camadas selecionadas",
            self._on_delete_layers_clicked,
            "Excluir camadas selecionadas",
            QStyle.SP_TrashIcon,
        )

        self.select_all_layers_button = make_button(
            None,
            "Selecionar todos",
            self._on_select_all_layers_clicked,
            "Selecionar todos",
            QStyle.SP_DialogYesButton,
            tool_button_style=Qt.ToolButtonTextBesideIcon,
        )
        self.select_all_layers_button.setIconSize(QSize(20, 20))

        self.clear_selection_button = make_button(
            None,
            "Limpar sele├º├úo",
            self._on_clear_selection_clicked,
            "Limpar sele├º├úo",
            QStyle.SP_DialogResetButton,
            tool_button_style=Qt.ToolButtonTextBesideIcon,
        )
        self.clear_selection_button.setIconSize(QSize(20, 20))

        self.btn_undo = make_button(
            ":/icons/undo.svg",
            "Desfazer (Ctrl+Z)",
            self.undo_stack.undo,
            "Desfazer (Ctrl+Z)",
        )
        self.btn_undo.setObjectName("btn_undo")
        self.btn_undo.setEnabled(False)

        self.btn_redo = make_button(
            ":/icons/redo.svg",
            "Refazer (Ctrl+Y)",
            self.undo_stack.redo,
            "Refazer (Ctrl+Y)",
        )
        self.btn_redo.setObjectName("btn_redo")
        self.btn_redo.setEnabled(False)

        self.btn_show_stacking_summary = make_button(
            ":/icons/stacking_summary.svg",
            "Abrir Resumo do Stacking",
            self._show_stacking_summary,
            "Abrir Resumo do Stacking",
        )
        self.btn_show_stacking_summary.setObjectName(
            "btn_show_stacking_summary"
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
            self._connect_material_sync(model)
            self._connect_header_band_model_signals(model)
            self._set_stacking_summary_model(model)
        else:
            self._set_stacking_summary_model(None)

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
        header.setSectionResizeMode(StackingTableModel.COL_SEQUENCE, QHeaderView.Fixed)
        header.setSectionResizeMode(StackingTableModel.COL_PLY, QHeaderView.Fixed)
        header.setSectionResizeMode(StackingTableModel.COL_PLY_TYPE, QHeaderView.Fixed)
        header.setSectionResizeMode(StackingTableModel.COL_MATERIAL, QHeaderView.Stretch)
        header.setSectionResizeMode(
            StackingTableModel.COL_ORIENTATION, QHeaderView.Stretch
        )
        header.setMinimumSectionSize(60)
        header.setFixedHeight(max(header.height(), header.sizeHint().height()))

        view.setColumnWidth(StackingTableModel.COL_NUMBER, 60)
        view.setColumnWidth(StackingTableModel.COL_SELECT, 120)
        view.setColumnWidth(StackingTableModel.COL_SEQUENCE, 110)
        view.setColumnWidth(StackingTableModel.COL_PLY, 90)
        view.setColumnWidth(StackingTableModel.COL_PLY_TYPE, 160)
        view.verticalHeader().setVisible(False)
        self._sync_header_band()

    def _connect_material_sync(self, model: StackingTableModel) -> None:
        key = id(model)
        if key in self._material_sync_models:
            return
        model.dataChanged.connect(self._on_stacking_material_changed)
        self._material_sync_models.add(key)

    def _on_stacking_material_changed(self, top_left, bottom_right, roles=None) -> None:  # noqa: ARG002
        if self._material_sync_guard or self._grid_model is None:
            return
        if bottom_right.column() < StackingTableModel.COL_MATERIAL or top_left.column() > StackingTableModel.COL_MATERIAL:
            return
        binding, model, current_laminate = self._stacking_binding_context()
        if model is None or current_laminate is None:
            return

        def provider(lam: Laminado) -> Optional[StackingTableModel]:
            if getattr(lam, "nome", None) == getattr(current_laminate, "nome", None):
                return model
            return None

        changed_any = False
        self._material_sync_guard = True
        try:
            for row in range(top_left.row(), bottom_right.row() + 1):
                idx = model.index(row, StackingTableModel.COL_MATERIAL)
                if not idx.isValid():
                    continue
                material = str(model.data(idx, Qt.EditRole) or "").strip()
                updated = sync_material_by_sequence(
                    self._grid_model,
                    row,
                    material,
                    stacking_model_provider=provider,
                )
                if updated:
                    changed_any = True
        finally:
            self._material_sync_guard = False

        if changed_any:
            self._refresh_virtual_stacking_view()
            self.update_stacking_summary_ui()
            self._mark_dirty()

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
        for signal in (model.modelReset, model.layoutChanged):
            try:
                signal.connect(self._sync_header_band, Qt.UniqueConnection)
            except TypeError:
                # Already connected
                pass

    def _set_stacking_summary_model(
        self, model: Optional[StackingTableModel]
    ) -> None:
        current = getattr(self, "_stacking_summary_model", None)
        if current is model:
            self._stacking_summary_model = model
            self.update_stacking_summary_ui()
            return

        if isinstance(current, StackingTableModel):
            for signal in (
                current.dataChanged,
                current.rowsInserted,
                current.rowsRemoved,
            ):
                try:
                    signal.disconnect(self.update_stacking_summary_ui)
                except (TypeError, RuntimeError):
                    pass
            try:
                current.modelReset.disconnect(self._handle_stacking_model_reset)
            except (TypeError, RuntimeError):
                pass

        self._stacking_summary_model = model

        if isinstance(model, StackingTableModel):
            model.dataChanged.connect(self.update_stacking_summary_ui)
            model.rowsInserted.connect(self.update_stacking_summary_ui)
            model.rowsRemoved.connect(self.update_stacking_summary_ui)
            model.modelReset.connect(self._handle_stacking_model_reset)

        self.update_stacking_summary_ui()

    def _stacking_binding_context(
        self,
    ) -> tuple[Optional[object], Optional[StackingTableModel], Optional[Laminado]]:
        binding = getattr(self, "_grid_binding", None)
        model: Optional[StackingTableModel] = None
        laminate: Optional[Laminado] = None
        if binding is not None:
            candidate_model = getattr(binding, "stacking_model", None)
            if isinstance(candidate_model, StackingTableModel):
                model = candidate_model
            current_name = getattr(binding, "_current_laminate", None)
            grid_model = getattr(binding, "model", None)
            if current_name and grid_model is not None:
                laminate = grid_model.laminados.get(current_name)
        return binding, model, laminate

    def _handle_stacking_model_reset(self) -> None:
        self._clear_undo_history()
        self.update_stacking_summary_ui()

    def _setup_undo_shortcuts(self) -> None:
        if not hasattr(self, "undo_stack"):
            return
        for shortcut in self._undo_shortcuts:
            shortcut.setParent(None)
        self._undo_shortcuts.clear()
        undo_shortcut = QShortcut(QKeySequence("Ctrl+Z"), self)
        undo_shortcut.activated.connect(self.undo_stack.undo)
        self._undo_shortcuts.append(undo_shortcut)
        for sequence in ("Ctrl+Y", "Ctrl+Shift+Z"):
            redo_shortcut = QShortcut(QKeySequence(sequence), self)
            redo_shortcut.activated.connect(self.undo_stack.redo)
            self._undo_shortcuts.append(redo_shortcut)
        self.undo_stack.canUndoChanged.connect(self._update_undo_buttons_state)
        self.undo_stack.canRedoChanged.connect(self._update_undo_buttons_state)

    def _update_undo_buttons_state(self) -> None:
        btn_undo = getattr(self, "btn_undo", None)
        btn_redo = getattr(self, "btn_redo", None)
        can_undo = self.undo_stack.canUndo() if hasattr(self, "undo_stack") else False
        can_redo = self.undo_stack.canRedo() if hasattr(self, "undo_stack") else False
        if isinstance(btn_undo, QAbstractButton):
            btn_undo.setEnabled(can_undo)
        if isinstance(btn_redo, QAbstractButton):
            btn_redo.setEnabled(can_redo)

    def _clear_undo_history(self) -> None:
        if hasattr(self, "undo_stack"):
            self.undo_stack.blockSignals(True)
            self.undo_stack.clear()
            self.undo_stack.blockSignals(False)
            self._update_undo_buttons_state()

    def _current_laminate_instance(self) -> Optional[Laminado]:
        binding, _, laminate = self._stacking_binding_context()
        if laminate is not None:
            return laminate
        if binding is not None:
            return None
        if self._grid_model is None:
            return None
        if not self._grid_model.laminados:
            return None
        if (
            hasattr(self, "laminate_name_combo")
            and self.laminate_name_combo.currentText()
        ):
            current_name = self.laminate_name_combo.currentText()
            return self._grid_model.laminados.get(current_name)
        return next(iter(self._grid_model.laminados.values()), None)

    def _render_stacking_summary(
        self, laminate: Optional[Laminado]
    ) -> str:
        layers = list(laminate.camadas) if laminate is not None else []
        oriented_layers = [camada for camada in layers if camada.orientacao is not None]
        total_layers = len(oriented_layers)
        structural_layers = sum(
            1
            for camada in oriented_layers
            if is_structural_ply_label(getattr(camada, "ply_type", DEFAULT_PLY_TYPE))
        )
        non_structural_layers = total_layers - structural_layers

        materials_counter: Counter[str] = Counter()
        orientations_counter: Counter[float] = Counter()

        for camada in oriented_layers:
            material_text = (camada.material or "").strip()
            if material_text:
                materials_counter[material_text] += 1
            normalized_orientation = _normalize_orientation_for_summary(
                camada.orientacao
            )
            if normalized_orientation is not None:
                orientations_counter[normalized_orientation] += 1

        def _pluralize(count: int) -> str:
            return "Ply" if count == 1 else "Plies"

        orientation_parts = "".join(
            f"[{count} {_pluralize(count)} a {format_orientation_value(angle)}]"
            for angle, count in sorted(orientations_counter.items(), key=lambda item: item[0])
        )
        materials_parts = "".join(
            f"[{count} {_pluralize(count)}({material})]"
            for material, count in sorted(
                materials_counter.items(), key=lambda item: (-item[1], item[0])
            )
        )

        if not orientation_parts:
            orientation_parts = "-"
        if not materials_parts:
            materials_parts = "-"

        lines = [
            f"Total de camadas:  {total_layers}",
            f"Camadas 'N\u00e3o Considerar':  {non_structural_layers}",
            f"Camadas 'Considerar':  {structural_layers}",
            "",
            "Orienta\u00e7\u00f5es:",
            orientation_parts,
            "",
            "Materiais:",
            materials_parts,
        ]
        return "\n".join(lines)

    def update_stacking_summary_ui(self, *args) -> None:  # noqa: ARG002
        dialog = getattr(self, "stacking_summary_dialog", None)
        if dialog is None or not dialog.isVisible():
            return
        laminate = self._current_laminate_instance()
        dialog.update_summary(self._render_stacking_summary(laminate))

    def _show_stacking_summary(self) -> None:
        dialog = getattr(self, "stacking_summary_dialog", None)
        if dialog is None:
            return
        laminate = self._current_laminate_instance()
        dialog.update_summary(self._render_stacking_summary(laminate))
        if dialog.isVisible():
            dialog.raise_()
            dialog.activateWindow()
        else:
            dialog.show()

    def _restore_stacking_summary_dialog_state(self) -> None:
        dialog = getattr(self, "stacking_summary_dialog", None)
        settings = getattr(self, "_settings", None)
        if dialog is None or settings is None:
            return
        geometry_value = settings.value("UI/StackingSummary/geometry")
        if isinstance(geometry_value, QByteArray):
            dialog.restoreGeometry(geometry_value)
        elif isinstance(geometry_value, (bytes, bytearray)):
            dialog.restoreGeometry(QByteArray(geometry_value))
        visible_value = settings.value("UI/StackingSummary/visible", False)
        if self._settings_value_to_bool(visible_value):
            self._show_stacking_summary()

    def _save_stacking_summary_dialog_state(self) -> None:
        dialog = getattr(self, "stacking_summary_dialog", None)
        settings = getattr(self, "_settings", None)
        if dialog is None or settings is None:
            return
        settings.setValue("UI/StackingSummary/geometry", dialog.saveGeometry())
        settings.setValue("UI/StackingSummary/visible", dialog.isVisible())

    @staticmethod
    def _settings_value_to_bool(
        value: object, default: bool = False
    ) -> bool:
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)):
            return bool(int(value))
        if isinstance(value, str):
            normalized = value.strip().lower()
            return normalized in {"1", "true", "yes", "on"}
        if isinstance(value, QByteArray):
            try:
                decoded = bytes(value).decode("utf-8")
            except Exception:
                return default
            return decoded.strip().lower() in {"1", "true", "yes", "on"}
        return default

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
        self.new_laminate_auto_rename_checkbox = QCheckBox(
            "Automatic Rename", view
        )
        self.new_laminate_auto_rename_checkbox.setChecked(True)
        self.new_laminate_auto_rename_checkbox.toggled.connect(
            self._on_new_laminate_auto_rename_toggled
        )
        self.new_laminate_auto_rename_checkbox.setEnabled(False)
        self.new_laminate_auto_rename_checkbox.hide()
        self.new_laminate_name_edit.setReadOnly(True)
        self._on_new_laminate_auto_rename_toggled(True)
        name_container = QHBoxLayout()
        name_container.setSpacing(6)
        name_container.addWidget(self.new_laminate_auto_rename_checkbox)
        name_container.addWidget(name_label)
        name_container.addWidget(self.new_laminate_name_edit, stretch=1)
        form_row.addLayout(name_container, stretch=1)

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

        tag_label = QLabel("Tag:", view)
        self.new_laminate_tag_edit = QLineEdit(view)
        self.new_laminate_tag_edit.setPlaceholderText("Opcional")
        self.new_laminate_tag_edit.textChanged.connect(
            lambda *_: self._update_new_laminate_auto_name()
        )
        form_row.addWidget(tag_label)
        form_row.addWidget(self.new_laminate_tag_edit)

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
        self.new_laminate_stacking_table.itemChanged.connect(
            self._on_new_laminate_item_changed
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
        if hasattr(self, "new_laminate_tag_edit"):
            self.new_laminate_tag_edit.clear()
        if hasattr(self, "new_laminate_color_combo"):
            default_idx = self.new_laminate_color_combo.findText(
                str(DEFAULT_COLOR_INDEX)
            )
            self.new_laminate_color_combo.setCurrentIndex(
                default_idx if default_idx >= 0 else 0
            )
        self.new_laminate_type_combo.setCurrentIndex(0)
        if hasattr(self, "new_laminate_auto_rename_checkbox"):
            self.new_laminate_auto_rename_checkbox.setChecked(True)

        table = self.new_laminate_stacking_table
        table.setRowCount(0)
        self._new_laminate_add_layer()
        table.setCurrentCell(0, 0)
        self._update_new_laminate_auto_name()

    def _new_laminate_add_layer(self) -> None:
        table = self.new_laminate_stacking_table
        row = table.rowCount()
        table.insertRow(row)
        self._apply_layer_row(table, row, ("", "Empty", True, False))
        table.setCurrentCell(row, 0)
        self._update_new_laminate_auto_name()

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
        self._update_new_laminate_auto_name()

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
        self._update_new_laminate_auto_name()

    def _collect_layer_row(
        self, table: QTableWidget, row: int
    ) -> tuple[str, str, bool, bool]:
        material = self._text(table.item(row, 0))
        orientation = self._text(table.item(row, 1))
        active = self._checkbox_value(table, row, 2)
        symmetry = self._checkbox_value(table, row, 3)
        return material, orientation, active, symmetry

    def _apply_orientation_highlight_item(self, item: Optional[QTableWidgetItem]) -> None:
        if item is None:
            return
        color = orientation_highlight_color(item.text())
        if color is None:
            item.setBackground(QColor())
        else:
            item.setBackground(color)
        text = (item.text() or "").strip().lower()
        if not text or text == "empty":
            if not item.text():
                item.setText("Empty")
            item.setForeground(QColor(160, 160, 160))
        else:
            item.setForeground(QColor())

    def _apply_layer_row(
        self,
        table: QTableWidget,
        row: int,
        data: tuple[str, str, bool, bool],
    ) -> None:
        material, orientation, active, symmetry = data
        table.setItem(row, 0, QTableWidgetItem(str(material)))
        orientation_item = QTableWidgetItem(str(orientation))
        table.setItem(row, 1, orientation_item)
        self._apply_orientation_highlight_item(orientation_item)

        active_checkbox = QCheckBox(table)
        active_checkbox.setChecked(active)
        table.setCellWidget(row, 2, self._wrap_checkbox(active_checkbox))

        symmetry_checkbox = QCheckBox(table)
        symmetry_checkbox.setChecked(symmetry)
        table.setCellWidget(row, 3, self._wrap_checkbox(symmetry_checkbox))

    def _on_new_laminate_item_changed(self, item: Optional[QTableWidgetItem]) -> None:
        if item is None:
            return
        table = getattr(self, "new_laminate_stacking_table", None)
        if not isinstance(table, QTableWidget):
            return
        row = item.row()
        column = item.column()
        if column == 1:
            self._apply_orientation_highlight_item(item)
            orientation_text = self._text(item)
            material_item = table.item(row, 0)
            if not orientation_text.strip() or orientation_text.strip().lower() == "empty":
                if material_item is not None:
                    material_item.setText("")
            elif material_item is not None and not self._text(material_item):
                suggestion = project_most_used_material(self._grid_model)
                if suggestion:
                    material_item.setText(suggestion)
        if column == 0:
            orientation_item = table.item(row, 1)
            if self._text(item) and (orientation_item is None or not self._text(orientation_item)):
                # Respect invariant: material requires orientation; clear material if orientation is empty.
                item.setText("")
                return
        self._update_new_laminate_auto_name()

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
            orientation_text = self._text(table.item(row, 1))
            if orientation_text.strip().lower() == "empty":
                orientation_text = ""
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
                    ply_label=f"Ply.{len(camadas) + 1}",
                    sequence=f"Seq.{len(camadas) + 1}",
                )
            )

        if not camadas:
            QMessageBox.warning(
                self,
                "Stacking obrigatorio",
                "Adicione ao menos uma camada ao laminado.",
            )
            return

        tag_text = ""
        tag_edit = getattr(self, "new_laminate_tag_edit", None)
        if isinstance(tag_edit, QLineEdit):
            tag_text = tag_edit.text().strip()
        laminado = Laminado(
            nome=name,
            tipo=tipo,
            color_index=color_index,
            tag=tag_text,
            celulas=[],
            camadas=camadas,
        )
        auto_checkbox = getattr(
            self, "new_laminate_auto_rename_checkbox", None
        )
        laminado.auto_rename_enabled = True

        self._grid_model.laminados[name] = laminado
        self._refresh_after_new_laminate(name)
        self._exit_creating_mode()
        self._mark_dirty()

    def _refresh_after_new_laminate(self, laminate_name: str) -> None:
        if self._grid_model is None:
            return
        self._clear_undo_history()
        self._clear_undo_history()
        self._clear_undo_history()
        bind_model_to_ui(self._grid_model, self)
        binding = getattr(self, "_grid_binding", None)
        if binding is not None:
            self._configure_stacking_table(binding)
        self._sync_all_auto_renamed_laminates()
        bind_cells_to_ui(self._grid_model, self)
        current_name = getattr(binding, "_current_laminate", None) if binding else None
        self._on_binding_laminate_changed(current_name)
        if hasattr(self, "laminate_name_combo"):
            self._reset_laminate_filter(clear_text=True)
            idx = self.laminate_name_combo.findText(laminate_name)
            if idx >= 0:
                self.laminate_name_combo.setCurrentIndex(idx)
        if binding is not None and hasattr(binding, "_apply_laminate"):
            try:
                binding._apply_laminate(laminate_name)  # type: ignore[attr-defined]
            except Exception as exc:  # pragma: no cover - defensive
                logger.warning("Nao foi possivel aplicar novo laminado: %s", exc)
        self._show_model_warnings()
        self.update_stacking_summary_ui()
        self._update_save_actions_enabled()

    def _open_duplicate_laminate_dialog(self) -> None:
        if self._grid_model is None or not self._grid_model.laminados:
            QMessageBox.information(
                self,
                "Duplicar laminado",
                "Nenhum laminado disponivel para duplicar.",
            )
            return
        dialog = DuplicateLaminateDialog(self)
        laminates = list(self._grid_model.laminados.values())
        names = [laminado.nome for laminado in laminates]
        dialog.set_laminates(names)
        if dialog.exec() != QDialog.Accepted:
            return
        source_name = dialog.selected_name()
        if not source_name:
            return
        source = self._grid_model.laminados.get(source_name)
        if source is None:
            QMessageBox.warning(
                self, "Duplicar laminado", "Laminado nao encontrado."
            )
            return
        clone = self._clone_laminate(source)
        new_name = auto_name_for_laminate(self._grid_model, clone)
        if not new_name:
            QMessageBox.warning(self, "Duplicar laminado", "Falha ao gerar nome automatico.")
            return
        clone.nome = new_name
        self._grid_model.laminados[new_name] = clone
        self._refresh_main_laminate_dropdown(select_name=new_name)
        QMessageBox.information(
            self,
            "Duplicar laminado",
            f"Laminado '{source_name}' duplicado como '{new_name}'.",
        )
        self._mark_dirty()

    def _on_laminate_combo_changed(self, index: int) -> None:
        """Handle selection of a laminate from the combo box to assign it to the current cell."""
        combo = self.laminate_name_combo
        selected_name = combo.itemText(index)
        
        if selected_name == NO_LAMINATE_COMBO_OPTION:
            binding = getattr(self, "_grid_binding", None)
            clear_func = getattr(binding, "set_current_cell_without_laminate", None)
            if callable(clear_func):
                try:
                    clear_func()
                except Exception as exc:  # pragma: no cover - defensive
                    logger.warning(
                        "Nao foi possivel limpar o laminado associado: %s", exc
                    )
            self._reset_laminate_filter(clear_text=True)
            return

        if hasattr(self, "_grid_binding"):
            self._grid_binding._on_laminate_selected(selected_name)

        self._reset_laminate_filter(clear_text=True)

    def _clone_laminate(self, laminado: Laminado) -> Laminado:
        clone = copy.deepcopy(laminado)
        clone.celulas = []
        return clone

    def _refresh_main_laminate_dropdown(
        self, select_name: Optional[str] = None
    ) -> None:
        combo = getattr(self, "laminate_name_combo", None)
        source = getattr(self, "_laminate_source_model", None)
        proxy = getattr(self, "_laminate_filter_model", None)
        if not isinstance(combo, QComboBox):
            return
        if self._grid_model is None:
            if isinstance(source, QStandardItemModel):
                source.clear()
                source.appendRow(QStandardItem(NO_LAMINATE_COMBO_OPTION))
            else:
                combo.clear()
                combo.addItem(NO_LAMINATE_COMBO_OPTION)
            self._reset_laminate_filter(clear_text=True)
            return
        
        # Natural sort key function
        def natural_sort_key(s):
            return [int(text) if text.isdigit() else text.lower()
                    for text in re.split('([0-9]+)', s)]

        names = [laminado.nome for laminado in self._grid_model.laminados.values()]
        sorted_names = sorted(names, key=natural_sort_key)

        if isinstance(source, QStandardItemModel) and isinstance(proxy, LaminateFilterProxy):
            source.blockSignals(True)
            source.clear()
            source.appendRow(QStandardItem(NO_LAMINATE_COMBO_OPTION))
            for name in sorted_names:
                item = QStandardItem(name)
                item.setEditable(False)
                source.appendRow(item)
            source.blockSignals(False)
            proxy.invalidate()
            if select_name:
                self._set_laminate_combo_selection(select_name)
            else:
                self._clear_laminate_combo_display()
        else:
            combo.blockSignals(True)
            combo.clear()
            combo.addItem(NO_LAMINATE_COMBO_OPTION)
            combo.addItems(sorted_names)
            if select_name:
                idx = combo.findText(select_name)
                if idx >= 0:
                    combo.setCurrentIndex(idx)
                else:
                    combo.setCurrentIndex(0)
            else:
                combo.setCurrentIndex(-1)
            combo.blockSignals(False)

        if select_name:
            self._reset_laminate_filter(clear_text=True)
        else:
            self._clear_laminate_combo_display()

    def _open_associated_cells_dialog(self) -> None:
        laminate = self._current_laminate_instance()
        if laminate is None:
            QMessageBox.information(
                self,
                "Celulas associadas",
                "Selecione um laminado para visualizar as celulas associadas.",
            )
            return
        if self._associated_cells_dialog is None:
            self._associated_cells_dialog = AssociatedCellsDialog(self)
        dialog = self._associated_cells_dialog
        cells = (
            self._current_associated_cells
            or self._associated_cells_for_laminate(laminate)
        )
        dialog.refresh_from_laminate(laminate.nome, cells)
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()

    def _associated_cells_for_laminate(
        self, laminate: Optional[Laminado]
    ) -> list[str]:
        if laminate is None or self._grid_model is None:
            return []
        mapped = [
            cell_id
            for cell_id in self._grid_model.celulas_ordenadas
            if self._grid_model.cell_to_laminate.get(cell_id) == laminate.nome
        ]
        if mapped:
            return mapped
        return list(laminate.celulas)

    def update_associated_cells_display(self, cells: Iterable[str]) -> None:
        self._current_associated_cells = [str(cell).strip() for cell in cells if str(cell).strip()]
        button = getattr(self, "btn_associated_cells", None)
        if isinstance(button, QPushButton):
            count = len(self._current_associated_cells)
            tooltip = "Abrir lista de celulas associadas ao laminado atual"
            if count:
                tooltip = f"{tooltip} ({count})"
            button.setToolTip(tooltip)

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

    def _stacking_orientation_token(self, value: object) -> Optional[float]:
        if value is None:
            return None
        try:
            return normalize_angle(value)
        except Exception:
            try:
                return normalize_angle(str(value))
            except Exception:
                return None

    def _stacking_orientations_match(
        self, left: Optional[float], right: Optional[float]
    ) -> bool:
        if left is None and right is None:
            return True
        if left is None or right is None:
            return False
        return math.isclose(left, right, abs_tol=1e-6)

    def _stacking_structural_rows(self, layers: list[Camada]) -> list[int]:
        return [
            idx
            for idx, camada in enumerate(layers)
            if normalize_ply_type_label(getattr(camada, "ply_type", DEFAULT_PLY_TYPE)) != PLY_TYPE_OPTIONS[1]
        ]

    def _stacking_is_unbalanced(
        self, layers: list[Camada], centers: list[int], structural_rows: Optional[list[int]] = None
    ) -> bool:
        if not centers:
            return False
        structural_rows = structural_rows or self._stacking_structural_rows(layers)
        center_min = min(centers)
        center_set = set(centers)
        pos45 = 0
        neg45 = 0
        for row in structural_rows:
            if row in center_set:
                continue
            if row > center_min:
                continue
            orientation = self._stacking_orientation_token(
                getattr(layers[row], "orientacao", None)
            )
            if orientation is None:
                continue
            if math.isclose(orientation, 45.0, abs_tol=1e-6):
                pos45 += 1
            elif math.isclose(orientation, -45.0, abs_tol=1e-6):
                neg45 += 1
        return pos45 != neg45

    def check_symmetry(self, *, show_messages: bool = False) -> None:
        """
        Atualiza o estado visual de simetria do laminado atual.

        A logica replica o comportamento do Virtual Stacking: destaca pares
        quebrados em vermelho, centros simetricos em verde e exibe o alerta de
        balanceamento quando aplicavel. Nenhum dialogo modal e exibido durante
        as atualizacoes automaticas.
        """
        _, model = self._get_stacking_view_and_model()
        if model is None:
            return
        if hasattr(model, "clear_all_highlights"):
            model.clear_all_highlights()

        try:
            layers = model.layers()  # type: ignore[assignment]
        except Exception:
            layers = []

        evaluation = evaluate_symmetry_for_layers(layers)
        centers = evaluation.centers
        structural_rows = evaluation.structural_rows
        symmetric = evaluation.is_symmetric
        status_text = ""

        if not structural_rows:
            symmetric = False
            status_text = "Laminado sem sequencias validas para verificar simetria."
        elif symmetric:
            if hasattr(model, "add_green_rows"):
                model.add_green_rows(centers)
            status_text = (
                f"Laminado simetrico com {len(structural_rows)} sequencias consideradas."
            )
        else:
            mismatch = evaluation.first_mismatch
            if mismatch is not None and hasattr(model, "add_red_rows"):
                model.add_red_rows(list(mismatch))
                status_text = (
                    f"Quebra de simetria nas camadas {mismatch[0] + 1} e {mismatch[1] + 1}."
                )
            elif structural_rows:
                status_text = "Quebra de simetria detectada."

        if hasattr(model, "set_unbalanced_warning"):
            try:
                model.set_unbalanced_warning(
                    self._stacking_is_unbalanced(layers, centers, structural_rows) if symmetric else False
                )
            except Exception:
                model.set_unbalanced_warning(False)

        if show_messages and status_text:
            QMessageBox.information(self, "Verificar simetria", status_text)
        elif status_text and self.statusBar():
            self.statusBar().showMessage(status_text, 4000)

    def on_new_laminate_from_paste(self) -> None:
        binding, model, laminate = self._stacking_binding_context()
        if binding is None or model is None or laminate is None:
            return
        if model.rowCount() > 0:
            QMessageBox.warning(
                self,
                "Laminado existente",
                "Para criar um novo laminado por colagem, primeiro delete todas as camadas do laminado atual.",
            )
            return

        dialog = NewLaminatePasteDialog(self)
        if dialog.exec() != QDialog.Accepted:
            return

        orientations = getattr(dialog, "result_orientations", []) or []
        if not orientations:
            return

        new_layers: list[Camada] = []
        for angle in orientations:
            if angle is None:
                new_layers.append(
                    Camada(
                        idx=0,
                        material="",
                        orientacao=None,
                        ativo=True,
                        simetria=False,
                        ply_type=DEFAULT_PLY_TYPE,
                    )
                )
                continue
            try:
                normalized = normalize_angle(angle)
            except ValueError:
                continue
            new_layers.append(
                Camada(
                    idx=0,
                    material="",
                    orientacao=normalized,
                    ativo=True,
                    simetria=False,
                    ply_type=DEFAULT_PLY_TYPE,
                )
            )
        if not new_layers:
            QMessageBox.information(
                self,
                "Colar stacking",
                "Nenhuma camada valida foi colada.",
            )
            return

        command = AppendLayersCommand(model, laminate, new_layers)
        self.undo_stack.push(command)
        model.clear_checks()
        if hasattr(binding, "_update_layers_count"):
            binding._update_layers_count()
        if self.statusBar():
            self.statusBar().showMessage(
                "Camadas adicionadas a partir da colagem.", 3000
            )

    def on_renumber_sequences(self) -> None:
        _, model, laminate = self._stacking_binding_context()
        if model is None or laminate is None:
            return
        layers_snapshot = model.layers()
        changes: list[LayerFieldChange] = []
        pattern = re.compile(r"^(?P<prefix>[A-Za-z][A-Za-z0-9_-]*)\.?(?P<number>\d+)$")
        seed_prefix, seed_separator = "Seq", "."
        for layer in layers_snapshot:
            text = str(getattr(layer, "sequence", "") or "").strip()
            match = pattern.fullmatch(text)
            if match:
                seed_prefix = match.group("prefix") or seed_prefix
                seed_separator = "." if "." in text else seed_separator
                break

        def _label_parts(label: str) -> tuple[str, str]:
            clean = str(label or "").strip()
            match = pattern.fullmatch(clean)
            if match:
                prefix = match.group("prefix") or seed_prefix
                separator = "." if "." in clean else seed_separator
                return prefix, separator
            return seed_prefix, seed_separator

        for row, layer in enumerate(layers_snapshot):
            prefix, separator = _label_parts(layer.sequence)
            expected = f"{prefix}{separator}{row + 1}"
            stored_value = layer.sequence or ""
            display_value = stored_value or expected
            if display_value == expected:
                continue
            changes.append(
                LayerFieldChange(
                    row=row,
                    column=COL_SEQUENCE,
                    old_value=stored_value,
                    new_value=expected,
                )
            )
        if not changes:
            QMessageBox.information(
                self,
                "Renumerar sequ├¬ncia",
                "A sequ├¬ncia j├í est├í atualizada.",
            )
            return
        command = BulkLayerEditCommand(
            model, laminate, changes, "Renumerar sequ├¬ncia"
        )
        self.undo_stack.push(command)
        self._update_save_actions_enabled()
        self.update_stacking_summary_ui()

    def on_bulk_change_material(self) -> None:
        rows = self._selected_row_indexes()
        if not rows:
            QMessageBox.information(
                self,
                "Sele├º├úo necess├íria",
                "Selecione pelo menos uma camada para trocar o material.",
            )
            return
        _, model = self._get_stacking_view_and_model()
        _, _, laminate = self._stacking_binding_context()
        if model is None or laminate is None:
            return
        materials = self.available_materials()
        dialog = BulkMaterialDialog(
            parent=self,
            available_materials=materials,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        material = dialog.cmb_material.currentText().strip()
        layers_snapshot = model.layers()
        changes: list[LayerFieldChange] = []
        for row in rows:
            if not (0 <= row < len(layers_snapshot)):
                continue
            old_value = layers_snapshot[row].material
            if old_value == material:
                continue
            changes.append(
                LayerFieldChange(
                    row=row,
                    column=COL_MATERIAL,
                    old_value=old_value,
                    new_value=material,
                )
            )
        if not changes:
            QMessageBox.information(
                self,
                "Trocar material",
                "Os materiais selecionados ja possuem o valor informado.",
            )
            return
        command = BulkLayerEditCommand(
            model, laminate, changes, "Trocar material (lote)"
        )
        self.undo_stack.push(command)
        self._update_save_actions_enabled()
        self.update_stacking_summary_ui()

    def on_bulk_change_orientation(self) -> None:
        rows = self._selected_row_indexes()
        if not rows:
            QMessageBox.information(
                self,
                "Sele├º├úo necess├íria",
                "Selecione pelo menos uma camada para trocar a orienta├º├úo.",
            )
            return
        _, model = self._get_stacking_view_and_model()
        _, _, laminate = self._stacking_binding_context()
        if model is None or laminate is None:
            return
        project = self._grid_model
        project_orientations = project_distinct_orientations(project)
        dialog = BulkOrientationDialog(
            parent=self,
            available_orientations=project_orientations,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        new_orientation = dialog.new_orientation
        if new_orientation is None:
            return
        layers_snapshot = model.layers()
        changes: list[LayerFieldChange] = []
        for row in rows:
            if not (0 <= row < len(layers_snapshot)):
                continue
            old_value = layers_snapshot[row].orientacao
            if old_value == new_orientation:
                continue
            changes.append(
                LayerFieldChange(
                    row=row,
                    column=COL_ORIENTATION,
                    old_value=old_value,
                    new_value=new_orientation,
                )
            )
        if not changes:
            QMessageBox.information(
                self,
                "Trocar orienta´┐¢´┐¢Ã£o",
                "As camadas selecionadas ja possuem a orienta´┐¢´┐¢Ã£o escolhida.",
            )
            return
        command = BulkLayerEditCommand(
            model, laminate, changes, "Trocar orienta´┐¢´┐¢Ã£o (lote)"
        )
        self.undo_stack.push(command)
        self._update_save_actions_enabled()

    def _on_add_layer_clicked(self) -> None:
        binding, model, laminate = self._stacking_binding_context()
        if binding is None or model is None or laminate is None:
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
        command = AddLayerCommand(model, laminate, target_row)
        self.undo_stack.push(command)
        if self.statusBar():
            self.statusBar().showMessage("Camada adicionada.", 3000)
        self._update_save_actions_enabled()

    def _on_delete_layers_clicked(self) -> None:
        binding, model, laminate = self._stacking_binding_context()
        if binding is None or model is None or laminate is None:
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
        command = DeleteLayersCommand(model, laminate, selected)
        if not command.has_changes:
            return
        self.undo_stack.push(command)
        if self.statusBar():
            self.statusBar().showMessage(
                f"{len(command.removed_rows)} camada(s) excluidas.", 3000
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

    def _selected_row_indexes(self) -> list[int]:
        view, model = self._get_stacking_view_and_model()
        rows: list[int] = []
        if model is not None and hasattr(model, "checked_rows"):
            try:
                rows = [row for row in model.checked_rows() if isinstance(row, int)]
            except Exception:
                rows = []
        rows = [row for row in rows if row >= 0]
        if rows:
            return sorted(dict.fromkeys(rows))
        if view is None:
            return []
        selection_model = view.selectionModel() if view is not None else None
        if selection_model is None:
            return []
        selected = selection_model.selectedRows()
        fallback_rows = sorted({index.row() for index in selected if index.isValid()})
        return fallback_rows

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
                "Nenhum item est├í selecionado.",
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
        self._move_selected_layer(-1, "acima", "Camada movida para cima.")

    def _on_move_down_clicked(self) -> None:
        self._move_selected_layer(1, "abaixo", "Camada movida para baixo.")

    def _move_selected_layer(
        self, direction: int, error_label: str, status_text: str
    ) -> None:
        binding, model, laminate = self._stacking_binding_context()
        if binding is None or model is None or laminate is None:
            return
        rows = binding.checked_rows()
        if not rows:
            self._handle_move_error("none", error_label)
            return
        if len(rows) > 1:
            self._handle_move_error("multi", error_label)
            return
        source = rows[0]
        target = source + direction
        if not (0 <= target < model.rowCount()):
            self._handle_move_error("edge", error_label)
            return
        command = MoveLayerCommand(model, laminate, source, target)
        self.undo_stack.push(command)
        if self.statusBar():
            self.statusBar().showMessage(status_text, 3000)
        self._update_save_actions_enabled()

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

    # Batch laminate import ----------------------------------------------------

    def _batch_template_path(self) -> Path:
        base_candidates = [
            package_path("..", "Template for Batch Upload.xlsx"),
            package_path("Template for Batch Upload.xlsx"),
            Path(__file__).resolve().parents[2] / "Template for Batch Upload.xlsx",
        ]
        for candidate in base_candidates:
            try:
                resolved = candidate.resolve()
            except Exception:
                resolved = candidate
            if resolved.exists():
                return resolved
        raise FileNotFoundError("Template for Batch Upload.xlsx nao encontrado.")

    def _open_with_default_app(self, path: Path) -> None:
        try:
            os.startfile(path)  # type: ignore[attr-defined]
            return
        except Exception:
            pass
        try:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
        except Exception:
            logger.warning("Nao foi possivel abrir o arquivo %s automaticamente.", path)

    def _prompt_batch_import_choice(self) -> Optional[str]:
        box = QMessageBox(self)
        box.setWindowTitle("Importar laminados em lote")
        box.setText(
            "Selecione como deseja proceder com a importacao em lote."
        )
        open_template = box.addButton("Abrir template", QMessageBox.ActionRole)
        choose_file = box.addButton(
            "Selecionar template preenchido", QMessageBox.AcceptRole
        )
        box.addButton(QMessageBox.Cancel)
        box.exec()
        clicked = box.clickedButton()
        if clicked is open_template:
            return "open"
        if clicked is choose_file:
            return "choose"
        return None

    def _save_blank_batch_template(self, base_template: Path) -> None:
        options = self._file_dialog_options()
        dest, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar template de lote",
            "Template for Batch Upload.xlsx",
            "Planilhas Excel (*.xlsx);;Todos os arquivos (*)",
            options=options,
        )
        if not dest:
            return
        try:
            create_blank_batch_template(base_template, destination=Path(dest), sheet_name="Sheet1")
        except Exception as exc:
            logger.error("Falha ao salvar template de lote: %s", exc, exc_info=True)
            QMessageBox.critical(
                self,
                "Erro",
                "Nao foi possivel salvar o template em branco.",
            )
            return
        QMessageBox.information(
            self,
            "Template salvo",
            f"Modelo salvo em '{Path(dest).name}'.",
        )

    def _select_filled_batch_file(self) -> Optional[Path]:
        options = self._file_dialog_options()
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar template preenchido",
            "",
            "Planilhas Excel (*.xlsx *.xls);;Todos os arquivos (*)",
            options=options,
        )
        return Path(path) if path else None

    def _build_layers_from_batch_entry(self, entry: BatchLaminateInput) -> list[Camada]:
        if not entry.orientations:
            return []
        base = list(entry.orientations)
        if entry.is_symmetric:
            mirror_source = base[:-1] if entry.center_is_single else base
            mirrored = list(reversed(mirror_source))
            full_stack = base + mirrored
        else:
            full_stack = base
        layers: list[Camada] = []
        for idx, angle in enumerate(full_stack):
            layers.append(
                Camada(
                    idx=idx,
                    material="",
                    orientacao=angle if angle is not None else None,
                    ativo=True,
                    simetria=False,
                    ply_type=DEFAULT_PLY_TYPE,
                    ply_label=f"Ply.{idx + 1}",
                    sequence=f"Seq.{idx + 1}",
                )
            )
        return layers

    def _apply_batch_entries(self, entries: list[BatchLaminateInput]) -> list[str]:
        if self._grid_model is None:
            self._grid_model = GridModel()
        created: list[str] = []
        for entry in entries:
            layers = self._build_layers_from_batch_entry(entry)
            if not layers:
                continue
            laminate = Laminado(
                nome="",
                tipo="SS",
                color_index=DEFAULT_COLOR_INDEX,
                tag=str(entry.tag or ""),
                celulas=[],
                camadas=layers,
            )
            laminate.auto_rename_enabled = True
            laminate.nome = _build_auto_name_from_layers(
                layers,
                model=self._grid_model,
                tag=laminate.tag,
                target=laminate,
            )
            self._grid_model.laminados[laminate.nome] = laminate
            self._ensure_unique_laminate_color(laminate)
            self._apply_auto_rename_if_needed(laminate, force=True)
            created.append(laminate.nome)

        if created:
            self._refresh_after_batch_import(created)
        return created

    def _refresh_after_batch_import(self, laminate_names: list[str]) -> None:
        if self._grid_model is None:
            return
        if self.ui_state == UiState.CREATING:
            self._exit_creating_mode()

        self._clear_undo_history()
        bind_model_to_ui(self._grid_model, self)
        binding = getattr(self, "_grid_binding", None)
        if binding is not None:
            self._configure_stacking_table(binding)
        self._sync_all_auto_renamed_laminates()
        bind_cells_to_ui(self._grid_model, self)

        target_name = laminate_names[0] if laminate_names else None
        if target_name:
            combo = getattr(self, "laminate_name_combo", None)
            if isinstance(combo, QComboBox):
                self._reset_laminate_filter(clear_text=True)
                idx = combo.findText(target_name)
                if idx >= 0:
                    combo.setCurrentIndex(idx)
            if binding is not None and hasattr(binding, "_apply_laminate"):
                try:
                    binding._apply_laminate(target_name)  # type: ignore[attr-defined]
                except Exception:
                    logger.debug("Nao foi possivel aplicar laminado importado.", exc_info=True)

        self._refresh_virtual_stacking_view()
        self.project_manager.capture_from_model(
            self._grid_model, self._collect_ui_state()
        )
        self.project_manager.mark_dirty(True)
        self._update_save_actions_enabled()
        self._update_window_title()

    def _import_batch_from_path(self, target_path: Path) -> None:
        try:
            entries = parse_batch_template(target_path)
        except Exception as exc:
            logger.error("Falha ao ler template de lote: %s", exc, exc_info=True)
            QMessageBox.critical(
                self,
                "Erro ao ler planilha",
                str(exc),
            )
            return

        if not entries:
            QMessageBox.information(
                self,
                "Nenhum laminado encontrado",
                "O arquivo informado nao contem laminados preenchidos.",
            )
            return

        total = len(entries)
        confirm = QMessageBox.question(
            self,
            "Confirmar importacao",
            f"{total} laminados encontrados. Deseja importar?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if confirm != QMessageBox.Yes:
            return

        created = self._apply_batch_entries(entries)
        QMessageBox.information(
            self,
            "Importacao concluida",
            f"{len(created)} laminados importados.",
        )

    def _on_batch_import_laminates(self, checked: bool = False) -> None:  # noqa: ARG002
        try:
            template_path = self._batch_template_path()
        except FileNotFoundError as exc:
            QMessageBox.critical(self, "Template ausente", str(exc))
            return

        choice = self._prompt_batch_import_choice()
        if choice is None:
            return

        if choice == "open":
            try:
                temp_copy = create_blank_batch_template(
                    template_path, sheet_name="Sheet1"
                )
            except Exception as exc:
                logger.error(
                    "Falha ao preparar template em branco: %s", exc, exc_info=True
                )
                QMessageBox.critical(
                    self,
                    "Erro",
                    "Nao foi possivel preparar o template de lote.",
                )
                return
            self._open_with_default_app(temp_copy)
            return

        if choice == "choose":
            selected = self._select_filled_batch_file()
            if selected is None:
                return
            self._import_batch_from_path(selected)


    def _load_spreadsheet(self, checked: bool = False) -> None:  # noqa: ARG002
        """Open an Excel file and populate the UI."""
        if not self._confirm_discard_changes():
            return

        previous_laminates = {}
        if self._grid_model is not None and getattr(self._grid_model, "laminados", None):
            previous_laminates = dict(self._grid_model.laminados)

        options = self._file_dialog_options()
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Carregar planilha do Grid Design",
            "",
            "Planilhas Excel (*.xlsx *.xls);;Todos os arquivos (*)",
            options=options,
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

        if previous_laminates:
            for name, laminate in previous_laminates.items():
                if name not in model.laminados:
                    model.laminados[name] = laminate

        self._grid_model = model
        self.project_manager.current_path = None

        if self.ui_state == UiState.CREATING:
            self._exit_creating_mode()

        self._clear_undo_history()
        bind_model_to_ui(self._grid_model, self)
        binding = getattr(self, "_grid_binding", None)
        if binding is not None:
            self._configure_stacking_table(binding)
        self._sync_all_auto_renamed_laminates()
        bind_cells_to_ui(self._grid_model, self)
        current_name = getattr(binding, "_current_laminate", None) if binding else None
        self._on_binding_laminate_changed(current_name)
        self._refresh_virtual_stacking_view()
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
        options = self._file_dialog_options()
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Abrir projeto GridLam",
            str(self.project_manager.current_path or ""),
            "Projetos GridLam (*.gridlam);;Todos os arquivos (*)",
            options=options,
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
        self._sync_all_auto_renamed_laminates()
        bind_cells_to_ui(self._grid_model, self)
        current_name = getattr(binding, "_current_laminate", None) if binding else None
        self._on_binding_laminate_changed(current_name)
        self._refresh_virtual_stacking_view()
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
                self._reset_laminate_filter(clear_text=True)
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
        else:
            self._clear_laminate_combo_display()

    def _snapshot_from_model(self) -> None:
        if self._grid_model is None:
            return
        if self._virtual_stacking_window is not None:
            try:
                self._virtual_stacking_window.persist_column_order()
            except Exception:
                pass
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
        options = self._file_dialog_options()
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar projeto",
            initial_path,
            "Projetos GridLam (*.gridlam);;Todos os arquivos (*)",
            options=options,
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

        try:
            ensure_layers_have_material(self._grid_model)
        except ValueError as exc:
            QMessageBox.warning(
                self,
                "Exportar planilha",
                str(exc),
            )
            return False

        if self._export_checks_thread is not None:
            QMessageBox.information(
                self,
                "Exportar planilha",
                "Uma an├ílise j├í est├í em andamento. Aguarde sua conclus├úo.",
            )
            return False

        self._start_export_verification()
        return True

    def _on_register_material(self, checked: bool = False) -> None:  # noqa: ARG002
        text, ok = QInputDialog.getText(
            self,
            "Cadastrar material",
            "Informe o material (texto completo) para disponibilizar nas listas:",
        )
        if not ok:
            return
        material = str(text or "").strip()
        if not material:
            QMessageBox.warning(
                self, "Material vazio", "Informe um material v├ílido para cadastro."
            )
            return

        existing_keys = {item.casefold() for item in self.available_materials()}
        updated = add_custom_material(material, settings=self._settings)
        added = material.casefold() not in existing_keys and material.casefold() in {
            item.casefold() for item in updated
        }

        if added:
            QMessageBox.information(
                self,
                "Material cadastrado",
                "Material adicionado e dispon├¡vel na lista de materiais.",
            )
        else:
            QMessageBox.information(
                self,
                "Material existente",
                "O material j├í estava cadastrado e segue dispon├¡vel.",
            )

        status_bar = self.statusBar()
        if status_bar:
            status_bar.showMessage("Materiais atualizados", 3000)

    def _start_export_verification(self) -> None:
        model = self._grid_model
        if model is None:
            return
        laminates = list(model.laminados.values())
        if not laminates:
            QMessageBox.information(
                self,
                "Exportar planilha",
                "Nenhum laminado dispon├¡vel para an├ílise.",
            )
            return

        snapshots = [copy.deepcopy(laminado) for laminado in laminates]
        worker = _LaminateChecksWorker(snapshots)
        thread = QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(self._handle_checks_success)
        worker.failed.connect(self._handle_checks_failure)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(self._on_checks_thread_finished)

        self._export_checks_thread = thread
        self._export_checks_worker = worker
        thread.start()

    def _handle_checks_success(self, report: ChecksReport) -> None:
        self._last_checks_report = report
        self._show_verification_report(report)

    def _handle_checks_failure(self, message: str) -> None:
        self._last_checks_report = None
        QMessageBox.critical(
            self,
            "Falha ao executar as verifica├º├Áes.",
            message or "Falha ao executar as verifica├º├Áes.",
        )

    def _on_checks_thread_finished(self) -> None:
        self._export_checks_thread = None
        self._export_checks_worker = None

    def _show_verification_report(self, report: ChecksReport) -> None:
        dialog = VerificationReportDialog(self)
        dialog.set_report(report)
        dialog.removeDuplicatesRequested.connect(
            lambda: self._handle_remove_duplicates_request(dialog)
        )
        result = dialog.exec()
        if result == QDialog.Accepted:
            self._continue_export_after_report()

    def _handle_remove_duplicates_request(
        self, dialog: VerificationReportDialog
    ) -> None:
        """Executa o fluxo de remo├º├úo de laminados duplicados sem associa├º├úo."""
        if self._grid_model is None:
            QMessageBox.information(
                dialog,
                "Remover Duplicados",
                "Nenhum projeto carregado para remover duplicados.",
            )
            return
        eligible = self._duplicates_without_cell_association()
        if not eligible:
            QMessageBox.information(
                dialog,
                "Remover Duplicados",
                "Nenhum laminado duplicado sem associa\u00e7\u00f5es foi encontrado.",
            )
            return
        cleanup_dialog = DuplicateRemovalDialog(eligible, dialog)
        if cleanup_dialog.exec() != QDialog.Accepted:
            return
        removed_names = self._remove_laminates_by_name(
            lam.nome for lam in eligible if lam.nome
        )
        if not removed_names:
            QMessageBox.information(
                dialog,
                "Remover Duplicados",
                "Nenhum laminado foi removido.",
            )
            return
        self._finalize_duplicate_cleanup(removed_names, dialog)

    def _duplicates_without_cell_association(self) -> list[Laminado]:
        """Retorna laminados duplicados que n\u00e3o est\u00e3o vinculados a nenhuma c\u00e9lula."""
        if self._grid_model is None or self._last_checks_report is None:
            return []
        associated_names: set[str] = {
            str(name).strip()
            for name in self._grid_model.cell_to_laminate.values()
            if str(name).strip()
        }
        for name, laminado in self._grid_model.laminados.items():
            if any((cell or "").strip() for cell in getattr(laminado, "celulas", [])):
                associated_names.add(name)
        ordered: OrderedDict[str, Laminado] = OrderedDict()
        for group in self._last_checks_report.duplicates:
            for lam_name in group.laminates:
                if not lam_name or lam_name in associated_names:
                    continue
                laminado = self._grid_model.laminados.get(lam_name)
                if laminado is None:
                    continue
                ordered.setdefault(lam_name, laminado)
        return list(ordered.values())

    def _remove_laminates_by_name(self, names: Iterable[str]) -> list[str]:
        """Remove laminados pelo nome e limpa mapeamentos residuais."""
        if self._grid_model is None:
            return []
        removed: list[str] = []
        for name in names:
            clean_name = (name or "").strip()
            if not clean_name:
                continue
            if clean_name in self._grid_model.laminados:
                del self._grid_model.laminados[clean_name]
                removed.append(clean_name)
        if not removed:
            return []
        cells_to_clear = [
            cell_id
            for cell_id, lam_name in self._grid_model.cell_to_laminate.items()
            if lam_name in removed
        ]
        for cell_id in cells_to_clear:
            self._grid_model.cell_to_laminate.pop(cell_id, None)
        self._mark_dirty()
        return removed

    def _finalize_duplicate_cleanup(
        self,
        removed_names: list[str],
        dialog: VerificationReportDialog,
    ) -> None:
        """Reaplica o binding, atualiza o relat\u00f3rio e informa o usu\u00e1rio."""
        if self._grid_model is None:
            return
        remaining_names = list(self._grid_model.laminados.keys())
        preferred_name = ""
        current_combo_value = ""
        if hasattr(self, "laminate_name_combo"):
            current_combo_value = self.laminate_name_combo.currentText().strip()
        if current_combo_value and current_combo_value in self._grid_model.laminados:
            preferred_name = current_combo_value
        elif remaining_names:
            preferred_name = remaining_names[0]
        self._refresh_after_new_laminate(preferred_name)
        snapshots = [copy.deepcopy(lam) for lam in self._grid_model.laminados.values()]
        new_report = run_all_checks(snapshots)
        self._last_checks_report = new_report
        dialog.set_report(new_report)
        QMessageBox.information(
            dialog,
            "Remover Duplicados",
            f"{len(removed_names)} laminado(s) duplicados removidos.",
        )

    def _continue_export_after_report(self) -> None:
        target_path = self._select_export_destination()
        if target_path is None:
            return
        self._export_model_to_path(target_path)

    def _select_export_destination(self) -> Optional[Path]:
        model = self._grid_model
        if model is None:
            return None

        source_path = model.source_excel_path
        if source_path:
            base_path = Path(source_path)
            suggested = base_path.with_name(f"{base_path.stem}_editado.xlsx")
        else:
            suggested = Path.cwd() / "grid_export.xlsx"

        options = self._file_dialog_options()
        path_str, _ = QFileDialog.getSaveFileName(
            self,
            "Exportar planilha do Grid Design",
            str(suggested),
            "Planilhas Excel (*.xlsx *.xls);;Todos os arquivos (*)",
            options=options,
        )
        if not path_str:
            return None
        return Path(path_str)

    def _export_model_to_path(self, target_path: Path) -> bool:
        model = self._grid_model
        if model is None:
            return False

        try:
            final_path = export_grid_xlsx(model, target_path)
        except ValueError as exc:
            QMessageBox.critical(self, "Falha na exporta├º├úo da planilha.", str(exc))
            return False
        except Exception as exc:  # pragma: no cover - defensivo
            logger.error("Falha ao exportar planilha: %s", exc, exc_info=True)
            QMessageBox.critical(
                self,
                "Falha na exporta├º├úo da planilha.",
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
        self._refresh_virtual_stacking_view()

    def _refresh_virtual_stacking_view(self) -> None:
        window = getattr(self, "_virtual_stacking_window", None)
        if window is None or not window.isVisible():
            return
        try:
            window.populate_from_project(self._grid_model)
        except Exception:  # pragma: no cover - defensive
            logger.debug("Falha ao atualizar Virtual Stacking.", exc_info=True)

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
            self._save_stacking_summary_dialog_state()
            event.accept()
        else:
            event.ignore()


@dataclass
class LayerFieldChange:
    row: int
    column: int
    old_value: object
    new_value: object


class _BaseStackingCommand(QUndoCommand):
    def __init__(
        self, model: StackingTableModel, laminate: Optional[Laminado], text: str
    ) -> None:
        super().__init__(text)
        self._model = model
        self._laminate = laminate

    def _sync_laminate(self) -> None:
        if self._laminate is not None:
            self._laminate.camadas = self._model.layers()


class AddLayerCommand(_BaseStackingCommand):
    def __init__(
        self,
        model: StackingTableModel,
        laminate: Laminado,
        insert_after: Optional[int],
    ) -> None:
        super().__init__(model, laminate, "Adicionar camada")
        self._insert_after = insert_after
        self._insert_row: Optional[int] = None
        self._template = Camada(
            idx=0,
            material="",
            orientacao=0,
            ativo=True,
            simetria=False,
            ply_type=DEFAULT_PLY_TYPE,
        )

    def redo(self) -> None:
        position = self._model.rowCount()
        if self._insert_after is not None:
            position = min(self._insert_after + 1, self._model.rowCount())
        self._insert_row = position
        self._model.insert_layer(position, copy.deepcopy(self._template))
        self._model.clear_checks()
        self._sync_laminate()

    def undo(self) -> None:
        if self._insert_row is None:
            return
        self._model.remove_rows([self._insert_row])
        self._model.clear_checks()
        self._sync_laminate()


class DeleteLayersCommand(_BaseStackingCommand):
    def __init__(
        self,
        model: StackingTableModel,
        laminate: Laminado,
        rows: Iterable[int],
    ) -> None:
        super().__init__(model, laminate, "Excluir camadas selecionadas")
        unique_rows = sorted({row for row in rows if row >= 0})
        snapshot = model.layers()
        self.removed_rows: list[tuple[int, Camada]] = [
            (row, copy.deepcopy(snapshot[row]))
            for row in unique_rows
            if row < len(snapshot)
        ]
        self.has_changes = bool(self.removed_rows)

    def redo(self) -> None:
        if not self.has_changes:
            return
        rows = [row for row, _ in self.removed_rows]
        self._model.remove_rows(rows)
        self._model.clear_checks()
        self._sync_laminate()

    def undo(self) -> None:
        if not self.has_changes:
            return
        for row, layer in self.removed_rows:
            self._model.insert_layer(row, copy.deepcopy(layer))
        self._model.clear_checks()
        self._sync_laminate()


class MoveLayerCommand(_BaseStackingCommand):
    def __init__(
        self,
        model: StackingTableModel,
        laminate: Laminado,
        source: int,
        target: int,
    ) -> None:
        super().__init__(model, laminate, "Mover camada")
        self._source = source
        self._target = target

    def redo(self) -> None:
        if self._model.move_row(self._source, self._target):
            self._sync_laminate()

    def undo(self) -> None:
        if self._model.move_row(self._target, self._source):
            self._sync_laminate()


class BulkLayerEditCommand(_BaseStackingCommand):
    def __init__(
        self,
        model: StackingTableModel,
        laminate: Laminado,
        changes: list[LayerFieldChange],
        text: str,
    ) -> None:
        super().__init__(model, laminate, text)
        self._changes = changes

    def redo(self) -> None:
        for change in self._changes:
            self._model.apply_field_value(change.row, change.column, change.new_value)
        self._sync_laminate()

    def undo(self) -> None:
        for change in reversed(self._changes):
            self._model.apply_field_value(change.row, change.column, change.old_value)
        self._sync_laminate()


class AppendLayersCommand(_BaseStackingCommand):
    def __init__(
        self,
        model: StackingTableModel,
        laminate: Laminado,
        layers: list[Camada],
    ) -> None:
        super().__init__(model, laminate, "Colar stacking")
        self._layers = [copy.deepcopy(layer) for layer in layers]
        self._start_row: Optional[int] = None

    def redo(self) -> None:
        if self._start_row is None:
            self._start_row = self._model.rowCount()
        start = self._start_row
        for offset, layer in enumerate(self._layers):
            self._model.insert_layer(start + offset, copy.deepcopy(layer))
        self._model.clear_checks()
        self._sync_laminate()

    def undo(self) -> None:
        if self._start_row is None:
            return
        rows = range(self._start_row, self._start_row + len(self._layers))
        self._model.remove_rows(rows)
        self._model.clear_checks()
        self._sync_laminate()
