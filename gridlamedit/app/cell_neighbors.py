"""Cell Neighbors editor window.

This lightweight UI lets users define neighbor relationships between
grid cells visually using square nodes and '+' buttons around them.

The scene is intentionally simple and self-contained to avoid coupling
with the rest of the app for now. It exposes one public method:

    get_neighbors_mapping() -> dict[str, dict[str, Optional[str]]]

which returns the current neighbor map in memory.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from typing import Dict, Optional, Tuple

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QPainterPath, QPen, QAction, QUndoStack, QUndoCommand, QLinearGradient, QRadialGradient, QBrush
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QGraphicsItem,
    QGraphicsLineItem,
    QGraphicsRectItem,
    QGraphicsScene,
    QGraphicsSimpleTextItem,
    QGraphicsView,
    QHBoxLayout,
    QVBoxLayout,
    QListWidget,
    QListWidgetItem,
    QSizePolicy,
    QMenu,
    QPushButton,
    QToolBar,
    QMessageBox,
    QLabel,
    QComboBox,
    QInputDialog,
)

try:
    # Optional: reference to the GridModel and orientation colour map
    from gridlamedit.io.spreadsheet import (
        GridModel,
        ORIENTATION_HIGHLIGHT_COLORS,
        DEFAULT_ORIENTATION_HIGHLIGHT,
        normalize_angle,
    )
    from gridlamedit.services.laminate_checks import evaluate_symmetry_for_layers
except Exception:  # pragma: no cover - optional import for loose coupling
    GridModel = object  # type: ignore
    ORIENTATION_HIGHLIGHT_COLORS = {45.0: QColor(193, 174, 255), 90.0: QColor(160, 196, 255), -45.0: QColor(176, 230, 176), 0.0: QColor(230, 230, 230)}
    DEFAULT_ORIENTATION_HIGHLIGHT = QColor(255, 236, 200)

    def normalize_angle(value):  # type: ignore
        return float(value)

    def evaluate_symmetry_for_layers(layers):  # type: ignore
        class _Eval:
            structural_rows: list[int] = []
            centers: list[int] = []
            is_symmetric: bool = False
            first_mismatch = None

        return _Eval()


CELL_SIZE = 80.0
PLUS_SIZE = 20.0
GAP = 28.0
MARGIN = 24.0

# Technical grayscale palette
COLOR_BG = QColor(248, 249, 250)  # Clean light gray
COLOR_CELL = QColor(33, 37, 41)  # Dark charcoal
COLOR_CELL_GRADIENT_START = QColor(52, 58, 64)  # Medium gray
COLOR_CELL_GRADIENT_END = QColor(33, 37, 41)  # Dark charcoal
COLOR_CELL_BORDER = QColor(108, 117, 125)  # Steel gray border
COLOR_TEXT = QColor(248, 249, 250)  # Light gray text
COLOR_PLUS = QColor(108, 117, 125)  # Steel gray
COLOR_PLUS_HOVER = QColor(173, 181, 189)  # Light steel on hover
COLOR_DASH = QColor(134, 142, 150)  # Medium steel gray lines
COLOR_CENTER_BORDER = QColor(250, 128, 114)  # Salmon highlight for central sequences

COLOR_AML_SOFT = QColor(78, 153, 223)  # Clear blue for Soft
COLOR_AML_QUASI = QColor(147, 112, 219)  # Purple for Quasi-iso
COLOR_AML_HARD = QColor(240, 148, 69)  # Amber for Hard
COLOR_AML_UNKNOWN = QColor(108, 117, 125)  # Neutral gray fallback

AML_TYPE_COLORS = {
    "Soft": COLOR_AML_SOFT,
    "Quasi-iso": COLOR_AML_QUASI,
    "Hard": COLOR_AML_HARD,
}

BASE_BORDER_WIDTH = 2
CENTER_BORDER_WIDTH = 4

DIR_OFFSETS: dict[str, tuple[int, int]] = {
    "up": (0, -1),
    "down": (0, 1),
    "left": (-1, 0),
    "right": (1, 0),
}


def opposite(dir_name: str) -> str:
    return {
        "up": "down",
        "down": "up",
        "left": "right",
        "right": "left",
    }[dir_name]


# Undo/Redo Commands
class AddNeighborCommand(QUndoCommand):
    """Command to add a neighbor relationship."""
    def __init__(self, window, src_cell: str, direction: str, dst_cell: str, src_pos: tuple, dst_pos: tuple):
        super().__init__(f"Adicionar vizinho {dst_cell} à {src_cell}")
        self.window = window
        self.src_cell = src_cell
        self.direction = direction
        self.dst_cell = dst_cell
        self.src_pos = src_pos
        self.dst_pos = dst_pos
    
    def redo(self):
        """Execute: add the neighbor relationship."""
        self.window._link_cells_internal(self.src_cell, self.direction, self.dst_cell)
        self.window._draw_connection_between(self.src_pos, self.dst_pos)
        # Update visual state
        for rec in self.window._nodes_by_grid.values():
            if rec.cell_id == self.src_cell or rec.cell_id == self.dst_cell:
                self.window._update_node_neighbors(rec)
        self.window._update_all_plus_buttons_visibility()
        # Update colors for current sequence if one is selected
        if self.window._current_sequence_index is not None:
            self.window.update_cell_colors_for_sequence(self.window._current_sequence_index)
    
    def undo(self):
        """Undo: remove the neighbor relationship."""
        opposite_dir = opposite(self.direction)
        if self.src_cell in self.window._neighbors:
            self.window._neighbors[self.src_cell][self.direction] = None
        if self.dst_cell in self.window._neighbors:
            self.window._neighbors[self.dst_cell][opposite_dir] = None
        # Remove line
        key = (self.src_pos, self.dst_pos) if self.src_pos <= self.dst_pos else (self.dst_pos, self.src_pos)
        line = self.window._lines_between_nodes.get(key)
        if line:
            self.window.scene.removeItem(line)
            del self.window._lines_between_nodes[key]
        # Update visual state
        for rec in self.window._nodes_by_grid.values():
            if rec.cell_id == self.src_cell or rec.cell_id == self.dst_cell:
                self.window._update_node_neighbors(rec)
        self.window._update_all_plus_buttons_visibility()
        # Update colors for current sequence if one is selected
        if self.window._current_sequence_index is not None:
            self.window.update_cell_colors_for_sequence(self.window._current_sequence_index)


class DeleteCellCommand(QUndoCommand):
    """Command to delete a cell."""
    def __init__(self, window, record):
        super().__init__(f"Deletar célula {record.cell_id}")
        self.window = window
        self.cell_id = record.cell_id
        self.grid_pos = record.grid_pos
        # Save the rect position, not item.pos() which may be (0,0)
        self.rect_topleft = record.item.rect().topLeft()
        self.neighbors = dict(window._neighbors.get(record.cell_id, {}))
    
    def redo(self):
        """Execute: delete the cell."""
        record = self.window._nodes_by_grid.get(self.grid_pos)
        if not record:
            return
        
        # Update neighbors
        for direction, neighbor_id in self.neighbors.items():
            if neighbor_id and neighbor_id in self.window._neighbors:
                opposite_dir = opposite(direction)
                self.window._neighbors[neighbor_id][opposite_dir] = None
                # Update visual state
                for rec in self.window._nodes_by_grid.values():
                    if rec.cell_id == neighbor_id:
                        self.window._update_node_neighbors(rec)
        
        # Remove from neighbors dict
        if self.cell_id in self.window._neighbors:
            del self.window._neighbors[self.cell_id]
        
        # Remove lines
        keys_to_remove = [key for key in self.window._lines_between_nodes.keys() if self.grid_pos in key]
        for key in keys_to_remove:
            line = self.window._lines_between_nodes[key]
            self.window.scene.removeItem(line)
            del self.window._lines_between_nodes[key]
        
        # Remove node
        self.window.scene.removeItem(record.item)
        del self.window._nodes_by_grid[self.grid_pos]
        self.window._update_all_plus_buttons_visibility()
        # Update colors for current sequence if one is selected
        if self.window._current_sequence_index is not None:
            self.window.update_cell_colors_for_sequence(self.window._current_sequence_index)
    
    def undo(self):
        """Undo: restore the cell."""
        # Recreate node at same position using rect coordinates
        rect = QRectF(self.rect_topleft.x(), self.rect_topleft.y(), CELL_SIZE, CELL_SIZE)
        item = CellNodeItem(rect)
        item.set_text(self.cell_id)
        self.window.scene.addItem(item)
        
        record = _NodeRecord(item=item, grid_pos=self.grid_pos, cell_id=self.cell_id)
        self.window._nodes_by_grid[self.grid_pos] = record
        
        # Setup callbacks
        def on_select_cell():
            self.window._prompt_select_cell(record)
        def on_add_neighbor(direction: str):
            self.window._handle_add_neighbor(record, direction)
        def on_delete_cell():
            self.window._delete_cell(record)
        def on_change_orientation():
            self.window._change_cell_orientation(record)
        
        item.on_select_cell = on_select_cell
        item.on_add_neighbor = on_add_neighbor
        item.on_delete_cell = on_delete_cell
        item.on_change_orientation = on_change_orientation
        
        # Restore neighbors
        self.window._neighbors[self.cell_id] = dict(self.neighbors)
        for direction, neighbor_id in self.neighbors.items():
            if neighbor_id:
                opposite_dir = opposite(direction)
                if neighbor_id in self.window._neighbors:
                    self.window._neighbors[neighbor_id][opposite_dir] = self.cell_id
                # Redraw connection
                neighbor_rec = None
                for rec in self.window._nodes_by_grid.values():
                    if rec.cell_id == neighbor_id:
                        neighbor_rec = rec
                        break
                if neighbor_rec:
                    self.window._draw_connection_between(self.grid_pos, neighbor_rec.grid_pos)
                    self.window._update_node_neighbors(neighbor_rec)
        
        self.window._update_node_neighbors(record)
        self.window._update_all_plus_buttons_visibility()
        # Update colors for current sequence if one is selected
        if self.window._current_sequence_index is not None:
            self.window.update_cell_colors_for_sequence(self.window._current_sequence_index)


class PlusButtonItem(QGraphicsRectItem):
    """Modern circular "+" button with hover effect."""

    def __init__(self, rect: QRectF, parent: Optional[QGraphicsItem] = None) -> None:
        super().__init__(rect, parent)
        self.setFlags(QGraphicsItem.ItemIsSelectable | QGraphicsItem.ItemIgnoresParentOpacity)
        self.setAcceptHoverEvents(True)
        
        # Create circular appearance with rounded rect
        self._original_brush = QBrush(COLOR_PLUS)
        self._hover_brush = QBrush(COLOR_PLUS_HOVER)
        self.setBrush(self._original_brush)
        
        # Modern border
        pen = QPen(COLOR_CELL_BORDER)
        pen.setWidth(2)
        self.setPen(pen)
        
        self.on_click = None

        self._label = QGraphicsSimpleTextItem("+", self)
        font = self._label.font()
        font.setBold(True)
        font.setPointSize(font.pointSize() + 2)
        self._label.setFont(font)
        self._label.setBrush(QColor(255, 255, 255))
        # Center inside the circle - adjusted for better centering
        text_rect = self._label.boundingRect()
        self._label.setPos(
            rect.x() + (rect.width() - text_rect.width()) / 2 - 0.5,
            rect.y() + (rect.height() - text_rect.height()) / 2 - 1,
        )

    def paint(self, painter, option, widget=None):
        """Custom paint for circular shape."""
        painter.setRenderHint(painter.RenderHint.Antialiasing)
        painter.setBrush(self.brush())
        painter.setPen(self.pen())
        # Draw as circle (ellipse)
        painter.drawEllipse(self.rect())

    def hoverEnterEvent(self, event):  # noqa: N802
        self.setBrush(self._hover_brush)
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event):  # noqa: N802
        self.setBrush(self._original_brush)
        super().hoverLeaveEvent(event)

    # Use mousePress to emulate a simple button
    def mousePressEvent(self, event):  # noqa: N802
        if callable(self.on_click):
            self.on_click()
        event.accept()


class CellNodeItem(QGraphicsRectItem):
    """Modern rounded node with gradient, border glow and hover effect.
    Only shows '+' buttons on sides without neighbors. Lines are drawn behind the block.
    """
    def __init__(self, rect: QRectF, *, parent: Optional[QGraphicsItem] = None) -> None:
        super().__init__(rect, parent)
        
        # Create gradient brush for modern look
        gradient = QLinearGradient(rect.topLeft(), rect.bottomRight())
        gradient.setColorAt(0, COLOR_CELL_GRADIENT_START)
        gradient.setColorAt(1, COLOR_CELL_GRADIENT_END)
        self._normal_brush = QBrush(gradient)
        self.setBrush(self._normal_brush)
        
        # Modern glowing border
        pen = QPen(COLOR_CELL_BORDER)
        pen.setWidth(BASE_BORDER_WIDTH)
        self.setPen(pen)
        # Hover events disabled to preserve orientation colors
        self.setAcceptHoverEvents(False)
        
        self._label = QGraphicsSimpleTextItem("Select\nCell", self)
        f: QFont = self._label.font()
        f.setBold(True)
        f.setPointSize(f.pointSize() + 2)
        self._label.setFont(f)
        self._label.setBrush(COLOR_TEXT)
        # AML type overlay at top-center
        self._aml_label = QGraphicsSimpleTextItem("", self)
        af: QFont = self._aml_label.font()
        af.setBold(True)
        af.setPointSize(max(8, af.pointSize() - 1))
        self._aml_label.setFont(af)
        self._aml_label.setBrush(COLOR_TEXT)
        # Orientation overlay in bottom-right
        self._orientation_label = QGraphicsSimpleTextItem("", self)
        of: QFont = self._orientation_label.font()
        of.setPointSize(max(8, of.pointSize() - 2))
        self._orientation_label.setFont(of)
        self._orientation_label.setBrush(COLOR_TEXT)
        self._recenter_label()
        self.on_select_cell = None
        self.on_add_neighbor = None
        self.on_delete_cell = None
        self.on_change_orientation = None
        self.plus_items: dict[str, PlusButtonItem] = {}
        self.plus_lines: dict[str, QGraphicsLineItem] = {}
        self._neighbors: dict[str, Optional[str]] = {"up": None, "down": None, "left": None, "right": None}
        self._create_plus_buttons()

    def set_text(self, text: str) -> None:
        self._label.setText(text)
        self._recenter_label()

    def paint(self, painter, option, widget=None):
        """Custom paint for rounded corners."""
        painter.setRenderHint(painter.RenderHint.Antialiasing)
        painter.setBrush(self.brush())
        painter.setPen(self.pen())
        # Draw rounded rectangle
        painter.drawRoundedRect(self.rect(), 8, 8)
        # Adjust text contrast according to current fill
        br = self.brush()
        color = br.color() if isinstance(br, QBrush) else COLOR_CELL
        self._update_text_contrast(color)
        self._recenter_label()

    def set_neighbor(self, direction: str, cell_id: Optional[str]) -> None:
        self._neighbors[direction] = cell_id
        self._update_plus_buttons()

    def _update_plus_buttons(self) -> None:
        """Hide plus buttons and their lines when a neighbor is set in that direction.
        Note: This only hides buttons when neighbors exist. Showing buttons is handled
        by _update_all_plus_buttons_visibility which checks cell availability."""
        for direction in DIR_OFFSETS.keys():
            has_neighbor = self._neighbors.get(direction)
            btn = self.plus_items.get(direction)
            line = self.plus_lines.get(direction)
            if has_neighbor:
                # Has neighbor - always hide
                if btn:
                    btn.setVisible(False)
                if line:
                    line.setVisible(False)
            # Don't show here - let _update_all_plus_buttons_visibility handle visibility

    def _recenter_label(self) -> None:
        rect = self.rect()
        tb = self._label.boundingRect()
        self._label.setPos(
            rect.x() + (rect.width() - tb.width()) / 2,
            rect.y() + (rect.height() - tb.height()) / 2,
        )
        ob = self._orientation_label.boundingRect()
        margin = 4.0
        self._orientation_label.setPos(
            rect.right() - ob.width() - margin,
            rect.bottom() - ob.height() - margin,
        )
        ab = self._aml_label.boundingRect()
        self._aml_label.setPos(
            rect.x() + (rect.width() - ab.width()) / 2,
            rect.y() + margin,
        )

    def _update_text_contrast(self, base_color: QColor) -> None:
        r, g, b = base_color.red(), base_color.green(), base_color.blue()
        luminance = 0.299 * r + 0.587 * g + 0.114 * b
        light = QColor(248, 249, 250)
        dark = QColor(33, 37, 41)
        chosen = light if luminance < 150 else dark
        self._label.setBrush(chosen)
        self._orientation_label.setBrush(chosen)
        self._aml_label.setBrush(chosen)

    def set_border_highlight(self, color: Optional[QColor], width: int) -> None:
        """Apply a border style keeping rounded corners."""
        pen_color = COLOR_CELL_BORDER if color is None else color
        pen = QPen(pen_color)
        pen.setWidth(width)
        pen.setJoinStyle(Qt.RoundJoin)
        self.setPen(pen)

    def _create_plus_buttons(self) -> None:
        """Create plus buttons and dashed lines from cell edge (not center) to button center."""
        rect = self.rect()
        cx = rect.center().x()
        cy = rect.center().y()
        half = CELL_SIZE / 2.0
        
        # Button positions (center of the button)
        offsets: dict[str, Tuple[float, float]] = {
            "up": (0.0, -half - GAP),
            "down": (0.0, half + GAP),
            "left": (-half - GAP, 0.0),
            "right": (half + GAP, 0.0),
        }
        
        for direction, (dx, dy) in offsets.items():
            # Button rectangle
            bx = cx + dx - PLUS_SIZE / 2.0
            by = cy + dy - PLUS_SIZE / 2.0
            button = PlusButtonItem(QRectF(bx, by, PLUS_SIZE, PLUS_SIZE), self)
            
            # Line from cell edge to button center
            # Start point: edge of the cell
            if direction == "up":
                line_start_x = cx
                line_start_y = rect.top()
            elif direction == "down":
                line_start_x = cx
                line_start_y = rect.bottom()
            elif direction == "left":
                line_start_x = rect.left()
                line_start_y = cy
            else:  # right
                line_start_x = rect.right()
                line_start_y = cy
            
            # End point: center of button
            line_end_x = cx + dx
            line_end_y = cy + dy
            
            line = QGraphicsLineItem(line_start_x, line_start_y, line_end_x, line_end_y, self)
            pen = QPen(COLOR_DASH)
            pen.setStyle(Qt.DashLine)
            pen.setWidth(2)
            pen.setCapStyle(Qt.PenCapStyle.RoundCap)
            line.setPen(pen)
            line.setZValue(-1)  # Draw lines behind everything
            line.setOpacity(0.6)  # Subtle transparency
            self.plus_items[direction] = button
            self.plus_lines[direction] = line
            
            def handler(dir_name=direction):
                if callable(self.on_add_neighbor):
                    self.on_add_neighbor(dir_name)
            button.on_click = handler

    def mouseDoubleClickEvent(self, event):  # noqa: N802
        if event.button() == Qt.LeftButton:
            if callable(self.on_select_cell):
                self.on_select_cell()
            event.accept()
        else:
            super().mouseDoubleClickEvent(event)

    def mousePressEvent(self, event):  # noqa: N802
        if event.button() == Qt.RightButton:
            self._show_context_menu(event.screenPos())
            event.accept()
        else:
            super().mousePressEvent(event)

    def _show_context_menu(self, pos) -> None:
        """Show context menu with change orientation and delete options."""
        menu = QMenu()
        
        # Add "Trocar orientação" option
        change_orientation_action = QAction("Trocar orienta\u00e7\u00e3o", menu)
        change_orientation_action.triggered.connect(self._handle_change_orientation)
        menu.addAction(change_orientation_action)
        
        menu.addSeparator()
        
        delete_action = QAction("Deletar C\u00e9lula", menu)
        delete_action.triggered.connect(self._handle_delete)
        menu.addAction(delete_action)
        menu.exec(pos)

    def _handle_delete(self) -> None:
        """Handle delete action from context menu."""
        if callable(self.on_delete_cell):
            self.on_delete_cell()

    def _handle_change_orientation(self) -> None:
        """Handle change orientation action from context menu."""
        if callable(self.on_change_orientation):
            self.on_change_orientation()


@dataclass
class _NodeRecord:
    item: CellNodeItem
    grid_pos: tuple[int, int]
    cell_id: Optional[str] = None


class CellNeighborsWindow(QDialog):
    """Dialog to define neighbor relationships among cells."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Cell Neighbors")
        self.resize(1100, 720)

        self._model: Optional[GridModel] = None
        self._project_manager = None  # Will be set via populate_from_project
        self._cells: list[str] = []
        self._undo_stack = QUndoStack(self)
        self._has_unsaved_changes = False

        # Main layout
        main_layout = QVBoxLayout(self)
        
        # Toolbar with buttons
        toolbar = QToolBar(self)
        toolbar.setMovable(False)
        
        # Save button
        self.save_button = QPushButton("Salvar", self)
        self.save_button.clicked.connect(self._save_to_project)
        self.save_button.setEnabled(False)  # Start disabled
        toolbar.addWidget(self.save_button)
        
        toolbar.addSeparator()
        
        # Undo button
        self.undo_button = QPushButton("Desfazer", self)
        self.undo_button.clicked.connect(self._undo_stack.undo)
        self.undo_button.setEnabled(False)
        toolbar.addWidget(self.undo_button)
        
        # Redo button
        self.redo_button = QPushButton("Refazer", self)
        self.redo_button.clicked.connect(self._undo_stack.redo)
        self.redo_button.setEnabled(False)
        toolbar.addWidget(self.redo_button)

        # Sequence selection label & combo box
        self.sequence_label = QLabel("Sequência:", self)
        toolbar.addWidget(self.sequence_label)
        self.sequence_combo = QComboBox(self)
        self.sequence_combo.addItem("Nenhuma")  # index 0 => no colouring
        self.sequence_combo.currentIndexChanged.connect(self._on_sequence_changed)
        toolbar.addWidget(self.sequence_combo)

        # AML type highlight toggle
        self.aml_toggle_button = QPushButton("Tipo AML", self)
        self.aml_toggle_button.setCheckable(True)
        self.aml_toggle_button.setToolTip("Colorir células pelo tipo de AML (Soft, Quasi-iso, Hard)")
        self.aml_toggle_button.toggled.connect(self._on_aml_toggle)
        toolbar.addWidget(self.aml_toggle_button)

        self._current_sequence_index: Optional[int] = None  # 1-based; None => no colouring
        self._aml_highlight_enabled = False
        self._previous_sequence_index_for_aml: int = 0
        
        # Connect undo stack signals
        self._undo_stack.canUndoChanged.connect(self._update_command_buttons)
        self._undo_stack.canRedoChanged.connect(self._update_command_buttons)
        self._undo_stack.indexChanged.connect(self._mark_as_modified)
        
        main_layout.addWidget(toolbar)
        
        # Graphics view
        self.view = QGraphicsView(self)
        self.view.setRenderHint(self.view.renderHints(), True)
        self.view.setBackgroundBrush(COLOR_BG)
        self.view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        main_layout.addWidget(self.view)

        self.scene = QGraphicsScene(self)
        self.scene.setSceneRect(0, 0, 2000, 1200)
        self.view.setScene(self.scene)
        self._setup_view_interaction()

        # Maintain nodes by grid position and by cell id
        self._nodes_by_grid: Dict[tuple[int, int], _NodeRecord] = {}
        self._lines_between_nodes: dict[tuple[tuple[int, int], tuple[int, int]], QGraphicsLineItem] = {}

        # neighbors mapping public structure
        self._neighbors: Dict[str, dict[str, Optional[str]]] = {}
        
        # Track disconnected blocks for removal warnings
        self._disconnected_highlight_rects: list[QGraphicsRectItem] = []

        # Add initial placeholder node with better positioning (away from edges)
        anchor_x = 300.0
        anchor_y = 200.0
        self._create_node((0, 0), QPointF(anchor_x, anchor_y))

    def _setup_view_interaction(self):
        # Enable pan with middle mouse and zoom with wheel
        self._is_panning = False
        self._last_pan_point = None
        view = self.view
        view.setDragMode(QGraphicsView.NoDrag)
        view.viewport().installEventFilter(self)

    def eventFilter(self, obj, event):
        from PySide6.QtCore import QEvent
        from PySide6.QtGui import QWheelEvent, QMouseEvent
        if obj is self.view.viewport():
            if event.type() == QEvent.MouseButtonPress:
                if event.button() == Qt.MiddleButton:
                    self._is_panning = True
                    self._last_pan_point = event.pos()
                    return True
            elif event.type() == QEvent.MouseMove:
                if self._is_panning and self._last_pan_point is not None:
                    delta = event.pos() - self._last_pan_point
                    self._last_pan_point = event.pos()
                    self.view.horizontalScrollBar().setValue(self.view.horizontalScrollBar().value() - delta.x())
                    self.view.verticalScrollBar().setValue(self.view.verticalScrollBar().value() - delta.y())
                    return True
            elif event.type() == QEvent.MouseButtonRelease:
                if event.button() == Qt.MiddleButton:
                    self._is_panning = False
                    self._last_pan_point = None
                    return True
            elif event.type() == QEvent.Wheel:
                wheel: QWheelEvent = event
                if wheel.modifiers() & Qt.ControlModifier:
                    return False
                angle = wheel.angleDelta().y()
                factor = 1.15 if angle > 0 else 1/1.15
                self.view.scale(factor, factor)
                return True
        return super().eventFilter(obj, event)

    def closeEvent(self, event) -> None:
        """Handle window close event - check for disconnected blocks."""
        # Check for disconnected blocks before closing
        if not self._check_and_handle_disconnected_blocks():
            # User cancelled - don't close
            event.ignore()
            return
        
        # Clear highlights and proceed with close
        self._clear_disconnected_highlights()
        event.accept()

    # -------- Public API ---------
    def populate_from_project(self, model: Optional[GridModel], project_manager=None) -> None:
        """Populate the window with cells from the project and load existing neighbors."""
        self._model = model
        self._project_manager = project_manager
        # Always start with AML highlight disabled for a fresh load
        self.aml_toggle_button.setChecked(False)
        cells = []
        if model is not None:
            try:
                cells = list(getattr(model, "celulas_ordenadas", []) or [])
            except Exception:
                cells = []
            if not cells:
                try:
                    cells = list(getattr(model, "cell_to_laminate", {}).keys())
                except Exception:
                    cells = []
            
            # Load existing neighbors from the model
            existing_neighbors = getattr(model, "cell_neighbors", {})
            if existing_neighbors:
                self._neighbors = {cell: dict(mapping) for cell, mapping in existing_neighbors.items()}
                # Rebuild the visual graph from saved neighbors
                self._rebuild_graph_from_neighbors()
        self._cells = cells
        # Update all plus buttons visibility after loading project
        self._update_all_plus_buttons_visibility()
        # Populate sequence combo (max layers among laminados)
        self._populate_sequence_combo()
        self._update_command_buttons()
        # Center view on cells after loading
        self._center_view_on_cells()

    def get_neighbors_mapping(self) -> Dict[str, dict[str, Optional[str]]]:
        """Return neighbors mapping including cells without any neighbors (disconnected cells)."""
        # Return a deep-ish copy to avoid external mutation surprises
        result: Dict[str, dict[str, Optional[str]]] = {}
        for cell, mapping in self._neighbors.items():
            result[cell] = dict(mapping)
        return result

    def _mark_as_modified(self) -> None:
        """Mark that there are unsaved changes."""
        self._has_unsaved_changes = True
        self._update_command_buttons()

    def _update_command_buttons(self, *args) -> None:
        """Enable/disable Save/Undo/Redo respecting AML lock state."""
        if self._aml_highlight_enabled:
            self.save_button.setEnabled(False)
            self.undo_button.setEnabled(False)
            self.redo_button.setEnabled(False)
            return
        self.save_button.setEnabled(self._has_unsaved_changes)
        self.undo_button.setEnabled(self._undo_stack.canUndo())
        self.redo_button.setEnabled(self._undo_stack.canRedo())

    # ---------- Sequence colouring ----------
    def _populate_sequence_combo(self) -> None:
        """Fill the sequence combo with available sequence numbers based on current model.
        Always keeps 'Nenhuma' as first option."""
        if not hasattr(self, "sequence_combo"):
            return
        # Preserve first item, clear the rest
        while self.sequence_combo.count() > 1:
            self.sequence_combo.removeItem(1)
        max_layers = 0
        if self._model is not None:
            try:
                for laminado in self._model.laminados.values():
                    max_layers = max(max_layers, len(getattr(laminado, "camadas", [])))
            except Exception:
                max_layers = 0
        for i in range(1, max_layers + 1):
            self.sequence_combo.addItem(str(i))
        # Reset selection & colours
        self.sequence_combo.setCurrentIndex(0)
        self._current_sequence_index = None
        if self._aml_highlight_enabled:
            self.update_cell_colors_for_aml()
        else:
            self.update_cell_colors_for_sequence(None)
        self._update_command_buttons()

    def _on_sequence_changed(self, combo_index: int) -> None:
        """Handle combo change: index 0 => no sequence colouring, else 1-based sequence number."""
        if self._aml_highlight_enabled:
            # Keep AML mode locked to the base ("Nenhuma") sequence
            if combo_index != 0:
                self.sequence_combo.blockSignals(True)
                self.sequence_combo.setCurrentIndex(0)
                self.sequence_combo.blockSignals(False)
            self._current_sequence_index = None
            self.update_cell_colors_for_aml()
            return
        if combo_index <= 0:
            self._current_sequence_index = None
        else:
            self._current_sequence_index = combo_index  # 1-based sequence
        self.update_cell_colors_for_sequence(self._current_sequence_index)

    def _on_aml_toggle(self, enabled: bool) -> None:
        if enabled:
            self._activate_aml_highlight()
        else:
            self._deactivate_aml_highlight()

    def _activate_aml_highlight(self) -> None:
        if self._aml_highlight_enabled:
            return
        self._aml_highlight_enabled = True
        self._previous_sequence_index_for_aml = self.sequence_combo.currentIndex()
        self.sequence_combo.blockSignals(True)
        self.sequence_combo.setCurrentIndex(0)
        self.sequence_combo.blockSignals(False)
        self._current_sequence_index = None
        self.sequence_combo.setEnabled(False)
        self.sequence_label.setEnabled(False)
        self._update_command_buttons()
        self.update_cell_colors_for_aml()

    def _deactivate_aml_highlight(self) -> None:
        if not self._aml_highlight_enabled:
            return
        self._aml_highlight_enabled = False
        self.sequence_combo.setEnabled(True)
        self.sequence_label.setEnabled(True)
        restore_index = self._previous_sequence_index_for_aml
        if restore_index < 0 or restore_index >= self.sequence_combo.count():
            restore_index = 0
        self.sequence_combo.blockSignals(True)
        self.sequence_combo.setCurrentIndex(restore_index)
        self.sequence_combo.blockSignals(False)
        self._update_command_buttons()
        self._on_sequence_changed(self.sequence_combo.currentIndex())

    def _apply_border_highlight(self, item: CellNodeItem, highlighted: bool) -> None:
        if highlighted:
            item.set_border_highlight(COLOR_CENTER_BORDER, CENTER_BORDER_WIDTH)
        else:
            item.set_border_highlight(None, BASE_BORDER_WIDTH)

    def update_cell_colors_for_sequence(self, sequence_index: Optional[int]) -> None:
        """Apply orientation-based colours to each cell for the given sequence.

        sequence_index: 1-based sequence number or None for reset.
        Rule: same orientation => same colour (reuses global orientation mapping).
        Cells lacking a layer at that sequence or orientation => reset to neutral brush.
        """
        if self._aml_highlight_enabled:
            self.update_cell_colors_for_aml()
            return
        symmetry_cache: dict[int, object] = {}
        for record in self._nodes_by_grid.values():
            item = record.item
            item._aml_label.setText("")
            # Reset when no sequence selected
            if sequence_index is None or not record.cell_id or self._model is None:
                item.setBrush(item._normal_brush)
                item._orientation_label.setText("")
                self._apply_border_highlight(item, False)
                item._recenter_label()
                continue
            # Resolve laminate for cell
            try:
                laminate_name = self._model.cell_to_laminate.get(record.cell_id)
                laminado = self._model.laminados.get(laminate_name) if laminate_name else None
                camadas = getattr(laminado, "camadas", []) if laminado else []
            except Exception:
                camadas = []
                laminado = None
            evaluation = None
            if laminado is not None:
                cache_key = id(laminado)
                evaluation = symmetry_cache.get(cache_key)
                if evaluation is None:
                    evaluation = evaluate_symmetry_for_layers(getattr(laminado, "camadas", []))
                    symmetry_cache[cache_key] = evaluation
            central_rows = set(getattr(evaluation, "centers", []) or [])
            layer_idx = sequence_index - 1
            is_center_sequence = layer_idx in central_rows
            if layer_idx < 0 or layer_idx >= len(camadas):
                item.setBrush(item._normal_brush)
                item._orientation_label.setText("")
                self._apply_border_highlight(item, False)
                item._recenter_label()
                continue
            camada = camadas[layer_idx]
            orient = getattr(camada, "orientacao", None)
            if orient is None:
                item.setBrush(item._normal_brush)
                item._orientation_label.setText("")
                self._apply_border_highlight(item, is_center_sequence)
                item._recenter_label()
                continue
            # Use exact match first, else fallback default highlight
            color = ORIENTATION_HIGHLIGHT_COLORS.get(float(orient), DEFAULT_ORIENTATION_HIGHLIGHT)
            item.setBrush(QBrush(color))
            # Show orientation text in degrees with sign
            try:
                val = float(orient)
                # Normalize decimal: show integers without .0
                label = f"{val:.0f}\u00b0" if abs(val - round(val)) < 1e-6 else f"{val:.1f}\u00b0"
            except Exception:
                label = str(orient)
            item._orientation_label.setText(label)
            item._recenter_label()
            self._apply_border_highlight(item, is_center_sequence)

    # ---------- AML colouring ----------
    def _aml_orientation_bucket(self, value: object) -> Optional[str]:
        if value is None:
            return None
        try:
            angle = float(normalize_angle(value))
        except Exception:
            return None
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

    def _classify_laminate_aml_type(self, laminado: Optional[object]) -> Optional[str]:
        if laminado is None:
            return None
        counts: Counter[str] = Counter()
        for camada in getattr(laminado, "camadas", []) or []:
            bucket = self._aml_orientation_bucket(getattr(camada, "orientacao", None))
            if bucket is None:
                continue
            counts[bucket] += 1
        total = sum(counts.values())
        if total <= 0:
            return None
        pct_zero = counts.get("0", 0) / total
        pct_45 = (counts.get("+45", 0) + counts.get("-45", 0)) / total
        pct_90 = counts.get("90", 0) / total
        threshold = 0.45
        if pct_zero >= threshold and pct_zero >= pct_45 and pct_zero >= pct_90:
            return "Hard"
        if pct_45 >= threshold and pct_45 >= pct_zero and pct_45 >= pct_90:
            return "Soft"
        return "Quasi-iso"

    def update_cell_colors_for_aml(self) -> None:
        """Colour cells by AML type and show the label on top."""
        # Command buttons stay locked while AML highlighting is active
        self._update_command_buttons()
        for record in self._nodes_by_grid.values():
            item = record.item
            item._orientation_label.setText("")
            if not record.cell_id or self._model is None:
                item.setBrush(item._normal_brush)
                item._aml_label.setText("")
                self._apply_border_highlight(item, False)
                item._recenter_label()
                continue

            laminate_name = self._model.cell_to_laminate.get(record.cell_id)
            laminado = self._model.laminados.get(laminate_name) if laminate_name else None
            aml_type = self._classify_laminate_aml_type(laminado)

            if aml_type is None:
                item.setBrush(item._normal_brush)
                item._aml_label.setText("")
                self._apply_border_highlight(item, False)
                item._recenter_label()
                continue

            color = AML_TYPE_COLORS.get(aml_type, COLOR_AML_UNKNOWN)
            item.setBrush(QBrush(color))
            item._aml_label.setText(aml_type)
            self._apply_border_highlight(item, False)
            item._recenter_label()

    def _save_to_project(self) -> None:
        """Save current neighbors mapping to the project (manual save)."""
        if self._aml_highlight_enabled:
            return
        # Check for disconnected blocks before saving
        if not self._check_and_handle_disconnected_blocks():
            return  # User cancelled or no valid blocks
        
        if self._model is not None:
            self._model.cell_neighbors = self.get_neighbors_mapping()
            if self._project_manager is not None:
                try:
                    # Capture and save the updated model
                    self._project_manager.capture_from_model(self._model)
                    if self._project_manager.current_path is not None:
                        self._project_manager.save()
                    # Mark as saved
                    self._has_unsaved_changes = False
                    self._update_command_buttons()
                except Exception as e:
                    import logging
                    logger = logging.getLogger(__name__)
                    logger.warning(f"Failed to save cell neighbors: {e}")

    def _expand_scene_rect(self) -> None:
        """Expand the scene rect to fit all nodes with a safety margin."""
        if not self._nodes_by_grid:
            return
        
        # Calculate bounding box of all cell items
        min_x = float('inf')
        min_y = float('inf')
        max_x = float('-inf')
        max_y = float('-inf')
        
        for record in self._nodes_by_grid.values():
            rect = record.item.rect()
            min_x = min(min_x, rect.left())
            min_y = min(min_y, rect.top())
            max_x = max(max_x, rect.right())
            max_y = max(max_y, rect.bottom())
        
        # Add margin for plus buttons and safety space (100px on each side)
        margin = 100.0
        new_rect = QRectF(
            min_x - margin,
            min_y - margin,
            (max_x - min_x) + 2 * margin,
            (max_y - min_y) + 2 * margin
        )
        
        # Expand scene rect if needed (never shrink)
        current_rect = self.scene.sceneRect()
        final_rect = current_rect.united(new_rect)
        self.scene.setSceneRect(final_rect)

    def _center_view_on_cells(self) -> None:
        """Center the viewport on all existing cells."""
        if not self._nodes_by_grid:
            return
        
        # Calculate bounding box of all cell items
        min_x = float('inf')
        min_y = float('inf')
        max_x = float('-inf')
        max_y = float('-inf')
        
        for record in self._nodes_by_grid.values():
            rect = record.item.sceneBoundingRect()
            min_x = min(min_x, rect.left())
            min_y = min(min_y, rect.top())
            max_x = max(max_x, rect.right())
            max_y = max(max_y, rect.bottom())
        
        # Calculate center point of all cells
        center_x = (min_x + max_x) / 2.0
        center_y = (min_y + max_y) / 2.0
        
        # Center the view on this point
        self.view.centerOn(center_x, center_y)

    def _get_used_cell_ids(self) -> set[str]:
        """Return set of cell IDs already used in the interface."""
        used = set()
        for record in self._nodes_by_grid.values():
            if record.cell_id:
                used.add(record.cell_id)
        return used

    def _has_available_cells(self) -> bool:
        """Check if there are any cells available for selection."""
        used_cells = self._get_used_cell_ids()
        available_cells = [cell for cell in self._cells if cell not in used_cells]
        return len(available_cells) > 0

    def _update_all_plus_buttons_visibility(self) -> None:
        """Update visibility of all '+' buttons based on cell availability."""
        has_available = self._has_available_cells()
        
        for record in self._nodes_by_grid.values():
            if not record.cell_id:
                # Node without cell assigned - show/hide all buttons based on availability
                for direction in DIR_OFFSETS.keys():
                    btn = record.item.plus_items.get(direction)
                    line = record.item.plus_lines.get(direction)
                    if btn:
                        btn.setVisible(has_available)
                    if line:
                        line.setVisible(has_available)
            else:
                # Node with cell assigned - check each direction
                neighbors = self._neighbors.get(record.cell_id, {})
                for direction in DIR_OFFSETS.keys():
                    has_neighbor = neighbors.get(direction) is not None
                    btn = record.item.plus_items.get(direction)
                    line = record.item.plus_lines.get(direction)
                    
                    if has_neighbor:
                        # Has neighbor - hide button
                        if btn:
                            btn.setVisible(False)
                        if line:
                            line.setVisible(False)
                    else:
                        # No neighbor - show button only if cells are available
                        if btn:
                            btn.setVisible(has_available)
                        if line:
                            line.setVisible(has_available)
        self._update_command_buttons()

    def _delete_cell(self, record: _NodeRecord) -> None:
        """Delete a cell from the scene and update all neighbors."""
        if not record.cell_id:
            # Nothing to delete if no cell assigned
            return
        
        command = DeleteCellCommand(self, record)
        self._undo_stack.push(command)

    def _prompt_custom_orientation(self) -> float | None:
        """Prompt user for a custom orientation value."""
        dialog = QInputDialog(self)
        dialog.setInputMode(QInputDialog.InputMode.DoubleInput)
        dialog.setWindowTitle("Outro valor")
        dialog.setLabelText("Informe a orientação (-100 a 100 graus):")
        dialog.setDoubleRange(-100.0, 100.0)
        dialog.setDoubleDecimals(1)
        dialog.setDoubleStep(1.0)
        dialog.setDoubleValue(0.0)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None
        return dialog.doubleValue()

    def _change_cell_orientation(self, record: _NodeRecord) -> None:
        """Change the orientation of a cell's layer for the current sequence."""
        # Check if there is a cell assigned
        if not record.cell_id:
            QMessageBox.information(
                self,
                "Célula não atribuída",
                "Esta célula não possui um ID atribuído. Selecione uma célula primeiro.",
            )
            return
        
        # Check if there is a sequence selected
        if self._current_sequence_index is None:
            QMessageBox.information(
                self,
                "Nenhuma sequência selecionada",
                "Selecione uma sequência no menu suspenso para trocar a orientação.",
            )
            return
        
        # Check if model is available
        if self._model is None:
            QMessageBox.warning(
                self,
                "Modelo não disponível",
                "Não há um modelo de projeto carregado.",
            )
            return
        
        # Get the laminate for this cell
        laminate_name = self._model.cell_to_laminate.get(record.cell_id)
        if not laminate_name:
            QMessageBox.warning(
                self,
                "Laminado não encontrado",
                f"Não foi possível encontrar o laminado associado à célula {record.cell_id}.",
            )
            return
        
        laminado = self._model.laminados.get(laminate_name)
        if not laminado:
            QMessageBox.warning(
                self,
                "Laminado não encontrado",
                f"O laminado '{laminate_name}' não existe no projeto.",
            )
            return
        
        # Get the layer index (0-based)
        layer_idx = self._current_sequence_index - 1
        if layer_idx < 0 or layer_idx >= len(laminado.camadas):
            QMessageBox.warning(
                self,
                "Camada não encontrada",
                f"A célula {record.cell_id} não possui uma camada na sequência {self._current_sequence_index}.",
            )
            return
        
        # Get current orientation
        current_camada = laminado.camadas[layer_idx]
        current_orientation = current_camada.orientacao
        
        # Use the same pattern as Virtual Stacking - QInputDialog.getItem
        options = ["Empty", "0", "45", "-45", "90", "Outro valor..."]
        
        # Find the default selection based on current orientation
        default_index = 0
        if current_orientation is not None:
            current_str = f"{current_orientation:g}"
            try:
                default_index = options.index(current_str)
            except ValueError:
                # Current orientation is not in the list, keep default at 0
                pass
        
        selected, ok = QInputDialog.getItem(
            self,
            "Editar orientação",
            "Selecione a orientação:",
            options,
            default_index,
            False,
        )
        
        if not ok:
            return
        
        selected = selected.strip()
        
        # Handle the selection
        if selected == "Outro valor...":
            # Prompt for custom value
            custom_value = self._prompt_custom_orientation()
            if custom_value is None:
                return
            new_orientation = custom_value
            print(f"DEBUG: Custom orientation selected: {new_orientation}")
        elif selected.lower() == "empty":
            new_orientation = None
            print(f"DEBUG: Empty orientation selected")
        else:
            try:
                new_orientation = float(selected)
                print(f"DEBUG: Standard orientation selected: {new_orientation}")
            except ValueError:
                QMessageBox.warning(
                    self,
                    "Orientação inválida",
                    f"Não foi possível converter '{selected}' para número.",
                )
                return
        
        print(f"DEBUG: Current orientation: {current_orientation}, New orientation: {new_orientation}")
        
        # Check if orientation actually changed
        if current_orientation == new_orientation:
            QMessageBox.information(
                self,
                "Sem alterações",
                "A orientação selecionada é igual à orientação atual.",
            )
            return
        
        # Update the model
        current_camada.orientacao = new_orientation
        print(f"DEBUG: Model updated. Layer {layer_idx} orientation is now: {current_camada.orientacao}")
        
        # Mark as modified
        self._mark_as_modified()
        
        # Update the cell colors
        self.update_cell_colors_for_sequence(self._current_sequence_index)
        
        # Notify other windows (Virtual Stacking and Main Window)
        if self._project_manager is not None:
            try:
                # Capture changes to project manager
                self._project_manager.capture_from_model(self._model)
                # Mark project as dirty
                if hasattr(self._model, 'mark_dirty'):
                    self._model.mark_dirty(True)
            except Exception:
                pass
        
        # Notify parent window (main window) to refresh
        parent_window = self.parent()
        if parent_window is not None:
            try:
                # Try to call the refresh method on the main window
                if hasattr(parent_window, '_on_virtual_stacking_changed'):
                    parent_window._on_virtual_stacking_changed([laminate_name])
            except Exception:
                pass
        
        # Show confirmation message
        if new_orientation is None:
            orientation_text = "vazia"
        else:
            orientation_text = f"{new_orientation}°"
        
        QMessageBox.information(
            self,
            "Orientação atualizada",
            f"A orientação da célula {record.cell_id} na sequência {self._current_sequence_index} foi atualizada para {orientation_text}.",
        )

    def _find_connected_components(self) -> list[set[str]]:
        """Find all connected components (blocks) of cells.
        Returns a list of sets, where each set contains cell IDs in one connected component."""
        # Get all cells that have been assigned
        all_cells = set()
        for record in self._nodes_by_grid.values():
            if record.cell_id:
                all_cells.add(record.cell_id)
        
        if not all_cells:
            return []
        
        visited = set()
        components = []
        
        for cell in all_cells:
            if cell in visited:
                continue
            
            # BFS to find all cells in this component
            component = set()
            queue = [cell]
            
            while queue:
                current = queue.pop(0)
                if current in visited:
                    continue
                visited.add(current)
                component.add(current)
                
                # Check all neighbors
                neighbors = self._neighbors.get(current, {})
                for neighbor_id in neighbors.values():
                    if neighbor_id and neighbor_id not in visited:
                        queue.append(neighbor_id)
            
            if component:
                components.append(component)
        
        return components

    def _check_and_handle_disconnected_blocks(self) -> bool:
        """Check for disconnected blocks, highlight them, and ask user for confirmation.
        Returns True if save should proceed, False if cancelled."""
        # Clear previous highlights
        self._clear_disconnected_highlights()
        
        components = self._find_connected_components()
        
        if len(components) <= 1:
            # No disconnected blocks
            return True
        
        # Find the largest component (to preserve)
        largest_component = max(components, key=len)
        disconnected_components = [c for c in components if c != largest_component]
        
        if not disconnected_components:
            return True
        
        # Highlight disconnected blocks with red rectangles
        disconnected_cells = []
        for component in disconnected_components:
            disconnected_cells.extend(component)
            self._highlight_disconnected_block(component)
        
        # Show warning dialog
        cell_list = ", ".join(sorted(disconnected_cells))
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Blocos Desconectados")
        msg.setText(f"Os seguintes blocos estão desconectados e serão removidos:")
        msg.setInformativeText(f"Células: {cell_list}\\n\\n"
                              f"O bloco principal com {len(largest_component)} célula(s) será preservado.\\n"
                              f"Blocos menores com {sum(len(c) for c in disconnected_components)} célula(s) serão removidos.")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.setDefaultButton(QMessageBox.Cancel)
        
        result = msg.exec()
        
        if result == QMessageBox.Ok:
            # Remove disconnected blocks
            self._remove_disconnected_blocks(disconnected_components)
            self._clear_disconnected_highlights()
            return True
        else:
            # User cancelled - clear highlights so they can continue editing
            self._clear_disconnected_highlights()
            return False

    def _highlight_disconnected_block(self, component: set[str]) -> None:
        """Highlight a disconnected block with a red rectangle."""
        # Find all nodes in this component
        nodes = []
        for record in self._nodes_by_grid.values():
            if record.cell_id in component:
                nodes.append(record)
        
        if not nodes:
            return
        
        # Calculate bounding box
        min_x = float('inf')
        min_y = float('inf')
        max_x = float('-inf')
        max_y = float('-inf')
        
        for record in nodes:
            rect = record.item.rect()
            min_x = min(min_x, rect.left())
            min_y = min(min_y, rect.top())
            max_x = max(max_x, rect.right())
            max_y = max(max_y, rect.bottom())
        
        # Add margin
        margin = 10
        highlight_rect = QGraphicsRectItem(
            min_x - margin,
            min_y - margin,
            max_x - min_x + 2 * margin,
            max_y - min_y + 2 * margin
        )
        
        # Red border, transparent fill
        pen = QPen(QColor(255, 0, 0))
        pen.setWidth(3)
        highlight_rect.setPen(pen)
        highlight_rect.setBrush(Qt.NoBrush)
        highlight_rect.setZValue(10)  # Draw on top
        
        self.scene.addItem(highlight_rect)
        self._disconnected_highlight_rects.append(highlight_rect)

    def _clear_disconnected_highlights(self) -> None:
        """Remove all red highlight rectangles."""
        for rect in self._disconnected_highlight_rects:
            self.scene.removeItem(rect)
        self._disconnected_highlight_rects.clear()

    def _remove_disconnected_blocks(self, components: list[set[str]]) -> None:
        """Remove all cells in the given disconnected components."""
        cells_to_remove = set()
        for component in components:
            cells_to_remove.update(component)
        
        # Find records to delete
        records_to_delete = []
        for record in list(self._nodes_by_grid.values()):
            if record.cell_id in cells_to_remove:
                records_to_delete.append(record)
        
        # Delete each record (this will also clean up neighbors)
        for record in records_to_delete:
            # Remove neighbors references
            if record.cell_id in self._neighbors:
                neighbors = self._neighbors[record.cell_id]
                for direction, neighbor_id in neighbors.items():
                    if neighbor_id and neighbor_id in self._neighbors:
                        opposite_dir = opposite(direction)
                        self._neighbors[neighbor_id][opposite_dir] = None
                del self._neighbors[record.cell_id]
            
            # Remove lines
            keys_to_remove = [key for key in self._lines_between_nodes.keys() if record.grid_pos in key]
            for key in keys_to_remove:
                line = self._lines_between_nodes[key]
                self.scene.removeItem(line)
                del self._lines_between_nodes[key]
            
            # Special case: never remove the origin node at (0,0)
            # Instead, just clear its cell assignment
            if record.grid_pos == (0, 0):
                record.cell_id = None
                record.item.set_text("Select\nCell")
                # Clear all neighbors for this node
                record.item.set_neighbor("up", None)
                record.item.set_neighbor("down", None)
                record.item.set_neighbor("left", None)
                record.item.set_neighbor("right", None)
            else:
                # Remove node normally
                self.scene.removeItem(record.item)
                del self._nodes_by_grid[record.grid_pos]
        
        # Update visual state
        self._update_all_plus_buttons_visibility()
        
        # Update colors for current sequence if one is selected
        if self._current_sequence_index is not None:
            self.update_cell_colors_for_sequence(self._current_sequence_index)

    def _rebuild_graph_from_neighbors(self) -> None:
        """Rebuild the visual node graph from saved neighbor relationships."""
        if not self._neighbors:
            return
        
        # Ensure the origin node (0,0) always exists
        if (0, 0) not in self._nodes_by_grid:
            # Recreate the origin node if it was somehow removed
            anchor_x = MARGIN
            anchor_y = MARGIN
            self._create_node((0, 0), QPointF(anchor_x, anchor_y))
        
        # Clear existing nodes except the initial one
        nodes_to_remove = [pos for pos in self._nodes_by_grid.keys() if pos != (0, 0)]
        for pos in nodes_to_remove:
            rec = self._nodes_by_grid[pos]
            self.scene.removeItem(rec.item)
            del self._nodes_by_grid[pos]
        
        # Clear connection lines
        for line in self._lines_between_nodes.values():
            self.scene.removeItem(line)
        self._lines_between_nodes.clear()
        
        # Build a graph of cells and their positions
        # Start from first cell at (0,0)
        cell_to_grid: Dict[str, tuple[int, int]] = {}
        cells_to_process = list(self._neighbors.keys())
        
        if not cells_to_process:
            return
        
        # Place first cell at origin
        first_cell = cells_to_process[0]
        cell_to_grid[first_cell] = (0, 0)
        initial_node = self._nodes_by_grid.get((0, 0))
        if initial_node:
            initial_node.cell_id = first_cell
            initial_node.item.set_text(first_cell)
        
        # BFS to position all connected cells
        queue = [first_cell]
        visited = {first_cell}
        
        while queue:
            current_cell = queue.pop(0)
            current_pos = cell_to_grid[current_cell]
            neighbors = self._neighbors.get(current_cell, {})
            
            for direction, neighbor_cell in neighbors.items():
                if not neighbor_cell or neighbor_cell in visited:
                    continue
                
                # Calculate neighbor position
                offset = DIR_OFFSETS[direction]
                neighbor_pos = (current_pos[0] + offset[0], current_pos[1] + offset[1])
                cell_to_grid[neighbor_cell] = neighbor_pos
                
                # Create node if doesn't exist
                if neighbor_pos not in self._nodes_by_grid:
                    origin_node = self._nodes_by_grid.get((0, 0))
                    if origin_node:
                        base = origin_node.item.rect().topLeft()
                        top_left = QPointF(
                            base.x() + (CELL_SIZE + GAP + PLUS_SIZE) * neighbor_pos[0],
                            base.y() + (CELL_SIZE + GAP + PLUS_SIZE) * neighbor_pos[1],
                        )
                        self._create_node(neighbor_pos, top_left)
                
                # Assign cell to node
                neighbor_node = self._nodes_by_grid[neighbor_pos]
                neighbor_node.cell_id = neighbor_cell
                neighbor_node.item.set_text(neighbor_cell)
                
                visited.add(neighbor_cell)
                queue.append(neighbor_cell)
        
        # Update visual state and draw connections
        for cell_id, grid_pos in cell_to_grid.items():
            record = self._nodes_by_grid.get(grid_pos)
            if record:
                self._update_node_neighbors(record)
                # Draw connections to neighbors (use internal to avoid undo stack on load)
                neighbors = self._neighbors.get(cell_id, {})
                for direction, neighbor_cell in neighbors.items():
                    if neighbor_cell:
                        neighbor_pos = cell_to_grid.get(neighbor_cell)
                        if neighbor_pos:
                            self._draw_connection_between(grid_pos, neighbor_pos)
        
        # Expand scene to fit all loaded nodes
        self._expand_scene_rect()
        
        # Update colors for current sequence if one is selected
        if self._current_sequence_index is not None:
            self.update_cell_colors_for_sequence(self._current_sequence_index)

    # -------- Internal helpers ---------
    def _create_node(self, grid_pos: tuple[int, int], top_left: QPointF) -> _NodeRecord:
        rect = QRectF(top_left.x(), top_left.y(), CELL_SIZE, CELL_SIZE)
        item = CellNodeItem(rect)
        self.scene.addItem(item)

        record = _NodeRecord(item=item, grid_pos=grid_pos, cell_id=None)
        self._nodes_by_grid[grid_pos] = record

        def on_select_cell():
            self._prompt_select_cell(record)

        def on_add_neighbor(direction: str):
            self._handle_add_neighbor(record, direction)

        def on_delete_cell():
            self._delete_cell(record)

        def on_change_orientation():
            self._change_cell_orientation(record)

        item.on_select_cell = on_select_cell
        item.on_add_neighbor = on_add_neighbor
        item.on_delete_cell = on_delete_cell
        item.on_change_orientation = on_change_orientation
        
        # Expand scene to accommodate new node
        self._expand_scene_rect()
        
        return record

    def _prompt_select_cell(self, record: _NodeRecord) -> None:
        if not self._cells:
            return
        
        # Filter out already used cells
        used_cells = self._get_used_cell_ids()
        available_cells = [cell for cell in self._cells if cell not in used_cells or cell == record.cell_id]
        
        if not available_cells:
            return
        
        dialog = SelectCellDialog(available_cells, current=record.cell_id, parent=self)
        if dialog.exec() != QDialog.Accepted:
            return
        selected = dialog.selected_cell()
        if not selected:
            return
        record.cell_id = selected
        record.item.set_text(record.cell_id)
        # Always ensure cell is in neighbors dict, even without connections
        self._ensure_cell_mapping_entry(record.cell_id)
        # Check for adjacent cells and auto-connect (without undo - this is part of cell selection)
        for dir_name, (dx, dy) in DIR_OFFSETS.items():
            neighbor_pos = (record.grid_pos[0] + dx, record.grid_pos[1] + dy)
            neighbor = self._nodes_by_grid.get(neighbor_pos)
            if neighbor and neighbor.cell_id:
                self._link_cells_internal(record.cell_id, dir_name, neighbor.cell_id)
                self._draw_connection_between(record.grid_pos, neighbor.grid_pos)
        self._update_node_neighbors(record)
        self._expand_scene_rect()
        # Update all plus buttons visibility after selecting a cell
        self._update_all_plus_buttons_visibility()
        # Update colors for current sequence if one is selected
        if self._current_sequence_index is not None:
            self.update_cell_colors_for_sequence(self._current_sequence_index)

    def _handle_add_neighbor(self, record: _NodeRecord, direction: str) -> None:
        # Check if there are available cells before proceeding
        if not self._has_available_cells():
            return
        
        # Cannot add neighbor if current node has no assigned cell yet
        if not record.cell_id:
            self._prompt_select_cell(record)
            if not record.cell_id:
                return
        offset = DIR_OFFSETS[direction]
        target_grid = (record.grid_pos[0] + offset[0], record.grid_pos[1] + offset[1])
        if target_grid in self._nodes_by_grid:
            neighbor = self._nodes_by_grid[target_grid]
            if not neighbor.cell_id:
                self._prompt_select_cell(neighbor)
                if not neighbor.cell_id:
                    return
            self._link_cells(record.cell_id, direction, neighbor.cell_id)
            self._draw_connection_between(record.grid_pos, neighbor.grid_pos)
            self._update_node_neighbors(record)
            self._update_node_neighbors(neighbor)
            self._expand_scene_rect()
            return

        # Create the node at aligned position
        origin_node = self._nodes_by_grid.get((0, 0))
        if not origin_node:
            return  # Cannot create node without origin reference
        base = origin_node.item.rect().topLeft()
        top_left = QPointF(
            base.x() + (CELL_SIZE + GAP + PLUS_SIZE) * target_grid[0],
            base.y() + (CELL_SIZE + GAP + PLUS_SIZE) * target_grid[1],
        )
        neighbor = self._create_node(target_grid, top_left)
        self._prompt_select_cell(neighbor)
        if not neighbor.cell_id:
            # User cancelled - remove the created node
            self.scene.removeItem(neighbor.item)
            del self._nodes_by_grid[target_grid]
            return
        self._link_cells(record.cell_id, direction, neighbor.cell_id)
        self._draw_connection_between(record.grid_pos, neighbor.grid_pos)
        self._update_node_neighbors(record)
        self._update_node_neighbors(neighbor)
        self._expand_scene_rect()
        # Update all plus buttons visibility after creating neighbor
        self._update_all_plus_buttons_visibility()
        # Update colors for current sequence if one is selected
        if self._current_sequence_index is not None:
            self.update_cell_colors_for_sequence(self._current_sequence_index)


    def _ensure_cell_mapping_entry(self, cell_id: str) -> None:
        """Ensure a cell has an entry in neighbors dict, even if it has no neighbors."""
        if cell_id not in self._neighbors:
            self._neighbors[cell_id] = {"up": None, "down": None, "left": None, "right": None}

    def _link_cells_internal(self, src: str, direction: str, dst: str) -> None:
        """Internal method to link cells without undo."""
        self._ensure_cell_mapping_entry(src)
        self._ensure_cell_mapping_entry(dst)
        self._neighbors[src][direction] = dst
        self._neighbors[dst][opposite(direction)] = src
        # Update visual state of nodes
        for rec in self._nodes_by_grid.values():
            if rec.cell_id == src:
                rec.item.set_neighbor(direction, dst)
            if rec.cell_id == dst:
                rec.item.set_neighbor(opposite(direction), src)

    def _link_cells(self, src: str, direction: str, dst: str) -> None:
        """Link cells with undo support."""
        # Find grid positions
        src_pos = None
        dst_pos = None
        for rec in self._nodes_by_grid.values():
            if rec.cell_id == src:
                src_pos = rec.grid_pos
            if rec.cell_id == dst:
                dst_pos = rec.grid_pos
        
        if src_pos and dst_pos:
            command = AddNeighborCommand(self, src, direction, dst, src_pos, dst_pos)
            self._undo_stack.push(command)

    def _draw_connection_between(self, a: tuple[int, int], b: tuple[int, int]) -> None:
        """Draw dashed line between two cells, from edge to edge (only for direct neighbors)."""
        # Calculate grid distance
        dx = b[0] - a[0]
        dy = b[1] - a[1]
        
        # Only draw if cells are direct neighbors (horizontal or vertical, distance of 1)
        if not ((abs(dx) == 1 and dy == 0) or (abs(dy) == 1 and dx == 0)):
            # Not direct neighbors - skip diagonal or distant connections
            return
        
        key = (a, b) if a <= b else (b, a)
        if key in self._lines_between_nodes:
            return
        
        a_item = self._nodes_by_grid.get(a)
        b_item = self._nodes_by_grid.get(b)
        if not a_item or not b_item:
            return
        
        a_rect = a_item.item.rect()
        b_rect = b_item.item.rect()
        a_center = a_rect.center()
        b_center = b_rect.center()
        
        # Calculate start and end points on the edges of the cells
        if dx != 0 and dy == 0:  # Horizontal connection
            if dx > 0:  # a is left of b
                start_x = a_rect.right()
                start_y = a_center.y()
                end_x = b_rect.left()
                end_y = b_center.y()
            else:  # a is right of b
                start_x = a_rect.left()
                start_y = a_center.y()
                end_x = b_rect.right()
                end_y = b_center.y()
        elif dy != 0 and dx == 0:  # Vertical connection
            if dy > 0:  # a is above b
                start_x = a_center.x()
                start_y = a_rect.bottom()
                end_x = b_center.x()
                end_y = b_rect.top()
            else:  # a is below b
                start_x = a_center.x()
                start_y = a_rect.top()
                end_x = b_center.x()
                end_y = b_rect.bottom()
        else:
            # Should never reach here given the check above
            return
        
        line = QGraphicsLineItem(start_x, start_y, end_x, end_y)
        pen = QPen(COLOR_DASH)
        pen.setStyle(Qt.DashLine)
        pen.setWidth(3)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        line.setPen(pen)
        line.setZValue(-2)  # Draw connection lines behind everything
        line.setOpacity(0.7)  # Subtle glow effect
        self.scene.addItem(line)
        self._lines_between_nodes[key] = line

    def _update_node_neighbors(self, record: _NodeRecord) -> None:
        """Update the visual state of plus buttons based on current neighbors."""
        if not record.cell_id:
            # No cell assigned yet, show all plus buttons
            for dir_name in DIR_OFFSETS.keys():
                record.item.set_neighbor(dir_name, None)
            return
        
        neighbors = self._neighbors.get(record.cell_id, {})
        for dir_name in DIR_OFFSETS.keys():
            cell_id = neighbors.get(dir_name)
            record.item.set_neighbor(dir_name, cell_id)


class SelectCellDialog(QDialog):
    """Dialog listing available cells for selection."""

    def __init__(self, cells: list[str], *, current: Optional[str] = None, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Selecionar Celula")
        self.resize(320, 400)
        layout = QVBoxLayout(self)
        self.list = QListWidget(self)
        for cell in cells:
            item = QListWidgetItem(cell)
            self.list.addItem(item)
            if current and cell == current:
                self.list.setCurrentItem(item)
        layout.addWidget(self.list)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.list.itemDoubleClicked.connect(lambda _i: self.accept())

    def selected_cell(self) -> Optional[str]:
        item = self.list.currentItem()
        return item.text() if item else None
