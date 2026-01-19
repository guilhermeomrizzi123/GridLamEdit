"""Cell Neighbors editor window.

This lightweight UI lets users define neighbor relationships between
grid cells visually using square nodes and '+' buttons around them.

The scene is intentionally simple and self-contained to avoid coupling
with the rest of the app for now. It exposes one public method:

    get_neighbors_mapping() -> dict[str, dict[str, list[str]]]

which returns the current neighbor map in memory (listas para suportar múltiplas conexões).
"""

from __future__ import annotations

from collections import Counter
import logging
import re
from dataclasses import dataclass
from typing import Dict, Optional, Tuple

from PySide6.QtCore import QPointF, QRectF, Qt, QSize, QLineF, QSignalBlocker
from PySide6.QtGui import QColor, QFont, QPainterPath, QPen, QAction, QUndoStack, QUndoCommand, QLinearGradient, QRadialGradient, QBrush, QIcon, QPainter
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QGraphicsItem,
    QGraphicsLineItem,
    QGraphicsRectItem,
    QGraphicsScene,
    QGraphicsSimpleTextItem,
    QGraphicsTextItem,
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

from gridlamedit.core.paths import package_path

try:
    # Optional: reference to the GridModel and orientation colour map
    from gridlamedit.io.spreadsheet import (
        GridModel,
        ORIENTATION_HIGHLIGHT_COLORS,
        DEFAULT_ORIENTATION_HIGHLIGHT,
        normalize_angle,
        CELL_ID_PATTERN,
    )
    from gridlamedit.services.laminate_checks import evaluate_symmetry_for_layers
except Exception:  # pragma: no cover - optional import for loose coupling
    GridModel = object  # type: ignore
    ORIENTATION_HIGHLIGHT_COLORS = {45.0: QColor(193, 174, 255), 90.0: QColor(160, 196, 255), -45.0: QColor(176, 230, 176), 0.0: QColor(230, 230, 230)}
    DEFAULT_ORIENTATION_HIGHLIGHT = QColor(255, 236, 200)

    CELL_ID_PATTERN = re.compile(r"^C\d+$", re.IGNORECASE)

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
COLOR_CONTOUR_TEXT = QColor(55, 65, 81)  # Dark gray for contour labels

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

DRAWING_ITEM_ROLE = 0
DRAWING_ITEM_TAG = "drawing_item"
DRAW_LINE_PEN = QPen(QColor(40, 40, 40), 2)
TEXTBOX_HANDLE_SIZE = 6

RESOURCES_ICONS_DIR = package_path("resources", "icons")


def _load_drawing_icon(filename: str) -> QIcon:
    icon_path = RESOURCES_ICONS_DIR / filename
    if icon_path.is_file():
        return QIcon(str(icon_path))
    return QIcon()

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
        self.window._recalculate_cell_neighbors_from_scene()
        # Update colors for current sequence if one is selected
        if self.window._current_sequence_index is not None:
            self.window.update_cell_colors_for_sequence(self.window._current_sequence_index)
    
    def undo(self):
        """Undo: remove the neighbor relationship."""
        self.window._remove_neighbor_relation(self.src_cell, self.direction, self.dst_cell)
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
        self.window._recalculate_cell_neighbors_from_scene()
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
        # Deep copy aggregated neighbors (sets)
        raw_neighbors = window._neighbors.get(record.cell_id, {})
        self.neighbors: dict[str, set[str]] = {
            direction: set(values or []) for direction, values in raw_neighbors.items()
        }
    
    def redo(self):
        """Execute: delete the cell."""
        record = self.window._nodes_by_grid.get(self.grid_pos)
        if not record:
            return
        
        # Update neighbors
        for direction, neighbor_ids in self.neighbors.items():
            for neighbor_id in neighbor_ids:
                if neighbor_id and neighbor_id in self.window._neighbors:
                    self.window._remove_neighbor_relation(self.cell_id, direction, neighbor_id)
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
        self.window._recalculate_cell_neighbors_from_scene()
        self.window._update_all_plus_buttons_visibility()
        # Update colors for current sequence if one is selected
        if self.window._current_sequence_index is not None:
            self.window.update_cell_colors_for_sequence(self.window._current_sequence_index)
    
    def undo(self):
        """Undo: restore the cell."""
        # Recreate node at same position using rect coordinates
        rect = QRectF(self.rect_topleft.x(), self.rect_topleft.y(), CELL_SIZE, CELL_SIZE)
        item = CellNodeItem(rect)
        self.window.scene.addItem(item)
        
        record = _NodeRecord(item=item, grid_pos=self.grid_pos, cell_id=self.cell_id)
        self.window._nodes_by_grid[self.grid_pos] = record
        self.window._update_node_cell_display(record)
        
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
        self.window._neighbors[self.cell_id] = {
            direction: set(values or []) for direction, values in self.neighbors.items()
        }
        for direction, neighbor_ids in self.neighbors.items():
            for neighbor_id in neighbor_ids:
                if neighbor_id:
                    opposite_dir = opposite(direction)
                    self.window._add_neighbor_relation(self.cell_id, direction, neighbor_id)
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
        self.window._recalculate_cell_neighbors_from_scene()
        self.window._update_all_plus_buttons_visibility()
        # Update colors for current sequence if one is selected
        if self.window._current_sequence_index is not None:
            self.window.update_cell_colors_for_sequence(self.window._current_sequence_index)


class AddNodeCommand(QUndoCommand):
    """Command to add a new cell node with its connections."""

    def __init__(
        self,
        window,
        grid_pos: tuple[int, int],
        rect_topleft: QPointF,
        cell_id: str,
        connections: list[tuple[str, str, str, tuple[int, int], tuple[int, int]]],
    ) -> None:
        super().__init__(f"Adicionar célula {cell_id}")
        self.window = window
        self.grid_pos = grid_pos
        self.rect_topleft = rect_topleft
        self.cell_id = cell_id
        self.connections = connections

    def redo(self):
        record = self.window._nodes_by_grid.get(self.grid_pos)
        if record is None:
            record = self.window._create_node(self.grid_pos, self.rect_topleft)
        record.cell_id = self.cell_id
        self.window._update_node_cell_display(record)
        self.window._ensure_cell_mapping_entry(self.cell_id)

        for src_cell, direction, dst_cell, src_pos, dst_pos in self.connections:
            self.window._link_cells_internal(src_cell, direction, dst_cell)
            self.window._draw_connection_between(src_pos, dst_pos)

        affected_cells = {self.cell_id}
        for src_cell, _, dst_cell, _, _ in self.connections:
            affected_cells.add(src_cell)
            affected_cells.add(dst_cell)
        for rec in self.window._nodes_by_grid.values():
            if rec.cell_id in affected_cells:
                self.window._update_node_neighbors(rec)

        self.window._recalculate_cell_neighbors_from_scene()
        self.window._update_all_plus_buttons_visibility()
        if self.window._current_sequence_index is not None:
            self.window.update_cell_colors_for_sequence(self.window._current_sequence_index)

    def undo(self):
        self.window._remove_node_and_connections(
            self.grid_pos,
            self.cell_id,
            self.connections,
        )


class AddDrawingItemCommand(QUndoCommand):
    """Command to add a drawing item to the scene."""

    def __init__(self, window, item: QGraphicsItem, text: str) -> None:
        super().__init__(text)
        self.window = window
        self.item = item

    def redo(self):
        if self.item.scene() is None:
            self.window.scene.addItem(self.item)

    def undo(self):
        if self.item.scene() is not None:
            self.window.scene.removeItem(self.item)


class RemoveDrawingItemCommand(QUndoCommand):
    """Command to remove a drawing item from the scene."""

    def __init__(self, window, item: QGraphicsItem, text: str) -> None:
        super().__init__(text)
        self.window = window
        self.item = item

    def redo(self):
        if self.item.scene() is not None:
            self.window.scene.removeItem(self.item)

    def undo(self):
        if self.item.scene() is None:
            self.window.scene.addItem(self.item)


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


class DrawingLineItem(QGraphicsLineItem):
    """Movable line item for drawing tool."""

    def __init__(self, line, parent: Optional[QGraphicsItem] = None) -> None:
        super().__init__(line, parent)
        self.setFlag(QGraphicsItem.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.ItemIsSelectable, True)
        self.setFlag(QGraphicsItem.ItemSendsGeometryChanges, True)
        self.setData(DRAWING_ITEM_ROLE, DRAWING_ITEM_TAG)


class TextBoxItem(QGraphicsRectItem):
    """Movable text box with small handle markers when selected."""

    def __init__(self, rect: QRectF, parent: Optional[QGraphicsItem] = None) -> None:
        super().__init__(rect, parent)
        self.setFlag(QGraphicsItem.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.ItemIsSelectable, True)
        self.setFlag(QGraphicsItem.ItemSendsGeometryChanges, True)
        self.setData(DRAWING_ITEM_ROLE, DRAWING_ITEM_TAG)

    def paint(self, painter: QPainter, option, widget=None) -> None:
        super().paint(painter, option, widget)
        if not self.isSelected():
            return
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        handle = TEXTBOX_HANDLE_SIZE
        half = handle / 2.0
        rect = self.rect()
        points = [
            rect.topLeft(),
            QPointF(rect.center().x(), rect.top()),
            rect.topRight(),
            QPointF(rect.left(), rect.center().y()),
            QPointF(rect.right(), rect.center().y()),
            rect.bottomLeft(),
            QPointF(rect.center().x(), rect.bottom()),
            rect.bottomRight(),
        ]
        painter.setBrush(QColor(52, 58, 64))
        painter.setPen(QPen(QColor(255, 255, 255), 1))
        for point in points:
            painter.drawRect(
                QRectF(point.x() - half, point.y() - half, handle, handle)
            )


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
        # Laminate name label (smaller, below cell id)
        self._laminate_label = QGraphicsSimpleTextItem("", self)
        lf: QFont = self._laminate_label.font()
        lf.setPointSize(max(8, f.pointSize() - 2))
        self._laminate_label.setFont(lf)
        self._laminate_label.setBrush(COLOR_TEXT)
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
        # Contour labels around the cell
        self._contour_labels: dict[str, QGraphicsSimpleTextItem] = {}
        contour_font = QFont(f)
        contour_font.setBold(False)
        contour_font.setPointSize(max(5, f.pointSize() - 7))
        for key in ("top", "right", "bottom", "left"):
            label = QGraphicsSimpleTextItem("", self)
            label.setFont(contour_font)
            label.setBrush(COLOR_CONTOUR_TEXT)
            label.setVisible(False)
            self._contour_labels[key] = label
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

    def set_laminate_text(self, text: str) -> None:
        self._laminate_label.setText(text or "")
        self._recenter_label()

    def set_contour_texts(self, contours: Tuple[str, str, str, str]) -> None:
        mapping = {
            "top": contours[0],
            "right": contours[1],
            "bottom": contours[2],
            "left": contours[3],
        }
        for key, value in mapping.items():
            label = self._contour_labels.get(key)
            if not label:
                continue
            text = (value or "").strip()
            label.setText(text)
            label.setVisible(bool(text))
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
        lam_text = self._laminate_label.text()
        spacing = 2.0
        if lam_text:
            self._laminate_label.setVisible(True)
            lb = self._laminate_label.boundingRect()
            total_h = tb.height() + spacing + lb.height()
            start_y = rect.y() + (rect.height() - total_h) / 2
            self._label.setPos(
                rect.x() + (rect.width() - tb.width()) / 2,
                start_y,
            )
            self._laminate_label.setPos(
                rect.x() + (rect.width() - lb.width()) / 2,
                start_y + tb.height() + spacing,
            )
        else:
            self._laminate_label.setVisible(False)
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
        self._position_contour_labels(rect)

    def _position_contour_labels(self, rect: QRectF) -> None:
        margin = 4.0
        top_label = self._contour_labels.get("top")
        if top_label and top_label.isVisible():
            tb = top_label.boundingRect()
            top_label.setRotation(0)
            top_label.setTransformOriginPoint(tb.center())
            top_label.setPos(
                rect.center().x() - tb.width() / 2,
                rect.top() - margin - tb.height(),
            )

        bottom_label = self._contour_labels.get("bottom")
        if bottom_label and bottom_label.isVisible():
            bb = bottom_label.boundingRect()
            bottom_label.setRotation(0)
            bottom_label.setTransformOriginPoint(bb.center())
            bottom_label.setPos(
                rect.center().x() - bb.width() / 2,
                rect.bottom() + margin,
            )

        right_label = self._contour_labels.get("right")
        if right_label and right_label.isVisible():
            rb = right_label.boundingRect()
            right_label.setTransformOriginPoint(rb.center())
            right_label.setRotation(90)
            target_x = rect.right() + margin + rb.height() / 2
            target_y = rect.center().y()
            right_label.setPos(target_x - rb.center().x(), target_y - rb.center().y())

        left_label = self._contour_labels.get("left")
        if left_label and left_label.isVisible():
            lb = left_label.boundingRect()
            left_label.setTransformOriginPoint(lb.center())
            left_label.setRotation(-90)
            target_x = rect.left() - margin - lb.height() / 2
            target_y = rect.center().y()
            left_label.setPos(target_x - lb.center().x(), target_y - lb.center().y())

    def _update_text_contrast(self, base_color: QColor) -> None:
        r, g, b = base_color.red(), base_color.green(), base_color.blue()
        luminance = 0.299 * r + 0.587 * g + 0.114 * b
        light = QColor(248, 249, 250)
        dark = QColor(33, 37, 41)
        chosen = light if luminance < 150 else dark
        self._label.setBrush(chosen)
        self._orientation_label.setBrush(chosen)
        self._aml_label.setBrush(chosen)
        self._laminate_label.setBrush(chosen)

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
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowMinMaxButtonsHint
            | Qt.WindowSystemMenuHint
        )
        self.setWindowFlag(Qt.Window, True)
        self.resize(1100, 720)

        self._model: Optional[GridModel] = None
        self._project_manager = None  # Will be set via populate_from_project
        self._cells: list[str] = []
        self._undo_stack = QUndoStack(self)
        self._undo_stack.setUndoLimit(3)
        self._has_unsaved_changes = False
        self._virtual_stacking_proxy = None

        # Main layout
        main_layout = QVBoxLayout(self)
        
        # Toolbar with buttons
        toolbar = QToolBar(self)
        toolbar.setMovable(False)
        toolbar.setToolButtonStyle(Qt.ToolButtonIconOnly)
        toolbar.setIconSize(QSize(20, 20))
        
        # Undo action (icon only)
        self.undo_action = QAction(QIcon(":/icons/undo.svg"), "Desfazer", self)
        self.undo_action.setToolTip("Desfazer (Ctrl+Z)")
        self.undo_action.triggered.connect(self._undo_stack.undo)
        self.undo_action.setEnabled(False)
        toolbar.addAction(self.undo_action)

        # Redo action (icon only)
        self.redo_action = QAction(QIcon(":/icons/redo.svg"), "Refazer", self)
        self.redo_action.setToolTip("Refazer (Ctrl+Y)")
        self.redo_action.triggered.connect(self._undo_stack.redo)
        self.redo_action.setEnabled(False)
        toolbar.addAction(self.redo_action)

        toolbar.addSeparator()

        # Sequence selection label & combo box
        self.sequence_label = QLabel("Sequência:", self)
        toolbar.addWidget(self.sequence_label)
        self.sequence_combo = QComboBox(self)
        self.sequence_combo.addItem("Nenhuma")  # index 0 => no colouring
        self.sequence_combo.currentIndexChanged.connect(self._on_sequence_changed)
        toolbar.addWidget(self.sequence_combo)

        # Reorder by neighborhood (same behavior as Virtual Stacking)
        self.reorder_neighbors_button = QPushButton("Reorder by Neighborhood", self)
        self.reorder_neighbors_button.setCheckable(True)
        self.reorder_neighbors_button.setToolTip(
            "Reorders sequences based on neighborhood and symmetry rules"
        )
        self.reorder_neighbors_button.toggled.connect(self._on_reorder_neighbors_toggle)
        toolbar.addWidget(self.reorder_neighbors_button)

        # AML type highlight toggle
        self.aml_toggle_button = QPushButton("Tipo AML", self)
        self.aml_toggle_button.setCheckable(True)
        self.aml_toggle_button.setToolTip("Colorir células pelo tipo de AML (Soft, Quasi-iso, Hard)")
        self.aml_toggle_button.toggled.connect(self._on_aml_toggle)
        toolbar.addWidget(self.aml_toggle_button)

        # Drawing tools palette
        self._drawing_tool: Optional[str] = None
        self._drawing_line_item: Optional[QGraphicsLineItem] = None
        self._drawing_line_start: Optional[QPointF] = None
        self._drawing_actions: dict[str, QAction] = {}

        drawing_toolbar = QToolBar(self)
        drawing_toolbar.setMovable(False)
        drawing_toolbar.setToolButtonStyle(Qt.ToolButtonIconOnly)
        drawing_toolbar.setIconSize(QSize(22, 22))
        drawing_toolbar.setWindowTitle("Ferramentas de Desenho")
        drawing_toolbar.addWidget(QLabel("Desenho:", self))

        line_action = QAction(_load_drawing_icon("line_tool.svg"), "Linha", self)
        line_action.setCheckable(True)
        line_action.setToolTip("Ferramenta de linha")
        line_action.toggled.connect(lambda checked, tool="line": self._on_drawing_tool_toggled(tool, checked))
        drawing_toolbar.addAction(line_action)
        self._drawing_actions["line"] = line_action

        text_action = QAction(_load_drawing_icon("text_tool.svg"), "Texto", self)
        text_action.setCheckable(True)
        text_action.setToolTip("Adicionar caixa de texto")
        text_action.toggled.connect(lambda checked, tool="text": self._on_drawing_tool_toggled(tool, checked))
        drawing_toolbar.addAction(text_action)
        self._drawing_actions["text"] = text_action

        erase_action = QAction(_load_drawing_icon("erase_tool.svg"), "Apagar", self)
        erase_action.setCheckable(True)
        erase_action.setToolTip("Apagar linhas ou caixas de texto")
        erase_action.toggled.connect(lambda checked, tool="erase": self._on_drawing_tool_toggled(tool, checked))
        drawing_toolbar.addAction(erase_action)
        self._drawing_actions["erase"] = erase_action

        self._current_sequence_index: Optional[int] = None  # 1-based; None => no colouring
        self._aml_highlight_enabled = False
        self._previous_sequence_index_for_aml: int = 0
        
        # Connect undo stack signals
        self._undo_stack.canUndoChanged.connect(self._update_command_buttons)
        self._undo_stack.canRedoChanged.connect(self._update_command_buttons)
        self._undo_stack.indexChanged.connect(self._mark_as_modified)
        
        main_layout.addWidget(toolbar)
        main_layout.addWidget(drawing_toolbar)
        
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

        # neighbors mapping aggregated por célula (cada direção guarda um set de IDs vizinhos)
        self._neighbors: Dict[str, dict[str, set[str]]] = {}
        
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

    def _on_drawing_tool_toggled(self, tool: str, checked: bool) -> None:
        if checked:
            for name, action in self._drawing_actions.items():
                if name != tool and action.isChecked():
                    action.blockSignals(True)
                    action.setChecked(False)
                    action.blockSignals(False)
            self._drawing_tool = tool
        else:
            if all(not action.isChecked() for action in self._drawing_actions.values()):
                self._drawing_tool = None
        if self._drawing_tool != "line":
            self._cancel_active_line()
        self._apply_drawing_cursor()

    def _apply_drawing_cursor(self) -> None:
        if self._drawing_tool == "line":
            cursor = Qt.CrossCursor
        elif self._drawing_tool == "text":
            cursor = Qt.IBeamCursor
        elif self._drawing_tool == "erase":
            cursor = Qt.PointingHandCursor
        else:
            cursor = Qt.ArrowCursor
        self.view.viewport().setCursor(cursor)

    def _cancel_active_line(self) -> None:
        if self._drawing_line_item is not None:
            self.scene.removeItem(self._drawing_line_item)
            self._drawing_line_item = None
            self._drawing_line_start = None

    def _start_drawing_line(self, scene_pos: QPointF) -> None:
        self._cancel_active_line()
        self._drawing_line_start = scene_pos
        line_item = DrawingLineItem(
            QLineF(scene_pos, scene_pos)
        )
        line_item.setPen(DRAW_LINE_PEN)
        line_item.setZValue(0.5)
        self.scene.addItem(line_item)
        self._drawing_line_item = line_item

    def _update_drawing_line(self, scene_pos: QPointF) -> None:
        if self._drawing_line_item is None or self._drawing_line_start is None:
            return
        self._drawing_line_item.setLine(
            self._drawing_line_start.x(),
            self._drawing_line_start.y(),
            scene_pos.x(),
            scene_pos.y(),
        )

    def _finish_drawing_line(self, scene_pos: QPointF) -> None:
        if self._drawing_line_item is None:
            return
        self._update_drawing_line(scene_pos)
        line_item = self._drawing_line_item
        self._drawing_line_item = None
        self._drawing_line_start = None
        self._undo_stack.push(AddDrawingItemCommand(self, line_item, "Adicionar linha"))
        self._mark_as_modified()

    def _add_text_box(self, scene_pos: QPointF) -> None:
        text, ok = QInputDialog.getText(self, "Adicionar texto", "Texto:")
        if not ok:
            return
        text = (text or "").strip()
        if not text:
            return
        text_item = QGraphicsTextItem(text)
        text_item.setDefaultTextColor(QColor(33, 37, 41))
        text_item.setTextInteractionFlags(Qt.NoTextInteraction)
        font = text_item.font()
        font.setPointSize(max(9, font.pointSize()))
        text_item.setFont(font)

        padding = 6
        text_rect = text_item.boundingRect()
        rect_item = TextBoxItem(
            QRectF(
                0,
                0,
                text_rect.width() + padding * 2,
                text_rect.height() + padding * 2,
            )
        )
        rect_item.setBrush(QColor(255, 255, 255, 230))
        rect_item.setPen(QPen(QColor(120, 120, 120), 1))
        rect_item.setZValue(1)
        text_item.setParentItem(rect_item)
        text_item.setPos(padding, padding)
        self.scene.addItem(rect_item)

        rect_item.setPos(
            scene_pos
            - QPointF(rect_item.rect().width() / 2.0, rect_item.rect().height() / 2.0)
        )
        self._undo_stack.push(AddDrawingItemCommand(self, rect_item, "Adicionar texto"))

    def _find_drawing_item(self, item: Optional[QGraphicsItem]) -> Optional[QGraphicsItem]:
        if item is None:
            return None
        if item.data(DRAWING_ITEM_ROLE) == DRAWING_ITEM_TAG:
            return item
        parent = item.parentItem()
        if parent and parent.data(DRAWING_ITEM_ROLE) == DRAWING_ITEM_TAG:
            return parent
        return None

    def _erase_at(self, scene_pos: QPointF) -> None:
        item = self.scene.itemAt(scene_pos, self.view.transform())
        target = self._find_drawing_item(item)
        if target is not None:
            self._undo_stack.push(RemoveDrawingItemCommand(self, target, "Remover desenho"))

    def eventFilter(self, obj, event):
        from PySide6.QtCore import QEvent
        from PySide6.QtGui import QWheelEvent, QMouseEvent
        if obj is self.view.viewport():
            if self._drawing_tool is not None:
                if event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                    scene_pos = self.view.mapToScene(event.pos())
                    if self._drawing_tool == "line":
                        self._start_drawing_line(scene_pos)
                        return True
                    if self._drawing_tool == "text":
                        self._add_text_box(scene_pos)
                        return True
                    if self._drawing_tool == "erase":
                        self._erase_at(scene_pos)
                        return True
                elif (
                    event.type() == QEvent.MouseMove
                    and self._drawing_tool == "line"
                    and self._drawing_line_item is not None
                ):
                    self._update_drawing_line(self.view.mapToScene(event.pos()))
                    return True
                elif (
                    event.type() == QEvent.MouseButtonRelease
                    and event.button() == Qt.LeftButton
                    and self._drawing_tool == "line"
                    and self._drawing_line_item is not None
                ):
                    self._finish_drawing_line(self.view.mapToScene(event.pos()))
                    return True
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
        self._neighbors = {}
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
            # Load existing neighbors from the model (prefer detailed node graph)
            node_payload = list(getattr(model, "cell_neighbor_nodes", []) or [])
            if node_payload:
                self._rebuild_graph_from_nodes_payload(node_payload)
                self._recalculate_cell_neighbors_from_scene()
            else:
                existing_neighbors = getattr(model, "cell_neighbors", {}) or {}
                if existing_neighbors:
                    self._neighbors = self._convert_legacy_neighbors(existing_neighbors)
                    # Rebuild the visual graph from saved neighbors
                    self._rebuild_graph_from_neighbors()
        self._cells = cells
        # Refresh cell labels/laminate tags using latest model data
        for rec in self._nodes_by_grid.values():
            self._update_node_cell_display(rec)
        # Update all plus buttons visibility after loading project
        self._update_all_plus_buttons_visibility()
        # Populate sequence combo (max layers among laminados)
        self._populate_sequence_combo()
        self._update_command_buttons()
        # Center view on cells after loading
        self._center_view_on_cells()

    def get_neighbors_mapping(self) -> Dict[str, dict[str, list[str]]]:
        """Return neighbors mapping including cells without any neighbors (disconnected cells).

        Values are stored as lists to support múltiplas conexões por direção e manter compatibilidade
        com serialização JSON.
        """
        result: Dict[str, dict[str, list[str]]] = {}
        for cell, mapping in self._neighbors.items():
            bucket: dict[str, list[str]] = {}
            for direction in DIR_OFFSETS.keys():
                values = mapping.get(direction, set()) if isinstance(mapping, dict) else set()
                bucket[direction] = sorted(list(values)) if values else []
            result[cell] = bucket
        return result

    def _convert_legacy_neighbors(self, mapping: dict) -> Dict[str, dict[str, set[str]]]:
        """Convert legacy mapping (str/None or list) into set-based structure."""
        normalized: Dict[str, dict[str, set[str]]] = {}
        for cell, directions in mapping.items():
            bucket = self._empty_neighbor_bucket()
            if not isinstance(directions, dict):
                continue
            for direction in DIR_OFFSETS.keys():
                raw_value = directions.get(direction)
                if raw_value is None:
                    bucket[direction] = set()
                elif isinstance(raw_value, (list, tuple, set)):
                    bucket[direction] = {str(v) for v in raw_value if v}
                else:
                    bucket[direction] = {str(raw_value)}
            normalized[cell] = bucket
        return normalized

    def _mark_as_modified(self) -> None:
        """Mark that there are unsaved changes."""
        self._has_unsaved_changes = True
        self._update_command_buttons()
        self._auto_save_if_needed()

    def _update_command_buttons(self, *args) -> None:
        """Enable/disable Undo/Redo respecting AML lock state."""
        if self._aml_highlight_enabled:
            self.undo_action.setEnabled(False)
            self.redo_action.setEnabled(False)
            return
        self.undo_action.setEnabled(self._undo_stack.canUndo())
        self.redo_action.setEnabled(self._undo_stack.canRedo())

    def _auto_save_if_needed(self) -> None:
        """Auto-save changes immediately when possible."""
        if self._aml_highlight_enabled:
            return
        if not self._has_unsaved_changes:
            return
        self._save_to_project()

    # ---------- Sequence colouring ----------
    def _find_center_sequence_indices(self, max_layers: int) -> set[int]:
        """Collect 1-based center sequence indices for laminates used by current cells."""
        centers: set[int] = set()
        if self._model is None or max_layers <= 0:
            return centers

        # Limit to laminates actually referenced by the current grid cells to mirror the UI highlight.
        try:
            cell_to_laminate = getattr(self._model, "cell_to_laminate", {}) or {}
        except Exception:
            cell_to_laminate = {}
        used_laminate_names = {
            cell_to_laminate.get(cell_id)
            for cell_id in self._cells
            if cell_id in cell_to_laminate
        }
        try:
            laminados = getattr(self._model, "laminados", {}) or {}
        except Exception:
            laminados = {}

        for name in used_laminate_names:
            if not name:
                continue
            laminado = laminados.get(name)
            if laminado is None:
                continue
            try:
                evaluation = evaluate_symmetry_for_layers(getattr(laminado, "camadas", []) or [])
                for idx in getattr(evaluation, "centers", []) or []:
                    seq_number = idx + 1  # 0-based -> 1-based
                    if 1 <= seq_number <= max_layers:
                        centers.add(seq_number)
            except Exception:
                continue
        return centers

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
        center_sequences = self._find_center_sequence_indices(max_layers)
        for i in range(1, max_layers + 1):
            label = str(i)
            self.sequence_combo.addItem(label)
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

    def _ensure_virtual_stacking_window(self):
        parent = self.parent()
        window = None
        if parent is not None and hasattr(parent, "_virtual_stacking_window"):
            window = getattr(parent, "_virtual_stacking_window", None)
            if window is None:
                try:
                    undo_stack = getattr(parent, "undo_stack", None)
                except Exception:
                    undo_stack = None
                from gridlamedit.app.virtualstacking import VirtualStackingWindow

                window = VirtualStackingWindow(parent, undo_stack=undo_stack)
                setattr(parent, "_virtual_stacking_window", window)

        if window is None:
            if self._virtual_stacking_proxy is None:
                from gridlamedit.app.virtualstacking import VirtualStackingWindow

                self._virtual_stacking_proxy = VirtualStackingWindow(self)
            window = self._virtual_stacking_proxy

        if self._model is not None:
            try:
                if getattr(window, "_project", None) is not self._model:
                    window.populate_from_project(self._model)
            except Exception:
                window.populate_from_project(self._model)

        return window

    def _on_reorder_neighbors_toggle(self, checked: bool) -> None:
        window = self._ensure_virtual_stacking_window()
        if window is None:
            return

        if hasattr(window, "btn_reorganize_neighbors"):
            blocker = QSignalBlocker(window.btn_reorganize_neighbors)
            window.btn_reorganize_neighbors.setChecked(checked)
            del blocker

        window._toggle_reorganizar_por_vizinhanca(checked)

        actual_checked = checked
        if hasattr(window, "btn_reorganize_neighbors"):
            actual_checked = window.btn_reorganize_neighbors.isChecked()
        else:
            actual_checked = bool(getattr(window, "_neighbors_reorder_snapshot", None))

        if actual_checked != checked:
            blocker = QSignalBlocker(self.reorder_neighbors_button)
            self.reorder_neighbors_button.setChecked(actual_checked)
            del blocker

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
        """Save current neighbors mapping to the project."""
        if self._aml_highlight_enabled:
            return
        # Check for disconnected blocks before saving
        if not self._check_and_handle_disconnected_blocks():
            return  # User cancelled or no valid blocks
        # Validate that duplicated IDs stay connected
        if not self._validate_duplicate_cells_connected():
            return
        # Keep aggregated neighbors in sync with what is drawn
        self._recalculate_cell_neighbors_from_scene()
        
        if self._model is not None:
            self._model.cell_neighbor_nodes = self._build_neighbor_nodes_payload()
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

    def _has_available_cells(self) -> bool:
        """Check if there are any cells available for selection."""
        return bool(self._cells)

    def _has_grid_connection(self, record: _NodeRecord, direction: str) -> bool:
        """Return True if there is a drawn connection from this node in the given direction."""
        dx, dy = DIR_OFFSETS[direction]
        neighbor_pos = (record.grid_pos[0] + dx, record.grid_pos[1] + dy)
        key = (record.grid_pos, neighbor_pos) if record.grid_pos <= neighbor_pos else (neighbor_pos, record.grid_pos)
        return key in self._lines_between_nodes

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
                for direction in DIR_OFFSETS.keys():
                    has_neighbor = self._has_grid_connection(record, direction)
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

    def _remove_line_between(self, a: tuple[int, int], b: tuple[int, int]) -> None:
        key = (a, b) if a <= b else (b, a)
        line = self._lines_between_nodes.get(key)
        if line:
            self.scene.removeItem(line)
            del self._lines_between_nodes[key]

    def _remove_node_and_connections(
        self,
        grid_pos: tuple[int, int],
        cell_id: str,
        connections: list[tuple[str, str, str, tuple[int, int], tuple[int, int]]],
    ) -> None:
        affected_cells = {cell_id}
        for src_cell, direction, dst_cell, src_pos, dst_pos in connections:
            self._remove_neighbor_relation(src_cell, direction, dst_cell)
            self._remove_line_between(src_pos, dst_pos)
            affected_cells.add(src_cell)
            affected_cells.add(dst_cell)

        record = self._nodes_by_grid.get(grid_pos)
        if record is not None:
            self.scene.removeItem(record.item)
            del self._nodes_by_grid[grid_pos]

        if not any(rec.cell_id == cell_id for rec in self._nodes_by_grid.values()):
            if cell_id in self._neighbors:
                del self._neighbors[cell_id]

        for rec in self._nodes_by_grid.values():
            if rec.cell_id in affected_cells:
                self._update_node_neighbors(rec)

        self._recalculate_cell_neighbors_from_scene()
        self._update_all_plus_buttons_visibility()
        if self._current_sequence_index is not None:
            self.update_cell_colors_for_sequence(self._current_sequence_index)

    def _collect_node_connections(
        self, record: _NodeRecord
    ) -> list[tuple[str, str, str, tuple[int, int], tuple[int, int]]]:
        connections: list[tuple[str, str, str, tuple[int, int], tuple[int, int]]] = []
        if not record.cell_id:
            return connections
        for direction, (dx, dy) in DIR_OFFSETS.items():
            if not self._has_grid_connection(record, direction):
                continue
            neighbor_pos = (record.grid_pos[0] + dx, record.grid_pos[1] + dy)
            neighbor_rec = self._nodes_by_grid.get(neighbor_pos)
            if neighbor_rec and neighbor_rec.cell_id:
                connections.append(
                    (
                        record.cell_id,
                        direction,
                        neighbor_rec.cell_id,
                        record.grid_pos,
                        neighbor_pos,
                    )
                )
        return connections

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
        
        adjacency = self._build_cell_adjacency_from_lines()
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
                
                # Check all neighbors (by cell ID)
                for neighbor_id in adjacency.get(current, set()):
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
                for direction, neighbor_ids in neighbors.items():
                    for neighbor_id in neighbor_ids:
                        self._remove_neighbor_relation(record.cell_id, direction, neighbor_id)
                self._neighbors.pop(record.cell_id, None)
            
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
                self._update_node_cell_display(record)
                # Clear all neighbors for this node
                record.item.set_neighbor("up", None)
                record.item.set_neighbor("down", None)
                record.item.set_neighbor("left", None)
                record.item.set_neighbor("right", None)
            else:
                # Remove node normally
                self.scene.removeItem(record.item)
                del self._nodes_by_grid[record.grid_pos]
        # Rebuild aggregated neighbors after removals
        self._recalculate_cell_neighbors_from_scene()
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
            self._update_node_cell_display(initial_node)
        
        # BFS to position all connected cells
        queue = [first_cell]
        visited = {first_cell}
        
        while queue:
            current_cell = queue.pop(0)
            current_pos = cell_to_grid[current_cell]
            neighbors = self._neighbors.get(current_cell, {})
            
            for direction, neighbor_cells in neighbors.items():
                for neighbor_cell in neighbor_cells:
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
                    self._update_node_cell_display(neighbor_node)
                    
                    visited.add(neighbor_cell)
                    queue.append(neighbor_cell)
        
        # Update visual state and draw connections
        for cell_id, grid_pos in cell_to_grid.items():
            record = self._nodes_by_grid.get(grid_pos)
            if record:
                self._update_node_neighbors(record)
                # Draw connections to neighbors (use internal to avoid undo stack on load)
                neighbors = self._neighbors.get(cell_id, {})
                for direction, neighbor_cells in neighbors.items():
                    for neighbor_cell in neighbor_cells:
                        if neighbor_cell:
                            neighbor_pos = cell_to_grid.get(neighbor_cell)
                            if neighbor_pos:
                                self._draw_connection_between(grid_pos, neighbor_pos)
        
        self._recalculate_cell_neighbors_from_scene()
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

    def _format_laminate_label(self, cell_id: Optional[str]) -> str:
        if not cell_id or self._model is None:
            return ""
        try:
            laminate_name = self._model.cell_to_laminate.get(cell_id)
            laminado = self._model.laminados.get(laminate_name) if laminate_name else None
            if laminado is None and hasattr(self._model, "laminados_da_celula"):
                candidates = self._model.laminados_da_celula(cell_id)
                laminado = candidates[0] if candidates else None
            if laminado is not None:
                base_name = getattr(laminado, "nome", laminate_name) or ""
                tag = str(getattr(laminado, "tag", "") or "").strip()
                if tag and f"({tag})" not in base_name:
                    return f"{base_name}({tag})"
                return base_name
            return laminate_name or ""
        except Exception:
            return ""

    def _format_contour_labels(self, cell_id: Optional[str]) -> Tuple[str, str, str, str]:
        if not cell_id or self._model is None:
            return ("", "", "", "")
        contours = getattr(self._model, "cell_contours", {}).get(cell_id, [])
        values: list[str] = []
        for value in contours:
            values.append(str(value) if value is not None else "")
        while len(values) < 4:
            values.append("")
        return (values[0], values[1], values[2], values[3])

    @staticmethod
    def _normalize_contour_value(value: str) -> str:
        return (value or "").strip().lower()

    def _build_preferred_neighbor_cells(
        self,
        source_cell_id: str,
        direction: str,
        available_cells: list[str],
    ) -> list[str]:
        if not source_cell_id:
            return []
        source_contours = self._format_contour_labels(source_cell_id)
        if not any(self._normalize_contour_value(value) for value in source_contours):
            return []
        src_top, src_right, src_bottom, src_left = (
            self._normalize_contour_value(source_contours[0]),
            self._normalize_contour_value(source_contours[1]),
            self._normalize_contour_value(source_contours[2]),
            self._normalize_contour_value(source_contours[3]),
        )

        preferred: list[str] = []
        for cell_id in available_cells:
            contours = self._format_contour_labels(cell_id)
            cand_top, cand_right, cand_bottom, cand_left = (
                self._normalize_contour_value(contours[0]),
                self._normalize_contour_value(contours[1]),
                self._normalize_contour_value(contours[2]),
                self._normalize_contour_value(contours[3]),
            )

            match = False
            if direction == "right":
                match = (
                    cand_left == src_right
                    and cand_top == src_top
                    and cand_bottom == src_bottom
                )
            elif direction == "left":
                match = (
                    cand_right == src_left
                    and cand_top == src_top
                    and cand_bottom == src_bottom
                )
            elif direction == "up":
                match = (
                    cand_bottom == src_top
                    and cand_left == src_left
                    and cand_right == src_right
                )
            elif direction == "down":
                match = (
                    cand_top == src_bottom
                    and cand_left == src_left
                    and cand_right == src_right
                )

            if match:
                preferred.append(cell_id)

        return preferred

    def _update_node_cell_display(self, record: _NodeRecord) -> None:
        if record.cell_id:
            record.item.set_text(record.cell_id)
            record.item.set_laminate_text(self._format_laminate_label(record.cell_id))
            record.item.set_contour_texts(self._format_contour_labels(record.cell_id))
        else:
            record.item.set_text("Select\nCell")
            record.item.set_laminate_text("")
            record.item.set_contour_texts(("", "", "", ""))

    def _prompt_select_cell(
        self,
        record: _NodeRecord,
        *,
        source_record: Optional[_NodeRecord] = None,
        direction: Optional[str] = None,
    ) -> None:
        if not self._cells:
            return
        
        # Allow reusing the same cell ID multiple times when defining neighborhoods
        available_cells = list(self._cells)
        
        if not available_cells:
            return
        
        preferred_cells: list[str] = []
        if (
            source_record
            and direction
            and source_record.cell_id
            and source_record.cell_id in self._cells
        ):
            preferred_cells = self._build_preferred_neighbor_cells(
                source_record.cell_id,
                direction,
                available_cells,
            )

        dialog = SelectCellDialog(
            available_cells,
            current=record.cell_id,
            preferred=preferred_cells,
            parent=self,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        selected = dialog.selected_cell()
        if not selected:
            return

        # If the same cell ID already exists elsewhere, require adjacency before assignment
        existing_positions = [pos for pos, rec in self._nodes_by_grid.items() if rec.cell_id == selected and pos != record.grid_pos]
        if existing_positions:
            is_adjacent = any(abs(pos[0] - record.grid_pos[0]) + abs(pos[1] - record.grid_pos[1]) == 1 for pos in existing_positions)
            if not is_adjacent:
                QMessageBox.warning(
                    self,
                    "Posicionamento inválido",
                    f"Células iguais devem ser vizinhas entre si. Coloque {selected} ao lado de outra {selected} ou conecte-as antes de reutilizar.",
                )
                return
        record.cell_id = selected
        self._update_node_cell_display(record)
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
        self._recalculate_cell_neighbors_from_scene()
        self._update_all_plus_buttons_visibility()
        # Update colors for current sequence if one is selected
        if self._current_sequence_index is not None:
            self.update_cell_colors_for_sequence(self._current_sequence_index)
        self._mark_as_modified()

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
                self._prompt_select_cell(
                    neighbor,
                    source_record=record,
                    direction=direction,
                )
                if not neighbor.cell_id:
                    return
            self._link_cells(
                record.cell_id,
                direction,
                neighbor.cell_id,
                src_pos=record.grid_pos,
                dst_pos=neighbor.grid_pos,
            )
            self._draw_connection_between(record.grid_pos, neighbor.grid_pos)
            self._update_node_neighbors(record)
            self._update_node_neighbors(neighbor)
            self._recalculate_cell_neighbors_from_scene()
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
        self._prompt_select_cell(
            neighbor,
            source_record=record,
            direction=direction,
        )
        if not neighbor.cell_id:
            # User cancelled - remove the created node
            self.scene.removeItem(neighbor.item)
            del self._nodes_by_grid[target_grid]
            return
        self._link_cells_internal(record.cell_id, direction, neighbor.cell_id)
        self._draw_connection_between(record.grid_pos, neighbor.grid_pos)
        self._update_node_neighbors(record)
        self._update_node_neighbors(neighbor)
        self._recalculate_cell_neighbors_from_scene()
        self._expand_scene_rect()
        # Update all plus buttons visibility after creating neighbor
        self._update_all_plus_buttons_visibility()
        # Update colors for current sequence if one is selected
        if self._current_sequence_index is not None:
            self.update_cell_colors_for_sequence(self._current_sequence_index)
        connections = self._collect_node_connections(neighbor)
        self._undo_stack.push(
            AddNodeCommand(
                self,
                neighbor.grid_pos,
                neighbor.item.rect().topLeft(),
                neighbor.cell_id,
                connections,
            )
        )


    def _empty_neighbor_bucket(self) -> dict[str, set[str]]:
        return {dir_name: set() for dir_name in DIR_OFFSETS.keys()}

    def _ensure_cell_mapping_entry(self, cell_id: str) -> None:
        """Ensure a cell has an entry in neighbors dict, even if it has no neighbors."""
        if cell_id not in self._neighbors:
            self._neighbors[cell_id] = self._empty_neighbor_bucket()

    def _add_neighbor_relation(self, src: str, direction: str, dst: str) -> None:
        """Add a bidirectional relation between src->dst respecting direction."""
        self._ensure_cell_mapping_entry(src)
        self._ensure_cell_mapping_entry(dst)
        self._neighbors[src][direction].add(dst)
        self._neighbors[dst][opposite(direction)].add(src)

    def _remove_neighbor_relation(self, src: str, direction: str, dst: str) -> None:
        """Remove a bidirectional relation if it exists."""
        if src in self._neighbors:
            self._neighbors[src].setdefault(direction, set()).discard(dst)
        opposite_dir = opposite(direction)
        if dst in self._neighbors:
            self._neighbors[dst].setdefault(opposite_dir, set()).discard(src)

    def _link_cells_internal(self, src: str, direction: str, dst: str) -> None:
        """Internal method to link cells without undo."""
        self._add_neighbor_relation(src, direction, dst)
        # Update visual state of nodes
        for rec in self._nodes_by_grid.values():
            if rec.cell_id == src:
                rec.item.set_neighbor(direction, dst)
            if rec.cell_id == dst:
                rec.item.set_neighbor(opposite(direction), src)

    def _link_cells(
        self,
        src: str,
        direction: str,
        dst: str,
        *,
        src_pos: tuple[int, int] | None = None,
        dst_pos: tuple[int, int] | None = None,
    ) -> None:
        """Link cells with undo support."""
        if src_pos is None or dst_pos is None:
            for rec in self._nodes_by_grid.values():
                if src_pos is None and rec.cell_id == src:
                    src_pos = rec.grid_pos
                if dst_pos is None and rec.cell_id == dst:
                    dst_pos = rec.grid_pos
                if src_pos is not None and dst_pos is not None:
                    break
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

    def _direction_between(self, a: tuple[int, int], b: tuple[int, int]) -> Optional[str]:
        dx = b[0] - a[0]
        dy = b[1] - a[1]
        for name, (ox, oy) in DIR_OFFSETS.items():
            if (dx, dy) == (ox, oy):
                return name
        return None

    def _recalculate_cell_neighbors_from_scene(self) -> None:
        """Rebuild aggregated neighbors from the currently drawn lines."""
        self._neighbors = {}
        for record in self._nodes_by_grid.values():
            if record.cell_id:
                self._ensure_cell_mapping_entry(record.cell_id)
        for (pos_a, pos_b), _line in list(self._lines_between_nodes.items()):
            rec_a = self._nodes_by_grid.get(pos_a)
            rec_b = self._nodes_by_grid.get(pos_b)
            if not rec_a or not rec_b or not rec_a.cell_id or not rec_b.cell_id:
                continue
            dir_ab = self._direction_between(pos_a, pos_b)
            dir_ba = self._direction_between(pos_b, pos_a)
            if dir_ab:
                self._add_neighbor_relation(rec_a.cell_id, dir_ab, rec_b.cell_id)
            if dir_ba:
                self._add_neighbor_relation(rec_b.cell_id, dir_ba, rec_a.cell_id)

    def _build_cell_adjacency_from_lines(self) -> dict[str, set[str]]:
        """Graph of cell IDs based on the drawn connections between nodes."""
        adjacency: dict[str, set[str]] = {}
        for record in self._nodes_by_grid.values():
            if record.cell_id:
                adjacency.setdefault(record.cell_id, set())
        for (pos_a, pos_b) in self._lines_between_nodes.keys():
            rec_a = self._nodes_by_grid.get(pos_a)
            rec_b = self._nodes_by_grid.get(pos_b)
            if not rec_a or not rec_b or not rec_a.cell_id or not rec_b.cell_id:
                continue
            adjacency.setdefault(rec_a.cell_id, set()).add(rec_b.cell_id)
            adjacency.setdefault(rec_b.cell_id, set()).add(rec_a.cell_id)
        return adjacency

    def _grid_adjacency(self) -> dict[tuple[int, int], set[tuple[int, int]]]:
        """Adjacency between node positions for reachability checks."""
        adjacency: dict[tuple[int, int], set[tuple[int, int]]] = {}
        for record in self._nodes_by_grid.values():
            adjacency.setdefault(record.grid_pos, set())
        for pos_a, pos_b in self._lines_between_nodes.keys():
            adjacency.setdefault(pos_a, set()).add(pos_b)
            adjacency.setdefault(pos_b, set()).add(pos_a)
        return adjacency

    def _cell_positions_by_id(self) -> dict[str, list[tuple[int, int]]]:
        mapping: dict[str, list[tuple[int, int]]] = {}
        for pos, record in self._nodes_by_grid.items():
            if record.cell_id:
                mapping.setdefault(record.cell_id, []).append(pos)
        return mapping

    def _validate_duplicate_cells_connected(self) -> bool:
        """Ensure that every duplicated cell ID belongs to a single connected component."""
        adjacency = self._grid_adjacency()
        positions_by_cell = self._cell_positions_by_id()
        problematic: list[str] = []
        for cell_id, positions in positions_by_cell.items():
            if len(positions) <= 1:
                continue
            # BFS from the first instance through the full graph (connections may pass through outras células)
            start = positions[0]
            visited: set[tuple[int, int]] = set()
            queue = [start]
            while queue:
                current = queue.pop(0)
                if current in visited:
                    continue
                visited.add(current)
                for neighbor in adjacency.get(current, set()):
                    if neighbor not in visited:
                        queue.append(neighbor)
            if not all(pos in visited for pos in positions):
                problematic.append(cell_id)
        if problematic:
            cell_list = ", ".join(sorted(problematic))
            QMessageBox.warning(
                self,
                "Células duplicadas desconectadas",
                f"As células duplicadas precisam estar conectadas entre si. Conecte as instâncias de: {cell_list}.",
            )
            return False
        return True

    def _build_neighbor_nodes_payload(self) -> list[dict[str, object]]:
        """Serialize the current scene with positions to persist múltiplas instâncias."""
        payload: list[dict[str, object]] = []
        for record in self._nodes_by_grid.values():
            if not record.cell_id:
                continue
            neighbors: dict[str, object] = {}
            for direction, (dx, dy) in DIR_OFFSETS.items():
                if not self._has_grid_connection(record, direction):
                    continue
                neighbor_pos = (record.grid_pos[0] + dx, record.grid_pos[1] + dy)
                neighbor_rec = self._nodes_by_grid.get(neighbor_pos)
                if neighbor_rec and neighbor_rec.cell_id:
                    neighbors[direction] = {
                        "grid": [neighbor_pos[0], neighbor_pos[1]],
                        "cell": neighbor_rec.cell_id,
                    }
            payload.append(
                {
                    "cell": record.cell_id,
                    "grid": [record.grid_pos[0], record.grid_pos[1]],
                    "neighbors": neighbors,
                }
            )
        return payload

    def _rebuild_graph_from_nodes_payload(self, payload: list[dict[str, object]]) -> None:
        """Rebuild scene preserving múltiplas instâncias e conexões explícitas."""
        if not payload:
            return

        # Ensure origin exists and capture its base coordinate
        origin = self._nodes_by_grid.get((0, 0))
        if origin is None:
            anchor_x = 300.0
            anchor_y = 200.0
            origin = self._create_node((0, 0), QPointF(anchor_x, anchor_y))
        base_top_left = origin.item.rect().topLeft()

        # Clear other nodes and lines
        nodes_to_remove = [pos for pos in self._nodes_by_grid.keys() if pos != (0, 0)]
        for pos in nodes_to_remove:
            rec = self._nodes_by_grid[pos]
            self.scene.removeItem(rec.item)
            del self._nodes_by_grid[pos]
        for line in list(self._lines_between_nodes.values()):
            self.scene.removeItem(line)
        self._lines_between_nodes.clear()
        self._neighbors = {}

        # Create all nodes
        for entry in payload:
            grid_raw = entry.get("grid", [0, 0])
            try:
                grid_pos = (int(grid_raw[0]), int(grid_raw[1]))
            except Exception:
                continue
            if grid_pos not in self._nodes_by_grid:
                top_left = QPointF(
                    base_top_left.x() + (CELL_SIZE + GAP + PLUS_SIZE) * grid_pos[0],
                    base_top_left.y() + (CELL_SIZE + GAP + PLUS_SIZE) * grid_pos[1],
                )
                self._create_node(grid_pos, top_left)
            rec = self._nodes_by_grid[grid_pos]
            rec.cell_id = str(entry.get("cell", "")) or None
            self._update_node_cell_display(rec)
            if rec.cell_id:
                self._ensure_cell_mapping_entry(rec.cell_id)

        # Create connections
        processed_edges: set[tuple[tuple[int, int], tuple[int, int]]] = set()
        for entry in payload:
            grid_raw = entry.get("grid", [0, 0])
            try:
                src_pos = (int(grid_raw[0]), int(grid_raw[1]))
            except Exception:
                continue
            neighbors = entry.get("neighbors", {}) or {}
            if not isinstance(neighbors, dict):
                continue
            for direction, data in neighbors.items():
                target_grid = None
                if isinstance(data, dict) and "grid" in data:
                    try:
                        target_grid = (int(data["grid"][0]), int(data["grid"][1]))
                    except Exception:
                        target_grid = None
                elif isinstance(data, (list, tuple)) and len(data) >= 2:
                    try:
                        target_grid = (int(data[0]), int(data[1]))
                    except Exception:
                        target_grid = None
                if target_grid is None:
                    continue
                key = (src_pos, target_grid) if src_pos <= target_grid else (target_grid, src_pos)
                if key in processed_edges:
                    continue
                processed_edges.add(key)
                self._draw_connection_between(src_pos, target_grid)
                src_rec = self._nodes_by_grid.get(src_pos)
                dst_rec = self._nodes_by_grid.get(target_grid)
                if src_rec and dst_rec and src_rec.cell_id and dst_rec.cell_id:
                    if direction in DIR_OFFSETS:
                        self._add_neighbor_relation(src_rec.cell_id, direction, dst_rec.cell_id)
                    else:
                        dir_guess = self._direction_between(src_pos, target_grid)
                        if dir_guess:
                            self._add_neighbor_relation(src_rec.cell_id, dir_guess, dst_rec.cell_id)
        self._recalculate_cell_neighbors_from_scene()
        # Update visuals and bounds
        for rec in self._nodes_by_grid.values():
            self._update_node_neighbors(rec)
        self._expand_scene_rect()
        if self._current_sequence_index is not None:
            self.update_cell_colors_for_sequence(self._current_sequence_index)

    def _update_node_neighbors(self, record: _NodeRecord) -> None:
        """Update the visual state of plus buttons based on current neighbors."""
        if not record.cell_id:
            # No cell assigned yet, show all plus buttons
            for dir_name in DIR_OFFSETS.keys():
                record.item.set_neighbor(dir_name, None)
            return
        
        for dir_name, (dx, dy) in DIR_OFFSETS.items():
            neighbor_pos = (record.grid_pos[0] + dx, record.grid_pos[1] + dy)
            neighbor_rec = self._nodes_by_grid.get(neighbor_pos)
            has_connection = self._has_grid_connection(record, dir_name)
            neighbor_id = neighbor_rec.cell_id if (has_connection and neighbor_rec and neighbor_rec.cell_id) else None
            record.item.set_neighbor(dir_name, neighbor_id)


class SelectCellDialog(QDialog):
    """Dialog listing available cells for selection."""

    def __init__(
        self,
        cells: list[str],
        *,
        current: Optional[str] = None,
        preferred: Optional[list[str]] = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Selecionar Celula")
        self.resize(320, 400)
        layout = QVBoxLayout(self)
        self.list = QListWidget(self)
        preferred = preferred or []
        filtered_preferred = [cell for cell in preferred if cell in cells]
        preferred_set = {cell for cell in filtered_preferred}

        ordered_cells: list[str] = []
        ordered_cells.extend(filtered_preferred)
        for cell in cells:
            if cell not in preferred_set:
                ordered_cells.append(cell)

        if filtered_preferred and len(ordered_cells) > len(filtered_preferred):
            for cell in filtered_preferred:
                item = QListWidgetItem(cell)
                self.list.addItem(item)
                if current and cell == current:
                    self.list.setCurrentItem(item)
            separator = QListWidgetItem("---")
            separator.setFlags(Qt.NoItemFlags)
            separator.setTextAlignment(Qt.AlignCenter)
            separator.setForeground(QColor(140, 140, 140))
            self.list.addItem(separator)
            for cell in cells:
                if cell not in preferred_set:
                    item = QListWidgetItem(cell)
                    self.list.addItem(item)
                    if current and cell == current:
                        self.list.setCurrentItem(item)
        else:
            for cell in ordered_cells:
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
