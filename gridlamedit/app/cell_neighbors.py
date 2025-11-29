"""Cell Neighbors editor window.

This lightweight UI lets users define neighbor relationships between
grid cells visually using square nodes and '+' buttons around them.

The scene is intentionally simple and self-contained to avoid coupling
with the rest of the app for now. It exposes one public method:

    get_neighbors_mapping() -> dict[str, dict[str, Optional[str]]]

which returns the current neighbor map in memory.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Optional, Tuple

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QPainterPath, QPen
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
)

try:
    # Optional: reference to the GridModel for populate convenience
    from gridlamedit.io.spreadsheet import GridModel
except Exception:  # pragma: no cover - optional import for loose coupling
    GridModel = object  # type: ignore


CELL_SIZE = 80.0
PLUS_SIZE = 18.0
GAP = 28.0
MARGIN = 24.0

COLOR_BG = QColor(220, 220, 220)
COLOR_CELL = QColor(27, 36, 44)
COLOR_TEXT = QColor(255, 255, 255)
COLOR_PLUS = QColor(128, 128, 128)
COLOR_DASH = QColor(140, 140, 140)


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


class PlusButtonItem(QGraphicsRectItem):
    """Small square "+" button around a node."""

    def __init__(self, rect: QRectF, parent: Optional[QGraphicsItem] = None) -> None:
        super().__init__(rect, parent)
        self.setFlags(QGraphicsItem.ItemIsSelectable | QGraphicsItem.ItemIgnoresParentOpacity)
        self.setBrush(COLOR_PLUS)
        self.setPen(Qt.NoPen)
        self.on_click = None

        self._label = QGraphicsSimpleTextItem("+", self)
        font = self._label.font()
        font.setBold(True)
        self._label.setFont(font)
        self._label.setBrush(COLOR_TEXT)
        # Center inside the square
        text_rect = self._label.boundingRect()
        self._label.setPos(
            rect.x() + (rect.width() - text_rect.width()) / 2,
            rect.y() + (rect.height() - text_rect.height()) / 2,
        )

    # Use mousePress to emulate a simple button
    def mousePressEvent(self, event):  # noqa: N802
        if callable(self.on_click):
            self.on_click()
        event.accept()


class CellNodeItem(QGraphicsRectItem):
    """Square node with central label and four plus buttons.
    Only shows '+' buttons on sides without neighbors. Lines are drawn behind the block.
    """
    def __init__(self, rect: QRectF, *, parent: Optional[QGraphicsItem] = None) -> None:
        super().__init__(rect, parent)
        self.setBrush(COLOR_CELL)
        pen = QPen(QColor(11, 92, 115))
        pen.setWidth(2)
        self.setPen(pen)
        self.setAcceptHoverEvents(True)
        self._label = QGraphicsSimpleTextItem("Select\nCell", self)
        f: QFont = self._label.font()
        f.setBold(True)
        f.setPointSize(f.pointSize() + 2)
        self._label.setFont(f)
        self._label.setBrush(COLOR_TEXT)
        self._recenter_label()
        self.on_select_cell = None
        self.on_add_neighbor = None
        self.plus_items: dict[str, PlusButtonItem] = {}
        self.plus_lines: dict[str, QGraphicsLineItem] = {}
        self._neighbors: dict[str, Optional[str]] = {"up": None, "down": None, "left": None, "right": None}
        self._create_plus_buttons()

    def set_text(self, text: str) -> None:
        self._label.setText(text)
        self._recenter_label()

    def set_neighbor(self, direction: str, cell_id: Optional[str]) -> None:
        self._neighbors[direction] = cell_id
        self._update_plus_buttons()

    def _update_plus_buttons(self) -> None:
        for direction in DIR_OFFSETS.keys():
            has_neighbor = self._neighbors.get(direction)
            btn = self.plus_items.get(direction)
            line = self.plus_lines.get(direction)
            if has_neighbor:
                if btn:
                    btn.setVisible(False)
                if line:
                    line.setVisible(False)
            else:
                if btn:
                    btn.setVisible(True)
                if line:
                    line.setVisible(True)

        def _recenter_label(self) -> None:
            rect = self.rect()
            tb = self._label.boundingRect()
            self._label.setPos(
                rect.x() + (rect.width() - tb.width()) / 2,
                rect.y() + (rect.height() - tb.height()) / 2,
            )

        def _create_plus_buttons(self) -> None:
            cx = self.rect().center().x()
            cy = self.rect().center().y()
            half = CELL_SIZE / 2.0
            offsets: dict[str, Tuple[float, float]] = {
                "up": (0.0, -half - GAP),
                "down": (0.0, half + GAP),
                "left": (-half - GAP, 0.0),
                "right": (half + GAP, 0.0),
            }
            for direction, (dx, dy) in offsets.items():
                bx = cx + dx - PLUS_SIZE / 2.0
                by = cy + dy - PLUS_SIZE / 2.0
                button = PlusButtonItem(QRectF(bx, by, PLUS_SIZE, PLUS_SIZE), self)
                line = QGraphicsLineItem(cx, cy, cx + dx, cy + dy, self)
                pen = QPen(COLOR_DASH)
                pen.setStyle(Qt.DashLine)
                pen.setWidth(2)
                line.setPen(pen)
                line.setZValue(-1)  # Draw lines behind everything
                self.plus_items[direction] = button
                self.plus_lines[direction] = line
                def handler(dir_name=direction):
                    if callable(self.on_add_neighbor):
                        self.on_add_neighbor(dir_name)
                button.on_click = handler

        def mousePressEvent(self, event):  # noqa: N802
            if callable(self.on_select_cell):
                self.on_select_cell()
            event.accept()


class CellNodeItem(QGraphicsRectItem):
    """Square node with central label and four plus buttons.

    This item is not draggable. It exposes callbacks for:
      - on_select_cell(): user clicked the square to select/change the cell
      - on_add_neighbor(direction): user clicked a "+" in the given direction
    """

    def __init__(self, rect: QRectF, *, parent: Optional[QGraphicsItem] = None) -> None:
        super().__init__(rect, parent)
        self.setBrush(COLOR_CELL)
        pen = QPen(QColor(11, 92, 115))
        pen.setWidth(2)
        self.setPen(pen)
        self.setAcceptHoverEvents(True)

        self._label = QGraphicsSimpleTextItem("Select\nCell", self)
        f: QFont = self._label.font()
        f.setBold(True)
        f.setPointSize(f.pointSize() + 2)
        self._label.setFont(f)
        self._label.setBrush(COLOR_TEXT)
        self._recenter_label()

        # Callbacks set by owner
        self.on_select_cell = None
        self.on_add_neighbor = None

        # Plus buttons and dashed lines to them
        self.plus_items: dict[str, PlusButtonItem] = {}
        self.plus_lines: dict[str, QGraphicsLineItem] = {}
        self._create_plus_buttons()

    def set_text(self, text: str) -> None:
        self._label.setText(text)
        self._recenter_label()

    def _recenter_label(self) -> None:
        rect = self.rect()
        tb = self._label.boundingRect()
        self._label.setPos(
            rect.x() + (rect.width() - tb.width()) / 2,
            rect.y() + (rect.height() - tb.height()) / 2,
        )

    def _create_plus_buttons(self) -> None:
        cx = self.rect().center().x()
        cy = self.rect().center().y()
        half = CELL_SIZE / 2.0
        d = HALF = half
        # mapping dir -> (button center offset x,y)
        offsets: dict[str, Tuple[float, float]] = {
            "up": (0.0, -HALF - GAP),
            "down": (0.0, HALF + GAP),
            "left": (-HALF - GAP, 0.0),
            "right": (HALF + GAP, 0.0),
        }
        for direction, (dx, dy) in offsets.items():
            bx = cx + dx - PLUS_SIZE / 2.0
            by = cy + dy - PLUS_SIZE / 2.0
            button = PlusButtonItem(QRectF(bx, by, PLUS_SIZE, PLUS_SIZE), self)
            # dashed line from node center to the button center
            line = QGraphicsLineItem(cx, cy, cx + dx, cy + dy, self)
            pen = QPen(COLOR_DASH)
            pen.setStyle(Qt.DashLine)
            pen.setWidth(2)
            line.setPen(pen)
            line.setZValue(-1)  # Draw lines behind everything
            self.plus_items[direction] = button
            self.plus_lines[direction] = line

            def handler(dir_name=direction):
                if callable(self.on_add_neighbor):
                    self.on_add_neighbor(dir_name)
            button.on_click = handler

    def mousePressEvent(self, event):  # noqa: N802
        if callable(self.on_select_cell):
            self.on_select_cell()
        event.accept()

    def hoverEnterEvent(self, event):  # noqa: N802
        self.setBrush(QColor(45, 60, 70))  # Slightly lighter for hover
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event):  # noqa: N802
        self.setBrush(COLOR_CELL)
        super().hoverLeaveEvent(event)


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
        self._cells: list[str] = []

        layout = QHBoxLayout(self)
        self.view = QGraphicsView(self)
        self.view.setRenderHint(self.view.renderHints(), True)
        self.view.setBackgroundBrush(COLOR_BG)
        self.view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout.addWidget(self.view)

        self.scene = QGraphicsScene(self)
        self.scene.setSceneRect(0, 0, 2000, 1200)
        self.view.setScene(self.scene)
        self._setup_view_interaction()

        # Maintain nodes by grid position and by cell id
        self._nodes_by_grid: Dict[tuple[int, int], _NodeRecord] = {}
        self._lines_between_nodes: dict[tuple[tuple[int, int], tuple[int, int]], QGraphicsLineItem] = {}

        # neighbors mapping public structure
        self._neighbors: Dict[str, dict[str, Optional[str]]] = {}

        # Add initial placeholder node near bottom-right
        anchor_x = self.scene.sceneRect().width() - (MARGIN + CELL_SIZE)
        anchor_y = self.scene.sceneRect().height() - (MARGIN + CELL_SIZE)
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

    # -------- Public API ---------
    def populate_from_project(self, model: Optional[GridModel]) -> None:
        self._model = model
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
        self._cells = cells

    def get_neighbors_mapping(self) -> Dict[str, dict[str, Optional[str]]]:
        # Return a deep-ish copy to avoid external mutation surprises
        result: Dict[str, dict[str, Optional[str]]] = {}
        for cell, mapping in self._neighbors.items():
            result[cell] = dict(mapping)
        return result

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

        item.on_select_cell = on_select_cell
        item.on_add_neighbor = on_add_neighbor
        return record

    def _prompt_select_cell(self, record: _NodeRecord) -> None:
        if not self._cells:
            return
        dialog = SelectCellDialog(self._cells, current=record.cell_id, parent=self)
        if dialog.exec() != QDialog.Accepted:
            return
        selected = dialog.selected_cell()
        if not selected:
            return
        record.cell_id = selected
        record.item.set_text(record.cell_id)
        self._ensure_cell_mapping_entry(record.cell_id)
        for dir_name, (dx, dy) in DIR_OFFSETS.items():
            neighbor_pos = (record.grid_pos[0] + dx, record.grid_pos[1] + dy)
            neighbor = self._nodes_by_grid.get(neighbor_pos)
            if neighbor and neighbor.cell_id:
                self._link_cells(record.cell_id, dir_name, neighbor.cell_id)
                self._draw_connection_between(record.grid_pos, neighbor.grid_pos)
        self._update_node_neighbors(record)

    def _handle_add_neighbor(self, record: _NodeRecord, direction: str) -> None:
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
            return

        # Create the node at aligned position
        base = self._nodes_by_grid[(0, 0)].item.rect().topLeft()
        top_left = QPointF(
            base.x() + (CELL_SIZE + GAP + PLUS_SIZE) * target_grid[0],
            base.y() + (CELL_SIZE + GAP + PLUS_SIZE) * target_grid[1],
        )
        neighbor = self._create_node(target_grid, top_left)
        self._prompt_select_cell(neighbor)
        if not neighbor.cell_id:
            return
        self._link_cells(record.cell_id, direction, neighbor.cell_id)
        self._draw_connection_between(record.grid_pos, neighbor.grid_pos)
        self._update_node_neighbors(record)
        self._update_node_neighbors(neighbor)


    def _ensure_cell_mapping_entry(self, cell_id: str) -> None:
        if cell_id not in self._neighbors:
            self._neighbors[cell_id] = {"up": None, "down": None, "left": None, "right": None}

    def _link_cells(self, src: str, direction: str, dst: str) -> None:
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

    def _draw_connection_between(self, a: tuple[int, int], b: tuple[int, int]) -> None:
        key = (a, b) if a <= b else (b, a)
        if key in self._lines_between_nodes:
            return
        a_item = self._nodes_by_grid.get(a)
        b_item = self._nodes_by_grid.get(b)
        if not a_item or not b_item:
            return
        ac = a_item.item.rect().center()
        bc = b_item.item.rect().center()
        line = QGraphicsLineItem(ac.x(), ac.y(), bc.x(), bc.y())
        pen = QPen(COLOR_DASH)
        pen.setStyle(Qt.DashLine)
        pen.setWidth(2)
        line.setPen(pen)
        line.setZValue(-2)  # Draw connection lines behind everything
        self.scene.addItem(line)
        self._lines_between_nodes[key] = line

    def _update_node_neighbors(self, record: _NodeRecord) -> None:
        neighbors = self._neighbors.get(record.cell_id)
        if neighbors:
            for dir_name, cell_id in neighbors.items():
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
