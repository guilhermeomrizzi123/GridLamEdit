"""Intermediate laminate suggestion window."""

from __future__ import annotations

from typing import Optional

from PySide6.QtCore import Qt, QSize, QEvent
from PySide6.QtGui import QAction, QColor, QPainter, QPen, QBrush, QFont, QWheelEvent
from PySide6.QtWidgets import (
    QDialog,
    QGraphicsScene,
    QGraphicsTextItem,
    QGraphicsView,
    QLabel,
    QStyle,
    QToolBar,
    QVBoxLayout,
    QSizePolicy,
)

from gridlamedit.io.spreadsheet import GridModel


class IntermediateLaminateWindow(QDialog):
    """Dialog to suggest an intermediate laminate between two cells."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Sugestão de Laminado Intermediário")
        self.setWindowFlags(
            self.windowFlags() | Qt.WindowMinMaxButtonsHint | Qt.WindowSystemMenuHint
        )
        self.setWindowFlag(Qt.Window, True)
        self.resize(1100, 720)

        self._model: Optional[GridModel] = None
        self._project_manager = None

        main_layout = QVBoxLayout(self)

        toolbar = QToolBar(self)
        toolbar.setMovable(False)
        toolbar.setToolButtonStyle(Qt.ToolButtonTextOnly)
        toolbar.setIconSize(QSize(18, 18))
        toolbar.setWindowTitle("Navegação")

        zoom_in_action = QAction(self.style().standardIcon(QStyle.SP_ArrowUp), "Zoom +", self)
        zoom_in_action.setToolTip("Aumentar zoom")
        zoom_in_action.triggered.connect(lambda: self._apply_zoom(1.15))
        toolbar.addAction(zoom_in_action)

        zoom_out_action = QAction(self.style().standardIcon(QStyle.SP_ArrowDown), "Zoom -", self)
        zoom_out_action.setToolTip("Diminuir zoom")
        zoom_out_action.triggered.connect(lambda: self._apply_zoom(1 / 1.15))
        toolbar.addAction(zoom_out_action)

        reset_zoom_action = QAction(self.style().standardIcon(QStyle.SP_BrowserReload), "Reset Zoom", self)
        reset_zoom_action.setToolTip("Restaurar zoom")
        reset_zoom_action.triggered.connect(self._reset_zoom)
        toolbar.addAction(reset_zoom_action)

        center_action = QAction(self.style().standardIcon(QStyle.SP_ArrowRight), "Centralizar", self)
        center_action.setToolTip("Centralizar visualização")
        center_action.triggered.connect(self._center_view)
        toolbar.addAction(center_action)

        toolbar.addSeparator()
        toolbar.addWidget(QLabel("Pan: botão do meio", self))

        main_layout.addWidget(toolbar)

        self.view = QGraphicsView(self)
        self.view.setRenderHint(QPainter.Antialiasing, True)
        self.view.setBackgroundBrush(QColor(248, 248, 248))
        self.view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.view.setDragMode(QGraphicsView.NoDrag)
        self.view.setInteractive(False)
        main_layout.addWidget(self.view)

        self.scene = QGraphicsScene(self)
        self.scene.setSceneRect(0, 0, 2000, 1200)
        self.view.setScene(self.scene)

        self._is_panning = False
        self._last_pan_point = None
        self._setup_view_interaction()

        self._build_schematic_scene()

    def populate_from_project(self, model: Optional[GridModel], project_manager=None) -> None:
        """Store references for future use (interface still in construction)."""
        self._model = model
        self._project_manager = project_manager

    def refresh_from_model(self) -> None:
        """Placeholder for future updates from the model."""
        return

    def _build_schematic_scene(self) -> None:
        self.scene.clear()

        width = 1800.0
        height = 920.0
        self.scene.setSceneRect(0, 0, width, height)

        margin_x = 160.0
        margin_y = 80.0
        block_width = width - 2 * margin_x
        block_height = 560.0
        block_top = margin_y + 60.0

        outline_pen = QPen(QColor(30, 30, 30))
        outline_pen.setWidthF(2.0)
        self.scene.addRect(
            margin_x,
            block_top,
            block_width,
            block_height,
            outline_pen,
            QBrush(Qt.NoBrush),
        )

        # Stringer dashed bands (3 bands aligned to each section center)
        dashed_pen = QPen(QColor(0, 90, 255))
        dashed_pen.setWidthF(5.0)
        dashed_pen.setStyle(Qt.DashLine)
        band_height = 60.0
        orange_height = band_height
        remaining = max(1.0, block_height - orange_height)
        green_height = remaining / 2.0
        blue_height = remaining / 2.0

        stringer_positions = [
            block_top,
            block_top + green_height,
            block_top + green_height + orange_height + blue_height - band_height,
        ]

        for idx, y in enumerate(stringer_positions, start=1):
            rect = self.scene.addRect(
                margin_x - 60.0,
                y,
                block_width + 120.0,
                band_height,
                dashed_pen,
                QBrush(Qt.NoBrush),
            )
            rect.setZValue(4)
            label = QGraphicsTextItem(f"STRINGER {idx}")
            label.setDefaultTextColor(QColor(60, 60, 60))
            label.setPos(margin_x - 130.0, y - 26.0)
            label.setZValue(4)
            self.scene.addItem(label)

        # Main colored blocks
        sections = [
            (
                QColor(153, 213, 92),
                "SELECIONAR CELULA COM MENOR ESPESSURA",
                green_height,
            ),
            (
                QColor(248, 190, 60),
                "NOVA CELULA COM NOVO LAMINADO",
                orange_height,
            ),
            (
                QColor(220, 220, 220),
                "SELECIONAR CELULA COM MAIOR ESPESSURA",
                blue_height,
            ),
        ]

        font = QFont()
        font.setPointSize(12)
        font.setBold(True)

        current_y = block_top
        for color, text, height in sections:
            y = current_y
            rect_item = self.scene.addRect(
                margin_x,
                y,
                block_width,
                height,
                outline_pen,
                QBrush(color),
            )
            rect_item.setZValue(1)
            label = QGraphicsTextItem(text)
            label.setFont(font)
            label.setDefaultTextColor(QColor(20, 20, 20))
            label_rect = label.boundingRect()
            label.setPos(
                margin_x + (block_width - label_rect.width()) / 2.0,
                y + (height - label_rect.height()) / 2.0,
            )
            label.setZValue(2)
            self.scene.addItem(label)
            current_y += height

        hint = QGraphicsTextItem(
            "Use a roda do mouse para zoom e o botão do meio para deslocar."
        )
        hint.setDefaultTextColor(QColor(90, 90, 90))
        hint.setPos(margin_x, margin_y - 10.0)
        self.scene.addItem(hint)

    def _setup_view_interaction(self) -> None:
        self.view.viewport().installEventFilter(self)

    def _apply_zoom(self, factor: float) -> None:
        self.view.scale(factor, factor)

    def _reset_zoom(self) -> None:
        self.view.resetTransform()

    def _center_view(self) -> None:
        rect = self.scene.sceneRect()
        self.view.centerOn(rect.center())

    def eventFilter(self, obj, event):
        if obj is self.view.viewport():
            if event.type() == QEvent.MouseButtonPress and event.button() == Qt.MiddleButton:
                self._is_panning = True
                self._last_pan_point = event.pos()
                return True
            if event.type() == QEvent.MouseMove and self._is_panning and self._last_pan_point is not None:
                delta = event.pos() - self._last_pan_point
                self._last_pan_point = event.pos()
                self.view.horizontalScrollBar().setValue(
                    self.view.horizontalScrollBar().value() - delta.x()
                )
                self.view.verticalScrollBar().setValue(
                    self.view.verticalScrollBar().value() - delta.y()
                )
                return True
            if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.MiddleButton:
                self._is_panning = False
                self._last_pan_point = None
                return True
            if event.type() == QEvent.Wheel:
                wheel: QWheelEvent = event
                if wheel.modifiers() & Qt.ControlModifier:
                    return False
                angle = wheel.angleDelta().y()
                factor = 1.15 if angle > 0 else 1 / 1.15
                self._apply_zoom(factor)
                return True
        return super().eventFilter(obj, event)
