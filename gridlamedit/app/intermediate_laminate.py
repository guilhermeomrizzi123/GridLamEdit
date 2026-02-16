"""Intermediate laminate suggestion window."""

from __future__ import annotations

from typing import Optional

from PySide6.QtCore import Qt, QSize, QEvent
from PySide6.QtGui import QAction, QColor, QPainter, QWheelEvent
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

        self._add_placeholder_message()

    def populate_from_project(self, model: Optional[GridModel], project_manager=None) -> None:
        """Store references for future use (interface still in construction)."""
        self._model = model
        self._project_manager = project_manager

    def refresh_from_model(self) -> None:
        """Placeholder for future updates from the model."""
        return

    def _add_placeholder_message(self) -> None:
        text = (
            "Interface em construção.\n"
            "Use a roda do mouse para zoom e o botão do meio para deslocar."
        )
        item = QGraphicsTextItem(text)
        item.setDefaultTextColor(QColor(90, 90, 90))
        item.setPos(60, 60)
        self.scene.addItem(item)

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
