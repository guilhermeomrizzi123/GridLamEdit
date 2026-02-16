"""Intermediate laminate suggestion window."""

from __future__ import annotations

from typing import Optional

from PySide6.QtCore import Qt, QSize, QEvent, QPoint
from PySide6.QtGui import QAction, QColor, QPainter, QPen, QBrush, QFont, QWheelEvent
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QGraphicsScene,
    QGraphicsProxyWidget,
    QGraphicsTextItem,
    QGraphicsView,
    QLabel,
    QMessageBox,
    QInputDialog,
    QPushButton,
    QHBoxLayout,
    QSpinBox,
    QStyle,
    QToolBar,
    QVBoxLayout,
    QSizePolicy,
)

from gridlamedit.io.spreadsheet import GridModel, count_oriented_layers
from gridlamedit.app.cell_neighbors import SelectCellDialog


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
        self._selected_min_cell: Optional[str] = None
        self._selected_max_cell: Optional[str] = None
        self._min_cell_button: Optional[QPushButton] = None
        self._max_cell_button: Optional[QPushButton] = None
        self._distance_button: Optional[QPushButton] = None
        self._distance_mm: Optional[float] = None
        self._dropoff_ratio_button: Optional[QPushButton] = None
        self._dropoff_ratio: Optional[tuple[int, int]] = None
        self._layer_thickness_button: Optional[QPushButton] = None
        self._layer_thickness_mm: Optional[float] = None
        self._cell_button_proxies: list[QGraphicsProxyWidget] = []
        self._summary_item: Optional[QGraphicsTextItem] = None

        main_layout = QVBoxLayout(self)

        top_bar = QHBoxLayout()
        top_bar.setSpacing(8)
        top_bar.setAlignment(Qt.AlignLeft)
        dropoff_button = self._build_cell_select_button("Razão de Drop Off")
        dropoff_button.setMinimumWidth(180)
        top_bar.addWidget(dropoff_button)
        thickness_button = self._build_cell_select_button("Espessura da Camada (mm)")
        thickness_button.setMinimumWidth(200)
        top_bar.addWidget(thickness_button)
        main_layout.addLayout(top_bar)
        self._dropoff_ratio_button = dropoff_button
        dropoff_button.setStyleSheet("")
        dropoff_button.clicked.connect(self._on_dropoff_ratio_clicked)
        self._layer_thickness_button = thickness_button
        thickness_button.setStyleSheet("")
        thickness_button.clicked.connect(self._on_layer_thickness_clicked)

        self.view = QGraphicsView(self)
        self.view.setRenderHint(QPainter.Antialiasing, True)
        self.view.setBackgroundBrush(QColor(248, 248, 248))
        self.view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.view.setDragMode(QGraphicsView.NoDrag)
        self.view.setInteractive(True)
        self.view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        main_layout.addWidget(self.view)

        self.scene = QGraphicsScene(self)
        self.scene.setSceneRect(-50000, -50000, 100000, 100000)
        self.view.setScene(self.scene)

        self._is_panning = False
        self._last_pan_point: Optional[QPoint] = None
        self._last_pan_scene_pos = None
        self._setup_view_interaction()

        self._build_schematic_scene()
        self._center_view()

    def populate_from_project(self, model: Optional[GridModel], project_manager=None) -> None:
        """Store references for future use (interface still in construction)."""
        self._model = model
        self._project_manager = project_manager

    def refresh_from_model(self) -> None:
        """Placeholder for future updates from the model."""
        return

    def _build_schematic_scene(self) -> None:
        self.scene.clear()
        self._cell_button_proxies.clear()

        width = 1800.0
        height = 920.0
        self.scene.setSceneRect(-50000, -50000, 100000, 100000)

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

        # Distance between stringers (red marker)
        red_pen = QPen(QColor(220, 38, 38))
        red_pen.setWidthF(3.0)
        top_y = block_top + band_height
        bottom_y = block_top + green_height
        marker_x = margin_x + block_width + 18.0
        red_main = self.scene.addLine(marker_x, top_y, marker_x, bottom_y, red_pen)
        red_top = self.scene.addLine(marker_x - 10.0, top_y, marker_x + 10.0, top_y, red_pen)
        red_bottom = self.scene.addLine(marker_x - 10.0, bottom_y, marker_x + 10.0, bottom_y, red_pen)
        red_main.setZValue(6)
        red_top.setZValue(6)
        red_bottom.setZValue(6)

        # Main colored blocks
        sections = [
            (
                QColor(153, 213, 92),
                "SELECIONAR CELULA COM MENOR ESPESSURA",
                green_height,
                "min",
            ),
            (
                QColor(248, 190, 60),
                "NOVA CELULA COM NOVO LAMINADO",
                orange_height,
                "label",
            ),
            (
                QColor(220, 220, 220),
                "SELECIONAR CELULA COM MAIOR ESPESSURA",
                blue_height,
                "max",
            ),
        ]

        font = QFont()
        font.setPointSize(12)
        font.setBold(True)

        current_y = block_top
        for color, text, height, mode in sections:
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
            if mode in {"min", "max"}:
                button = self._build_cell_select_button(text)
                proxy = QGraphicsProxyWidget()
                proxy.setWidget(button)
                button_width = button.sizeHint().width()
                button_height = button.sizeHint().height()
                button.setFixedSize(button_width, button_height)
                proxy.setPos(
                    margin_x + (block_width - button_width) / 2.0,
                    y + (height - button_height) / 2.0,
                )
                proxy.setZValue(2)
                self.scene.addItem(proxy)
                self._cell_button_proxies.append(proxy)
                if mode == "min":
                    self._min_cell_button = button
                    button.clicked.connect(lambda _=False: self._select_cell("min"))
                else:
                    self._max_cell_button = button
                    button.clicked.connect(lambda _=False: self._select_cell("max"))
            else:
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


        # Distance button label
        distance_button = self._build_cell_select_button("?(mm)")
        distance_button.setStyleSheet(
            distance_button.styleSheet()
            + "QPushButton { text-align: center; padding: 4px 10px; }"
        )
        distance_proxy = QGraphicsProxyWidget()
        distance_proxy.setWidget(distance_button)
        distance_width = distance_button.sizeHint().width()
        distance_height = distance_button.sizeHint().height()
        distance_button.setFixedSize(distance_width, distance_height)
        distance_proxy.setPos(marker_x + 18.0, top_y + (bottom_y - top_y) / 2.0 - distance_height / 2.0)
        distance_proxy.setZValue(5)
        self.scene.addItem(distance_proxy)
        self._cell_button_proxies.append(distance_proxy)
        self._distance_button = distance_button
        distance_button.clicked.connect(self._on_distance_button_clicked)

        # Summary box
        summary_x = margin_x - 10.0
        summary_y = block_top + block_height + 40.0
        summary_width = 620.0
        summary_height = 180.0
        summary_rect = self.scene.addRect(
            summary_x,
            summary_y,
            summary_width,
            summary_height,
            QPen(QColor(160, 160, 160), 1.2),
            QBrush(QColor(252, 252, 252)),
        )
        summary_rect.setZValue(1)

        summary_html = (
            "<div style='font-family:Segoe UI; font-size:11pt; color:#222;'>"
            "<div style='font-weight:600; margin-bottom:6px;'>Resumo</div>"
            "<table style='border-collapse:collapse;'>"
            "<tr><td>Espaço para Drop Off</td><td style='padding-left:16px;'>—</td></tr>"
            "<tr><td>Razão de Drop Off</td><td style='padding-left:16px;'>—</td></tr>"
            "<tr><td>Diferença de Camadas Entre Células</td><td style='padding-left:16px;'>—</td></tr>"
            "</table>"
            "<div style='margin:8px 0; border-bottom:1px solid #cfcfcf;'></div>"
            "<div style='color:#555;'>Nesse espaço iremos adicionar o resultado<br>da análise.</div>"
            "</div>"
        )
        summary_item = QGraphicsTextItem()
        summary_item.setHtml(summary_html)
        summary_item.setDefaultTextColor(QColor(20, 20, 20))
        summary_item.setPos(summary_x + 12.0, summary_y + 10.0)
        summary_item.setZValue(2)
        self.scene.addItem(summary_item)
        self._summary_item = summary_item
        self._update_summary_box()
        self._center_view()

    def _build_cell_select_button(self, text: str) -> QPushButton:
        button = QPushButton(text)
        button.installEventFilter(self)
        return button

    def _select_cell(self, mode: str) -> None:
        cells = self._available_cells()
        if not cells:
            QMessageBox.information(
                self,
                "Selecionar Celula",
                "Carregue um projeto com celulas para selecionar.",
            )
            return
        current = self._selected_min_cell if mode == "min" else self._selected_max_cell
        dialog = SelectCellDialog(cells, current=current, parent=self)
        if dialog.exec() != QDialog.Accepted:
            return
        selected = dialog.selected_cell()
        if not selected:
            return
        if mode == "min":
            self._selected_min_cell = selected
            if self._min_cell_button is not None:
                self._min_cell_button.setText(selected)
        else:
            self._selected_max_cell = selected
            if self._max_cell_button is not None:
                self._max_cell_button.setText(selected)
        self._update_summary_box()

    def _available_cells(self) -> list[str]:
        if self._model is None:
            return []
        cells = list(self._model.celulas_ordenadas or [])
        if not cells:
            cells = sorted(self._model.cell_to_laminate.keys())
        return cells

    def _setup_view_interaction(self) -> None:
        self.view.viewport().installEventFilter(self)
        self.view.installEventFilter(self)
        self.view.setMouseTracking(True)

    def _apply_zoom(self, factor: float) -> None:
        self.view.scale(factor, factor)

    def _reset_zoom(self) -> None:
        self.view.resetTransform()

    def _center_view(self) -> None:
        if self.scene is None or self.view is None:
            return
        rect = self.scene.itemsBoundingRect()
        if rect.isNull():
            rect = self.scene.sceneRect()
        self.view.centerOn(rect.center())

    def eventFilter(self, obj, event):
        if not hasattr(self, "view"):
            return super().eventFilter(obj, event)
        if obj is self.view or obj is self.view.viewport() or isinstance(obj, QPushButton):
            if event.type() == QEvent.MouseButtonPress and event.button() == Qt.MiddleButton:
                global_pos = event.globalPosition().toPoint()
                self._start_pan(global_pos)
                return True
            if event.type() == QEvent.MouseMove and self._is_panning and self._last_pan_point is not None:
                self._pan_to(event.globalPosition().toPoint())
                return True
            if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.MiddleButton:
                self._end_pan()
                return True
            if obj is self.view.viewport() and event.type() == QEvent.Wheel:
                wheel: QWheelEvent = event
                if wheel.modifiers() & Qt.ControlModifier:
                    return False
                angle = wheel.angleDelta().y()
                factor = 1.15 if angle > 0 else 1 / 1.15
                self._apply_zoom(factor)
                return True
        return super().eventFilter(obj, event)

    def _start_pan(self, global_pos: QPoint) -> None:
        self._is_panning = True
        self._last_pan_point = global_pos
        viewport_pos = self.view.viewport().mapFromGlobal(global_pos)
        self._last_pan_scene_pos = self.view.mapToScene(viewport_pos)
        self.view.viewport().setCursor(Qt.ClosedHandCursor)
        self.view.viewport().grabMouse()

    def _pan_to(self, global_pos: QPoint) -> None:
        if self._last_pan_point is None:
            self._last_pan_point = global_pos
            return
        viewport_pos = self.view.viewport().mapFromGlobal(global_pos)
        current_scene_pos = self.view.mapToScene(viewport_pos)
        if self._last_pan_scene_pos is None:
            self._last_pan_scene_pos = current_scene_pos
            return
        delta = current_scene_pos - self._last_pan_scene_pos
        self._last_pan_scene_pos = current_scene_pos
        self.view.centerOn(self.view.mapToScene(self.view.viewport().rect().center()) - delta)

    def _end_pan(self) -> None:
        self._is_panning = False
        self._last_pan_point = None
        self._last_pan_scene_pos = None
        self.view.viewport().releaseMouse()
        self.view.viewport().unsetCursor()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self._center_view()

    def _on_distance_button_clicked(self) -> None:
        current = self._distance_mm if self._distance_mm is not None else 0.0
        value, ok = QInputDialog.getDouble(
            self,
            "Distância entre Stringers",
            "Espaço disponível para rampa de drop-off (mm):",
            current,
            0.0,
            100000.0,
            2,
        )
        if not ok:
            return
        self._distance_mm = float(value)
        if self._distance_button is not None:
            if self._distance_mm.is_integer():
                label = f"{int(self._distance_mm)} mm"
            else:
                label = f"{self._distance_mm:.2f} mm"
            self._distance_button.setText(label)
            self._distance_button.setFixedSize(
                self._distance_button.sizeHint().width(),
                self._distance_button.sizeHint().height(),
            )
        self._update_summary_box()

    def _on_dropoff_ratio_clicked(self) -> None:
        current = self._dropoff_ratio or (1, 20)
        dialog = DropoffRatioDialog(current[0], current[1], parent=self)
        if dialog.exec() != QDialog.Accepted:
            return
        ratio = dialog.selected_ratio()
        if ratio is None:
            return
        self._dropoff_ratio = ratio
        if self._dropoff_ratio_button is not None:
            num, den = ratio
            self._dropoff_ratio_button.setText(f"Razão de Drop Off: {num}/{den}")
            self._dropoff_ratio_button.setFixedSize(
                self._dropoff_ratio_button.sizeHint().width(),
                self._dropoff_ratio_button.sizeHint().height(),
            )
        self._update_summary_box()

    def _on_layer_thickness_clicked(self) -> None:
        current = self._layer_thickness_mm if self._layer_thickness_mm is not None else 0.0
        value, ok = QInputDialog.getDouble(
            self,
            "Espessura da Camada",
            "Informe a espessura da camada (mm):",
            current,
            0.0,
            1000.0,
            3,
        )
        if not ok:
            return
        self._layer_thickness_mm = float(value)
        if self._layer_thickness_button is not None:
            if self._layer_thickness_mm.is_integer():
                label = f"Espessura: {int(self._layer_thickness_mm)} mm"
            else:
                label = f"Espessura: {self._layer_thickness_mm:.3f} mm"
            self._layer_thickness_button.setText(label)
            self._layer_thickness_button.setFixedSize(
                self._layer_thickness_button.sizeHint().width(),
                self._layer_thickness_button.sizeHint().height(),
            )
        self._update_summary_box()

    def _update_summary_box(self) -> None:
        if self._summary_item is None:
            return
        if self._distance_mm is None:
            distance_text = "—"
        else:
            distance_text = (
                f"{int(self._distance_mm)} mm"
                if self._distance_mm.is_integer()
                else f"{self._distance_mm:.2f} mm"
            )
        if self._dropoff_ratio is None:
            ratio_text = "—"
        else:
            ratio_text = f"{self._dropoff_ratio[0]}/{self._dropoff_ratio[1]}"

        if self._layer_thickness_mm is None:
            thickness_text = "—"
        else:
            thickness_text = (
                f"{int(self._layer_thickness_mm)} mm"
                if self._layer_thickness_mm.is_integer()
                else f"{self._layer_thickness_mm:.3f} mm"
            )

        diff_text = "—"
        ramp_text = "—"
        if self._model is not None and self._selected_min_cell and self._selected_max_cell:
            count_a = self._count_oriented_layers_for_cell(self._selected_min_cell)
            count_b = self._count_oriented_layers_for_cell(self._selected_max_cell)
            if count_a is not None and count_b is not None:
                diff_layers = abs(count_a - count_b)
                diff_text = str(diff_layers)
                if (
                    self._dropoff_ratio is not None
                    and diff_layers > 0
                    and self._layer_thickness_mm is not None
                    and self._layer_thickness_mm > 0
                ):
                    num, den = self._dropoff_ratio
                    if num > 0:
                        ramp_length = diff_layers * self._layer_thickness_mm * (den / num)
                        if abs(ramp_length - int(ramp_length)) < 1e-6:
                            ramp_text = f"{int(ramp_length)} mm"
                        else:
                            ramp_text = f"{ramp_length:.2f} mm"

        summary_html = (
            "<div style='font-family:Segoe UI; font-size:11pt; color:#222;'>"
            "<div style='font-weight:600; margin-bottom:6px;'>Resumo</div>"
            "<table style='border-collapse:collapse;'>"
            f"<tr><td>Espaço para Drop Off</td><td style='padding-left:16px;'>{distance_text}</td></tr>"
            f"<tr><td>Razão de Drop Off</td><td style='padding-left:16px;'>{ratio_text}</td></tr>"
            f"<tr><td>Espessura da Camada</td><td style='padding-left:16px;'>{thickness_text}</td></tr>"
            f"<tr><td>Diferença de Camadas Entre Células</td><td style='padding-left:16px;'>{diff_text}</td></tr>"
            f"<tr><td>Comprimento da Rampa de Drop Off</td><td style='padding-left:16px;'>{ramp_text}</td></tr>"
            "</table>"
            "<div style='margin:8px 0; border-bottom:1px solid #cfcfcf;'></div>"
            "<div style='color:#555;'>Nesse espaço iremos adicionar o resultado<br>da análise.</div>"
            "</div>"
        )
        self._summary_item.setHtml(summary_html)

    def _count_oriented_layers_for_cell(self, cell_id: str) -> Optional[int]:
        if self._model is None:
            return None
        laminate_name = self._model.cell_to_laminate.get(cell_id)
        if not laminate_name:
            return None
        laminate = self._model.laminados.get(laminate_name)
        if laminate is None:
            return None
        try:
            return int(count_oriented_layers(laminate.camadas))
        except Exception:
            return None


class DropoffRatioDialog(QDialog):
    """Dialog to capture drop-off ratio (e.g., 1/20)."""

    def __init__(self, numerator: int = 1, denominator: int = 20, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Razão de Drop Off")
        self.resize(320, 160)

        layout = QVBoxLayout(self)

        helper = QLabel("Informe a razão de drop-off (ex.: 1/20):", self)
        layout.addWidget(helper)

        row = QHBoxLayout()
        self.numerator_spin = QSpinBox(self)
        self.numerator_spin.setRange(1, 20)
        self.numerator_spin.setValue(max(1, int(numerator)))
        row.addWidget(self.numerator_spin)

        slash = QLabel("/", self)
        row.addWidget(slash)

        self.denominator_spin = QSpinBox(self)
        self.denominator_spin.setRange(1, 500)
        self.denominator_spin.setValue(max(1, int(denominator)))
        row.addWidget(self.denominator_spin)
        row.addStretch(1)
        layout.addLayout(row)

        preview = QLabel(self)
        preview.setText(self._preview_text())
        layout.addWidget(preview)

        def _update_preview() -> None:
            preview.setText(self._preview_text())

        self.numerator_spin.valueChanged.connect(_update_preview)
        self.denominator_spin.valueChanged.connect(_update_preview)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _preview_text(self) -> str:
        return f"Razão: {self.numerator_spin.value()}/{self.denominator_spin.value()}"

    def selected_ratio(self) -> Optional[tuple[int, int]]:
        return (self.numerator_spin.value(), self.denominator_spin.value())
