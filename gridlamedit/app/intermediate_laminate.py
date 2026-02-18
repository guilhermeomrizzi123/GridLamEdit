"""Intermediate laminate suggestion window."""

from __future__ import annotations

from typing import Optional
import math
import copy
import re

from PySide6.QtCore import Qt, QSize, QEvent, QPoint, QRectF
from PySide6.QtGui import QAction, QColor, QPainter, QPen, QBrush, QFont, QWheelEvent, QTextDocument
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
    QFormLayout,
    QSpinBox,
    QStyle,
    QToolBar,
    QVBoxLayout,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QAbstractItemView,
    QStyledItemDelegate,
    QComboBox,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QStyleOptionHeader,
)

from gridlamedit.io.spreadsheet import (
    GridModel,
    Camada,
    Laminado,
    count_oriented_layers,
    normalize_angle,
    format_orientation_value,
    orientation_highlight_color,
)
from gridlamedit.services.laminate_checks import (
    evaluate_symmetry_for_layers,
    evaluate_laminate_balance_clt,
)
from gridlamedit.services.laminate_service import auto_name_for_laminate


class IntermediateLaminateWindow(QDialog):
    """Dialog to suggest an intermediate laminate between two cells."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Novo Laminado Intermediário")
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
        self._create_intermediate_button: Optional[QPushButton] = None
        self._cell_button_proxies: list[QGraphicsProxyWidget] = []
        self._summary_item: Optional[QGraphicsTextItem] = None
        self._summary_rect = None
        self._summary_min_height = 180.0
        self._interface_rects: list[QRectF] = []
        self._intermediate_preview_button: Optional[QPushButton] = None
        self._intermediate_preview_proxy: Optional[QGraphicsProxyWidget] = None
        self._created_intermediate_laminate: Optional[Laminado] = None
        self._created_min_laminate: Optional[Laminado] = None
        self._created_max_laminate: Optional[Laminado] = None
        self._created_min_cell_id: Optional[str] = None
        self._created_max_cell_id: Optional[str] = None
        self._min_cell_label = "SELECIONAR LAMINADO COM MENOR ESPESSURA"
        self._max_cell_label = "SELECIONAR LAMINADO COM MAIOR ESPESSURA"
        self._distance_label = "?(mm)"
        self._reduce_layers_needed: Optional[int] = None
        self._analysis_requires_reduction: bool = False

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
        create_button = self._build_cell_select_button("Criar Laminado Intermediário")
        create_button.setMinimumWidth(240)
        create_button.setEnabled(False)
        top_bar.addWidget(create_button)
        top_bar.addStretch(1)
        clear_button = self._build_cell_select_button("Limpar")
        clear_button.setMinimumWidth(120)
        top_bar.addWidget(clear_button)
        main_layout.addLayout(top_bar)
        self._dropoff_ratio_button = dropoff_button
        dropoff_button.setStyleSheet("")
        dropoff_button.clicked.connect(self._on_dropoff_ratio_clicked)
        self._layer_thickness_button = thickness_button
        thickness_button.setStyleSheet("")
        thickness_button.clicked.connect(self._on_layer_thickness_clicked)
        self._create_intermediate_button = create_button
        create_button.setStyleSheet("")
        create_button.clicked.connect(self._on_create_intermediate_clicked)
        clear_button.setStyleSheet("")
        clear_button.clicked.connect(self._on_clear_clicked)

        self.view = QGraphicsView(self)
        self.view.setRenderHint(QPainter.Antialiasing, True)
        self.view.setBackgroundBrush(QColor(255, 255, 255))
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
        self._interface_rects.clear()
        self._intermediate_preview_button = None
        self._intermediate_preview_proxy = None

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
            self._interface_rects.append(rect.rect())
            label = QGraphicsTextItem(f"INTERFACE {idx}")
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
            (QColor(153, 213, 92), self._min_cell_label, green_height, "min"),
            (QColor(230, 230, 230), "", orange_height, "label"),
            (QColor(220, 220, 220), self._max_cell_label, blue_height, "max"),
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
                button_font = button.font()
                button_font.setPointSize(max(12, button_font.pointSize() + 2))
                button.setFont(button_font)
                proxy = QGraphicsProxyWidget()
                proxy.setWidget(button)
                min_button_width = 360
                min_button_height = 38
                button_width = max(button.sizeHint().width(), min_button_width)
                button_height = max(button.sizeHint().height(), min_button_height)
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
                if text:
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
        distance_button = self._build_cell_select_button(self._distance_label)
        distance_button.setStyleSheet(
            distance_button.styleSheet()
            + "QPushButton { text-align: center; padding: 4px 10px; background-color: #ffffff; }"
        )
        distance_button.setMinimumWidth(140)
        distance_button.setMinimumHeight(34)
        distance_proxy = QGraphicsProxyWidget()
        distance_proxy.setWidget(distance_button)
        distance_width = max(distance_button.sizeHint().width(), 140)
        distance_height = max(distance_button.sizeHint().height(), 34)
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
        summary_height = self._summary_min_height
        summary_rect = self.scene.addRect(
            summary_x,
            summary_y,
            summary_width,
            summary_height,
            QPen(QColor(160, 160, 160), 1.2),
            QBrush(QColor(255, 255, 255)),
        )
        summary_rect.setZValue(1)
        self._summary_rect = summary_rect

        summary_html = (
            "<div style='font-family:Segoe UI; font-size:11pt; color:#1f2937;'>"
            "<div style='font-size:12pt; font-weight:700; letter-spacing:0.4px; margin-bottom:6px;'>RESUMO</div>"
            "<table style='border-collapse:collapse;'>"
            "<tr><td>Espaço para Drop Off</td><td style='padding-left:16px;'>—</td></tr>"
            "<tr><td>Razão de Drop Off</td><td style='padding-left:16px;'>—</td></tr>"
            "<tr><td>Diferença de Camadas Entre Laminados</td><td style='padding-left:16px;'>—</td></tr>"
            "</table>"
            "<div style='margin:8px 0; border-bottom:1px solid #d1d5db;'></div>"
            "<div style='font-weight:600; margin-bottom:4px;'>Conclusão:</div>"
            "<div style='color:#475569;'>---</div>"
            "</div>"
        )
        summary_item = QGraphicsTextItem()
        summary_item.setHtml(summary_html)
        summary_item.setDefaultTextColor(QColor(20, 20, 20))
        summary_item.setTextWidth(summary_width - 24.0)
        summary_item.setPos(summary_x + 12.0, summary_y + 10.0)
        summary_item.setZValue(2)
        self.scene.addItem(summary_item)
        self._summary_item = summary_item
        self._update_summary_box()
        self._ensure_intermediate_preview_button()
        self._center_view()

    def _ensure_intermediate_preview_button(self) -> None:
        if len(self._interface_rects) < 2:
            return
        target_rect = self._interface_rects[1]
        if self._intermediate_preview_button is None:
            button = self._build_cell_select_button("")
            button_font = button.font()
            button_font.setPointSize(max(12, button_font.pointSize() + 2))
            button.setFont(button_font)
            button.clicked.connect(self._on_intermediate_preview_clicked)
            proxy = QGraphicsProxyWidget()
            proxy.setWidget(button)
            proxy.setZValue(6)
            self.scene.addItem(proxy)
            self._intermediate_preview_button = button
            self._intermediate_preview_proxy = proxy
        else:
            button = self._intermediate_preview_button
            proxy = self._intermediate_preview_proxy
        if button is None or proxy is None:
            return
        created = self._created_intermediate_laminate
        if created is None:
            button.hide()
            return
        button.setText(created.nome)
        min_button_width = 360
        min_button_height = 38
        width = max(button.sizeHint().width(), min_button_width)
        height = max(button.sizeHint().height(), min_button_height)
        button.setFixedSize(width, height)
        center_x = target_rect.x() + target_rect.width() / 2.0
        center_y = target_rect.y() + target_rect.height() / 2.0
        proxy.setPos(center_x - width / 2.0, center_y - height / 2.0)
        button.show()

    def _on_intermediate_preview_clicked(self) -> None:
        if (
            self._created_intermediate_laminate is None
            or self._created_min_laminate is None
            or self._created_max_laminate is None
            or self._created_min_cell_id is None
            or self._created_max_cell_id is None
        ):
            return
        dialog = IntermediateLaminatePreviewDialog(
            min_laminate=self._created_min_laminate,
            intermediate_laminate=self._created_intermediate_laminate,
            max_laminate=self._created_max_laminate,
            min_cell_id=self._created_min_cell_id,
            max_cell_id=self._created_max_cell_id,
            model=self._model,
            parent=self,
        )
        dialog.exec()

    def _notify_intermediate_created(
        self,
        laminate: Laminado,
        min_laminate: Laminado,
        max_laminate: Laminado,
        min_cell_id: str,
        max_cell_id: str,
    ) -> None:
        self._created_intermediate_laminate = laminate
        self._created_min_laminate = min_laminate
        self._created_max_laminate = max_laminate
        self._created_min_cell_id = min_cell_id
        self._created_max_cell_id = max_cell_id
        self._ensure_intermediate_preview_button()

    def _build_cell_select_button(self, text: str) -> QPushButton:
        button = QPushButton(text)
        button.installEventFilter(self)
        return button

    def _select_cell(self, mode: str) -> None:
        laminates = self._available_laminates()
        if not laminates:
            QMessageBox.information(
                self,
                "Selecionar Laminado",
                "Carregue um projeto com laminados para selecionar.",
            )
            return
        current = self._selected_min_cell if mode == "min" else self._selected_max_cell
        dialog = LaminadoSelectDialog(laminates, current=current, parent=self)
        if dialog.exec() != QDialog.Accepted:
            return
        selected = dialog.selected_laminate()
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

    def _available_laminates(self) -> list[str]:
        if self._model is None:
            return []
        names = [name for name in self._model.laminados.keys() if str(name).strip()]
        return sorted(names, key=self._laminate_sort_key)

    @staticmethod
    def _laminate_sort_key(name: str) -> tuple[str, float, str]:
        text = str(name or "").strip()
        match = re.match(r"^([A-Za-z]+)?\s*([0-9]+(?:\.[0-9]+)?)?", text)
        prefix = match.group(1) if match else ""
        number_text = match.group(2) if match else ""
        try:
            number = float(number_text) if number_text else float("inf")
        except Exception:
            number = float("inf")
        remainder = text[match.end():].strip() if match else text
        return (prefix.upper(), number, remainder)

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
        diff_layers_value: Optional[int] = None
        ramp_length_value: Optional[float] = None
        step_length_value: Optional[float] = None
        if self._model is not None and self._selected_min_cell and self._selected_max_cell:
            count_a = self._count_oriented_layers_for_laminate(self._selected_min_cell)
            count_b = self._count_oriented_layers_for_laminate(self._selected_max_cell)
            if count_a is not None and count_b is not None:
                diff_layers = abs(count_a - count_b)
                diff_layers_value = diff_layers
                diff_text = str(diff_layers)
                if (
                    self._dropoff_ratio is not None
                    and diff_layers > 0
                    and self._layer_thickness_mm is not None
                    and self._layer_thickness_mm > 0
                ):
                    num, den = self._dropoff_ratio
                    if num > 0:
                        step_length_value = self._layer_thickness_mm * (den / num)
                        ramp_length = diff_layers * step_length_value
                        ramp_length_value = ramp_length
                        if abs(ramp_length - int(ramp_length)) < 1e-6:
                            ramp_text = f"{int(ramp_length)} mm"
                        else:
                            ramp_text = f"{ramp_length:.2f} mm"

        analysis_html = "---"
        self._analysis_requires_reduction = False
        self._reduce_layers_needed = None
        center_note = ""
        if (
            self._distance_mm is not None
            and ramp_length_value is not None
            and diff_layers_value is not None
            and step_length_value is not None
            and step_length_value > 0
        ):
            if self._distance_mm < ramp_length_value:
                max_layers = int(self._distance_mm / step_length_value)
                reduce_by = max(0, diff_layers_value - max_layers)
                if reduce_by % 2 == 1:
                    min_laminate = self._laminate_for_cell(self._selected_min_cell)
                    max_laminate = self._laminate_for_cell(self._selected_max_cell)
                    if min_laminate is not None and max_laminate is not None:
                        if not self._has_center_excess(max_laminate, min_laminate):
                            reduce_by += 1
                            center_note = (
                                "<br>Para manter simetria e balanceamento, "
                                "o número de camadas removidas foi ajustado para um número par."
                            )
                if reduce_by == 1:
                    reduce_by = 2
                self._analysis_requires_reduction = reduce_by > 0
                self._reduce_layers_needed = reduce_by
                if reduce_by > 0 and diff_layers_value < 2:
                    self._analysis_requires_reduction = False
                    analysis_html = (
                        "O espaço para drop off não é suficiente para o "
                        "comprimento da rampa de drop off.<br>"
                        "A diferença entre os laminados é menor que 2 camadas, "
                        "portanto não é possível propor a remoção mínima de duas camadas."
                    )
                else:
                    analysis_html = (
                        "O espaço para drop off não é suficiente para o "
                        "comprimento da rampa de drop off.<br>"
                        f"Reduza a diferença de camadas entre células em {reduce_by} "
                        "para que a rampa caiba no espaço disponível."
                        f"{center_note}"
                        "<br>Use o comando de criar laminado intermediário como opção de solução."
                    )
            else:
                analysis_html = "O espaço para drop off é suficiente para a rampa de drop off."

        self._update_create_button_state()

        summary_html = (
            "<div style='font-family:Segoe UI; font-size:11pt; color:#1f2937;'>"
            "<div style='font-size:12pt; font-weight:700; letter-spacing:0.4px; margin-bottom:6px;'>RESUMO</div>"
            "<table style='border-collapse:collapse;'>"
            f"<tr><td>Espaço para Drop Off</td><td style='padding-left:16px;'>{distance_text}</td></tr>"
            f"<tr><td>Razão de Drop Off</td><td style='padding-left:16px;'>{ratio_text}</td></tr>"
            f"<tr><td>Espessura da Camada</td><td style='padding-left:16px;'>{thickness_text}</td></tr>"
            f"<tr><td>Diferença de Camadas Entre Laminados</td><td style='padding-left:16px;'>{diff_text}</td></tr>"
            f"<tr><td>Comprimento da Rampa de Drop Off</td><td style='padding-left:16px;'>{ramp_text}</td></tr>"
            "</table>"
            "<div style='margin:8px 0; border-bottom:1px solid #d1d5db;'></div>"
            "<div style='font-weight:600; margin-bottom:4px;'>Conclusão:</div>"
            f"<div style='color:#475569;'>{analysis_html}</div>"
            "</div>"
        )
        self._summary_item.setHtml(summary_html)
        if self._summary_rect is not None:
            rect = self._summary_rect.rect()
            self._summary_item.setTextWidth(rect.width() - 24.0)
            text_height = self._summary_item.boundingRect().height() + 20.0
            new_height = max(self._summary_min_height, text_height)
            if abs(new_height - rect.height()) > 0.5:
                self._summary_rect.setRect(rect.x(), rect.y(), rect.width(), new_height)

    def _on_clear_clicked(self) -> None:
        self._distance_mm = None
        self._dropoff_ratio = None
        self._layer_thickness_mm = None
        self._selected_min_cell = None
        self._selected_max_cell = None
        self._created_intermediate_laminate = None
        self._created_min_laminate = None
        self._created_max_laminate = None
        self._created_min_cell_id = None
        self._created_max_cell_id = None
        if self._distance_button is not None:
            self._distance_button.setText(self._distance_label)
            self._distance_button.setFixedSize(
                self._distance_button.sizeHint().width(),
                self._distance_button.sizeHint().height(),
            )
        if self._dropoff_ratio_button is not None:
            self._dropoff_ratio_button.setText("Razão de Drop Off")
            self._dropoff_ratio_button.setFixedSize(
                self._dropoff_ratio_button.sizeHint().width(),
                self._dropoff_ratio_button.sizeHint().height(),
            )
        if self._layer_thickness_button is not None:
            self._layer_thickness_button.setText("Espessura da Camada (mm)")
            self._layer_thickness_button.setFixedSize(
                self._layer_thickness_button.sizeHint().width(),
                self._layer_thickness_button.sizeHint().height(),
            )
        if self._min_cell_button is not None:
            self._min_cell_button.setText(self._min_cell_label)
        if self._max_cell_button is not None:
            self._max_cell_button.setText(self._max_cell_label)
        if self._create_intermediate_button is not None:
            self._create_intermediate_button.setEnabled(False)
        self._update_summary_box()
        self._ensure_intermediate_preview_button()

    def _count_oriented_layers_for_laminate(self, laminate_name: str) -> Optional[int]:
        if self._model is None:
            return None
        laminate = self._model.laminados.get(laminate_name)
        if laminate is None:
            return None
        try:
            return int(count_oriented_layers(laminate.camadas))
        except Exception:
            return None

    def _update_create_button_state(self) -> None:
        if self._create_intermediate_button is None:
            return
        enabled = (
            self._model is not None
            and self._selected_min_cell is not None
            and self._selected_max_cell is not None
            and self._distance_mm is not None
            and self._dropoff_ratio is not None
            and self._layer_thickness_mm is not None
            and self._analysis_requires_reduction
            and (self._reduce_layers_needed or 0) > 0
        )
        self._create_intermediate_button.setEnabled(enabled)

    def _on_create_intermediate_clicked(self) -> None:
        if self._model is None:
            QMessageBox.information(self, "Criar Laminado Intermediário", "Nenhum projeto carregado.")
            return
        if not self._selected_min_cell or not self._selected_max_cell:
            QMessageBox.information(
                self,
                "Criar Laminado Intermediário",
                "Selecione os laminados de menor e maior espessura.",
            )
            return
        if not self._analysis_requires_reduction or not self._reduce_layers_needed:
            QMessageBox.information(
                self,
                "Criar Laminado Intermediário",
                "A análise não indica necessidade de redução de camadas.",
            )
            return

        min_laminate = self._laminate_for_cell(self._selected_min_cell)
        max_laminate = self._laminate_for_cell(self._selected_max_cell)
        if min_laminate is None or max_laminate is None:
            QMessageBox.warning(
                self,
                "Criar Laminado Intermediário",
                "Não foi possível localizar os laminados selecionados.",
            )
            return

        min_count = count_oriented_layers(min_laminate.camadas)
        max_count = count_oriented_layers(max_laminate.camadas)
        if max_count <= min_count:
            QMessageBox.information(
                self,
                "Criar Laminado Intermediário",
                "O laminado de maior espessura deve ter mais camadas orientadas do que o laminado de menor espessura.",
            )
            return

        allow_non_45 = True
        if not self._has_excess_45_layers(max_laminate, min_laminate):
            reply = QMessageBox.question(
                self,
                "Criar Laminado Intermediário",
                "Não existem camadas excedentes em 45° ou -45° para remoção. "
                "Deseja permitir remoção em outras orientações?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return
            allow_non_45 = True

        new_laminate, error = self._build_intermediate_laminate(
            max_laminate,
            min_laminate,
            self._reduce_layers_needed,
            allow_non_45,
        )
        if new_laminate is None:
            QMessageBox.warning(
                self,
                "Criar Laminado Intermediário",
                error or "Não foi possível gerar um laminado intermediário válido.",
            )
            return

        dialog = IntermediateLaminatePreviewDialog(
            min_laminate=min_laminate,
            intermediate_laminate=new_laminate,
            max_laminate=max_laminate,
            min_cell_id=self._selected_min_cell,
            max_cell_id=self._selected_max_cell,
            model=self._model,
            parent=self,
        )
        dialog.exec()

    def _laminate_for_cell(self, cell_id: str) -> Optional[Laminado]:
        if self._model is None:
            return None
        return self._model.laminados.get(cell_id)

    def _has_center_excess(self, max_laminate: Laminado, min_laminate: Laminado) -> bool:
        max_layers = list(getattr(max_laminate, "camadas", []) or [])
        min_layers = list(getattr(min_laminate, "camadas", []) or [])
        if len(max_layers) % 2 == 0:
            return False
        center_idx = len(max_layers) // 2
        if not (0 <= center_idx < len(max_layers)):
            return False

        def _orientation_token(layer: Camada) -> Optional[float]:
            try:
                value = getattr(layer, "orientacao", None)
            except Exception:
                return None
            if value is None:
                return None
            try:
                return normalize_angle(value)
            except Exception:
                return None

        max_token = _orientation_token(max_layers[center_idx])
        min_token = None
        if 0 <= center_idx < len(min_layers):
            min_token = _orientation_token(min_layers[center_idx])
        if max_token is None:
            return False
        return max_token != min_token

    def _has_excess_45_layers(self, max_laminate: Laminado, min_laminate: Laminado) -> bool:
        def _orientation_token(layer: Camada) -> Optional[float]:
            try:
                value = getattr(layer, "orientacao", None)
            except Exception:
                return None
            if value is None:
                return None
            try:
                return normalize_angle(value)
            except Exception:
                return None

        def _count_tokens(layers_list: list[Camada]) -> dict[float, int]:
            counts: dict[float, int] = {}
            for layer in layers_list:
                token = _orientation_token(layer)
                if token is None:
                    continue
                counts[token] = counts.get(token, 0) + 1
            return counts

        max_counts = _count_tokens(list(getattr(max_laminate, "camadas", []) or []))
        min_counts = _count_tokens(list(getattr(min_laminate, "camadas", []) or []))
        for angle in (45.0, -45.0):
            excess = max_counts.get(angle, 0) - min_counts.get(angle, 0)
            if excess > 0:
                return True
        return False

    def _build_intermediate_laminate(
        self,
        max_laminate: Laminado,
        min_laminate: Laminado,
        reduce_by: int,
        allow_non_45: bool,
    ) -> tuple[Optional[Laminado], Optional[str]]:
        if reduce_by <= 0:
            return None, "Nenhuma redução de camadas é necessária."
        layers = list(getattr(max_laminate, "camadas", []) or [])
        if not layers:
            return None, "O laminado de maior espessura não possui camadas."

        def _orientation_token(layer: Camada) -> Optional[float]:
            try:
                value = getattr(layer, "orientacao", None)
            except Exception:
                return None
            if value is None:
                return None
            try:
                return normalize_angle(value)
            except Exception:
                return None

        def _count_tokens(layers_list: list[Camada]) -> dict[float, int]:
            counts: dict[float, int] = {}
            for layer in layers_list:
                token = _orientation_token(layer)
                if token is None:
                    continue
                counts[token] = counts.get(token, 0) + 1
            return counts

        max_counts = _count_tokens(list(getattr(max_laminate, "camadas", []) or []))
        min_counts = _count_tokens(list(getattr(min_laminate, "camadas", []) or []))
        excess_counts: dict[float, int] = {}
        for token, count in max_counts.items():
            excess = count - min_counts.get(token, 0)
            if excess > 0:
                excess_counts[token] = excess

        total_excess = sum(excess_counts.values())
        if reduce_by > total_excess:
            return (
                None,
                "Não existem camadas suficientes com orientações excedentes para atender a redução necessária.",
            )

        preferred_pairs: list[tuple[int, int, float, float, int]] = []
        other_pairs: list[tuple[int, int, float, float, int]] = []
        min_layers = list(getattr(min_laminate, "camadas", []) or [])

        def _row_priority(idx: int, jdx: int, token_a: float, token_b: float) -> int:
            score = 0
            min_token_a = None
            min_token_b = None
            if 0 <= idx < len(min_layers):
                min_token_a = _orientation_token(min_layers[idx])
            if 0 <= jdx < len(min_layers):
                min_token_b = _orientation_token(min_layers[jdx])
            if min_token_a != token_a:
                score += 1
            if min_token_b != token_b:
                score += 1
            return score
        for idx in range(len(layers) // 2):
            jdx = len(layers) - 1 - idx
            token_a = _orientation_token(layers[idx])
            token_b = _orientation_token(layers[jdx])
            if token_a is None or token_b is None:
                continue
            if token_a not in excess_counts or token_b not in excess_counts:
                continue
            is_preferred = (
                math.isclose(token_a, 45.0, abs_tol=1e-6)
                or math.isclose(token_a, -45.0, abs_tol=1e-6)
            ) and (
                math.isclose(token_b, 45.0, abs_tol=1e-6)
                or math.isclose(token_b, -45.0, abs_tol=1e-6)
            )
            priority = _row_priority(idx, jdx, token_a, token_b)
            if is_preferred:
                preferred_pairs.append((idx, jdx, token_a, token_b, priority))
            else:
                other_pairs.append((idx, jdx, token_a, token_b, priority))

        preferred_pairs.sort(key=lambda item: item[4], reverse=True)
        other_pairs.sort(key=lambda item: item[4], reverse=True)


        center_idx: Optional[int] = None
        center_token: Optional[float] = None
        if len(layers) % 2 == 1:
            center_idx = len(layers) // 2
            center_token = _orientation_token(layers[center_idx])
            if center_token not in excess_counts:
                center_idx = None
                center_token = None

        use_center = reduce_by % 2 == 1
        if use_center and center_idx is None:
            return (
                None,
                "Não foi possível remover uma camada central para manter a simetria requerida.",
            )

        remaining_pairs = (reduce_by - (1 if use_center else 0)) // 2
        if remaining_pairs < 0:
            return None, "Número de camadas inválido para redução."

        def _apply_removal(indices: set[int]) -> list[Camada]:
            new_layers: list[Camada] = []
            for idx, layer in enumerate(layers):
                new_layer = copy.deepcopy(layer)
                if idx in indices:
                    new_layer.orientacao = None
                    new_layer.material = ""
                    new_layer.ativo = True
                new_layers.append(new_layer)
            for idx, layer in enumerate(new_layers):
                layer.idx = idx
            return new_layers

        def _search(
            pairs_list: list[tuple[int, int, float, float, int]],
            pair_idx: int,
            remaining: int,
            available: dict[float, int],
            selected_pairs: list[tuple[int, int, float, float, int]],
        ) -> Optional[set[int]]:
            if remaining == 0:
                indices: set[int] = set()
                if use_center and center_idx is not None:
                    indices.add(center_idx)
                for i, j, _, _, _ in selected_pairs:
                    indices.update([i, j])
                new_layers = _apply_removal(indices)
                if not evaluate_symmetry_for_layers(new_layers).is_symmetric:
                    return None
                if not evaluate_laminate_balance_clt(new_layers).is_balanced:
                    return None
                return indices

            if pair_idx >= len(pairs_list):
                return None
            if remaining > (len(pairs_list) - pair_idx):
                return None

            idx, jdx, token_a, token_b, _priority = pairs_list[pair_idx]
            can_include = False
            if token_a == token_b:
                can_include = available.get(token_a, 0) >= 2
            else:
                can_include = available.get(token_a, 0) >= 1 and available.get(token_b, 0) >= 1

            if can_include:
                if token_a == token_b:
                    available[token_a] -= 2
                else:
                    available[token_a] -= 1
                    available[token_b] -= 1
                selected_pairs.append((idx, jdx, token_a, token_b, _priority))
                found = _search(pairs_list, pair_idx + 1, remaining - 1, available, selected_pairs)
                if found is not None:
                    return found
                selected_pairs.pop()
                if token_a == token_b:
                    available[token_a] += 2
                else:
                    available[token_a] += 1
                    available[token_b] += 1

            return _search(pairs_list, pair_idx + 1, remaining, available, selected_pairs)

        def _try_with_pairs(pairs_list: list[tuple[int, int, float, float, int]]) -> Optional[set[int]]:
            if remaining_pairs > 0 and not pairs_list:
                return None
            available_counts = dict(excess_counts)
            if use_center and center_token is not None:
                if available_counts.get(center_token, 0) < 1:
                    return None
                available_counts[center_token] -= 1
            return _search(pairs_list, 0, remaining_pairs, available_counts, [])

        indices = _try_with_pairs(preferred_pairs)
        if indices is None and allow_non_45:
            indices = _try_with_pairs(preferred_pairs + other_pairs)
        if indices is None:
            return (
                None,
                "Não foi possível construir um laminado intermediário simétrico e balanceado com as restrições fornecidas.",
            )

        new_layers = _apply_removal(indices)
        new_name = f"{str(getattr(max_laminate, 'nome', '') or '').strip()} - Intermediário"
        new_name = new_name.strip(" -") or "Intermediário"
        new_laminate = Laminado(
            nome=new_name,
            tipo=getattr(max_laminate, "tipo", ""),
            color_index=getattr(max_laminate, "color_index", 1),
            tag=getattr(max_laminate, "tag", ""),
            celulas=[],
            camadas=new_layers,
            auto_rename_enabled=False,
        )
        return new_laminate, None


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


class LaminadoSelectDialog(QDialog):
    def __init__(self, laminates: list[str], current: Optional[str] = None, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Selecionar Laminado")
        self.resize(420, 520)
        self._all_laminates = laminates
        self._selected: str = ""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        self.search_input = QLineEdit(self)
        self.search_input.setPlaceholderText("Buscar laminado...")
        layout.addWidget(self.search_input)

        self.list_widget = QListWidget(self)
        layout.addWidget(self.list_widget, stretch=1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.search_input.textChanged.connect(self._apply_filter)
        self.list_widget.itemDoubleClicked.connect(lambda *_: self._on_accept())

        self._apply_filter("")
        if current:
            self._select_current(current)

    def _apply_filter(self, text: str) -> None:
        needle = text.strip().lower()
        self.list_widget.clear()
        for name in self._all_laminates:
            if needle and needle not in str(name).lower():
                continue
            self.list_widget.addItem(QListWidgetItem(str(name)))
        if self.list_widget.count() > 0:
            self.list_widget.setCurrentRow(0)

    def _select_current(self, name: str) -> None:
        for idx in range(self.list_widget.count()):
            item = self.list_widget.item(idx)
            if item and item.text() == name:
                self.list_widget.setCurrentRow(idx)
                break

    def _on_accept(self) -> None:
        item = self.list_widget.currentItem()
        self._selected = item.text() if item else ""
        if not self._selected:
            QMessageBox.information(self, "Selecionar Laminado", "Selecione um laminado.")
            return
        self.accept()

    def selected_laminate(self) -> str:
        return self._selected


class CreateIntermediateLaminateDialog(QDialog):
    def __init__(
        self,
        model: Optional[GridModel],
        base_tag: str,
        laminate_template: Laminado,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Criar Laminado")
        self.resize(420, 200)
        self._model = model
        self._laminate_template = laminate_template

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)

        self.name_edit = QLineEdit(self)
        self.name_edit.setReadOnly(True)
        self.name_edit.setPlaceholderText("Nome automático")

        self.tag_edit = QLineEdit(self)
        self.tag_edit.setPlaceholderText("Tag do novo laminado")

        form.addRow("Nome:", self.name_edit)
        form.addRow("Tag:", self.tag_edit)

        layout.addLayout(form)

        buttons = QDialogButtonBox(parent=self)
        self.create_button = QPushButton("Criar", self)
        buttons.addButton(self.create_button, QDialogButtonBox.AcceptRole)
        buttons.addButton(QDialogButtonBox.Cancel)
        buttons.rejected.connect(self.reject)
        self.create_button.clicked.connect(self.accept)
        layout.addWidget(buttons)

        suggestion = self._suggest_tag(base_tag)
        if suggestion:
            self.tag_edit.setText(suggestion)
        self.tag_edit.textChanged.connect(self._update_name_preview)
        self._update_name_preview()

    def _suggest_tag(self, base_tag: str) -> str:
        base = str(base_tag or "").strip()
        if not base:
            return ""
        existing = set()
        if self._model is not None:
            for lam in self._model.laminados.values():
                tag = str(getattr(lam, "tag", "") or "").strip()
                if tag:
                    existing.add(tag)
        idx = 1
        while True:
            candidate = f"{base}.{idx}"
            if candidate not in existing:
                return candidate
            idx += 1

    def _update_name_preview(self) -> None:
        temp = copy.deepcopy(self._laminate_template)
        temp.tag = self.tag_edit.text().strip()
        name = auto_name_for_laminate(self._model, temp)
        display_name = re.sub(r"\s*\([^)]*\)\s*$", "", name or "").strip()
        self.name_edit.setText(display_name)

    def selected_tag(self) -> str:
        return self.tag_edit.text().strip()

    def selected_name(self) -> str:
        return self.name_edit.text().strip()


class RichTextHeaderView(QHeaderView):
    """HeaderView que suporta HTML/Rich Text por seção."""

    def __init__(self, orientation: Qt.Orientation, parent=None) -> None:
        super().__init__(orientation, parent)
        self._section_html: dict[int, str] = {}
        self.setDefaultAlignment(Qt.AlignCenter)

    def set_section_html(self, logical_index: int, html: str) -> None:
        self._section_html[logical_index] = html
        self.updateSection(logical_index)

    def paintSection(self, painter: QPainter, rect, logicalIndex: int) -> None:
        if not rect.isValid():
            return
        painter.save()
        html = self._section_html.get(logicalIndex)
        option = QStyleOptionHeader()
        self.initStyleOption(option)
        option.rect = rect
        option.section = logicalIndex
        if html:
            option.text = ""
        else:
            option.text = str(
                self.model().headerData(logicalIndex, self.orientation(), Qt.DisplayRole)
                or ""
            )
        self.style().drawControl(QStyle.CE_Header, option, painter, self)

        if html:
            doc = QTextDocument()
            doc.setDefaultFont(self.font())
            doc.setHtml(html)
            doc.setTextWidth(rect.width() - 8.0)
            text_height = doc.size().height()
            y = rect.y() + max(0.0, (rect.height() - text_height) / 2.0)
            painter.translate(rect.x() + 4.0, y)
            doc.drawContents(painter)

        painter.restore()


class IntermediateLaminatePreviewDialog(QDialog):
    """Dialog to preview min/intermediate/max laminates."""

    def __init__(
        self,
        *,
        min_laminate: Laminado,
        intermediate_laminate: Laminado,
        max_laminate: Laminado,
        min_cell_id: str,
        max_cell_id: str,
        model: Optional[GridModel] = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Laminado Intermediário")
        self.resize(980, 640)
        self._model = model
        self._min_laminate = min_laminate
        self._max_laminate = max_laminate
        self._min_cell_id = min_cell_id
        self._max_cell_id = max_cell_id
        self._template_laminate = intermediate_laminate

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        table = QTableWidget(self)
        table.setColumnCount(4)
        header = RichTextHeaderView(Qt.Horizontal, table)
        header_font = QFont("Segoe UI", 9)
        header_font.setBold(False)
        header.setFont(header_font)
        table.setHorizontalHeader(header)

        min_layers = list(getattr(min_laminate, "camadas", []) or [])
        mid_layers = list(getattr(intermediate_laminate, "camadas", []) or [])
        max_layers = list(getattr(max_laminate, "camadas", []) or [])
        self._mid_layers = mid_layers
        row_count = max(len(min_layers), len(mid_layers), len(max_layers))
        table.setRowCount(row_count)

        def _orientation_counts(layers: list[Camada]) -> tuple[dict[str, int], int]:
            counts: dict[str, int] = {"0": 0, "+45": 0, "-45": 0, "90": 0, "other": 0}
            total = 0
            for layer in layers:
                value = getattr(layer, "orientacao", None)
                if value is None:
                    continue
                try:
                    angle = normalize_angle(value)
                except Exception:
                    continue
                total += 1
                if abs(angle) < 1e-6:
                    counts["0"] += 1
                elif abs(angle - 90.0) < 1e-6:
                    counts["90"] += 1
                elif abs(angle - 45.0) < 1e-6:
                    counts["+45"] += 1
                elif abs(angle + 45.0) < 1e-6:
                    counts["-45"] += 1
                else:
                    counts["other"] += 1
            return counts, total

        def _classify_laminate_type(counts: dict[str, int], total: int) -> tuple[str, float]:
            if total <= 0:
                return "-", 0.0
            pct_zero = counts.get("0", 0) / total
            pct_45 = (counts.get("+45", 0) + counts.get("-45", 0)) / total
            pct_90 = counts.get("90", 0) / total
            threshold = 0.45
            if pct_zero >= threshold and pct_zero >= pct_45 and pct_zero >= pct_90:
                return "Hard", pct_zero
            if pct_45 >= threshold and pct_45 >= pct_zero and pct_45 >= pct_90:
                return "Soft", pct_45
            dominant_pct = max(pct_zero, pct_45, pct_90)
            return "Quasi-isotropic", dominant_pct

        def _header_label_html(
            laminate: Laminado,
            cell_id: str,
            layers: Optional[list[Camada]] = None,
        ) -> str:
            layers = list(layers if layers is not None else getattr(laminate, "camadas", []) or [])
            counts, total = _orientation_counts(layers)
            lam_type, pct = _classify_laminate_type(counts, total)
            pct_text = f"({pct * 100:.0f}%)" if pct > 0 else ""
            balance = evaluate_laminate_balance_clt(layers).is_balanced
            symmetry = evaluate_symmetry_for_layers(layers).is_symmetric
            balance_text = "Sim" if balance else "Não"
            symmetry_text = "Sim" if symmetry else "Não"
            balance_color = "#2563eb" if balance else "#dc2626"
            symmetry_color = "#2563eb" if symmetry else "#dc2626"
            return (
                "<div style='text-align:center; font-weight:400;'>"
                f"<div>{cell_id}</div>"
                f"<div>Tipo: {lam_type}{pct_text}</div>"
                f"<div>Balanceado: <span style='color:{balance_color};'>{balance_text}</span></div>"
                f"<div>Simétrico: <span style='color:{symmetry_color};'>{symmetry_text}</span></div>"
                "</div>"
            )

        def _apply_header_labels() -> None:
            table.setHorizontalHeaderLabels([
                "Sequence",
                "",
                "",
                "",
            ])
            header.set_section_html(1, _header_label_html(min_laminate, min_cell_id))
            header.set_section_html(
                2,
                _header_label_html(
                    intermediate_laminate,
                    "Novo Laminado Intermediário",
                    mid_layers,
                ),
            )
            header.set_section_html(3, _header_label_html(max_laminate, max_cell_id))

        _apply_header_labels()

        def _layer_text(layer: Optional[Camada]) -> str:
            if layer is None:
                return ""
            orientation = getattr(layer, "orientacao", None)
            if orientation is None:
                return "Empty"
            return format_orientation_value(orientation)

        def _apply_cell_style(item: QTableWidgetItem, orientation: Optional[float]) -> None:
            item.setTextAlignment(Qt.AlignCenter)
            bg_color = orientation_highlight_color(orientation)
            if bg_color is not None:
                item.setBackground(QBrush(bg_color))
            else:
                item.setBackground(QBrush())
            if orientation is None:
                item.setForeground(QBrush(QColor(160, 160, 160)))
            else:
                try:
                    angle = normalize_angle(orientation)
                except Exception:
                    angle = None
                if angle is not None and abs(float(angle) - 90.0) <= 1e-9:
                    item.setForeground(QBrush(QColor(255, 255, 255)))
                else:
                    item.setForeground(QBrush())

        for row in range(row_count):
            seq_item = QTableWidgetItem(f"Seq.{row + 1}")
            seq_item.setFlags(seq_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(row, 0, seq_item)

            min_layer = min_layers[row] if row < len(min_layers) else None
            min_item = QTableWidgetItem(_layer_text(min_layer))
            min_item.setFlags(min_item.flags() & ~Qt.ItemIsEditable)
            min_orientation = getattr(min_layer, "orientacao", None) if min_layer else None
            min_item.setData(Qt.UserRole, min_orientation)
            _apply_cell_style(min_item, min_orientation)
            table.setItem(row, 1, min_item)

            mid_layer = mid_layers[row] if row < len(mid_layers) else None
            mid_item = QTableWidgetItem(_layer_text(mid_layer))
            mid_item.setFlags(mid_item.flags() | Qt.ItemIsEditable)
            mid_orientation = getattr(mid_layer, "orientacao", None) if mid_layer else None
            mid_item.setData(Qt.UserRole, mid_orientation)
            _apply_cell_style(mid_item, mid_orientation)
            table.setItem(row, 2, mid_item)

            max_layer = max_layers[row] if row < len(max_layers) else None
            max_item = QTableWidgetItem(_layer_text(max_layer))
            max_item.setFlags(max_item.flags() & ~Qt.ItemIsEditable)
            max_orientation = getattr(max_layer, "orientacao", None) if max_layer else None
            max_item.setData(Qt.UserRole, max_orientation)
            _apply_cell_style(max_item, max_orientation)
            table.setItem(row, 3, max_item)

        table.verticalHeader().setVisible(False)
        header = table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignCenter)
        header_font = header.font()
        header_font.setBold(True)
        header.setFont(header_font)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setMinimumHeight(76)
        table.setAlternatingRowColors(True)
        table.setSelectionMode(QTableWidget.SingleSelection)
        table.setSelectionBehavior(QTableWidget.SelectItems)
        table.setEditTriggers(QAbstractItemView.DoubleClicked)
        table.setFocusPolicy(Qt.StrongFocus)

        replaced_rows: set[int] = set()
        for idx in range(row_count):
            max_layer = max_layers[idx] if idx < len(max_layers) else None
            mid_layer = mid_layers[idx] if idx < len(mid_layers) else None
            max_orient = getattr(max_layer, "orientacao", None) if max_layer else None
            mid_orient = getattr(mid_layer, "orientacao", None) if mid_layer else None
            if max_orient is not None and mid_orient is None:
                replaced_rows.add(idx)

        delegate = _IntermediateOrientationDelegate(replaced_rows, table)
        table.setItemDelegateForColumn(2, delegate)

        symmetry_color = QColor(250, 128, 114)

        def _center_rows() -> set[int]:
            total = row_count
            if total <= 0:
                return set()
            if total % 2 == 1:
                return {total // 2}
            return {total // 2 - 1, total // 2}

        def _apply_symmetry_highlight() -> set[int]:
            rows = _center_rows()
            for row in range(row_count):
                for col in range(table.columnCount()):
                    item = table.item(row, col)
                    if item is None:
                        continue
                    if row in rows:
                        item.setBackground(QBrush(symmetry_color))
                    else:
                        if col == 0:
                            item.setBackground(QBrush())
                        else:
                            orientation = item.data(Qt.UserRole)
                            _apply_cell_style(item, orientation)
            return rows

        symmetry_rows = _apply_symmetry_highlight()

        def _on_item_changed(item: QTableWidgetItem) -> None:
            if item.column() != 2:
                return
            row = item.row()
            text = item.text().strip()
            if text.lower() in {"", "empty"}:
                new_value: Optional[float] = None
            else:
                try:
                    new_value = normalize_angle(text.replace("°", ""))
                except Exception:
                    QMessageBox.warning(
                        self,
                        "Orientação inválida",
                        "Informe uma orientação válida (ex.: 45, -45, 0, 90) ou 'Empty'.",
                    )
                    prev = item.data(Qt.UserRole)
                    table.blockSignals(True)
                    item.setText("Empty" if prev is None else format_orientation_value(prev))
                    table.blockSignals(False)
                    return

            if 0 <= row < len(mid_layers):
                mid_layers[row].orientacao = new_value
            item.setData(Qt.UserRole, new_value)
            table.blockSignals(True)
            item.setText("Empty" if new_value is None else format_orientation_value(new_value))
            table.blockSignals(False)
            _apply_cell_style(item, new_value)

            max_layer = max_layers[row] if row < len(max_layers) else None
            max_orient = getattr(max_layer, "orientacao", None) if max_layer else None
            if max_orient is not None and new_value is None:
                replaced_rows.add(row)
            else:
                replaced_rows.discard(row)
            table.blockSignals(True)
            _apply_symmetry_highlight()
            _apply_header_labels()
            table.blockSignals(False)
            table.viewport().update()

        table.itemChanged.connect(_on_item_changed)
        layout.addWidget(table, stretch=1)

        buttons = QDialogButtonBox(parent=self)
        create_button = QPushButton("Criar Laminado", self)
        buttons.addButton(create_button, QDialogButtonBox.ActionRole)
        buttons.addButton(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        create_button.clicked.connect(self._on_create_laminate_clicked)
        layout.addWidget(buttons)

    def _find_main_window(self):
        widget = self.parent()
        while widget is not None:
            if hasattr(widget, "_refresh_after_new_laminate"):
                return widget
            widget = widget.parent()
        return None

    def _on_create_laminate_clicked(self) -> None:
        if self._model is None:
            QMessageBox.information(self, "Criar Laminado", "Nenhum projeto carregado.")
            return
        base_tag = str(getattr(self._max_laminate, "tag", "") or "").strip()
        dialog = CreateIntermediateLaminateDialog(
            self._model,
            base_tag,
            self._template_laminate,
            parent=self,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        new_tag = dialog.selected_tag()

        new_layers = copy.deepcopy(self._mid_layers)
        for idx, layer in enumerate(new_layers):
            layer.idx = idx

        new_laminate = Laminado(
            nome="",
            tipo=str(getattr(self._max_laminate, "tipo", "") or ""),
            color_index=getattr(self._max_laminate, "color_index", 1),
            tag=new_tag,
            celulas=[],
            camadas=new_layers,
            auto_rename_enabled=True,
        )
        new_name = auto_name_for_laminate(self._model, new_laminate)
        if not new_name:
            QMessageBox.warning(self, "Criar Laminado", "Falha ao gerar nome automático.")
            return
        new_laminate.nome = new_name
        self._model.laminados[new_name] = new_laminate

        if isinstance(self.parent(), IntermediateLaminateWindow):
            try:
                self.parent()._notify_intermediate_created(
                    new_laminate,
                    self._min_laminate,
                    self._max_laminate,
                    self._min_cell_id,
                    self._max_cell_id,
                )
            except Exception:
                pass

        main_window = self._find_main_window()
        if main_window is not None:
            try:
                main_window._refresh_after_new_laminate(new_name)
            except Exception:
                pass
            if hasattr(main_window, "_mark_dirty"):
                try:
                    main_window._mark_dirty()
                except Exception:
                    pass

        QMessageBox.information(
            self,
            "Criar Laminado",
            f"Laminado '{new_name}' criado e disponível na lista de seleção.",
        )
        self.close()
        if isinstance(self.parent(), QDialog):
            try:
                self.parent().show()
                self.parent().raise_()
                self.parent().activateWindow()
            except Exception:
                pass


class _IntermediateOrientationDelegate(QStyledItemDelegate):
    def __init__(self, replaced_rows: set[int], parent=None) -> None:
        super().__init__(parent)
        self._replaced_rows = replaced_rows
        self._common_orientations = [
            "45°",
            "-45°",
            "0°",
            "90°",
            "30°",
            "-30°",
            "60°",
            "-60°",
        ]

    def createEditor(self, parent, option, index):  # type: ignore[override]
        combo = QComboBox(parent)
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.addItems(self._common_orientations)
        line_edit = combo.lineEdit()
        if line_edit is not None:
            line_edit.setPlaceholderText("Digite a orientação (ex.: 45)")
        return combo

    def setEditorData(self, editor, index) -> None:  # type: ignore[override]
        text = str(index.data(Qt.DisplayRole) or "")
        if isinstance(editor, QComboBox):
            editor.setCurrentText(text)

    def setModelData(self, editor, model, index) -> None:  # type: ignore[override]
        if isinstance(editor, QComboBox):
            text = editor.currentText().strip()
        else:
            text = ""
        model.setData(index, text, Qt.EditRole)

    def paint(self, painter: QPainter, option, index) -> None:  # type: ignore[override]
        super().paint(painter, option, index)
        if index.column() != 2:
            return
        if index.row() not in self._replaced_rows:
            return
        painter.save()
        pen = QPen(QColor(220, 53, 69))
        pen.setWidth(2)
        painter.setPen(pen)
        rect = option.rect.adjusted(4, 4, -4, -4)
        painter.drawRect(rect)
        painter.restore()
