"""Qt delegates used across the GridLamEdit UI."""

from __future__ import annotations

from typing import Callable, Iterable, Optional

from PySide6.QtCore import Qt, QRect, QRegularExpression
from PySide6.QtGui import QRegularExpressionValidator, QPainter, QPen, QColor
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QInputDialog,
    QStyledItemDelegate,
    QWidget,
    QStyle,
    QStyleOptionButton,
)

from gridlamedit.io.spreadsheet import (
    DEFAULT_PLY_TYPE,
    PLY_TYPE_OPTIONS,
    ORIENTATION_SYMMETRY_ROLE,
    normalize_angle,
    normalize_ply_type_label,
)


class _BaseComboDelegate(QStyledItemDelegate):
    """Combo-box based delegate that defers option listing to a provider callable."""

    def __init__(
        self,
        parent: Optional[QWidget] = None,
        *,
        items_provider: Optional[Callable[[], Iterable[str]]] = None,
    ) -> None:
        super().__init__(parent)
        self._items_provider = items_provider or (lambda: [])

    def _items(self) -> list[str]:
        return [str(item) for item in self._items_provider()]

    def createEditor(self, parent: QWidget, option, index):  # noqa: D401, N802
        editor = QComboBox(parent)
        editor.setEditable(False)
        editor.addItems(self._items())
        return editor

    def setEditorData(self, editor: QWidget, index):  # noqa: N802
        if isinstance(editor, QComboBox):
            current = index.data(Qt.EditRole) or index.data(Qt.DisplayRole)
            if current is None:
                return
            text = str(current)
            items = self._items()
            if text not in items:
                items = [text, *items]
                editor.clear()
                editor.addItems(items)
            target_index = editor.findText(text)
            if target_index >= 0:
                editor.setCurrentIndex(target_index)

    def setModelData(self, editor: QWidget, model, index):  # noqa: N802
        if isinstance(editor, QComboBox):
            model.setData(index, editor.currentText(), Qt.EditRole)


class MaterialComboDelegate(_BaseComboDelegate):
    """Delegate para edicao inline da coluna de material."""


class OrientationComboDelegate(QStyledItemDelegate):
    """Delegate que restringe a edicao de orientacao a um combo controlado."""

    EMPTY_LABEL = "Empty"
    CUSTOM_TOKEN = "__custom__"

    def __init__(
        self,
        parent: Optional[QWidget] = None,
        *,
        items_provider: Optional[Callable[[], Iterable[str]]] = None,
    ) -> None:
        super().__init__(parent)
        # items_provider is kept for API compatibility, even though options are fixed.
        self._items_provider = items_provider or (lambda: [])

    def _option_items(self) -> list[tuple[str, object]]:
        return [
            (self.EMPTY_LABEL, None),
            ("0", 0.0),
            ("45", 45.0),
            ("-45", -45.0),
            ("90", 90.0),
            ("Outro valor...", self.CUSTOM_TOKEN),
        ]

    def _prompt_custom_orientation(self, parent: QWidget) -> Optional[float]:
        dialog = QInputDialog(parent)
        dialog.setInputMode(QInputDialog.DoubleInput)
        dialog.setWindowTitle("Outro valor")
        dialog.setLabelText("Informe a orientacao (-100 a 100 graus):")
        dialog.setDoubleRange(-100.0, 100.0)
        dialog.setDoubleDecimals(1)
        dialog.setDoubleStep(1.0)
        dialog.setDoubleValue(0.0)
        dialog.setTextValue("")
        if dialog.exec() != dialog.Accepted:
            return None
        return dialog.doubleValue()

    def _coerce_orientation(self, value: object) -> Optional[float]:
        if value is None:
            return None
        text = str(value).strip()
        if not text:
            return None
        if text.lower() == self.EMPTY_LABEL.lower():
            return None
        try:
            return normalize_angle(value)
        except Exception:
            try:
                return normalize_angle(text)
            except Exception:
                return None

    def _find_index_for_value(self, editor: QComboBox, value: Optional[float]) -> int:
        for idx in range(editor.count()):
            data = editor.itemData(idx)
            if value is None and data is None:
                return idx
            if isinstance(data, float) and value is not None:
                try:
                    if abs(data - value) <= 1e-9:
                        return idx
                except Exception:
                    continue
        return editor.count() - 1  # Custom option.

    def createEditor(self, parent: QWidget, option, index):  # noqa: D401, N802
        editor = QComboBox(parent)
        editor.setEditable(False)
        for text, data in self._option_items():
            editor.addItem(text, data)
        editor.setProperty("allowCustomPrompt", False)
        editor.setProperty("pendingCustomOrientation", None)
        editor.currentIndexChanged.connect(
            lambda _idx, ed=editor: self._on_index_changed(ed)
        )
        return editor

    def setEditorData(self, editor: QWidget, index):  # noqa: N802
        if not isinstance(editor, QComboBox):
            return
        current = index.data(Qt.EditRole) or index.data(Qt.DisplayRole)
        value = self._coerce_orientation(current)
        editor.setProperty("currentOrientation", value)
        editor.setCurrentIndex(self._find_index_for_value(editor, value))
        editor.setProperty("allowCustomPrompt", True)

    def setModelData(self, editor: QWidget, model, index):  # noqa: N802
        if not isinstance(editor, QComboBox):
            return
        data = editor.currentData()
        if data == self.CUSTOM_TOKEN:
            pending_value = editor.property("pendingCustomOrientation")
            value = (
                float(pending_value)
                if isinstance(pending_value, (int, float))
                else self._prompt_custom_orientation(editor)
            )
            editor.setProperty("pendingCustomOrientation", None)
            if value is None:
                return
            model.setData(index, value, Qt.EditRole)
            return
        if data is None:
            model.setData(index, "", Qt.EditRole)
        else:
            model.setData(index, float(data), Qt.EditRole)

    def _restore_previous_selection(self, editor: QComboBox) -> None:
        previous_value = self._coerce_orientation(editor.property("currentOrientation"))
        editor.setProperty("allowCustomPrompt", False)
        editor.setCurrentIndex(self._find_index_for_value(editor, previous_value))
        editor.setProperty("allowCustomPrompt", True)

    def _on_index_changed(self, editor: QComboBox) -> None:
        if not editor.property("allowCustomPrompt"):
            return
        if editor.currentData() != self.CUSTOM_TOKEN:
            editor.setProperty("pendingCustomOrientation", None)
            return
        value = self._prompt_custom_orientation(editor)
        if value is None:
            self._restore_previous_selection(editor)
            return
        editor.setProperty("pendingCustomOrientation", value)
        self.commitData.emit(editor)
        self.closeEditor.emit(editor)

    def paint(self, painter: QPainter, option, index):  # noqa: N802
        super().paint(painter, option, index)
        if not bool(index.data(ORIENTATION_SYMMETRY_ROLE)):
            return
        pen = QPen(QColor("orange"))
        pen.setStyle(Qt.DashLine)
        painter.save()
        painter.setPen(pen)
        rect = option.rect.adjusted(1, 1, -1, -1)
        painter.drawRect(rect)
        painter.restore()


class PlyTypeComboDelegate(QStyledItemDelegate):
    """Delegate que apresenta o tipo de ply como combo."""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.items = list(PLY_TYPE_OPTIONS)

    def createEditor(self, parent: QWidget, option, index):  # noqa: N802
        cb = QComboBox(parent)
        cb.addItems(self.items)
        cb.setEditable(False)
        return cb

    def setEditorData(self, editor: QWidget, index):  # noqa: N802
        if not isinstance(editor, QComboBox):
            return
        current = (
            index.data(Qt.EditRole)
            or index.data(Qt.DisplayRole)
            or DEFAULT_PLY_TYPE
        )
        normalized = normalize_ply_type_label(current)
        idx = editor.findText(normalized)
        editor.setCurrentIndex(idx if idx >= 0 else 0)

    def setModelData(self, editor: QWidget, model, index):  # noqa: N802
        if not isinstance(editor, QComboBox):
            return
        selection = normalize_ply_type_label(editor.currentText())
        model.setData(index, selection, Qt.EditRole)


class CenteredCheckBoxDelegate(QStyledItemDelegate):
    """Renderiza checkboxes centralizados nas celulas."""

    def paint(self, painter, option, index):  # noqa: N802
        value = index.data(Qt.CheckStateRole)
        if value is None:
            super().paint(painter, option, index)
            return

        opt = QStyleOptionButton()
        opt.state |= QStyle.State_Enabled
        if value == Qt.Checked:
            opt.state |= QStyle.State_On
        else:
            opt.state |= QStyle.State_Off
        opt.rect = self._indicator_rect(option)
        QApplication.style().drawControl(QStyle.CE_CheckBox, opt, painter)

    def _indicator_rect(self, option: QStyleOptionButton) -> QRect:
        indicator = QApplication.style().subElementRect(
            QStyle.SE_CheckBoxIndicator, option, None
        )
        x = option.rect.x() + (option.rect.width() - indicator.width()) // 2
        y = option.rect.y() + (option.rect.height() - indicator.height()) // 2
        return QRect(x, y, indicator.width(), indicator.height())
