"""Qt delegates used across the GridLamEdit UI."""

from __future__ import annotations

from typing import Callable, Iterable, Optional

from PySide6.QtCore import Qt, QRect, QEvent
from PySide6.QtWidgets import QApplication, QComboBox, QStyledItemDelegate, QWidget, QStyle, QStyleOptionButton


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


class OrientationComboDelegate(_BaseComboDelegate):
    """Delegate para edicao inline da coluna de orientacao."""


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
