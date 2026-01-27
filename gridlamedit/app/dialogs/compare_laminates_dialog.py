"""Dialog for comparing two laminates."""

from __future__ import annotations

from typing import Iterable

from PySide6.QtCore import Signal
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


class CompareLaminatesDialog(QDialog):
    """Dialog used to select two laminates and show comparison report."""

    compare_requested = Signal(str, str)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Comparar Laminados")
        self.setModal(False)
        self.resize(720, 520)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        selection_layout = QGridLayout()
        selection_layout.setHorizontalSpacing(12)
        selection_layout.setVerticalSpacing(8)

        label_a = QLabel("Laminado A:", self)
        label_b = QLabel("Laminado B:", self)
        self.laminate_a_combo = QComboBox(self)
        self.laminate_b_combo = QComboBox(self)
        self.laminate_a_combo.setMinimumWidth(240)
        self.laminate_b_combo.setMinimumWidth(240)

        selection_layout.addWidget(label_a, 0, 0)
        selection_layout.addWidget(self.laminate_a_combo, 0, 1)
        selection_layout.addWidget(label_b, 1, 0)
        selection_layout.addWidget(self.laminate_b_combo, 1, 1)

        layout.addLayout(selection_layout)

        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(8)
        self.compare_button = QPushButton("Comparar", self)
        self.close_button = QPushButton("Fechar", self)
        self.compare_button.clicked.connect(self._on_compare_clicked)
        self.close_button.clicked.connect(self.close)

        buttons_layout.addWidget(self.compare_button)
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.close_button)

        layout.addLayout(buttons_layout)

        self.report_view = QTextEdit(self)
        self.report_view.setReadOnly(True)
        self.report_view.setLineWrapMode(QTextEdit.NoWrap)
        self.report_view.setMinimumHeight(240)
        layout.addWidget(self.report_view, stretch=1)

    def set_laminate_names(
        self,
        names: Iterable[str],
        *,
        select_a: str | None = None,
        select_b: str | None = None,
    ) -> None:
        name_list = [str(name) for name in names]
        self.laminate_a_combo.blockSignals(True)
        self.laminate_b_combo.blockSignals(True)
        self.laminate_a_combo.clear()
        self.laminate_b_combo.clear()
        self.laminate_a_combo.addItems(name_list)
        self.laminate_b_combo.addItems(name_list)

        if select_a in name_list:
            self.laminate_a_combo.setCurrentText(select_a)
        if select_b in name_list:
            self.laminate_b_combo.setCurrentText(select_b)
        elif name_list and self.laminate_b_combo.currentText() == self.laminate_a_combo.currentText():
            if len(name_list) > 1:
                self.laminate_b_combo.setCurrentIndex(1)

        self.laminate_a_combo.blockSignals(False)
        self.laminate_b_combo.blockSignals(False)

    def selected_names(self) -> tuple[str, str]:
        return self.laminate_a_combo.currentText().strip(), self.laminate_b_combo.currentText().strip()

    def set_report(self, text: str) -> None:
        self.report_view.setPlainText(text)
        self.report_view.moveCursor(QTextCursor.Start)

    def _on_compare_clicked(self) -> None:
        name_a, name_b = self.selected_names()
        self.compare_requested.emit(name_a, name_b)
