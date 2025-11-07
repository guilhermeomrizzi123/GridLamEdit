"""Modeless dialog used to show the laminate stacking summary."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QTextOption
from PySide6.QtWidgets import QDialog, QPlainTextEdit, QVBoxLayout, QWidget


class StackingSummaryDialog(QDialog):
    """Displays the stacking summary in its own resizable window."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setModal(False)
        self.setWindowTitle("Resumo do Stacking")
        self._setup_ui()

    def _setup_ui(self) -> None:
        self.setSizeGripEnabled(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        self.summary_view = QPlainTextEdit(self)
        self.summary_view.setObjectName("stackingSummaryDialogText")
        self.summary_view.setReadOnly(True)
        self.summary_view.setLineWrapMode(QPlainTextEdit.NoWrap)
        self.summary_view.setWordWrapMode(QTextOption.NoWrap)
        self.summary_view.document().setDefaultFont(QFont("Consolas", 10))
        self.summary_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.summary_view.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.summary_view.setMinimumSize(360, 240)
        layout.addWidget(self.summary_view)

        self.resize(520, 480)

    def update_summary(self, text: str) -> None:
        """Populate the summary area with the latest rendered text."""
        if not isinstance(text, str):
            text = "" if text is None else str(text)
        self.summary_view.setPlainText(text)
