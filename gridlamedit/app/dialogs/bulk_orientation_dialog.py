"""Dialog to define an orientation applied in bulk to selected layers."""

from __future__ import annotations

import re

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QLineEdit,
    QLabel,
    QMessageBox,
    QVBoxLayout,
    QWidget,
)

from gridlamedit.io.spreadsheet import normalize_angle

_ORIENTATION_PATTERN = re.compile(r"^[+-]?\d+(?:\.\d+)?[°º]?$")


class BulkOrientationDialog(QDialog):
    """Collects a new orientation for all selected layers."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.new_orientation: int | None = None
        self._setup_ui()

    def _setup_ui(self) -> None:
        self.setWindowTitle("Trocar orientação (camadas selecionadas)")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        instructions = QLabel(
            "Informe a nova orientação (ex.: 0, 45, -45, 90).",
            self,
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        self.edt_orientation = QLineEdit(self)
        self.edt_orientation.setPlaceholderText("0")
        layout.addWidget(self.edt_orientation)

        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, self
        )
        apply_button = button_box.button(QDialogButtonBox.Ok)
        apply_button.setText("Aplicar")
        cancel_button = button_box.button(QDialogButtonBox.Cancel)
        cancel_button.setText("Cancelar")
        button_box.accepted.connect(self._on_accept_clicked)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.resize(360, 160)

    def _on_accept_clicked(self) -> None:
        text = self.edt_orientation.text().strip()
        normalized = text.replace("º", "°")
        if not normalized or not _ORIENTATION_PATTERN.match(normalized):
            self._show_invalid_message()
            return
        cleaned = normalized.replace("°", "")
        try:
            number = float(cleaned)
        except ValueError:
            self._show_invalid_message()
            return
        rounded = int(round(number))
        try:
            self.new_orientation = normalize_angle(rounded)
        except ValueError:
            self._show_invalid_message()
            return
        self.accept()

    def _show_invalid_message(self) -> None:
        QMessageBox.warning(
            self,
            "Dados inválidos",
            "Informe uma orientação válida (ex.: 0, 45, -45, 90).",
        )

