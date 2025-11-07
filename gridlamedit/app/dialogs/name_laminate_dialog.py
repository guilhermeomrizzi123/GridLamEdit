"""Dialog used to collect a unique name for the cloned laminate."""

from __future__ import annotations

import re
from typing import Iterable, Sequence

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

_NAME_PATTERN = re.compile(r"^[A-Za-z0-9_\-\s\.]+$")


class NameLaminateDialog(QDialog):
    """Collects and validates the new laminate name."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Nome do novo laminado")
        self.edt_name = QLineEdit(self)
        self._button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, self
        )
        self._existing_names: set[str] = set()
        self._setup_ui()

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        instructions = QLabel(
            "Informe um nome unico para o novo laminado.", self
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        self.edt_name.setPlaceholderText("Ex.: WEB-RIB-26")
        layout.addWidget(self.edt_name)

        create_button = self._button_box.button(QDialogButtonBox.Ok)
        create_button.setText("Criar")
        cancel_button = self._button_box.button(QDialogButtonBox.Cancel)
        cancel_button.setText("Cancelar")
        self._button_box.accepted.connect(self._validate_and_accept)
        self._button_box.rejected.connect(self.reject)
        layout.addWidget(self._button_box)

        self.resize(360, 160)

    def set_existing_names(
        self, names: Sequence[str] | Iterable[str]
    ) -> None:
        """Set the reference list used to avoid duplicates."""
        self._existing_names = {
            str(name).strip().lower() for name in names if str(name).strip()
        }

    def result_name(self) -> str:
        """Return the trimmed input text."""
        return self.edt_name.text().strip()

    def _validate_and_accept(self) -> None:
        text = self.result_name()
        if not text:
            QMessageBox.warning(
                self, "Nome obrigatorio", "Informe um nome para o laminado."
            )
            return
        if not _NAME_PATTERN.match(text):
            QMessageBox.warning(
                self,
                "Nome invalido",
                "Use apenas letras, numeros, espacos, '.', '_' ou '-'.",
            )
            return
        if text.lower() in self._existing_names:
            QMessageBox.warning(
                self,
                "Nome duplicado",
                f"Ja existe um laminado chamado '{text}'. Escolha outro nome.",
            )
            return
        self.accept()

