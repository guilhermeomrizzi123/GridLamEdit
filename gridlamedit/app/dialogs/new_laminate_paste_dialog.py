"""Dialog to create a laminate by pasting orientations."""

from __future__ import annotations

import re
from typing import Iterable

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from gridlamedit.io.spreadsheet import normalize_angle

_TOKEN_PATTERN = re.compile(r"^[+-]?\d+(?:\.\d+)?°?$")
_SPLIT_PATTERN = re.compile(r"[,\s;\/|]+")


class NewLaminatePasteDialog(QDialog):
    """Dialog that parses pasted orientations to build a laminate."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.result_orientations: list[int] = []
        self._setup_ui()

    def _setup_ui(self) -> None:
        self.setWindowTitle("Criar laminado por colagem")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        instructions = QLabel(
            (
                "Cole abaixo as orientações (ex.: 0, 45, -45, 90). "
                "Você pode colar com Ctrl+V. Caracteres não compatíveis serão ignorados."
            ),
            self,
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        self.txt_paste = QTextEdit(self)
        self.txt_paste.setAcceptRichText(False)
        self.txt_paste.setMinimumHeight(220)
        layout.addWidget(self.txt_paste)

        self.cb_symmetric = QCheckBox("Criar laminado simétrico", self)
        layout.addWidget(self.cb_symmetric)

        self.cb_last_layer_center = QCheckBox(
            "Considerar última camada como camada central do laminado", self
        )
        self.cb_last_layer_center.setEnabled(False)
        layout.addWidget(self.cb_last_layer_center)

        self.cb_symmetric.toggled.connect(self._on_symmetric_toggled)

        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, self
        )
        ok_button = button_box.button(QDialogButtonBox.Ok)
        ok_button.setText("Adicionar camadas")
        cancel_button = button_box.button(QDialogButtonBox.Cancel)
        cancel_button.setText("Cancelar")
        button_box.accepted.connect(self._on_accept_clicked)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.resize(520, 380)

    def _on_accept_clicked(self) -> None:
        orientations_base = self._parse_orientations(self.txt_paste.toPlainText())
        if not orientations_base:
            QMessageBox.warning(
                self,
                "Dados inválidos",
                "Nenhuma orientação válida encontrada.",
            )
            return

        if self.cb_symmetric.isChecked():
            include_center = self.cb_last_layer_center.isChecked()
            full_sequence = self._build_symmetric_sequence(
                orientations_base, include_center=include_center
            )
        else:
            full_sequence = orientations_base

        self.result_orientations = full_sequence
        self.accept()

    def _parse_orientations(self, raw_text: str) -> list[int]:
        tokens = _SPLIT_PATTERN.split(raw_text or "")
        orientations: list[int] = []
        for token in tokens:
            cleaned = token.strip()
            if not cleaned:
                continue
            if not _TOKEN_PATTERN.match(cleaned):
                continue
            normalized_token = cleaned.replace("°", "").replace("º", "")
            try:
                number = float(normalized_token)
            except ValueError:
                continue
            rounded = int(round(number))
            try:
                orientation = normalize_angle(rounded)
            except ValueError:
                continue
            orientations.append(orientation)
        return orientations

    def _on_symmetric_toggled(self, checked: bool) -> None:
        self.cb_last_layer_center.setEnabled(checked)
        if not checked:
            self.cb_last_layer_center.setChecked(False)

    def _build_symmetric_sequence(
        self, base: Iterable[int], *, include_center: bool
    ) -> list[int]:
        items = list(base)
        if not items:
            return []
        if include_center:
            mirrored = items[:-1]
            center = items[-1]
            return mirrored + [center] + list(reversed(mirrored))
        return items + list(reversed(items))
