"""Dialog to define an orientation applied in bulk to selected layers."""

from __future__ import annotations

import re
from typing import Iterable, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
    QVBoxLayout,
    QWidget,
)

from gridlamedit.io.spreadsheet import normalize_angle

try:
    from gridlamedit.app.services.excel_io import (
        _normalize_angle as legacy_normalize_angle,
    )
except ImportError:  # pragma: no cover - fallback
    legacy_normalize_angle = None

_ORIENTATION_PATTERN = re.compile(
    r"^[+-]?\d+(?:\.\d+)?(?:\N{DEGREE SIGN}|\N{MASCULINE ORDINAL INDICATOR})?$"
)
_DEGREE_TOKENS = {"\N{DEGREE SIGN}", "\N{MASCULINE ORDINAL INDICATOR}"}
_BASE_ORIENTATIONS = [0, 45, -45, 90]


class BulkOrientationDialog(QDialog):
    """Collects a new orientation for all selected layers."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        available_orientations: Sequence[int] | Iterable[int] | None = None,
    ) -> None:
        super().__init__(parent)
        self.new_orientation: int | None = None
        self._setup_ui()
        self._populate_orientations(available_orientations)

    def _setup_ui(self) -> None:
        self.setWindowTitle("Trocar orientação (camadas selecionadas)")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        instructions = QLabel(
            "Selecione ou informe a nova orientação (ex.: 0, 45, -45, 90).",
            self,
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        self.cmb_orientation = QComboBox(self)
        self.cmb_orientation.setEditable(True)
        self.cmb_orientation.setInsertPolicy(QComboBox.NoInsert)
        self.cmb_orientation.setSizeAdjustPolicy(
            QComboBox.AdjustToContentsOnFirstShow
        )
        layout.addWidget(self.cmb_orientation)

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

        self.resize(360, 180)

    def _populate_orientations(
        self, orientations: Sequence[int] | Iterable[int] | None
    ) -> None:
        ordered: list[int] = []
        for value in _BASE_ORIENTATIONS:
            if value not in ordered:
                ordered.append(value)

        extras: set[int] = set()
        if orientations:
            for item in orientations:
                normalized = self._coerce_orientation_value(item)
                if normalized is not None:
                    extras.add(normalized)
        for value in sorted(extras):
            if value not in ordered:
                ordered.append(value)

        self.cmb_orientation.clear()
        for value in ordered:
            self.cmb_orientation.addItem(str(value))

    def _coerce_orientation_value(self, value: object) -> int | None:
        candidate = value
        if legacy_normalize_angle is not None:
            try:
                candidate = legacy_normalize_angle(value)
            except Exception:  # pragma: no cover - defensive
                candidate = value
        try:
            number = float(candidate)
        except (TypeError, ValueError):
            try:
                number = float(str(candidate).strip())
            except (TypeError, ValueError):
                return None
        return int(round(number))

    def _on_accept_clicked(self) -> None:
        text = self.cmb_orientation.currentText().strip()
        if not text or not _ORIENTATION_PATTERN.match(text):
            self._show_invalid_message()
            return
        cleaned = text
        for token in _DEGREE_TOKENS:
            cleaned = cleaned.replace(token, "")
        cleaned = cleaned.strip()
        candidate_value: object = cleaned
        if legacy_normalize_angle is not None:
            try:
                candidate_value = legacy_normalize_angle(cleaned)
            except Exception:  # pragma: no cover - defensive
                candidate_value = cleaned
        try:
            number = float(candidate_value)
        except (TypeError, ValueError):
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
