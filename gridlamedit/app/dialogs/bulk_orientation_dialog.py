"""Dialog to define an orientation applied in bulk to selected layers."""

from __future__ import annotations

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

from gridlamedit.io.spreadsheet import (
    ORIENTATION_MAX,
    ORIENTATION_MIN,
    format_orientation_value,
    normalize_angle,
)

try:
    from gridlamedit.app.services.excel_io import (
        _normalize_angle as legacy_normalize_angle,
    )
except ImportError:  # pragma: no cover - fallback
    legacy_normalize_angle = None

_DEFAULT_SUGGESTIONS = [0.0, 45.0, -45.0, 90.0]


class BulkOrientationDialog(QDialog):
    """Collects a new orientation for all selected layers."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        available_orientations: Sequence[float] | Iterable[float] | None = None,
    ) -> None:
        super().__init__(parent)
        self.new_orientation: float | None = None
        self._setup_ui()
        self._populate_orientations(available_orientations)

    def _setup_ui(self) -> None:
        self.setWindowTitle("Trocar orienta\u00e7\u00e3o (camadas selecionadas)")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        instructions = QLabel(
            (
                "Informe a nova orienta\u00e7\u00e3o (qualquer valor entre "
                f"{ORIENTATION_MIN:.0f}\N{DEGREE SIGN} e {ORIENTATION_MAX:.0f}\N{DEGREE SIGN})."
            ),
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

        self.resize(400, 190)

    def _populate_orientations(
        self, orientations: Sequence[float] | Iterable[float] | None
    ) -> None:
        ordered: list[float] = []
        for value in _DEFAULT_SUGGESTIONS:
            if value not in ordered:
                ordered.append(float(value))

        extras: set[float] = set()
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
            self.cmb_orientation.addItem(format_orientation_value(value))

    def _coerce_orientation_value(self, value: object) -> float | None:
        candidate = value
        if legacy_normalize_angle is not None:
            try:
                candidate = legacy_normalize_angle(value)
            except Exception:  # pragma: no cover - defensive
                candidate = value
        try:
            return normalize_angle(candidate)
        except (TypeError, ValueError):
            try:
                return normalize_angle(str(candidate))
            except (TypeError, ValueError):
                return None

    def _on_accept_clicked(self) -> None:
        text = self.cmb_orientation.currentText().strip()
        if not text:
            self._show_invalid_message()
            return
        try:
            self.new_orientation = normalize_angle(text)
        except ValueError:
            self._show_invalid_message()
            return
        self.accept()

    def _show_invalid_message(self) -> None:
        QMessageBox.warning(
            self,
            "Dados inv\u00e1lidos",
            (
                "Informe uma orienta\u00e7\u00e3o num\u00e9rica em graus "
                f"entre {ORIENTATION_MIN:.0f}\N{DEGREE SIGN} e {ORIENTATION_MAX:.0f}\N{DEGREE SIGN}."
            ),
        )
