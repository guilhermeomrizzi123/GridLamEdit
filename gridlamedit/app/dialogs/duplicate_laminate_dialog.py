"""Dialog used to choose which laminate will be duplicated."""

from __future__ import annotations

from typing import Iterable, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QVBoxLayout,
    QWidget,
)


class DuplicateLaminateDialog(QDialog):
    """Simple selector that lists all available laminates."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Duplicar laminado")
        self.cmb_existing = QComboBox(self)
        self._button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, self
        )
        self._setup_ui()

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        instructions = QLabel(
            "Selecione o laminado que deseja duplicar.", self
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        self.cmb_existing.setEditable(False)
        self.cmb_existing.setSizeAdjustPolicy(
            QComboBox.AdjustToContentsOnFirstShow
        )
        layout.addWidget(self.cmb_existing)

        duplicate_button = self._button_box.button(QDialogButtonBox.Ok)
        duplicate_button.setText("Duplicar")
        cancel_button = self._button_box.button(QDialogButtonBox.Cancel)
        cancel_button.setText("Cancelar")
        self._button_box.accepted.connect(self.accept)
        self._button_box.rejected.connect(self.reject)
        layout.addWidget(self._button_box)

        self.resize(360, 180)

    def set_laminates(self, names: Sequence[str] | Iterable[str]) -> None:
        """Populate the combo box with laminate names."""
        self.cmb_existing.blockSignals(True)
        self.cmb_existing.clear()
        for name in names:
            clean = str(name).strip()
            if clean:
                self.cmb_existing.addItem(clean)
        self.cmb_existing.blockSignals(False)
        has_items = self.cmb_existing.count() > 0
        self.cmb_existing.setEnabled(has_items)
        duplicate_button = self._button_box.button(QDialogButtonBox.Ok)
        if duplicate_button is not None:
            duplicate_button.setEnabled(has_items)

    def selected_name(self) -> str:
        """Return the current laminate selection."""
        return self.cmb_existing.currentText().strip()

