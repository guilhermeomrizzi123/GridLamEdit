"""Dialog for selecting a material to apply to multiple layers."""

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


class BulkMaterialDialog(QDialog):
    """Collects the material to apply to all selected layers."""

    def __init__(
        self,
        parent: QWidget | None = None,
        materials: Sequence[str] | Iterable[str] | None = None,
    ) -> None:
        super().__init__(parent)
        self.selected_material: str = ""
        self._setup_ui(materials)

    def _setup_ui(self, materials: Sequence[str] | Iterable[str] | None) -> None:
        self.setWindowTitle("Trocar material (camadas selecionadas)")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        instructions = QLabel(
            "Escolha o material para aplicar em todas as camadas selecionadas.",
            self,
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        self.cmb_material = QComboBox(self)
        self.cmb_material.setEditable(True)
        self.cmb_material.setInsertPolicy(QComboBox.NoInsert)
        self.cmb_material.setSizeAdjustPolicy(
            QComboBox.AdjustToContentsOnFirstShow
        )
        layout.addWidget(self.cmb_material)

        if materials:
            for item in materials:
                text = (item or "").strip()
                if text:
                    self.cmb_material.addItem(text)

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
        self.selected_material = self.cmb_material.currentText().strip()
        self.accept()

