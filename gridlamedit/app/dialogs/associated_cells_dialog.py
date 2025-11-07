"""Dialog to present the cells linked to the currently selected laminate."""

from __future__ import annotations

from typing import Iterable, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QListWidget,
    QListWidgetItem,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QVBoxLayout,
    QWidget,
)


class AssociatedCellsDialog(QDialog):
    """Shows the list of cells associated with a laminate."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Celulas associadas")
        self._current_laminate_name: str = ""
        self._cells_view = QListWidget(self)
        self._button_box = QDialogButtonBox(QDialogButtonBox.Close, self)
        self._header_label = QLabel(self)
        self._setup_ui()

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        self._header_label.setWordWrap(True)
        layout.addWidget(self._header_label)

        self._cells_view.setSelectionMode(QListWidget.NoSelection)
        layout.addWidget(self._cells_view)

        close_button = self._button_box.button(QDialogButtonBox.Close)
        close_button.setText("Fechar")
        self._button_box.rejected.connect(self.reject)
        self._button_box.accepted.connect(self.accept)
        layout.addWidget(self._button_box)

        self.resize(320, 280)

    def set_cells(self, cells: Sequence[str] | Iterable[str]) -> None:
        """Populate the list widget with the provided cells."""
        self._cells_view.clear()
        added = False
        for cell in cells:
            text = str(cell).strip()
            if not text:
                continue
            self._cells_view.addItem(QListWidgetItem(text))
            added = True
        if not added:
            empty = QListWidgetItem("(nenhuma celula associada)")
            empty.setFlags(empty.flags() & ~Qt.ItemIsEnabled)
            self._cells_view.addItem(empty)

    def refresh_from_laminate(
        self, laminate_name: str, cells: Sequence[str] | Iterable[str]
    ) -> None:
        """Update title and cell listing based on laminate info."""
        self._current_laminate_name = laminate_name.strip()
        if self._current_laminate_name:
            self._header_label.setText(
                f"Celulas associadas ao laminado '{self._current_laminate_name}'."
            )
        else:
            self._header_label.setText(
                "Celulas associadas ao laminado atual."
            )
        self.set_cells(cells)

