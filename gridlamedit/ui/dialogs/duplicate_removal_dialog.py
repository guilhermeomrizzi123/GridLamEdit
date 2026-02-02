"""Dialog used to confirm duplicate laminate removals."""

from __future__ import annotations

from typing import Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QLabel,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from gridlamedit.io.spreadsheet import Laminado


class DuplicateRemovalDialog(QDialog):
    """Confirms the removal of duplicate laminates with no cell associations."""

    def __init__(
        self,
        laminates: Sequence[Laminado],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._laminates = [lam for lam in laminates if isinstance(lam, Laminado)]
        self.setWindowTitle("Remover laminados duplicados")
        self.setModal(True)
        self._build_ui()
        self._populate_tree()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        info = QLabel(
            "Os laminados listados abaixo s\u00e3o duplicados id\u00eanticos (incluindo tag) "
            "e sem associa\u00e7\u00f5es."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        self._tree = QTreeWidget(self)
        self._tree.setColumnCount(4)
        self._tree.setHeaderLabels(["Nome", "Tipo", "Tag", "Cor"])
        self._tree.setRootIsDecorated(False)
        self._tree.setSelectionMode(QTreeWidget.NoSelection)
        layout.addWidget(self._tree)

        self.button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, self
        )
        ok_button = self.button_box.button(QDialogButtonBox.Ok)
        ok_button.setText("Remover Laminados")
        cancel_button = self.button_box.button(QDialogButtonBox.Cancel)
        cancel_button.setText("Cancelar")
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

        self.resize(520, 360)

    def _populate_tree(self) -> None:
        self._tree.clear()
        if not self._laminates:
            QTreeWidgetItem(self._tree, ["(nenhum laminado eleg\u00edvel)", "", "", ""])
            self._tree.setEnabled(False)
            return
        self._tree.setEnabled(True)
        # Exibe cada laminado com informacoes essenciais para confer\u00eancia.
        for laminado in self._laminates:
            color_value = getattr(laminado, "color_index", "") or ""
            QTreeWidgetItem(
                self._tree,
                [
                    laminado.nome or "",
                    laminado.tipo or "",
                    str(getattr(laminado, "tag", "") or ""),
                    str(color_value),
                ],
            )


__all__ = ["DuplicateRemovalDialog"]
