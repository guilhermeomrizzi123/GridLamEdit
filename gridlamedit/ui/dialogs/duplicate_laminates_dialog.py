"""Dialog to review and remove identical laminates."""

from __future__ import annotations

from typing import Iterable, Sequence

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from gridlamedit.services.laminate_checks import DuplicateGroup


class DuplicateLaminatesDialog(QDialog):
    """Lists identical laminates and allows deleting a single selection."""

    deleteRequested = Signal(str)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Laminados Duplicados")
        self.setModal(False)
        self._block_item_change = False
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        info = QLabel(
            "Selecione um laminado duplicado (apenas um por vez) e clique em "
            "'Deletar Selecionado'.",
            self,
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        self._tree = QTreeWidget(self)
        self._tree.setHeaderLabels(["Laminado (marcar para remover)"])
        self._tree.setRootIsDecorated(True)
        self._tree.setSelectionMode(QTreeWidget.NoSelection)
        self._tree.itemChanged.connect(self._on_item_changed)
        layout.addWidget(self._tree, stretch=1)

        self._delete_button = QPushButton("Deletar Selecionado", self)
        self._delete_button.clicked.connect(self._on_delete_clicked)

        self._button_box = QDialogButtonBox(QDialogButtonBox.Close, Qt.Horizontal, self)
        self._button_box.rejected.connect(self.close)

        layout.addWidget(self._delete_button)
        layout.addWidget(self._button_box)

        self.resize(520, 420)

    def set_duplicate_groups(self, groups: Sequence[DuplicateGroup]) -> None:
        self._block_item_change = True
        self._tree.clear()

        if not groups:
            placeholder = QTreeWidgetItem(["Nenhum laminado duplicado encontrado."])
            placeholder.setFlags(placeholder.flags() & ~Qt.ItemIsSelectable)
            self._tree.addTopLevelItem(placeholder)
            self._tree.setEnabled(False)
            self._delete_button.setEnabled(False)
            self._block_item_change = False
            return

        self._tree.setEnabled(True)
        self._delete_button.setEnabled(True)

        for idx, group in enumerate(groups, start=1):
            parent = QTreeWidgetItem([f"Grupo {idx}"])
            parent.setFlags(parent.flags() & ~Qt.ItemIsSelectable & ~Qt.ItemIsUserCheckable)
            if group.summary:
                parent.setToolTip(0, group.summary)
            for name in group.laminates:
                item = QTreeWidgetItem(parent, [name])
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                item.setCheckState(0, Qt.Unchecked)
            self._tree.addTopLevelItem(parent)

        self._tree.expandAll()
        self._block_item_change = False

    def selected_laminate_name(self) -> str | None:
        for item in self._iter_checkable_items():
            if item.checkState(0) == Qt.Checked:
                name = item.text(0).strip()
                return name if name else None
        return None

    def _iter_checkable_items(self) -> Iterable[QTreeWidgetItem]:
        top_count = self._tree.topLevelItemCount()
        for i in range(top_count):
            parent = self._tree.topLevelItem(i)
            if parent is None:
                continue
            child_count = parent.childCount()
            for j in range(child_count):
                child = parent.child(j)
                if child is not None and child.flags() & Qt.ItemIsUserCheckable:
                    yield child

    def _on_item_changed(self, item: QTreeWidgetItem, column: int) -> None:
        if self._block_item_change or column != 0:
            return
        if item.checkState(0) != Qt.Checked:
            return
        self._block_item_change = True
        for other in self._iter_checkable_items():
            if other is not item and other.checkState(0) == Qt.Checked:
                other.setCheckState(0, Qt.Unchecked)
        self._block_item_change = False

    def _on_delete_clicked(self) -> None:
        selected = self.selected_laminate_name()
        if not selected:
            QMessageBox.information(
                self,
                "Remover laminado",
                "Selecione um laminado duplicado para remover.",
            )
            return
        self.deleteRequested.emit(selected)


__all__ = ["DuplicateLaminatesDialog"]
