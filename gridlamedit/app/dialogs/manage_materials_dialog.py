"""Dialog used to manage registered materials."""

from __future__ import annotations

from typing import Iterable, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QGridLayout,
    QHeaderView,
    QLabel,
    QInputDialog,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
)

from gridlamedit.services.material_registry import (
    load_custom_materials,
    remove_custom_material,
    update_custom_material,
)


class ManageMaterialsDialog(QDialog):
    """View and edit the custom material registry."""

    def __init__(self, parent: QWidget | None = None, settings=None) -> None:
        super().__init__(parent)
        self._settings = settings
        self.setWindowTitle("Materiais cadastrados")
        self._table = QTableWidget(self)
        self._btn_edit = QPushButton("Editar", self)
        self._btn_delete = QPushButton("Excluir", self)
        self._btn_close = QPushButton("Fechar", self)
        self._setup_ui()
        self._load_materials()

    def _setup_ui(self) -> None:
        layout = QGridLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        instructions = QLabel(
            "Edite ou remova materiais cadastrados manualmente.", self
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions, 0, 0, 1, 3)

        self._table.setColumnCount(1)
        self._table.setHorizontalHeaderLabels(["Material"])
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SingleSelection)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.verticalHeader().setVisible(False)
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        self._table.itemSelectionChanged.connect(self._update_buttons_state)
        layout.addWidget(self._table, 1, 0, 1, 3)

        self._btn_edit.clicked.connect(self._on_edit)
        self._btn_delete.clicked.connect(self._on_delete)
        self._btn_close.clicked.connect(self.accept)

        layout.addWidget(self._btn_edit, 2, 0)
        layout.addWidget(self._btn_delete, 2, 1)
        layout.addWidget(self._btn_close, 2, 2)

        self.resize(640, 420)

    def _load_materials(self) -> None:
        materials = load_custom_materials(self._settings)
        self._table.setRowCount(len(materials))
        for row, material in enumerate(materials):
            item = QTableWidgetItem(str(material))
            item.setFlags(item.flags() ^ Qt.ItemIsEditable)
            self._table.setItem(row, 0, item)
        self._update_buttons_state()

    def _update_buttons_state(self) -> None:
        has_selection = self._table.currentRow() >= 0
        has_items = self._table.rowCount() > 0
        self._btn_edit.setEnabled(has_selection)
        self._btn_delete.setEnabled(has_selection)
        if not has_items:
            self._btn_edit.setEnabled(False)
            self._btn_delete.setEnabled(False)

    def _selected_material(self) -> str | None:
        row = self._table.currentRow()
        if row < 0:
            return None
        item = self._table.item(row, 0)
        if item is None:
            return None
        text = str(item.text()).strip()
        return text or None

    def _on_edit(self) -> None:
        current = self._selected_material()
        if not current:
            QMessageBox.information(
                self,
                "Editar material",
                "Selecione um material para editar.",
            )
            return
        text, ok = QInputDialog.getText(
            self,
            "Editar material",
            "Atualize o texto do material:",
            text=current,
        )
        if not ok:
            return
        updated_text = str(text or "").strip()
        if not updated_text:
            QMessageBox.warning(
                self,
                "Material vazio",
                "Informe um material válido para atualização.",
            )
            return
        update_custom_material(current, updated_text, settings=self._settings)
        self._load_materials()
        self._select_material(updated_text)

    def _on_delete(self) -> None:
        current = self._selected_material()
        if not current:
            QMessageBox.information(
                self,
                "Excluir material",
                "Selecione um material para excluir.",
            )
            return
        result = QMessageBox.question(
            self,
            "Excluir material",
            "Deseja excluir o material selecionado?",
        )
        if result != QMessageBox.Yes:
            return
        remove_custom_material(current, settings=self._settings)
        self._load_materials()

    def _select_material(self, material: str) -> None:
        target = str(material or "").casefold()
        for row in range(self._table.rowCount()):
            item = self._table.item(row, 0)
            if item and str(item.text()).casefold() == target:
                self._table.setCurrentCell(row, 0)
                break

    def set_materials(self, items: Sequence[str] | Iterable[str]) -> None:
        """Populate the table with a custom list of materials."""
        materials = [str(item).strip() for item in items if str(item).strip()]
        self._table.setRowCount(len(materials))
        for row, material in enumerate(materials):
            item = QTableWidgetItem(material)
            item.setFlags(item.flags() ^ Qt.ItemIsEditable)
            self._table.setItem(row, 0, item)
        self._update_buttons_state()
