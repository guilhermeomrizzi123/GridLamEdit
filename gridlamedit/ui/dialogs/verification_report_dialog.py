"""Dialog that summarizes laminate verification results."""

from __future__ import annotations

from typing import Iterable, Sequence

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QGroupBox,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QTabWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from gridlamedit.services.laminate_checks import (
    ChecksReport,
    DuplicateGroup,
    SymmetryResult,
)


class VerificationReportDialog(QDialog):
    """Modal dialog that displays laminate verification data before export."""

    removeDuplicatesRequested = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Relatório de Verificação de Laminados")
        self.setModal(True)
        self._report: ChecksReport | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        self._tabs = QTabWidget(self)
        self._tabs.setTabPosition(QTabWidget.North)
        self._build_symmetry_tab()
        self._build_duplicates_tab()
        layout.addWidget(self._tabs)

        self.button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, self
        )
        remove_duplicates_button = QPushButton("Remover Duplicados", self)
        remove_duplicates_button.setObjectName("btnRemoveDuplicatesReport")
        remove_duplicates_button.clicked.connect(self._emit_remove_duplicates_request)
        # Mantem o botao alinhado ao fluxo de exportacao para incentivar a limpeza antes da saida.
        self.button_box.addButton(remove_duplicates_button, QDialogButtonBox.ActionRole)
        export_button = self.button_box.button(QDialogButtonBox.Ok)
        export_button.setText("Exportar")
        export_button.setDefault(True)
        cancel_button = self.button_box.button(QDialogButtonBox.Cancel)
        cancel_button.setText("Cancelar")
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

        self.resize(640, 420)

    def _build_symmetry_tab(self) -> None:
        tab = QWidget(self)
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(12)

        self._symmetric_list = QListWidget(tab)
        self._symmetric_list.setObjectName("lstSymmetric")
        self._symmetric_list.setSelectionMode(QListWidget.NoSelection)
        layout.addWidget(self._wrap_group_box("Simétricos", self._symmetric_list))

        self._asymmetric_list = QListWidget(tab)
        self._asymmetric_list.setObjectName("lstAsymmetric")
        self._asymmetric_list.setSelectionMode(QListWidget.NoSelection)
        layout.addWidget(self._wrap_group_box("Não Simétricos", self._asymmetric_list))

        layout.addStretch(1)
        self._tabs.addTab(tab, "Simetria")

    def _build_duplicates_tab(self) -> None:
        tab = QWidget(self)
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        self._duplicates_tree = QTreeWidget(tab)
        self._duplicates_tree.setHeaderLabel("Grupos de laminados duplicados")
        self._duplicates_tree.setColumnCount(1)
        self._duplicates_tree.setRootIsDecorated(True)
        self._duplicates_tree.setSelectionMode(QTreeWidget.NoSelection)
        layout.addWidget(self._duplicates_tree)

        self._tabs.addTab(tab, "Duplicados")

    @staticmethod
    def _wrap_group_box(title: str, content: QWidget) -> QGroupBox:
        group = QGroupBox(title)
        group_layout = QVBoxLayout(group)
        group_layout.setContentsMargins(8, 12, 8, 8)
        group_layout.addWidget(content)
        return group

    def set_report(self, report: ChecksReport) -> None:
        """Populate the dialog widgets using the provided report."""

        self._report = report
        self._populate_symmetry(report.symmetry)
        self._populate_duplicates(report.duplicates)

    def _populate_symmetry(self, symmetry: SymmetryResult) -> None:
        self._fill_list(self._symmetric_list, symmetry.symmetric)
        self._fill_list(self._asymmetric_list, symmetry.not_symmetric)

    @staticmethod
    def _fill_list(widget: QListWidget, entries: Iterable[str]) -> None:
        widget.clear()
        names = [name for name in entries if name]
        if not names:
            QListWidgetItem("(nenhum laminado)", widget)
            widget.setEnabled(False)
            return
        widget.setEnabled(True)
        for name in names:
            QListWidgetItem(name, widget)

    def _populate_duplicates(self, groups: Sequence[DuplicateGroup]) -> None:
        tree = self._duplicates_tree
        tree.clear()
        if not groups:
            root = QTreeWidgetItem(["Nenhum duplicado identificado"])
            tree.addTopLevelItem(root)
            tree.setEnabled(False)
            tree.expandAll()
            return

        tree.setEnabled(True)
        for idx, group in enumerate(groups, start=1):
            title = f"Grupo{idx}"
            parent = QTreeWidgetItem([title])
            tooltip = group.summary or group.signature
            if tooltip:
                parent.setToolTip(0, tooltip)
            for name in group.laminates:
                QTreeWidgetItem(parent, [name])
            tree.addTopLevelItem(parent)
        tree.expandAll()

    def _emit_remove_duplicates_request(self) -> None:
        """Propaga o clique no botao `Remover Duplicados` para o chamador."""
        self.removeDuplicatesRequested.emit()


__all__ = ["VerificationReportDialog"]
