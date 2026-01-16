"""Dialog that displays laminate reassociation results."""

from __future__ import annotations

from typing import Iterable

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QGroupBox,
    QListWidget,
    QListWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from gridlamedit.services.laminate_reassociation import ReassociationReport


class ReassociationReportDialog(QDialog):
    """Modal dialog that summarizes reassociation results."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Relatorio de Reassociacao")
        self.setModal(True)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        self._tabs = QTabWidget(self)
        self._tabs.setTabPosition(QTabWidget.North)

        self._lst_reassociated = QListWidget(self)
        self._tabs.addTab(
            self._wrap_group_box("Reassociados", self._lst_reassociated),
            "Reassociados",
        )

        self._lst_conflicts = QListWidget(self)
        self._tabs.addTab(
            self._wrap_group_box("Conflitos", self._lst_conflicts),
            "Conflitos",
        )

        self._lst_missing = QListWidget(self)
        self._tabs.addTab(
            self._wrap_group_box("Sem contorno", self._lst_missing),
            "Sem contorno",
        )

        self._lst_not_found = QListWidget(self)
        self._tabs.addTab(
            self._wrap_group_box("Nao encontrados", self._lst_not_found),
            "Nao encontrados",
        )

        self._lst_unmapped = QListWidget(self)
        self._tabs.addTab(
            self._wrap_group_box("Novas celulas sem laminado", self._lst_unmapped),
            "Sem laminado",
        )

        layout.addWidget(self._tabs)

        self._button_box = QDialogButtonBox(QDialogButtonBox.Ok, Qt.Horizontal, self)
        self._button_box.accepted.connect(self.accept)
        layout.addWidget(self._button_box)

        self.resize(720, 480)

    @staticmethod
    def _wrap_group_box(title: str, content: QWidget) -> QGroupBox:
        group = QGroupBox(title)
        group_layout = QVBoxLayout(group)
        group_layout.setContentsMargins(8, 12, 8, 8)
        group_layout.addWidget(content)
        return group

    def set_report(self, report: ReassociationReport) -> None:
        self._populate_reassociated(report)
        self._populate_conflicts(report)
        self._populate_missing(report)
        self._populate_not_found(report)
        self._populate_unmapped(report)

    def _populate_reassociated(self, report: ReassociationReport) -> None:
        self._fill_list(
            self._lst_reassociated,
            (
                f"{entry.laminate}: {entry.old_cell} -> {entry.new_cell}"
                for entry in report.reassociated
            ),
            empty_label="Nenhuma reassociacao realizada.",
        )

    def _populate_conflicts(self, report: ReassociationReport) -> None:
        self._fill_list(
            self._lst_conflicts,
            (
                f"{issue.laminate} ({issue.old_cell}): {issue.details}"
                for issue in report.conflicts
            ),
            empty_label="Nenhum conflito identificado.",
        )

    def _populate_missing(self, report: ReassociationReport) -> None:
        self._fill_list(
            self._lst_missing,
            (
                f"{issue.laminate} ({issue.old_cell}): {issue.details}"
                for issue in report.missing_contours
            ),
            empty_label="Todas as celulas antigas possuam contornos.",
        )

    def _populate_not_found(self, report: ReassociationReport) -> None:
        self._fill_list(
            self._lst_not_found,
            (
                f"{issue.laminate} ({issue.old_cell}): {issue.details}"
                for issue in report.not_found
            ),
            empty_label="Nenhuma celula equivalente perdida.",
        )

    def _populate_unmapped(self, report: ReassociationReport) -> None:
        self._fill_list(
            self._lst_unmapped,
            (f"{cell_id}" for cell_id in report.unmapped_new_cells),
            empty_label="Todas as celulas possuem laminado.",
        )

    @staticmethod
    def _fill_list(widget: QListWidget, entries: Iterable[str], *, empty_label: str) -> None:
        widget.clear()
        entries_list = [entry for entry in entries if entry]
        if not entries_list:
            QListWidgetItem(empty_label, widget)
            widget.setEnabled(False)
            return
        widget.setEnabled(True)
        for entry in entries_list:
            QListWidgetItem(entry, widget)


__all__ = ["ReassociationReportDialog"]
