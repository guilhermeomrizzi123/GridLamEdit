"""Dialog for comparing two laminates."""

from __future__ import annotations

import re
import unicodedata
from typing import Iterable

from PySide6.QtCore import QSortFilterProxyModel, Qt, Signal
from PySide6.QtGui import QStandardItem, QStandardItemModel, QTextCursor
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


class CompareLaminatesDialog(QDialog):
    """Dialog used to select two laminates and show comparison report."""

    compare_requested = Signal(str, str)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Comparar Laminados")
        self.setModal(False)
        self.resize(720, 520)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        selection_layout = QGridLayout()
        selection_layout.setHorizontalSpacing(12)
        selection_layout.setVerticalSpacing(8)

        label_a = QLabel("Laminado A:", self)
        label_b = QLabel("Laminado B:", self)
        self.laminate_a_combo = QComboBox(self)
        self.laminate_b_combo = QComboBox(self)
        self.laminate_a_combo.setMinimumWidth(240)
        self.laminate_b_combo.setMinimumWidth(240)
        self.laminate_a_combo.setEditable(True)
        self.laminate_b_combo.setEditable(True)
        self.laminate_a_combo.setInsertPolicy(QComboBox.NoInsert)
        self.laminate_b_combo.setInsertPolicy(QComboBox.NoInsert)
        self.laminate_a_combo.setCompleter(None)
        self.laminate_b_combo.setCompleter(None)
        self._init_filter_combo(
            self.laminate_a_combo,
            "_laminate_a_source_model",
            "_laminate_a_proxy",
            self._on_filter_a_changed,
            self._select_first_visible_a,
        )
        self._init_filter_combo(
            self.laminate_b_combo,
            "_laminate_b_source_model",
            "_laminate_b_proxy",
            self._on_filter_b_changed,
            self._select_first_visible_b,
        )

        selection_layout.addWidget(label_a, 0, 0)
        selection_layout.addWidget(self.laminate_a_combo, 0, 1)
        selection_layout.addWidget(label_b, 1, 0)
        selection_layout.addWidget(self.laminate_b_combo, 1, 1)

        layout.addLayout(selection_layout)

        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(8)
        self.compare_button = QPushButton("Comparar", self)
        self.close_button = QPushButton("Fechar", self)
        self.compare_button.clicked.connect(self._on_compare_clicked)
        self.close_button.clicked.connect(self.close)

        buttons_layout.addWidget(self.compare_button)
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.close_button)

        layout.addLayout(buttons_layout)

        self.report_view = QTextEdit(self)
        self.report_view.setReadOnly(True)
        self.report_view.setLineWrapMode(QTextEdit.NoWrap)
        self.report_view.setMinimumHeight(240)
        layout.addWidget(self.report_view, stretch=1)

    def set_laminate_names(
        self,
        names: Iterable[str],
        *,
        select_a: str | None = None,
        select_b: str | None = None,
    ) -> None:
        name_list = [str(name) for name in names]
        self.laminate_a_combo.blockSignals(True)
        self.laminate_b_combo.blockSignals(True)
        self._populate_combo_model(self._laminate_a_source_model, name_list)
        self._populate_combo_model(self._laminate_b_source_model, name_list)
        self._reset_filter(self.laminate_a_combo, self._laminate_a_proxy, clear_text=True)
        self._reset_filter(self.laminate_b_combo, self._laminate_b_proxy, clear_text=True)

        if select_a in name_list:
            self._set_combo_selection(self.laminate_a_combo, self._laminate_a_proxy, name_list, select_a)
        else:
            self.laminate_a_combo.setCurrentIndex(-1)

        if select_b in name_list:
            self._set_combo_selection(self.laminate_b_combo, self._laminate_b_proxy, name_list, select_b)
        else:
            self.laminate_b_combo.setCurrentIndex(-1)

        self.laminate_a_combo.blockSignals(False)
        self.laminate_b_combo.blockSignals(False)

    def selected_names(self) -> tuple[str, str]:
        return self.laminate_a_combo.currentText().strip(), self.laminate_b_combo.currentText().strip()

    def set_report(self, text: str) -> None:
        self.report_view.setPlainText(text)
        self.report_view.moveCursor(QTextCursor.Start)

    def _on_compare_clicked(self) -> None:
        name_a, name_b = self.selected_names()
        self.compare_requested.emit(name_a, name_b)

    def _init_filter_combo(
        self,
        combo: QComboBox,
        source_attr: str,
        proxy_attr: str,
        filter_handler,
        select_first_handler,
    ) -> None:
        source_model = QStandardItemModel(combo)
        proxy_model = LaminateFilterProxy(combo)
        proxy_model.setSourceModel(source_model)
        proxy_model.setFilterKeyColumn(0)
        combo.setModel(proxy_model)
        combo.setModelColumn(0)

        line_edit = combo.lineEdit()
        if line_edit is not None:
            line_edit.setPlaceholderText("Digite Buscar Laminado")
            line_edit.setCompleter(None)
            line_edit.textEdited.connect(filter_handler)
            line_edit.returnPressed.connect(select_first_handler)

        setattr(self, source_attr, source_model)
        setattr(self, proxy_attr, proxy_model)
        self._reset_filter(combo, proxy_model, clear_text=True)

    def _populate_combo_model(self, model: QStandardItemModel, names: list[str]) -> None:
        model.clear()
        for name in names:
            model.appendRow(QStandardItem(name))

    def _reset_filter(self, combo: QComboBox, proxy: "LaminateFilterProxy", *, clear_text: bool) -> None:
        proxy.set_filter_text("")
        if clear_text:
            line_edit = combo.lineEdit()
            if line_edit is not None:
                line_edit.blockSignals(True)
                line_edit.clear()
                line_edit.blockSignals(False)

    def _on_filter_a_changed(self, text: str) -> None:
        self._apply_filter(self.laminate_a_combo, self._laminate_a_proxy, text)

    def _on_filter_b_changed(self, text: str) -> None:
        self._apply_filter(self.laminate_b_combo, self._laminate_b_proxy, text)

    def _apply_filter(self, combo: QComboBox, proxy: "LaminateFilterProxy", text: str) -> None:
        proxy.set_filter_text(text)
        if text.strip():
            view = combo.view()
            if view is not None and not view.isVisible():
                combo.showPopup()

    def _select_first_visible_a(self) -> None:
        self._select_first_visible(self.laminate_a_combo, self._laminate_a_proxy)

    def _select_first_visible_b(self) -> None:
        self._select_first_visible(self.laminate_b_combo, self._laminate_b_proxy)

    def _select_first_visible(self, combo: QComboBox, proxy: "LaminateFilterProxy") -> None:
        for row in range(proxy.rowCount()):
            idx = proxy.index(row, 0)
            text = str(idx.data() or "")
            if text:
                combo.blockSignals(True)
                combo.setCurrentIndex(row)
                combo.blockSignals(False)
                break

    def _set_combo_selection(
        self,
        combo: QComboBox,
        proxy: "LaminateFilterProxy",
        names: list[str],
        target: str,
    ) -> None:
        if target not in names:
            return
        self._reset_filter(combo, proxy, clear_text=True)
        source = proxy.sourceModel()
        if not isinstance(source, QStandardItemModel):
            return
        match_row = None
        for row in range(source.rowCount()):
            idx = source.index(row, 0)
            if str(idx.data() or "") == target:
                match_row = row
                break
        if match_row is None:
            return
        proxy_idx = proxy.mapFromSource(source.index(match_row, 0))
        if proxy_idx.isValid():
            combo.blockSignals(True)
            combo.setCurrentIndex(proxy_idx.row())
            combo.blockSignals(False)


class LaminateFilterProxy(QSortFilterProxyModel):
    """Case-insensitive filter that matches raw or normalized text."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._filter_text: str = ""
        self._filter_norm: str = ""
        self.setFilterCaseSensitivity(Qt.CaseInsensitive)

    @staticmethod
    def _normalize(text: str) -> str:
        base = unicodedata.normalize("NFKD", text)
        stripped = "".join(ch for ch in base if not unicodedata.combining(ch))
        return re.sub(r"[^0-9A-Za-z]+", "", stripped).lower()

    def set_filter_text(self, text: str) -> None:
        self._filter_text = text.strip()
        self._filter_norm = self._normalize(self._filter_text) if self._filter_text else ""
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row: int, source_parent) -> bool:  # type: ignore[override]
        if not self._filter_text:
            return True
        index = self.sourceModel().index(source_row, self.filterKeyColumn(), source_parent)
        raw_text = str(index.data() or "")

        if self._filter_text.lower() in raw_text.lower():
            return True

        if self._filter_norm:
            norm_text = self._normalize(raw_text)
            if self._filter_norm in norm_text:
                return True

        return False
