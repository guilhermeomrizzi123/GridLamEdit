"""Dialog for quickly creating laminates with cell association."""

from __future__ import annotations

import logging
from typing import Iterable, Optional, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLineEdit,
    QMessageBox,
    QVBoxLayout,
    QWidget,
)

from gridlamedit.io.spreadsheet import GridModel
from gridlamedit.services.laminate_service import (
    LaminateCreationError,
    auto_name_for_layers,
    create_laminate_with_association,
)

logger = logging.getLogger(__name__)


class NewLaminateDialog(QDialog):
    """Dialog used to collect laminate metadata before creation."""

    def __init__(
        self,
        grid_model: GridModel,
        *,
        color_options: Sequence[str],
        type_options: Sequence[str],
        cell_options: Sequence[str],
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Criar novo laminado")
        self.setModal(True)
        self.setObjectName("dlgNewLaminate")
        self._grid_model = grid_model
        self._color_options = [str(option) for option in color_options]
        self._type_options = self._prepare_type_options(type_options)
        self._cell_options = [str(option) for option in cell_options]
        if not self._color_options:
            self._color_options = [str(index) for index in range(1, 151)]
        if not self._type_options:
            self._type_options = ["SS", "Core", "Skin", "Custom"]
        self.created_laminate = None
        self._build_ui()

    def _prepare_type_options(
        self, source: Optional[Sequence[str]]
    ) -> list[str]:
        ordered = ["SS"]
        for option in source or []:
            text = str(option).strip()
            if not text:
                continue
            if text.upper() == "SS":
                continue
            if text not in ordered:
                ordered.append(text)
        return ordered

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        form_layout = QFormLayout()
        form_layout.setLabelAlignment(Qt.AlignRight)
        form_layout.setFormAlignment(Qt.AlignLeft | Qt.AlignTop)
        form_layout.setSpacing(10)

        self.edt_nome = QLineEdit(self)
        self.edt_nome.setObjectName("edtNomeLaminado")
        self.edt_nome.setPlaceholderText("Automatic Rename")
        self.edt_nome.setClearButtonEnabled(False)
        self.edt_nome.setReadOnly(True)
        self.edt_nome.setText("Automatic Rename")
        self.edt_nome.setFocusPolicy(Qt.NoFocus)
        self.edt_nome.setCursor(Qt.ArrowCursor)
        self.edt_nome.setStyleSheet(
            "QLineEdit { color: gray; background-color: #f0f0f0; }"
        )
        form_layout.addRow("Nome:", self.edt_nome)

        self.edt_tag = QLineEdit(self)
        self.edt_tag.setPlaceholderText("Opcional")
        self.edt_tag.textChanged.connect(self._update_auto_name)
        form_layout.addRow("Tag:", self.edt_tag)

        # Automatic Rename option removed from UI; name is always auto-generated.

        self.cmb_cor = QComboBox(self)
        self.cmb_cor.setObjectName("cmbCor")
        self.cmb_cor.addItems(self._color_options or [str(index) for index in range(1, 151)])
        form_layout.addRow("Cor:", self.cmb_cor)

        self.cmb_tipo = QComboBox(self)
        self.cmb_tipo.setObjectName("cmbTipo")
        self.cmb_tipo.addItems(self._type_options or ["SS", "Core", "Skin", "Custom"])
        self.cmb_tipo.setEditable(False)
        form_layout.addRow("Tipo:", self.cmb_tipo)

        self.cmb_celula = QComboBox(self)
        self.cmb_celula.setObjectName("cmbCelula")
        self._populate_cells()
        form_layout.addRow("Célula associada:", self.cmb_celula)

        layout.addLayout(form_layout)

        self.button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, self
        )
        create_button = self.button_box.button(QDialogButtonBox.Ok)
        create_button.setText("Criar")
        create_button.setDefault(True)
        cancel_button = self.button_box.button(QDialogButtonBox.Cancel)
        cancel_button.setText("Cancelar")
        self.button_box.accepted.connect(self._handle_create)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

        self._update_auto_name()
        self.edt_nome.clearFocus()

    def refresh_options(
        self,
        *,
        color_options: Optional[Iterable[str]] = None,
        type_options: Optional[Iterable[str]] = None,
        cell_options: Optional[Iterable[str]] = None,
    ) -> None:
        """Update combo box options without recreating the dialog."""
        if color_options is not None:
            self._color_options = [str(option) for option in color_options]
            self.cmb_cor.clear()
            self.cmb_cor.addItems(self._color_options)
        if type_options is not None:
            self._type_options = self._prepare_type_options(type_options)
            self.cmb_tipo.clear()
            self.cmb_tipo.addItems(self._type_options or ["SS", "Core", "Skin", "Custom"])
        if cell_options is not None:
            self._cell_options = [str(option) for option in cell_options]
            self._populate_cells()

    def reset_fields(self) -> None:
        """Reset user inputs to defaults."""
        self._update_auto_name()
        self.edt_tag.clear()
        if self.cmb_cor.count():
            self.cmb_cor.setCurrentIndex(0)
        if self.cmb_tipo.count():
            self.cmb_tipo.setCurrentIndex(0)
        if self.cmb_celula.count():
            self.cmb_celula.setCurrentIndex(0)
        self.created_laminate = None

    def set_grid_model(self, grid_model: GridModel) -> None:
        """Update the dialog to point to a new GridModel instance."""
        self._grid_model = grid_model
        self._populate_cells()

    def _populate_cells(self) -> None:
        self.cmb_celula.clear()
        if not self._cell_options:
            self.cmb_celula.setEnabled(False)
            return
        self.cmb_celula.setEnabled(True)
        model = self._grid_model
        cell_to_laminate = model.cell_to_laminate if model is not None else {}
        for cell_id in self._cell_options:
            text = cell_id
            mapped = cell_to_laminate.get(cell_id) if cell_to_laminate else None
            if mapped:
                text = f"{cell_id} | {mapped}"
            self.cmb_celula.addItem(text, cell_id)

    def _handle_create(self) -> None:
        # Name is always auto-generated; ignore user input
        auto_name = auto_name_for_layers(
            self._grid_model,
            layer_count=0,
            tag=self.edt_tag.text(),
        )
        nome = (auto_name or "").strip()
        if not nome:
            QMessageBox.warning(
                self,
                "Campos obrigatórios",
                "O nome será gerado automaticamente, mas não foi possível determinar um nome.",
            )
            return
        cor = self.cmb_cor.currentData()
        if cor is None:
            cor = self.cmb_cor.currentText()
        tipo = self.cmb_tipo.currentText().strip()
        if not tipo:
            QMessageBox.warning(
                self, "Campos obrigatórios", "Selecione um tipo de laminado válido."
            )
            return
        cell_data = self.cmb_celula.currentData()
        cell_text = self.cmb_celula.currentText()
        celula = cell_data or cell_text
        celula = str(celula or "").strip()
        if not celula:
            QMessageBox.warning(self, "Campos obrigatórios", "Selecione uma célula para associar.")
            return
        tag_value = self.edt_tag.text()
        try:
            laminado = create_laminate_with_association(
                self._grid_model,
                nome,
                cor or self.cmb_cor.currentText(),
                tipo,
                celula,
                tag=tag_value,
            )
        except LaminateCreationError as exc:
            QMessageBox.warning(self, "Não foi possível criar", str(exc))
            return
        except Exception as exc:  # pragma: no cover - defensive
            logger.exception("Falha inesperada ao criar laminado: %s", exc)
            QMessageBox.critical(
                self,
                "Erro inesperado",
                "Ocorreu um erro ao criar o laminado. Verifique os logs para mais detalhes.",
            )
            return

        laminado.auto_rename_enabled = True
        self.created_laminate = laminado
        self.accept()

    def _update_auto_name(self) -> None:
        # Display-only text indicating automatic naming; field is non-editable
        self.edt_nome.setText("Automatic Rename")
