"""Spreadsheet import pipeline and UI binding for GridLamEdit."""

from __future__ import annotations

import logging
import math
import re
import unicodedata
import zipfile
from collections import OrderedDict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Protocol

import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QHeaderView,
    QListWidget,
    QListWidgetItem,
    QSizePolicy,
    QTableView,
)

from gridlamedit.app.delegates import MaterialComboDelegate, OrientationComboDelegate

_OPEN_EXCEL_EXCEPTIONS: tuple[type[BaseException], ...] = (zipfile.BadZipFile,)
try:  # pragma: no cover - dependente de openpyxl
    from openpyxl.utils.exceptions import InvalidFileException
except Exception:  # pragma: no cover - fallback quando openpyxl nAo disponAvel
    InvalidFileException = None  # type: ignore[assignment]
else:  # pragma: no cover - depende de openpyxl
    _OPEN_EXCEL_EXCEPTIONS = _OPEN_EXCEL_EXCEPTIONS + (InvalidFileException,)  # type: ignore[assignment]

try:  # pragma: no cover - dependente de xlrd
    import xlrd  # type: ignore
except Exception:  # pragma: no cover - fallback quando xlrd nAo disponAvel
    xlrd = None  # type: ignore[assignment]

logger = logging.getLogger(__name__)

ALLOWED_ANGLES = {-90, -45, 0, 45, 90}

LAMINATE_ALIASES = ("Laminate", "Laminate Name", "Laminado", "Nome")
COLOR_ALIASES = ("Color", "Colour", "Cor", "ColorIdx", "Color Index")
TYPE_ALIASES = ("Type", "Tipo")
MATERIAL_ALIASES = ("Material",)
ORIENTATION_ALIASES = ("Orientation", "Orientacao", "Orientacao", "Angle", "Angulo", "Angulo")
ACTIVE_ALIASES = ("Active", "Ativo", "Status")
SYMMETRY_ALIASES = ("Symmetry", "Simetria")
INDEX_ALIASES = ("Index", "#", "Idx", "Ordem", "SequAancia", "Sequencia")

CELL_MAPPING_ALIASES = ("Cell", "Cells", "Celula", "CAlula", "C")
CELL_ID_PATTERN = re.compile(r"^C\d+$", re.IGNORECASE)
NO_LAMINATE_LABEL = "(sem laminado)"

@dataclass
class Camada:
    """Representa uma camada do laminado."""

    idx: int
    material: str
    orientacao: int
    ativo: bool
    simetria: bool


@dataclass
class Laminado:
    """Agregado de metadados e camadas de um laminado."""

    nome: str
    cor_hex: str
    tipo: str
    celulas: list[str] = field(default_factory=list)
    camadas: list[Camada] = field(default_factory=list)


@dataclass
class GridModel:
    """Modelo raiz carregado da planilha do Grid Design."""

    laminados: Dict[str, Laminado] = field(default_factory=OrderedDict)
    celulas_ordenadas: list[str] = field(default_factory=list)
    cell_to_laminate: Dict[str, str] = field(default_factory=dict)
    source_excel_path: Optional[str] = None
    dirty: bool = False

    def mark_dirty(self, value: bool = True) -> None:
        self.dirty = value

    def laminados_da_celula(self, cell_id: str) -> list[Laminado]:
        """Retorna laminados associados a uma celula."""
        mapped_name = self.cell_to_laminate.get(cell_id)
        if mapped_name:
            laminado = self.laminados.get(mapped_name)
            return [laminado] if laminado is not None else []
        return [
            laminado
            for laminado in self.laminados.values()
            if cell_id in laminado.celulas
        ]


def normalize_angle(value: object) -> int:
    """Normaliza a orientacao para inteiro dentro do conjunto permitido."""
    if value is None:
        raise ValueError("orientacao ausente")

    if isinstance(value, bool):
        raise ValueError(f"valor booleano invalido para orientacao: {value!r}")

    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            raise ValueError("orientacao ausente")
        angle = int(round(float(value)))
    else:
        text = str(value).strip()
        if not text:
            raise ValueError("orientacao ausente")
        cleaned = re.sub(r"[^\d\-]+", "", text)
        if not cleaned:
            raise ValueError(f"orientacao invalida: {value!r}")
        angle = int(cleaned)

    if angle not in ALLOWED_ANGLES:
        raise ValueError(
            f"orientacao {angle} fora do conjunto permitido {sorted(ALLOWED_ANGLES)}"
        )
    return angle


def normalize_bool(value: object) -> bool:
    """Normaliza textos 'Sim/NAo' ou 'Yes/No' (e similares) em booleano."""
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return False
        return float(value) != 0.0

    text = str(value).strip()
    if not text:
        return False

    normalized = _normalize_header(text)
    true_values = {"yes", "y", "true", "1", "sim", "s"}
    false_values = {"no", "n", "false", "0", "nao"}

    if normalized in true_values:
        return True
    if normalized in false_values:
        return False

    logger.warning("Valor booleano desconhecido '%s'; assumindo False.", value)
    return False


def normalize_color(value: object, default: str = "#FFFFFF") -> str:
    """Normaliza cores em #RRGGBB, aceitando nomes padrao."""
    if value is None:
        return default
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return default
        value = str(value)

    text = str(value).strip()
    if not text:
        return default

    if re.fullmatch(r"#?[0-9a-fA-F]{6}", text):
        formatted = text.upper()
        return formatted if formatted.startswith("#") else f"#{formatted}"

    color = QColor(text)
    if color.isValid():
        return color.name().upper()

    logger.warning("Cor invalida '%s'; usando %s.", value, default)
    return default


def load_grid_spreadsheet(path: str) -> GridModel:
    """Carrega a planilha do Grid Design em um GridModel."""
    file_path = Path(path)
    if not file_path.exists():
        raise ValueError(f"Arquivo '{path}' nAo encontrado.")

    ext = file_path.suffix.lower()
    if ext not in {".xls", ".xlsx"}:
        raise ValueError("Formato de planilha nAo suportado (use .xls ou .xlsx).")

    workbook = _open_workbook(file_path, ext)
    if "Planilha1" not in workbook.sheet_names:
        logger.error("Aba Planilha1 ausente.")
        raise ValueError(
            "A planilha deve conter a aba 'Planilha1' no formato exportado pelo Grid Design."
        )

    logger.info("Aba Planilha1 localizada - iniciando leitura.")
    df = _parse_sheet(workbook, "Planilha1")
    df = df.dropna(how="all").reset_index(drop=True)
    if df.empty:
        raise ValueError("Planilha1 nAo contAm dados para importar.")

    # Extrai celulas e mapeamento utilizando a nova funcao pAoblica.
    celulas_ordenadas = parse_cells_from_planilha1(df)
    cells_info = _extract_cells_section(df)
    cell_to_laminate = cells_info.mapping
    separator_idx = cells_info.separator_idx
    config_section = df.iloc[separator_idx + 1 :]

    if config_section.empty:
        raise ValueError("Planilha1: secao de configuracao de laminados estA vazia.")

    laminados = _parse_configuration_section(config_section)

    for cell_id, laminate_name in cell_to_laminate.items():
        laminado = laminados.get(laminate_name)
        if laminado is None:
            logger.warning(
                "CAlula '%s' associa laminado '%s', que nAo estA definido na secao de configuracao.",
                cell_id,
                laminate_name,
            )
            continue
        if cell_id not in laminado.celulas:
            laminado.celulas.append(cell_id)

    total_camadas = sum(len(laminado.camadas) for laminado in laminados.values())
    logger.info("Celulas importadas: %s", ", ".join(celulas_ordenadas))
    logger.info(
        "Planilha carregada com %d laminados, %d camadas e %d celulas.",
        len(laminados),
        total_camadas,
        len(celulas_ordenadas),
    )

    model = GridModel(
        laminados=laminados,
        celulas_ordenadas=celulas_ordenadas,
        cell_to_laminate=dict(cell_to_laminate),
    )
    model.source_excel_path = str(file_path)
    return model


def save_grid_spreadsheet(path: str, model: GridModel) -> None:
    """Persistir o GridModel no layout Planilha1 esperado pelo import."""
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    rows: list[list[object]] = []

    # Secao de celulas.
    rows.append(["Cells", "Laminate"])
    for cell_id in model.celulas_ordenadas:
        rows.append([cell_id, model.cell_to_laminate.get(cell_id, "")])
    rows.append(["#"])

    # Secao de configuracao por laminado.
    laminados_iter = model.laminados.values()
    for laminado in laminados_iter:
        rows.append(["Name", laminado.nome])
        rows.append(["ColorIdx", laminado.cor_hex or "#FFFFFF"])
        rows.append(["Type", laminado.tipo or ""])
        rows.append(["Stacking"])

        for camada in laminado.camadas:
            rows.append(
                [
                    camada.material,
                    camada.orientacao,
                    "Sim" if camada.ativo else "NAo",
                    "Sim" if camada.simetria else "NAo",
                ]
            )
        rows.append(["#"])

    if output_path.suffix.lower() == ".xls":
        try:
            import xlwt  # type: ignore
        except ImportError as exc:  # pragma: no cover - dependAancia opcional
            raise ValueError(
                "Salvar em '.xls' requer a dependAancia 'xlwt'. "
                "Instale 'xlwt' ou salve o arquivo com extensAo '.xlsx'."
            ) from exc

        book = xlwt.Workbook()
        sheet = book.add_sheet("Planilha1")
        for r_idx, row in enumerate(rows):
            for c_idx, value in enumerate(row):
                sheet.write(r_idx, c_idx, value)
        book.save(str(output_path))
        return

    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Planilha1"
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    wb.save(output_path)


class StackingTableModel(QAbstractTableModel):
    """Apresenta as camadas do laminado na tabela de stacking."""

    headers = ["Material", "Orientacao"]

    def __init__(
        self,
        camadas: list[Camada] | None = None,
        change_callback: Optional[Callable[[list[Camada]], None]] = None,
    ) -> None:
        super().__init__()
        self._camadas: list[Camada] = list(camadas or [])
        self._change_callback = change_callback

    def update_layers(self, camadas: list[Camada]) -> None:
        self.beginResetModel()
        self._camadas = list(camadas)
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        if parent.isValid():
            return 0
        return len(self._camadas)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        if parent.isValid():
            return 0
        return len(self.headers)

    def headerData(  # noqa: N802
        self,
        section: int,
        orientation: Qt.Orientation,
        role: int = Qt.DisplayRole,
    ) -> Optional[str]:
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal and 0 <= section < len(self.headers):
            return self.headers[section]
        if orientation == Qt.Vertical:
            return str(section + 1)
        return None

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Optional[str]:  # noqa: N802
        if not index.isValid() or not (0 <= index.row() < len(self._camadas)):
            return None

        camada = self._camadas[index.row()]

        if role == Qt.DisplayRole:
            coluna = index.column()
            if coluna == 0:
                return camada.material
            if coluna == 1:
                return f"{camada.orientacao}\N{DEGREE SIGN}"
        elif role == Qt.TextAlignmentRole:
            if index.column() == 0:
                return int(Qt.AlignVCenter | Qt.AlignLeft)
            return int(Qt.AlignVCenter | Qt.AlignCenter)
        elif role == Qt.EditRole:
            if index.column() == 0:
                return camada.material
            if index.column() == 1:
                return f"{camada.orientacao}\N{DEGREE SIGN}"

        return None

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:  # noqa: N802
        if not index.isValid():
            return Qt.NoItemFlags
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable

    def setData(
        self,
        index: QModelIndex,
        value: object,
        role: int = Qt.EditRole,
    ) -> bool:  # noqa: N802
        if not index.isValid() or role not in (Qt.EditRole, Qt.DisplayRole):
            return False
        camada = self._camadas[index.row()]
        column = index.column()
        if column == 0:
            new_value = str(value).strip()
            if new_value == camada.material:
                return False
            camada.material = new_value
        elif column == 1:
            text = str(value).strip()
            cleaned = text.replace("\N{DEGREE SIGN}", "").replace("º", "").strip()
            try:
                angle = normalize_angle(cleaned)
            except ValueError:
                return False
            if angle == camada.orientacao:
                return False
            camada.orientacao = angle
        else:
            return False
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        if self._change_callback:
            self._change_callback(self._camadas)
        return True


def bind_model_to_ui(model: GridModel, ui) -> None:
    """Efetua o binding do modelo carregado com os widgets da UI."""
    binding = getattr(ui, "_grid_binding", None)
    if binding is not None:
        binding.teardown()
    ui._grid_binding = _GridUiBinding(model, ui)  # type: ignore[attr-defined]


def bind_cells_to_ui(model: GridModel, ui) -> None:
    """
    Popula o listbox de celulas com base em ``model.celulas_ordenadas``.
    """
    list_widget = getattr(ui, "lstCelulas", None)
    if not isinstance(list_widget, QListWidget):
        list_widget = getattr(ui, "cells_list", None)
    if not isinstance(list_widget, QListWidget):
        return

    list_widget.blockSignals(True)
    list_widget.clear()
    for cell_id in model.celulas_ordenadas:
        item = QListWidgetItem(_format_cell_label(model, cell_id))
        item.setData(Qt.UserRole, cell_id)
        list_widget.addItem(item)
    list_widget.blockSignals(False)

    if model.celulas_ordenadas:
        list_widget.setCurrentRow(0)


def _format_cell_label(model: GridModel, cell_id: str) -> str:
    laminate_name = model.cell_to_laminate.get(cell_id) or ""
    if laminate_name:
        laminate = model.laminados.get(laminate_name)
        laminate_name = laminate.nome if laminate is not None else laminate_name
    else:
        laminados = model.laminados_da_celula(cell_id)
        laminate_name = laminados[0].nome if laminados else NO_LAMINATE_LABEL
    return f"{cell_id} | {laminate_name}"


# Helpers ------------------------------------------------------------------ #


class _WorkbookProtocol(Protocol):
    sheet_names: list[str]

    def parse(self, sheet_name: str, *args, **kwargs) -> pd.DataFrame:
        ...


def _open_workbook(file_path: Path, ext: str) -> _WorkbookProtocol:
    if ext == ".xls":
        if xlrd is None:
            raise ValueError(
                "Leitura de arquivos .xls requer a dependAancia 'xlrd==1.2.0'."
            )
        return _XlrdWorkbook(file_path)

    try:
        workbook = pd.ExcelFile(file_path, engine="openpyxl")
    except _OPEN_EXCEL_EXCEPTIONS as exc:
        return _open_with_xlrd(file_path, exc)
    except ValueError as exc:
        if "not a zip file" in str(exc).lower():
            return _open_with_xlrd(file_path, exc)
        raise ValueError(f"NAo foi possAvel abrir '{file_path}': {exc}") from exc
    except Exception as exc:  # pragma: no cover - protecao
        raise ValueError(f"NAo foi possAvel abrir '{file_path}': {exc}") from exc
    return workbook  # type: ignore[return-value]


def _parse_sheet(workbook: _WorkbookProtocol, sheet_name: str) -> pd.DataFrame:
    parse = getattr(workbook, "parse")
    try:
        return parse(sheet_name, header=None, dtype=object)
    except TypeError:
        return parse(sheet_name)


def _normalize_header(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_only = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    ascii_only = ascii_only.lower()
    return re.sub(r"[^a-z0-9]+", "", ascii_only)


def _find_separator_row(df: pd.DataFrame) -> Optional[int]:
    for idx, row in df.iterrows():
        for value in row:
            if _is_blank(value):
                continue
            if _is_separator_token(value):
                return idx
            break
    return None


def _is_separator_token(value: object) -> bool:
    text = str(value).strip().lower()
    text = text.replace(" ", "")
    return text == "#"


def _find_cells_header_row(df: pd.DataFrame, separator_idx: int) -> Optional[int]:
    normalized_aliases = {_normalize_header(alias) for alias in CELL_MAPPING_ALIASES}
    for idx in range(separator_idx):
        row = df.iloc[idx]
        found_cells = False
        for value in row:
            if _is_blank(value):
                continue
            normalized = _normalize_header(str(value))
            if normalized in normalized_aliases:
                found_cells = True
                break
        if found_cells:
            return idx
    return None


def _first_non_blank_value(row: pd.Series) -> Optional[str]:
    for value in row:
        if _is_blank(value):
            continue
        return str(value)
    return None


def _extract_cells_section(df: pd.DataFrame) -> _CellsSection:
    """Localiza a secao de celulas em Planilha1 e retorna dados estruturados."""
    separator_idx = _find_separator_row(df)
    if separator_idx is None:
        raise ValueError(
            "NAo foi possAvel localizar a linha separadora '#' em Planilha1."
        )

    cells_header_idx = _find_cells_header_row(df, separator_idx)
    if cells_header_idx is None:
        raise ValueError("NAo foi possAvel localizar a secao 'Cells' em Planilha1.")

    header_row = df.iloc[cells_header_idx]
    header_keys = [
        _normalize_header(str(value)) if not _is_blank(value) else ""
        for value in header_row
    ]

    cell_col = _find_column(header_keys, CELL_MAPPING_ALIASES)
    laminate_col = _find_column(header_keys, LAMINATE_ALIASES)
    data_start = cells_header_idx + 1

    if cell_col is None:
        cell_col = 0

    cells_ordered: list[str] = []
    seen_cells: set[str] = set()
    cell_to_laminate: Dict[str, str] = {}

    for row_idx in range(data_start, separator_idx):
        row = df.iloc[row_idx]
        cell_value = row.iloc[cell_col] if cell_col < len(row) else None
        if _is_blank(cell_value):
            continue
        cell_id = str(cell_value).strip().upper()
        if not CELL_ID_PATTERN.match(cell_id):
            logger.warning(
                "Linha %d da secao 'Cells': celula '%s' invalida, ignorada.",
                row_idx + 1,
                cell_value,
            )
            continue
        if cell_id not in seen_cells:
            cells_ordered.append(cell_id)
            seen_cells.add(cell_id)

        if laminate_col is not None and laminate_col < len(row):
            laminate_value = row.iloc[laminate_col]
            if not _is_blank(laminate_value) and cell_id not in cell_to_laminate:
                cell_to_laminate[cell_id] = str(laminate_value).strip()

    if not cells_ordered:
        raise ValueError("NAo hA celulas vAlidas entre 'Cells' e '#' em Planilha1.")

    logger.info(
        "Secao 'Cells' processada com %d celulas distintas.",
        len(cells_ordered),
    )
    return _CellsSection(cells=cells_ordered, mapping=cell_to_laminate, separator_idx=separator_idx)


def _parse_configuration_section(
    df: pd.DataFrame,
) -> OrderedDict[str, Laminado]:
    df = df.reset_index(drop=True)
    laminados: OrderedDict[str, Laminado] = OrderedDict()

    normalized_color_aliases = {_normalize_header(alias) for alias in COLOR_ALIASES}
    normalized_type_aliases = {_normalize_header(alias) for alias in TYPE_ALIASES}

    idx = 0
    while idx < len(df):
        row = df.iloc[idx]
        label_raw = row.iloc[0] if len(row) > 0 else None
        if _is_blank(label_raw):
            idx += 1
            continue

        if _is_separator_token(label_raw):
            idx += 1
            continue

        label = str(label_raw).strip()
        normalized_label = _normalize_header(label)
        if normalized_label != "name":
            idx += 1
            continue

        name_value = row.iloc[1] if len(row) > 1 else None
        if _is_blank(name_value):
            raise ValueError(f"Linha {idx + 1}: laminado sem nome definido.")
        laminate_name = str(name_value).strip()
        idx += 1

        color = "#FFFFFF"
        laminate_type = ""

        while idx < len(df):
            row = df.iloc[idx]
            first_val = row.iloc[0] if len(row) > 0 else None

            if _is_blank(first_val):
                idx += 1
                continue

            if _is_separator_token(first_val):
                break

            normalized = _normalize_header(str(first_val))
            if normalized == "name":
                break

            if normalized == "coloridx":
                color = "#FFFFFF"
                idx += 1
                continue
            if normalized in normalized_color_aliases:
                color = normalize_color(row.iloc[1] if len(row) > 1 else None)
                idx += 1
                continue

            if normalized in normalized_type_aliases:
                laminate_type = _string_or_empty(row.iloc[1] if len(row) > 1 else "")
                idx += 1
                continue

            if normalized == "stacking":
                idx += 1
                break

            logger.warning(
                "Linha %d: rA3tulo '%s' desconhecido na configuracao de laminados; ignorando.",
                idx + 1,
                first_val,
            )
            idx += 1

        layers: list[Camada] = []
        while idx < len(df):
            row = df.iloc[idx]
            first_val = row.iloc[0] if len(row) > 0 else None

            if not _is_blank(first_val) and (
                _is_separator_token(first_val)
                or _normalize_header(str(first_val)) == "name"
            ):
                break

            if _is_blank(first_val):
                idx += 1
                continue

            material = str(first_val).strip()
            orientation_raw = row.iloc[1] if len(row) > 1 else 0
            try:
                orientation = normalize_angle(orientation_raw)
            except ValueError as exc:
                raise ValueError(
                    f"Linha {idx + 1}: laminado '{laminate_name}' possui orientacao invalida ({exc})."
                ) from exc

            active_raw = row.iloc[2] if len(row) > 2 else None
            symmetry_raw = row.iloc[3] if len(row) > 3 else None
            active = normalize_bool(active_raw) if not _is_blank(active_raw) else True
            symmetry = (
                normalize_bool(symmetry_raw) if not _is_blank(symmetry_raw) else False
            )

            layers.append(
                Camada(
                    idx=len(layers),
                    material=material,
                    orientacao=orientation,
                    ativo=active,
                    simetria=symmetry,
                )
            )
            idx += 1

        if not layers:
            logger.warning(
                "Laminado '%s' nAo possui camadas definidas na secao de stacking.",
                laminate_name,
            )

        laminados[laminate_name] = Laminado(
            nome=laminate_name,
            cor_hex=color,
            tipo=laminate_type,
            celulas=[],
            camadas=layers,
        )

    if not laminados:
        raise ValueError(
            "Planilha1: secao de configuracao nAo definiu nenhum laminado."
        )

    return laminados


def _find_column(header_keys: list[str], aliases: Iterable[str]) -> Optional[int]:
    normalized_aliases = {_normalize_header(alias) for alias in aliases}
    for idx, key in enumerate(header_keys):
        if key in normalized_aliases:
            return idx
    return None


def _require_column(
    header_keys: list[str], aliases: Iterable[str], display_name: str
) -> int:
    column = _find_column(header_keys, aliases)
    if column is None:
        raise ValueError(
            f"Coluna obrigatA3ria '{display_name}' ausente em Planilha1 (secao de configuracao abaixo de '#')."
        )
    return column


class _GridUiBinding:
    """Encapsula o binding entre GridModel e widgets do MainWindow."""

    def __init__(self, model: GridModel, ui) -> None:
        self.model = model
        self.ui = ui
        self._updating = False
        self._current_laminate: Optional[str] = None
        self._current_cell_id: Optional[str] = None
        self._laminates_by_cell: dict[str, list[str]] = self._build_cell_index()
        self.stacking_model = StackingTableModel(change_callback=self._on_layers_modified)
        self._cells_widget: Optional[QListWidget] = None

        self._setup_widgets()
        self._connect_signals()

    def teardown(self) -> None:
        """Remove conexAes quando um novo binding for aplicado."""
        try:
            cells_widget = getattr(self.ui, "lstCelulas", None)
            if cells_widget is None:
                cells_widget = getattr(self.ui, "cells_list", None)
            if isinstance(cells_widget, QListWidget):
                cells_widget.currentItemChanged.disconnect(self._on_cell_item_changed)
        except (AttributeError, TypeError):
            pass
        try:
            self.ui.laminate_name_combo.currentTextChanged.disconnect(self._on_laminate_selected)
        except (AttributeError, TypeError):
            pass

    # Internal helpers -------------------------------------------------- #

    def _build_cell_index(self) -> dict[str, list[str]]:
        if self.model.cell_to_laminate:
            mapping: dict[str, list[str]] = {}
            for cell, lam in self.model.cell_to_laminate.items():
                if lam:
                    mapping.setdefault(cell, []).append(lam)
            return mapping
        mapping: dict[str, list[str]] = {}
        for name, laminado in self.model.laminados.items():
            for cell_id in laminado.celulas:
                mapping.setdefault(cell_id, []).append(name)
        return mapping

    def _setup_widgets(self) -> None:
        name_combo = getattr(self.ui, "laminate_name_combo", None)
        if isinstance(name_combo, QComboBox):
            name_combo.blockSignals(True)
            name_combo.clear()
            for laminado in self.model.laminados.values():
                name_combo.addItem(laminado.nome)
            name_combo.blockSignals(False)

        color_combo = getattr(self.ui, "laminate_color_combo", None)
        if isinstance(color_combo, QComboBox):
            color_combo.blockSignals(True)
            unique_colors = list(
                OrderedDict(
                    (laminado.cor_hex, None) for laminado in self.model.laminados.values()
                ).keys()
            )
            for color in unique_colors:
                if color not in [color_combo.itemText(i) for i in range(color_combo.count())]:
                    color_combo.addItem(color)
            color_combo.blockSignals(False)

        type_combo = getattr(self.ui, "laminate_type_combo", None)
        if isinstance(type_combo, QComboBox):
            type_combo.blockSignals(True)
            unique_types = [
                laminado.tipo for laminado in self.model.laminados.values() if laminado.tipo
            ]
            for tipo in OrderedDict((item, None) for item in unique_types):
                if tipo not in [type_combo.itemText(i) for i in range(type_combo.count())]:
                    type_combo.addItem(tipo)
            type_combo.blockSignals(False)

        table = getattr(self.ui, "layers_table", None)
        if isinstance(table, QTableView):
            table.setModel(self.stacking_model)
            table.setEditTriggers(
                QAbstractItemView.DoubleClicked
                | QAbstractItemView.SelectedClicked
                | QAbstractItemView.EditKeyPressed
            )
            table.setSelectionBehavior(QAbstractItemView.SelectItems)
            table.setSelectionMode(QAbstractItemView.SingleSelection)
            table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            header = table.horizontalHeader()
            header.setSectionResizeMode(QHeaderView.Stretch)
            table.verticalHeader().setVisible(False)
            self._install_delegates(table)

        cells_widget = getattr(self.ui, "lstCelulas", None)
        if not isinstance(cells_widget, QListWidget):
            cells_widget = getattr(self.ui, "cells_list", None)
        if isinstance(cells_widget, QListWidget):
            self._cells_widget = cells_widget

    def _connect_signals(self) -> None:
        cells_widget = self._cells_widget
        if isinstance(cells_widget, QListWidget):
            cells_widget.currentItemChanged.connect(self._on_cell_item_changed)

        name_combo = getattr(self.ui, "laminate_name_combo", None)
        if isinstance(name_combo, QComboBox):
            name_combo.currentTextChanged.connect(self._on_laminate_selected)

    def _on_cell_item_changed(
        self,
        current: Optional[QListWidgetItem],
        previous: Optional[QListWidgetItem],  # noqa: ARG002
    ) -> None:
        if current is None:
            return
        cell_id = current.data(Qt.UserRole)
        if not cell_id:
            text = current.text().split("|")[0].strip()
            cell_id = text
        self._on_cell_selected(str(cell_id))

    def _on_cell_selected(self, cell_id: str) -> None:
        if not cell_id:
            return
        self._current_cell_id = cell_id
        mapped = self.model.cell_to_laminate.get(cell_id)
        if mapped and mapped in self.model.laminados:
            self._apply_laminate(mapped)
            return
        candidates = self._laminates_by_cell.get(cell_id, [])
        if candidates:
            self._apply_laminate(candidates[0])
            return
        if self.model.laminados:
            first_name = next(iter(self.model.laminados))
            self._apply_laminate(first_name)

    def _on_laminate_selected(self, laminate_name: str) -> None:
        if self._updating:
            return
        if not laminate_name or laminate_name not in self.model.laminados:
            return
        self._apply_laminate(laminate_name)
        self._update_cell_mapping(laminate_name)

    def _apply_laminate(self, laminate_name: str) -> None:
        if self._updating:
            return
        laminado = self.model.laminados.get(laminate_name)
        if laminado is None:
            return

        self._updating = True
        try:
            name_combo = getattr(self.ui, "laminate_name_combo", None)
            if isinstance(name_combo, QComboBox):
                name_combo.setEditText(laminado.nome)

            color_combo = getattr(self.ui, "laminate_color_combo", None)
            if isinstance(color_combo, QComboBox):
                color_combo.setEditText(laminado.cor_hex)

            type_combo = getattr(self.ui, "laminate_type_combo", None)
            if isinstance(type_combo, QComboBox):
                type_combo.setEditText(laminado.tipo)

            associated_cells = getattr(self.ui, "associated_cells", None)
            if hasattr(associated_cells, "setPlainText"):
                associated_cells.setPlainText(
                    ", ".join(self._cells_for_laminate(laminado.nome))
                )

            self.stacking_model.update_layers(laminado.camadas)
            self._current_laminate = laminado.nome
        finally:
            self._updating = False

    def _cells_for_laminate(self, laminate_name: str) -> list[str]:
        mapped = [
            cell
            for cell in self.model.celulas_ordenadas
            if self.model.cell_to_laminate.get(cell) == laminate_name
        ]
        if mapped:
            return mapped
        laminado = self.model.laminados.get(laminate_name)
        if laminado is not None:
            return list(laminado.celulas)
        return []

    def _update_cell_mapping(self, laminate_name: str) -> None:
        cell_id = self._current_cell_id
        if not cell_id:
            return
        old = self.model.cell_to_laminate.get(cell_id)
        if old == laminate_name:
            return
        if old:
            old_lam = self.model.laminados.get(old)
            if old_lam and cell_id in old_lam.celulas:
                old_lam.celulas.remove(cell_id)
        new_lam = self.model.laminados.get(laminate_name)
        if new_lam is not None and cell_id not in new_lam.celulas:
            new_lam.celulas.append(cell_id)
        if laminate_name:
            self.model.cell_to_laminate[cell_id] = laminate_name
        else:
            self.model.cell_to_laminate.pop(cell_id, None)
        if old and old != laminate_name:
            self._update_associated_cells_text(old)
        self._update_associated_cells_text(laminate_name)
        self._laminates_by_cell = self._build_cell_index()
        if hasattr(self.ui, "_mark_dirty"):
            self.ui._mark_dirty()
        self._refresh_cell_item_label(cell_id)

    def _update_associated_cells_text(self, laminate_name: str) -> None:
        lam = self.model.laminados.get(laminate_name)
        if lam is None:
            return
        cells = self._cells_for_laminate(laminate_name)
        lam.celulas = cells
        widget = getattr(self.ui, "associated_cells", None)
        if (
            widget is not None
            and hasattr(widget, "setPlainText")
            and self._current_laminate == laminate_name
        ):
            widget.setPlainText(", ".join(cells))

    def _install_delegates(self, table: QTableView) -> None:
        self._material_delegate = MaterialComboDelegate(
            table, items_provider=self._material_options
        )
        self._orientation_delegate = OrientationComboDelegate(
            table, items_provider=self._orientation_options
        )
        table.setItemDelegateForColumn(0, self._material_delegate)
        table.setItemDelegateForColumn(1, self._orientation_delegate)

    def _material_options(self) -> list[str]:
        materials: list[str] = []
        seen: set[str] = set()
        for laminado in self.model.laminados.values():
            for camada in laminado.camadas:
                if camada.material and camada.material not in seen:
                    seen.add(camada.material)
                    materials.append(camada.material)
        if not materials:
            materials.append("")
        return materials

    def _orientation_options(self) -> list[str]:
        preferred_order = [0, 45, -45, 90, -90]
        options: list[str] = []
        added: set[int] = set()
        for angle in preferred_order:
            if angle in ALLOWED_ANGLES and angle not in added:
                options.append(f"{angle}\N{DEGREE SIGN}")
                added.add(angle)
        for angle in sorted(ALLOWED_ANGLES):
            if angle not in added:
                options.append(f"{angle}\N{DEGREE SIGN}")
        return options

    def _refresh_cell_item_label(self, cell_id: str) -> None:
        if self._cells_widget is None:
            return
        for idx in range(self._cells_widget.count()):
            item = self._cells_widget.item(idx)
            if item.data(Qt.UserRole) == cell_id:
                self._cells_widget.blockSignals(True)
                item.setText(_format_cell_label(self.model, cell_id))
                self._cells_widget.blockSignals(False)
                break

    def _on_layers_modified(self, _layers: list[Camada]) -> None:
        if hasattr(self.ui, "_mark_dirty"):
            self.ui._mark_dirty()


def _open_with_xlrd(file_path: Path, cause: Exception) -> _WorkbookProtocol:
    """Fallback que usa xlrd 1.2.x para planilhas .xls renomeadas."""
    if xlrd is None:
        raise ValueError(
            "A planilha parece utilizar o formato legado (.xls renomeado). "
            "Instale a dependAancia 'xlrd==1.2.0' para habilitar o fallback."
        ) from cause
    logger.warning(
        "Utilizando fallback xlrd para abrir '%s' devido a erro: %s",
        file_path,
        cause,
    )
    return _XlrdWorkbook(file_path)


class _XlrdWorkbook:
    """Wrapper simples para oferecer interface parecida com pandas.ExcelFile."""

    def __init__(self, file_path: Path) -> None:
        try:
            self._book = xlrd.open_workbook(file_path)  # type: ignore[call-arg]
        except Exception as exc:
            raise ValueError(f"NAo foi possAvel abrir '{file_path}': {exc}") from exc
        self.sheet_names = list(self._book.sheet_names())

    def parse(self, sheet_name: str) -> pd.DataFrame:
        try:
            sheet = self._book.sheet_by_name(sheet_name)
        except Exception as exc:
            raise ValueError(f"Aba '{sheet_name}' nAo pA de ser lida: {exc}") from exc

        rows: list[list[object]] = []
        for row_idx in range(sheet.nrows):
            row: list[object] = []
            for col_idx in range(sheet.ncols):
                value = sheet.cell_value(row_idx, col_idx)
                row.append(value)
            rows.append(row)

        while rows and all(_is_blank(cell) for cell in rows[0]):
            rows.pop(0)

        if not rows:
            return pd.DataFrame()

        max_len = max(len(row) for row in rows)
        normalized_rows = [
            row + [None] * (max_len - len(row)) if len(row) < max_len else row
            for row in rows
        ]
        return pd.DataFrame(normalized_rows)


def _is_blank(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    if isinstance(value, float):
        return math.isnan(value)
    return False


def _resolve_int(value: object, default: Optional[int] = None) -> Optional[int]:
    if _is_blank(value):
        return default
    try:
        if isinstance(value, float) and value.is_integer():
            return int(value)
        return int(str(value).strip())
    except (TypeError, ValueError):
        logger.warning("Valor de Indice invalido '%s'; usando padrao.", value)
        return default


def _string_or_empty(value: object) -> str:
    if _is_blank(value):
        return ""
    return str(value).strip()


@dataclass
class _CellsSection:
    cells: list[str]
    mapping: Dict[str, str]
    separator_idx: int


def parse_cells_from_planilha1(df: pd.DataFrame) -> list[str]:
    """
    Retorna a lista de celulas listadas entre a linha 'Cells' e a linha separadora '#'.
    """
    sanitized = df.dropna(how="all").reset_index(drop=True)
    section = _extract_cells_section(sanitized)
    return section.cells
