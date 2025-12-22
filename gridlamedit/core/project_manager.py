"""Project file management for GridLamEdit."""

from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable, Optional

from gridlamedit.io.spreadsheet import (
    Camada,
    DEFAULT_COLOR_INDEX,
    DEFAULT_PLY_TYPE,
    DEFAULT_ROSETTE_LABEL,
    GridModel,
    Laminado,
    PLY_TYPE_OPTIONS,
    normalize_angle,
    normalize_color_index,
    normalize_hex_color,
)

logger = logging.getLogger(__name__)


def _serialize_model(model: GridModel) -> dict:
    laminates_data: list[dict] = []
    for laminate in model.laminados.values():
        color_value = getattr(laminate, "color_index", DEFAULT_COLOR_INDEX)
        color_index = DEFAULT_COLOR_INDEX
        color_hex: Optional[str] = None
        if isinstance(color_value, str):
            normalized = normalize_hex_color(color_value)
            if normalized:
                color_hex = normalized
            else:
                try:
                    color_index = int(float(color_value))
                except (TypeError, ValueError):
                    color_index = DEFAULT_COLOR_INDEX
        else:
            try:
                color_index = int(color_value or DEFAULT_COLOR_INDEX)
            except (TypeError, ValueError):
                color_index = DEFAULT_COLOR_INDEX
        laminate_dict = {
            "nome": laminate.nome,
            "color_index": color_index,
            "tipo": laminate.tipo,
            "tag": getattr(laminate, "tag", ""),
            "celulas": list(laminate.celulas),
            "camadas": [
                {
                    "idx": layer.idx,
                    "material": layer.material,
                    "orientacao": layer.orientacao,
                    "ativo": layer.ativo,
                    "simetria": layer.simetria,
                    "ply_type": getattr(layer, "ply_type", DEFAULT_PLY_TYPE),
                    "ply_label": getattr(
                        layer, "ply_label", f"Ply.{layer.idx + 1}"
                    ),
                    "sequence": getattr(layer, "sequence", f"Seq.{layer.idx + 1}"),
                    "rosette": getattr(layer, "rosette", DEFAULT_ROSETTE_LABEL),
                }
                for layer in laminate.camadas
            ],
            "auto_rename_enabled": bool(
                getattr(laminate, "auto_rename_enabled", True)
            ),
        }
        if color_hex:
            laminate_dict["color_hex"] = color_hex
        laminates_data.append(laminate_dict)

    return {
        "celulas_ordenadas": list(model.celulas_ordenadas),
        "cell_to_laminate": dict(model.cell_to_laminate),
        # Grafo detalhado com posições (compatível com múltiplas instâncias da mesma célula)
        "cell_neighbor_nodes": list(getattr(model, "cell_neighbor_nodes", [])),
        "cell_neighbors": dict(getattr(model, "cell_neighbors", {})),
        "laminados": laminates_data,
        "source_excel_path": getattr(model, "source_excel_path", None),
    }


def _deserialize_model(data: dict) -> GridModel:
    model = GridModel()
    model.celulas_ordenadas = list(data.get("celulas_ordenadas", []))
    model.cell_to_laminate = dict(data.get("cell_to_laminate", {}))
    model.cell_neighbor_nodes = list(data.get("cell_neighbor_nodes", []))
    model.cell_neighbors = dict(data.get("cell_neighbors", {}))
    model.source_excel_path = data.get("source_excel_path")

    laminates = {}
    compat_warnings: list[str] = []
    for lam_data in data.get("laminados", []):
        lam_name = str(lam_data.get("nome", ""))
        color_value: int | str = DEFAULT_COLOR_INDEX
        raw_hex = lam_data.get("color_hex")
        normalized_hex = normalize_hex_color(raw_hex) if raw_hex else None
        if normalized_hex:
            color_value = normalized_hex
        else:
            raw_color_idx = lam_data.get("color_index", None)
            if raw_color_idx is None:
                legacy_hex = lam_data.get("cor_hex")
                if legacy_hex:
                    message = (
                        f"Laminado '{lam_name or '(sem nome)'}' traz cor hexadecimal '{legacy_hex}'; "
                        f"convertendo para indice {DEFAULT_COLOR_INDEX}."
                    )
                    logger.warning(message)
                    compat_warnings.append(message)
            else:
                color_index = normalize_color_index(raw_color_idx, DEFAULT_COLOR_INDEX)
                color_value = color_index
                try:
                    parsed_value = int(float(raw_color_idx))
                except (TypeError, ValueError):
                    parsed_value = None
                if parsed_value is None or parsed_value != color_index:
                    message = (
                        f"Laminado '{lam_name or '(sem nome)'}' possui indice de cor invalido "
                        f"'{raw_color_idx}'; usando {color_index}."
                    )
                    compat_warnings.append(message)

        layers: list[Camada] = []
        for index, layer in enumerate(lam_data.get("camadas", [])):
            ply_type = layer.get("ply_type")
            if ply_type not in PLY_TYPE_OPTIONS:
                ply_type = (
                    PLY_TYPE_OPTIONS[1]
                    if bool(layer.get("nao_estrutural", False))
                    else DEFAULT_PLY_TYPE
                )
            orientation_raw = layer.get("orientacao", None)
            orientation_value: Optional[float]
            try:
                orientation_value = normalize_angle(orientation_raw)
            except (TypeError, ValueError):
                orientation_value = None
            sequence_value = str(layer.get("sequence", "") or "").strip()
            if not sequence_value:
                sequence_value = f"Seq.{index + 1}"
            ply_label_value = str(layer.get("ply_label", "") or "").strip()
            if not ply_label_value:
                ply_label_value = f"Ply.{index + 1}"
            rosette_value = str(layer.get("rosette", "") or "").strip() or DEFAULT_ROSETTE_LABEL
            layers.append(
                Camada(
                    idx=int(layer.get("idx", index)),
                    material=str(layer.get("material", "")),
                    orientacao=orientation_value,
                    ativo=bool(layer.get("ativo", True)),
                    simetria=bool(layer.get("simetria", False)),
                    ply_type=str(ply_type),
                    ply_label=ply_label_value,
                    sequence=sequence_value,
                    rosette=rosette_value,
                )
            )
        laminate = Laminado(
            nome=lam_name,
            tipo=str(lam_data.get("tipo", "")),
            color_index=color_value,
            tag=str(lam_data.get("tag", "") or ""),
            celulas=list(lam_data.get("celulas", [])),
            camadas=layers,
            auto_rename_enabled=bool(
                lam_data.get("auto_rename_enabled", True)
            ),
        )
        laminates[laminate.nome] = laminate

    model.laminados = laminates
    if compat_warnings:
        model.compat_warnings.extend(compat_warnings)
    return model


class ProjectManager:
    """Handle saving and loading of GridLamEdit project files."""

    VERSION = "1.0"

    def __init__(
        self, dirty_callback: Optional[Callable[[bool], None]] = None
    ) -> None:
        self.current_path: Optional[Path] = None
        self.snapshot: dict = {}
        self.is_dirty: bool = False
        self._dirty_callback = dirty_callback

    # Snapshot helpers -------------------------------------------------

    def capture_from_model(
        self, model: GridModel, ui_state: Optional[dict] = None
    ) -> None:
        """Capture the current model state into the project snapshot."""
        self.snapshot = {
            "grid": _serialize_model(model),
            "ui_state": ui_state or {},
        }

    def build_model(self) -> GridModel:
        """Build a GridModel instance from the last loaded snapshot."""
        grid_data = self.snapshot.get("grid")
        if not grid_data:
            raise ValueError("Project snapshot is empty.")
        model = _deserialize_model(grid_data)
        model.dirty = self.is_dirty
        return model

    def get_ui_state(self) -> dict:
        return dict(self.snapshot.get("ui_state", {}))

    # Dirty state ------------------------------------------------------

    def mark_dirty(self, value: bool = True) -> None:
        if self.is_dirty == value:
            return
        self.is_dirty = value
        if self._dirty_callback:
            try:
                self._dirty_callback(self.is_dirty)
            except Exception as exc:  # pragma: no cover - defensive
                logger.warning("Dirty callback failed: %s", exc)

    # Save / load ------------------------------------------------------

    def save(self, path: Optional[Path] = None) -> None:
        if path is not None:
            self.current_path = Path(path)
        if self.current_path is None:
            raise ValueError("Nenhum arquivo de projeto informado para salvar.")
        if not self.snapshot:
            raise ValueError("Nenhum estado capturado para salvar.")

        project_data = {
            "version": self.VERSION,
            "saved_at_utc": datetime.now(tz=timezone.utc)
            .replace(microsecond=0)
            .isoformat(),
            "source_excel_path": self.snapshot.get("grid", {}).get(
                "source_excel_path"
            ),
            "grid": self.snapshot.get("grid", {}),
            "ui_state": self.snapshot.get("ui_state", {}),
        }

        self.current_path.parent.mkdir(parents=True, exist_ok=True)
        with self.current_path.open("w", encoding="utf-8") as handle:
            json.dump(project_data, handle, indent=2)

        self.mark_dirty(False)

    def load(self, path: Path) -> None:
        path = Path(path)
        with path.open("r", encoding="utf-8") as handle:
            data = json.load(handle)

        version = data.get("version")
        if version != self.VERSION:
            logger.warning(
                "Projeto na versao %s diferente da suportada %s.",
                version,
                self.VERSION,
            )

        self.snapshot = {
            "grid": data.get("grid", {}),
            "ui_state": data.get("ui_state", {}),
        }
        if "source_excel_path" in data:
            self.snapshot["grid"]["source_excel_path"] = data["source_excel_path"]

        self.current_path = path
        self.mark_dirty(False)
