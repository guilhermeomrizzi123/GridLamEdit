"""Project file management for GridLamEdit."""

from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable, Optional

from gridlamedit.io.spreadsheet import Camada, GridModel, Laminado

logger = logging.getLogger(__name__)


def _serialize_model(model: GridModel) -> dict:
    laminates_data: list[dict] = []
    for laminate in model.laminados.values():
        laminate_dict = {
            "nome": laminate.nome,
            "cor_hex": laminate.cor_hex,
            "tipo": laminate.tipo,
            "celulas": list(laminate.celulas),
            "camadas": [
                {
                    "idx": layer.idx,
                    "material": layer.material,
                    "orientacao": layer.orientacao,
                    "ativo": layer.ativo,
                    "simetria": layer.simetria,
                    "nao_estrutural": getattr(layer, "nao_estrutural", False),
                }
                for layer in laminate.camadas
            ],
        }
        laminates_data.append(laminate_dict)

    return {
        "celulas_ordenadas": list(model.celulas_ordenadas),
        "cell_to_laminate": dict(model.cell_to_laminate),
        "laminados": laminates_data,
        "source_excel_path": getattr(model, "source_excel_path", None),
    }


def _deserialize_model(data: dict) -> GridModel:
    model = GridModel()
    model.celulas_ordenadas = list(data.get("celulas_ordenadas", []))
    model.cell_to_laminate = dict(data.get("cell_to_laminate", {}))
    model.source_excel_path = data.get("source_excel_path")

    laminates = {}
    for lam_data in data.get("laminados", []):
        layers = [
            Camada(
                idx=int(layer.get("idx", index)),
                material=str(layer.get("material", "")),
                orientacao=int(layer.get("orientacao", 0)),
                ativo=bool(layer.get("ativo", True)),
                simetria=bool(layer.get("simetria", False)),
                nao_estrutural=bool(layer.get("nao_estrutural", False)),
            )
            for index, layer in enumerate(lam_data.get("camadas", []))
        ]
        laminate = Laminado(
            nome=str(lam_data.get("nome", "")),
            cor_hex=str(lam_data.get("cor_hex", "#FFFFFF")),
            tipo=str(lam_data.get("tipo", "")),
            celulas=list(lam_data.get("celulas", [])),
            camadas=layers,
        )
        laminates[laminate.nome] = laminate

    model.laminados = laminates
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
