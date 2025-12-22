from __future__ import annotations

from pathlib import Path

from gridlamedit.io.spreadsheet import (
    Camada,
    DEFAULT_COLOR_INDEX,
    DEFAULT_PLY_TYPE,
    GridModel,
    Laminado,
)
from gridlamedit.services.laminate_batch_import import BatchLaminateInput, parse_batch_template
from gridlamedit.services.laminate_service import auto_name_for_layers, count_oriented_layers


def _build_layers_from_entry(entry: BatchLaminateInput) -> list[Camada]:
    base = list(entry.orientations)
    if not base:
        return []
    if entry.is_symmetric:
        mirror_source = base[:-1] if entry.center_is_single else base
        mirrored = list(reversed(mirror_source))
        full_stack = base + mirrored
    else:
        full_stack = base
    return [
        Camada(
            idx=idx,
            material="",
            orientacao=angle if angle is not None else None,
            ativo=True,
            simetria=False,
            ply_type=DEFAULT_PLY_TYPE,
            ply_label=f"Ply.{idx + 1}",
            sequence=f"Seq.{idx + 1}",
        )
        for idx, angle in enumerate(full_stack)
    ]


def test_batch_template_creates_all_tags() -> None:
    template_path = Path(__file__).resolve().parents[1] / "Template Preenchido Upper Skin Grid_RevA.xlsx"
    assert template_path.exists()

    entries = parse_batch_template(template_path)
    assert entries  # ensures at least one laminate parsed

    created = []
    model = GridModel()
    for entry in entries:
        layers = _build_layers_from_entry(entry)
        if not layers:
            continue
        laminate = Laminado(
            nome="",
            tipo="SS",
            color_index=DEFAULT_COLOR_INDEX,
            tag=str(entry.tag or ""),
            celulas=[],
            camadas=layers,
        )
        laminate.auto_rename_enabled = True
        laminate.nome = auto_name_for_layers(
            model,
            layer_count=count_oriented_layers(layers),
            tag=laminate.tag,
            target=laminate,
        )
        model.laminados[laminate.nome] = laminate
        created.append(laminate)

    assert len(model.laminados) == len(created)

    parsed_tags = {str(entry.tag).strip() for entry in entries if str(entry.tag).strip()}
    created_tags = {str(lam.tag).strip() for lam in created if str(lam.tag).strip()}
    assert parsed_tags.issubset(created_tags)
    assert "4109" in created_tags
