"""Round-trip tests for Excel IO service."""

from __future__ import annotations

from pathlib import Path

from gridlamedit.app.models import Cell, GridDoc, Laminate, Layer
from gridlamedit.app.services.excel_io import load_grid_xlsx, save_grid_xlsx


def _sample_doc() -> GridDoc:
    laminate_a = Laminate(
        name="LAM-A",
        color="#FF0000",
        type="core",
        layers=[
            Layer(index=0, material="Carbon", angle_deg=0.0, active=True),
            Layer(index=1, material="Glass", angle_deg=45.0, active=True),
        ],
        symmetry_index=1,
    )
    laminate_b = Laminate(
        name="LAM-B",
        color="#00FF00",
        type="skin",
        layers=[
            Layer(index=0, material="Kevlar", angle_deg=-45.0, active=True),
            Layer(index=1, material="Basalt", angle_deg=90.0, active=False),
            Layer(index=2, material="Foam", angle_deg=0.0, active=True),
        ],
        symmetry_index=None,
    )

    doc = GridDoc(
        cells=[
            Cell(id="C1", laminate_name="LAM-A"),
            Cell(id="C2", laminate_name="LAM-B"),
            Cell(id="C3", laminate_name=None),
        ],
        laminates={"LAM-A": laminate_a, "LAM-B": laminate_b},
    )
    doc.ensure_associations()
    return doc


def test_excel_round_trip(tmp_path: Path) -> None:
    doc = _sample_doc()
    output_path = tmp_path / "grid_roundtrip.xlsx"

    save_grid_xlsx(doc, output_path)
    assert output_path.exists()

    loaded = load_grid_xlsx(output_path)

    assert set(loaded.laminates.keys()) == set(doc.laminates.keys())
    assert len(loaded.cells) == len(doc.cells)

    for original_cell, loaded_cell in zip(doc.cells, loaded.cells):
        assert original_cell.id == loaded_cell.id
        assert original_cell.laminate_name == loaded_cell.laminate_name

    for name, original_laminate in doc.laminates.items():
        loaded_laminate = loaded.get_laminate(name)
        assert loaded_laminate is not None
        assert loaded_laminate.color == original_laminate.color
        assert loaded_laminate.type == original_laminate.type
        assert (
            loaded_laminate.symmetry_index == original_laminate.symmetry_index
        )
        assert len(loaded_laminate.layers) == len(original_laminate.layers)
        for orig_layer, loaded_layer in zip(
            original_laminate.layers, loaded_laminate.layers
        ):
            assert orig_layer.material == loaded_layer.material
            assert orig_layer.angle_deg == loaded_layer.angle_deg
            assert orig_layer.active == loaded_layer.active

    loaded.ensure_associations()
    for name, original_laminate in doc.laminates.items():
        loaded_laminate = loaded.get_laminate(name)
        assert loaded_laminate is not None
        assert loaded_laminate.associated_cells == original_laminate.associated_cells
