"""Tests for GridLamEdit data models."""

from __future__ import annotations

from gridlamedit.app.models import Cell, GridDoc, Laminate, Layer


def make_base_layers() -> list[Layer]:
    return [
        Layer(index=0, material="Carbon", angle_deg=0.0, active=True),
        Layer(index=1, material="Glass", angle_deg=45.0, active=True),
        Layer(index=2, material="Kevlar", angle_deg=-45.0, active=False),
    ]


def test_laminate_layer_operations() -> None:
    laminate = Laminate(
        name="LAM-1",
        color="#FFFFFF",
        type="structural",
        layers=make_base_layers(),
    )

    new_layer = Layer(index=99, material="Basalt", angle_deg=90.0, active=True)
    assert laminate.add_layer(new_layer, pos=1) is True
    assert [layer.material for layer in laminate.layers] == [
        "Carbon",
        "Basalt",
        "Glass",
        "Kevlar",
    ]
    assert [layer.index for layer in laminate.layers] == [0, 1, 2, 3]

    appended_layer = Layer(index=88, material="Foam", angle_deg=0.0, active=True)
    assert laminate.add_layer(appended_layer) is True
    assert laminate.layers[-1].material == "Foam"
    assert [layer.index for layer in laminate.layers] == [0, 1, 2, 3, 4]

    assert laminate.add_layer(Layer(0, "Err", 0.0, True), pos=10) is False

    assert laminate.move_layer(0, 4) is True
    assert [layer.material for layer in laminate.layers] == [
        "Basalt",
        "Glass",
        "Kevlar",
        "Foam",
        "Carbon",
    ]
    assert [layer.index for layer in laminate.layers] == [0, 1, 2, 3, 4]
    assert laminate.move_layer(9, 0) is False
    assert laminate.move_layer(0, 9) is False

    assert laminate.duplicate_layer(2) is True
    assert len(laminate.layers) == 6
    assert laminate.layers[2].material == laminate.layers[3].material == "Kevlar"
    assert laminate.layers[2] is not laminate.layers[3]
    assert [layer.index for layer in laminate.layers] == [0, 1, 2, 3, 4, 5]
    assert laminate.duplicate_layer(99) is False

    assert laminate.remove_layer(3) is True
    assert len(laminate.layers) == 5
    assert [layer.index for layer in laminate.layers] == [0, 1, 2, 3, 4]
    assert laminate.remove_layer(42) is False

    laminate.set_symmetry_index(2)
    assert laminate.symmetry_index == 2
    laminate.set_symmetry_index(None)
    assert laminate.symmetry_index is None


def test_grid_doc_associations() -> None:
    laminate_a = Laminate(name="A", color="#FF0000", type="core")
    laminate_b = Laminate(name="B", color="#00FF00", type="skin")

    doc = GridDoc(
        cells=[
            Cell(id="C1", laminate_name="A"),
            Cell(id="C2", laminate_name="B"),
            Cell(id="C3", laminate_name=None),
        ],
        laminates={"A": laminate_a, "B": laminate_b},
    )

    doc.ensure_associations()
    assert laminate_a.associated_cells == ["C1"]
    assert laminate_b.associated_cells == ["C2"]

    assert doc.reassign_cell("C2", "A") is True
    assert laminate_b.associated_cells == []
    assert laminate_a.associated_cells == ["C1", "C2"]

    assert doc.reassign_cell("C4", "A") is False
    assert doc.reassign_cell("C1", "Z") is False

    assert doc.reassign_cell("C3", "B") is True
    assert laminate_b.associated_cells == ["C3"]

    # Duplicate associations should not appear after reconciliation.
    doc.ensure_associations()
    assert laminate_a.associated_cells == ["C1", "C2"]
    assert laminate_b.associated_cells == ["C3"]
