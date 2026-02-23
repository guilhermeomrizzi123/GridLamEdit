from __future__ import annotations

from collections import OrderedDict

from gridlamedit.io.spreadsheet import GridModel, Laminado
from gridlamedit.services.laminate_reassociation import (
    reassociate_laminates_by_contours,
    transfer_neighbor_metadata_after_reassociation,
)


def test_reassociation_preserves_neighbor_drawings_and_remaps_neighbors() -> None:
    old_model = GridModel()
    old_model.celulas_ordenadas = ["OLD_A", "OLD_B"]
    old_model.cell_to_laminate = {
        "OLD_A": "LAM_A",
        "OLD_B": "LAM_B",
    }
    old_model.cell_contours = {
        "OLD_A": ["C1", "C2", "C3"],
        "OLD_B": ["D1", "D2", "D3"],
    }
    old_model.laminados = OrderedDict(
        {
            "LAM_A": Laminado(nome="LAM_A", tipo="SS", celulas=["OLD_A"]),
            "LAM_B": Laminado(nome="LAM_B", tipo="SS", celulas=["OLD_B"]),
        }
    )

    old_model.cell_neighbor_nodes = [
        {
            "cell": "OLD_A",
            "grid": [0, 0],
            "neighbors": {
                "right": {"grid": [1, 0], "cell": "OLD_B"},
            },
        },
        {
            "cell": "OLD_B",
            "grid": [1, 0],
            "neighbors": {
                "left": {"grid": [0, 0], "cell": "OLD_A"},
            },
        },
    ]
    old_model.cell_neighbors = {
        "OLD_A": {"right": ["OLD_B"]},
        "OLD_B": {"left": ["OLD_A"]},
    }
    old_model.cell_neighbor_drawings = [
        {
            "type": "line",
            "p1": [10.0, 10.0],
            "p2": [25.0, 35.0],
            "color": "#112233",
            "width": 2.0,
        },
        {
            "type": "text",
            "pos": [5.0, 8.0],
            "size": [120.0, 32.0],
            "text": "observacao",
            "font_size": 11.0,
            "text_color": "#445566",
        },
    ]

    new_model = GridModel()
    new_model.celulas_ordenadas = ["NEW_1", "NEW_2"]
    new_model.cell_contours = {
        "NEW_1": ["C1", "C2", "C3"],
        "NEW_2": ["D1", "D2", "D3"],
    }
    new_model.laminados = OrderedDict(
        {
            "LAM_A": Laminado(nome="LAM_A", tipo="SS", celulas=[]),
            "LAM_B": Laminado(nome="LAM_B", tipo="SS", celulas=[]),
        }
    )

    report = reassociate_laminates_by_contours(old_model, new_model, apply=True)
    transfer_neighbor_metadata_after_reassociation(
        old_model,
        new_model,
        report.reassociated,
    )

    assert new_model.cell_to_laminate == {
        "NEW_1": "LAM_A",
        "NEW_2": "LAM_B",
    }

    assert new_model.cell_neighbor_drawings == old_model.cell_neighbor_drawings

    assert new_model.cell_neighbor_nodes[0]["cell"] == "NEW_1"
    assert (
        new_model.cell_neighbor_nodes[0]["neighbors"]["right"]["cell"]
        == "NEW_2"
    )
    assert new_model.cell_neighbor_nodes[1]["cell"] == "NEW_2"
    assert (
        new_model.cell_neighbor_nodes[1]["neighbors"]["left"]["cell"]
        == "NEW_1"
    )

    assert new_model.cell_neighbors["NEW_1"]["right"] == ["NEW_2"]
    assert new_model.cell_neighbors["NEW_2"]["left"] == ["NEW_1"]
