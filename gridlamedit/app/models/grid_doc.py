"""Container aggregating cells and laminates."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from .cell import Cell
from .laminate import Laminate


@dataclass
class GridDoc:
    """In-memory document holding grid cells and laminate definitions."""

    cells: list[Cell] = field(default_factory=list)
    laminates: dict[str, Laminate] = field(default_factory=dict)

    def get_laminate(self, name: str) -> Optional[Laminate]:
        """Return the laminate with ``name`` if present."""
        return self.laminates.get(name)

    def ensure_associations(self) -> None:
        """Rebuild laminate associations from the current cell assignments."""
        for laminate in self.laminates.values():
            laminate.associated_cells.clear()

        for cell in self.cells:
            if not cell.laminate_name:
                continue
            laminate = self.laminates.get(cell.laminate_name)
            if laminate is None:
                continue
            if cell.id not in laminate.associated_cells:
                laminate.associated_cells.append(cell.id)

    def reassign_cell(self, cell_id: str, laminate_name: str) -> bool:
        """Assign ``cell_id`` to ``laminate_name`` updating associations."""
        cell = next((item for item in self.cells if item.id == cell_id), None)
        if cell is None:
            return False

        laminate = self.laminates.get(laminate_name)
        if laminate is None:
            return False

        if cell.laminate_name:
            current = self.laminates.get(cell.laminate_name)
            if current is not None and cell.id in current.associated_cells:
                current.associated_cells = [
                    cid for cid in current.associated_cells if cid != cell.id
                ]

        cell.laminate_name = laminate_name
        if cell.id not in laminate.associated_cells:
            laminate.associated_cells.append(cell.id)
        return True
