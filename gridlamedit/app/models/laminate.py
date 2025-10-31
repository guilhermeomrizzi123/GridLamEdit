"""Laminate model aggregating multiple layers."""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Optional

from .layer import Layer


@dataclass
class Laminate:
    """Laminate definition with its layers and associated cells."""

    name: str
    color: str
    type: str
    layers: list[Layer] = field(default_factory=list)
    associated_cells: list[str] = field(default_factory=list)
    symmetry_index: Optional[int] = None

    def add_layer(self, layer: Layer, pos: Optional[int] = None) -> bool:
        """Insert a new layer at ``pos`` or append when ``pos`` is ``None``."""
        if pos is None:
            self.layers.append(layer)
        else:
            if pos < 0 or pos > len(self.layers):
                return False
            self.layers.insert(pos, layer)
        self._reindex_layers()
        return True

    def remove_layer(self, pos: int) -> bool:
        """Remove the layer at the given index."""
        if pos < 0 or pos >= len(self.layers):
            return False
        del self.layers[pos]
        self._reindex_layers()
        return True

    def move_layer(self, src: int, dst: int) -> bool:
        """Move the layer from ``src`` to ``dst`` position."""
        if (
            src < 0
            or src >= len(self.layers)
            or dst < 0
            or dst >= len(self.layers)
        ):
            return False
        if src == dst:
            return True
        layer = self.layers.pop(src)
        self.layers.insert(dst, layer)
        self._reindex_layers()
        return True

    def duplicate_layer(self, pos: int) -> bool:
        """Duplicate the layer at ``pos`` inserting the copy right after it."""
        if pos < 0 or pos >= len(self.layers):
            return False
        original = self.layers[pos]
        clone = replace(original)
        self.layers.insert(pos + 1, clone)
        self._reindex_layers()
        return True

    def set_symmetry_index(self, pos: Optional[int]) -> None:
        """Assign the stored symmetry index (no validation)."""
        self.symmetry_index = pos

    def _reindex_layers(self) -> None:
        """Ensure layer indexes follow their position."""
        for idx, layer in enumerate(self.layers):
            layer.index = idx
