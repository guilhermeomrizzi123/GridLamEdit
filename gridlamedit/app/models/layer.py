"""Layer model for GridLamEdit."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class Layer:
    """Single laminate layer definition."""

    index: int
    material: str
    angle_deg: float
    active: bool = True

    def __str__(self) -> str:
        return (
            "Layer("
            f"index={self.index}, "
            f"material='{self.material}', "
            f"angle_deg={self.angle_deg}, "
            f"active={self.active}"
            ")"
        )

    __repr__ = __str__
