"""Grid cell model linking to laminates."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass
class Cell:
    """Single grid cell and its laminate association."""

    id: str
    laminate_name: Optional[str] = None
