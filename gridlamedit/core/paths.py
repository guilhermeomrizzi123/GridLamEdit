"""Utilities for resolving data paths both in development and frozen builds."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, Union
import sys

Pathish = Union[str, "Path"]

_PACKAGE_ROOT = Path(__file__).resolve().parent.parent


def is_frozen() -> bool:
    """Return True when running from a PyInstaller-built executable."""
    return bool(getattr(sys, "frozen", False))


def _frozen_package_root() -> Path:
    """Base directory for bundled assets when running as a frozen app."""
    base = getattr(sys, "_MEIPASS", None)
    if base is None:
        return _PACKAGE_ROOT
    return Path(base) / "gridlamedit"


def package_root() -> Path:
    """Return the root directory of the ``gridlamedit`` package."""
    if is_frozen():
        return _frozen_package_root()
    return _PACKAGE_ROOT


def package_path(*relative_parts: Pathish) -> Path:
    """
    Build a path inside the project/package, compatible with PyInstaller.

    Parameters
    ----------
    relative_parts:
        Path components relative to the package root.
    """
    root = package_root()
    parts: Iterable[Pathish] = relative_parts or ()
    for part in parts:
        root = root / part
    return root


def resource_path(*relative_parts: Pathish) -> Path:
    """
    Alias for package_path kept for semantic clarity when pointing to assets.
    """
    return package_path(*relative_parts)
