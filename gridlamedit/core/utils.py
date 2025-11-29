"""Utility functions for GridLamEdit."""

from __future__ import annotations

import re


def natural_sort_key(text: str) -> list:
    """Return a key for natural sorting, e.g. L1, L2, L10, L25, L25.1.
    
    This function splits the input text into numeric and non-numeric parts
    and returns a list where numbers are converted to floats for proper
    numerical comparison.
    
    Args:
        text: The string to generate a sort key for.
    
    Returns:
        A list of mixed string/float elements for comparison.
    
    Example:
        >>> sorted(['L10', 'L2', 'L1'], key=natural_sort_key)
        ['L1', 'L2', 'L10']
    """
    parts = re.split(r"(\d+\.?\d*)", text)
    result: list = []
    for part in parts:
        part_stripped = part.strip()
        if not part_stripped:
            continue
        try:
            result.append(float(part_stripped))
        except ValueError:
            result.append(part_stripped.lower())
    return result


__all__ = ["natural_sort_key"]
