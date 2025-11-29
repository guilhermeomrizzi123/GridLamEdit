"""Tests for natural sort function used in laminate dropdown."""

from __future__ import annotations

from gridlamedit.core.utils import natural_sort_key


def test_natural_sort_basic_order():
    """Test basic natural sorting of laminate names."""
    names = ["L25.1", "L1", "L10", "L2", "L35", "L25"]
    sorted_names = sorted(names, key=natural_sort_key)
    assert sorted_names == ["L1", "L2", "L10", "L25", "L25.1", "L35"]


def test_natural_sort_with_letters():
    """Test natural sorting with mixed prefixes."""
    names = ["A10", "A2", "B1", "A1"]
    sorted_names = sorted(names, key=natural_sort_key)
    assert sorted_names == ["A1", "A2", "A10", "B1"]


def test_natural_sort_empty_list():
    """Test sorting an empty list."""
    names: list[str] = []
    sorted_names = sorted(names, key=natural_sort_key)
    assert sorted_names == []


def test_natural_sort_single_item():
    """Test sorting a single item."""
    names = ["L1"]
    sorted_names = sorted(names, key=natural_sort_key)
    assert sorted_names == ["L1"]


def test_natural_sort_case_insensitive():
    """Test that sorting is case insensitive."""
    names = ["l10", "L2", "l1"]
    sorted_names = sorted(names, key=natural_sort_key)
    # After sorting, order should be based on numerical values
    assert sorted_names[0].upper() == "L1"
    assert sorted_names[1].upper() == "L2"
    assert sorted_names[2].upper() == "L10"


def test_natural_sort_decimal_numbers():
    """Test sorting with decimal numbers like L25.1, L25.2."""
    names = ["L25.3", "L25.1", "L25.2", "L25"]
    sorted_names = sorted(names, key=natural_sort_key)
    assert sorted_names == ["L25", "L25.1", "L25.2", "L25.3"]
