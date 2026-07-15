"""Utilities for glass area and pane calculations."""
from typing import Optional


def calculate_pane_count(bays_wide: int, bays_tall: int, total_count: int = 1) -> int:
    """
    Calculate total number of glass panes.

    Args:
        bays_wide: Number of bays horizontally
        bays_tall: Number of bays vertically
        total_count: Number of elevations (instances)

    Returns:
        Total number of panes
    """
    panes_per_elevation = bays_wide * bays_tall
    return panes_per_elevation * total_count


def calculate_pane_dimensions(
    opening_width: float,
    opening_height: float,
    bays_wide: int,
    bays_tall: int,
    custom_bay_widths: Optional[list] = None,
    custom_bay_heights: Optional[list] = None,
) -> dict:
    """
    Calculate average pane dimensions (useful for display).

    Args:
        opening_width: Total opening width in inches
        opening_height: Total opening height in inches
        bays_wide: Number of bays horizontally
        bays_tall: Number of bays vertically
        custom_bay_widths: Optional custom bay widths in inches
        custom_bay_heights: Optional custom bay heights in inches

    Returns:
        dict with 'avg_width_in', 'avg_height_in', 'avg_width_ft', 'avg_height_ft'
    """
    # Use custom dimensions if provided, otherwise divide equally
    if custom_bay_widths and len(custom_bay_widths) == bays_wide:
        avg_width_in = sum(custom_bay_widths) / len(custom_bay_widths)
    else:
        avg_width_in = opening_width / bays_wide

    if custom_bay_heights and len(custom_bay_heights) == bays_tall:
        avg_height_in = sum(custom_bay_heights) / len(custom_bay_heights)
    else:
        avg_height_in = opening_height / bays_tall

    return {
        'avg_width_in': round(avg_width_in, 2),
        'avg_height_in': round(avg_height_in, 2),
        'avg_width_ft': round(avg_width_in / 12, 2),
        'avg_height_ft': round(avg_height_in / 12, 2),
        'formatted': f"{round(avg_width_in / 12, 2)}' × {round(avg_height_in / 12, 2)}'",
    }


def format_glass_display(
    total_sqft: float,
    pane_count: int,
    pane_dimensions: Optional[dict] = None,
    show_panes: bool = False,
    show_dimensions: bool = False,
) -> str:
    """
    Format glass display string based on options.

    Args:
        total_sqft: Total glass area in square feet
        pane_count: Number of panes
        pane_dimensions: Dict from calculate_pane_dimensions
        show_panes: Include pane count in display
        show_dimensions: Include dimensions in display

    Returns:
        Formatted string like: "24.5 sqft" or "24.5 sqft (8 panes, 2.5' × 3.0')"
    """
    base = f"{total_sqft:.2f} sqft"

    parts = []
    if show_panes:
        parts.append(f"{pane_count} panes")
    if show_dimensions and pane_dimensions:
        parts.append(pane_dimensions['formatted'])

    if parts:
        return f"{base} ({', '.join(parts)})"
    return base
