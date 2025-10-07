"""Shared Material 3 inspired expressive theme primitives for the desktop UI."""

from __future__ import annotations

import flet as ft


EXPRESSIVE_PALETTE: dict[str, str] = {
    "primary": "#5A6BFF",
    "on_primary": "#FFFFFF",
    "primary_container": "#2B2F7A",
    "on_primary_container": "#E2E6FF",
    "secondary": "#FF8A64",
    "on_secondary": "#2A170F",
    "secondary_container": "#442920",
    "on_secondary_container": "#FFDACC",
    "tertiary": "#20D3B9",
    "on_tertiary": "#002922",
    "tertiary_container": "#163F3A",
    "on_tertiary_container": "#A5F5E8",
    "background": "#0F172A",
    "surface": "#131C31",
    "surface_dim": "#0C1426",
    "surface_high": "#1D2742",
    "surface_variant": "#1F2A45",
    "on_surface": "#E2E8F0",
    "on_surface_variant": "#A3B2D8",
    "outline": "#3E4B73",
    "outline_variant": "#2C3759",
    "inverse_surface": "#F4F7FF",
    "inverse_on_surface": "#111321",
    "error": "#FF5468",
    "on_error": "#FFFFFF",
    "success": "#4ADE80",
    "warning": "#FACC15",
}


def primary_surface_gradient() -> ft.LinearGradient:
    """Primary expressive gradient for hero areas."""
    return ft.LinearGradient(
        begin=ft.alignment.top_left,
        end=ft.alignment.bottom_right,
        colors=[
            "#3B5BFF",
            "#5A6BFF",
            "#8A5BFF",
        ],
    )


def elevated_surface_gradient() -> ft.LinearGradient:
    """Soft elevated gradient for cards and panels."""
    return ft.LinearGradient(
        begin=ft.alignment.top_center,
        end=ft.alignment.bottom_center,
        colors=[
            "#1B2440",
            "#1F2C4F",
        ],
    )


def accent_glow_gradient() -> ft.RadialGradient:
    """Subtle accent glow for floating actions."""
    return ft.RadialGradient(
        center=ft.Alignment(0, 0),
        radius=1.2,
        colors=[
            "#5A6BFF",
            "#5A6BFF11",
        ],
    )


__all__ = [
    "EXPRESSIVE_PALETTE",
    "primary_surface_gradient",
    "elevated_surface_gradient",
    "accent_glow_gradient",
]

