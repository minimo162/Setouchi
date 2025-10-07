"""Shared Material 3 inspired expressive theme primitives for the desktop UI."""

from __future__ import annotations

import flet as ft


EXPRESSIVE_PALETTE: dict[str, str] = {
    "primary": "#2C7BE5",
    "on_primary": "#FFFFFF",
    "primary_container": "#D9E7FF",
    "on_primary_container": "#082E5A",
    "secondary": "#37A6C1",
    "on_secondary": "#FFFFFF",
    "secondary_container": "#C7F0FF",
    "on_secondary_container": "#003441",
    "tertiary": "#6CC68D",
    "on_tertiary": "#00391C",
    "tertiary_container": "#CFEFD7",
    "on_tertiary_container": "#00210C",
    "background": "#F5FAFF",
    "surface": "#F7FBFF",
    "surface_dim": "#E6EFF7",
    "surface_high": "#FFFFFF",
    "surface_variant": "#E0E7F5",
    "on_surface": "#1F2A37",
    "on_surface_variant": "#4A6075",
    "outline": "#7A8EA2",
    "outline_variant": "#C5D3E1",
    "inverse_surface": "#1F2A37",
    "inverse_on_surface": "#F0F4FA",
    "error": "#BA1A1A",
    "on_error": "#FFFFFF",
    "success": "#1C9A5B",
    "warning": "#F7B628",
}


def primary_surface_gradient() -> ft.LinearGradient:
    """Primary expressive gradient for hero areas."""
    return ft.LinearGradient(
        begin=ft.alignment.top_left,
        end=ft.alignment.bottom_right,
        colors=[
            "#36A6FF",
            "#2C7BE5",
            "#37A6C1",
        ],
    )


def elevated_surface_gradient() -> ft.LinearGradient:
    """Soft elevated gradient for cards and panels."""
    return ft.LinearGradient(
        begin=ft.alignment.top_center,
        end=ft.alignment.bottom_center,
        colors=[
            "#EBF4FF",
            "#DCE9F8",
        ],
    )


def accent_glow_gradient() -> ft.RadialGradient:
    """Subtle accent glow for floating actions."""
    return ft.RadialGradient(
        center=ft.Alignment(0, 0),
        radius=1.2,
        colors=[
            "#2C7BE5",
            "#2C7BE511",
        ],
    )


__all__ = [
    "EXPRESSIVE_PALETTE",
    "primary_surface_gradient",
    "elevated_surface_gradient",
    "accent_glow_gradient",
]
