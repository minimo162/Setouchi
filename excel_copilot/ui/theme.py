"""Shared Material 3 inspired expressive theme primitives for the desktop UI."""

from __future__ import annotations

import flet as ft


EXPRESSIVE_PALETTE: dict[str, str] = {
    "primary": "#2563EB",
    "on_primary": "#FFFFFF",
    "primary_container": "#E0E7FF",
    "on_primary_container": "#1E3A8A",
    "secondary": "#14B8A6",
    "on_secondary": "#FFFFFF",
    "secondary_container": "#CCFBF1",
    "on_secondary_container": "#0F766E",
    "tertiary": "#16A34A",
    "on_tertiary": "#FFFFFF",
    "tertiary_container": "#DCFCE7",
    "on_tertiary_container": "#14532D",
    "background": "#FFFFFF",
    "surface": "#FFFFFF",
    "surface_dim": "#F8FAFC",
    "surface_high": "#FFFFFF",
    "surface_variant": "#F1F5F9",
    "on_surface": "#0F172A",
    "on_surface_variant": "#475569",
    "outline": "#CBD5E1",
    "outline_variant": "#E2E8F0",
    "inverse_surface": "#0F172A",
    "inverse_on_surface": "#F8FAFC",
    "error": "#DC2626",
    "on_error": "#FFFFFF",
    "error_container": "#FEE2E2",
    "on_error_container": "#7F1D1D",
    "success": "#047857",
    "warning": "#F59E0B",
}


def primary_surface_gradient() -> ft.LinearGradient:
    """Primary expressive gradient for hero areas."""
    return ft.LinearGradient(
        begin=ft.alignment.top_left,
        end=ft.alignment.bottom_right,
        colors=[
            "#60A5FA",
            "#2563EB",
            "#14B8A6",
        ],
    )


def elevated_surface_gradient() -> ft.LinearGradient:
    """Soft elevated gradient for cards and panels."""
    return ft.LinearGradient(
        begin=ft.alignment.top_center,
        end=ft.alignment.bottom_center,
        colors=[
            "#FFFFFF",
            "#F8FAFC",
        ],
    )


def accent_glow_gradient() -> ft.RadialGradient:
    """Subtle accent glow for floating actions."""
    return ft.RadialGradient(
        center=ft.Alignment(0, 0),
        radius=1.2,
        colors=[
            "#2563EB",
            "#2563EB11",
        ],
    )


__all__ = [
    "EXPRESSIVE_PALETTE",
    "primary_surface_gradient",
    "elevated_surface_gradient",
    "accent_glow_gradient",
]
