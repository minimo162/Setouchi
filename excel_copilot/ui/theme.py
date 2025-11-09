"""Setouchi Pearl デザインシステムのテーマプリミティブ。"""

from __future__ import annotations

import flet as ft


# Apple のハードウェア仕上げや visionOS のフォググラスを意識した、
# 涼しいメタリックブルーとオーロラライクなアクセントを持つ新パレット。
EXPRESSIVE_PALETTE: dict[str, str] = {
    "primary": "#0A84FF",
    "on_primary": "#F4FBFF",
    "primary_container": "#CFE7FF",
    "on_primary_container": "#00274A",
    "secondary": "#7D7AFF",
    "on_secondary": "#F8F7FF",
    "secondary_container": "#E5E3FF",
    "on_secondary_container": "#1B0E5C",
    "tertiary": "#31DFC3",
    "on_tertiary": "#01211D",
    "tertiary_container": "#C2FFF1",
    "on_tertiary_container": "#01362F",
    "background": "#F5F7FB",
    "surface": "#FFFFFF",
    "surface_high": "#F3F7FF",
    "surface_dim": "#E8EDF5",
    "surface_variant": "#DDE2F1",
    "on_surface": "#040A17",
    "on_surface_variant": "#4B5166",
    "outline": "#C4CBDA",
    "outline_variant": "#E1E6F3",
    "inverse_surface": "#0F1422",
    "inverse_on_surface": "#E4EBFF",
    "error": "#FF5B5B",
    "on_error": "#FFFFFF",
    "error_container": "#FFE3E0",
    "on_error_container": "#5E1412",
    "success": "#2BD37D",
    "warning": "#F7B849",
    "info": "#4B73FF",
    "shadow": "#040714",
}


TYPE_SCALE: dict[str, dict[str, object]] = {
    "hero": {"size": 24, "weight": ft.FontWeight.W_600},
    "title": {"size": 18, "weight": ft.FontWeight.W_600},
    "subtitle": {"size": 14, "weight": ft.FontWeight.W_500},
    "body": {"size": 13, "weight": ft.FontWeight.W_400},
    "caption": {"size": 12, "weight": ft.FontWeight.W_400},
}


def primary_surface_gradient() -> ft.LinearGradient:
    """アプリ全体のヒーローブロック向けグラデーション。"""

    return ft.LinearGradient(
        begin=ft.alignment.top_left,
        end=ft.alignment.bottom_right,
        colors=[
            "#53E4FF",
            "#0A84FF",
            "#7D7AFF",
            "#FF87D1",
        ],
    )


def elevated_surface_gradient() -> ft.LinearGradient:
    """カードやモーダルのベースに使う柔らかな艶感。"""

    return ft.LinearGradient(
        begin=ft.alignment.top_center,
        end=ft.alignment.bottom_center,
        colors=[
            "#FFFFFF",
            "#F4F7FF",
            "#EBF1FF",
        ],
    )


def accent_glow_gradient() -> ft.RadialGradient:
    """浮遊感を出すアクセントグロー。"""

    return ft.RadialGradient(
        center=ft.Alignment(0, 0),
        radius=1.25,
        colors=[
            "#7D7AFFDD",
            "#31DFC300",
        ],
    )


def floating_shadow(level: str = "md") -> ft.BoxShadow:
    """Apple ライクな奥行きを付与するソフトシャドウ。"""

    levels = {
        "sm": {"y": 4, "blur": 14, "spread": 0, "opacity": 0.08},
        "md": {"y": 10, "blur": 28, "spread": 0, "opacity": 0.10},
        "lg": {"y": 22, "blur": 46, "spread": 2, "opacity": 0.12},
    }
    spec = levels.get(level, levels["md"])
    return ft.BoxShadow(
        spread_radius=spec["spread"],
        blur_radius=spec["blur"],
        color=ft.Colors.with_opacity(spec["opacity"], EXPRESSIVE_PALETTE["shadow"]),
        offset=ft.Offset(0, spec["y"]),
    )


def glass_surface(opacity: float = 0.78) -> str:
    """ガラス質感の背景色。"""

    opacity = max(0.0, min(opacity, 1.0))
    return ft.Colors.with_opacity(opacity, EXPRESSIVE_PALETTE["surface_high"])


def glass_border(alpha: float = 0.32) -> ft.Border:
    """角丸カード向けの繊細なボーダー。"""

    alpha = max(0.0, min(alpha, 1.0))
    return ft.border.all(1, ft.Colors.with_opacity(alpha, EXPRESSIVE_PALETTE["outline"]))


__all__ = [
    "EXPRESSIVE_PALETTE",
    "TYPE_SCALE",
    "primary_surface_gradient",
    "elevated_surface_gradient",
    "accent_glow_gradient",
    "floating_shadow",
    "glass_surface",
    "glass_border",
]
