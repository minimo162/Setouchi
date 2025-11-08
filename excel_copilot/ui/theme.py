"""Setouchi Pearl デザインシステムのテーマプリミティブ。"""

from __future__ import annotations

import flet as ft


# Apple の Human Interface Guidelines をヒントに、柔らかいニュートラルと
# 深海のようなアクセントカラーを組み合わせたカラーパレット。
EXPRESSIVE_PALETTE: dict[str, str] = {
    "primary": "#0071E3",
    "on_primary": "#FFFFFF",
    "primary_container": "#D8ECFF",
    "on_primary_container": "#01305A",
    "secondary": "#5E5CE6",
    "on_secondary": "#FFFFFF",
    "secondary_container": "#E3E0FF",
    "on_secondary_container": "#21105E",
    "tertiary": "#30D158",
    "on_tertiary": "#00250D",
    "tertiary_container": "#C9F9D6",
    "on_tertiary_container": "#06391A",
    "background": "#F4F5FA",
    "surface": "#FFFFFF",
    "surface_high": "#FCFDFF",
    "surface_dim": "#EEF1F7",
    "surface_variant": "#E4E7F1",
    "on_surface": "#0A0C18",
    "on_surface_variant": "#4B4F64",
    "outline": "#D3D7E4",
    "outline_variant": "#E8EBF4",
    "inverse_surface": "#1D2033",
    "inverse_on_surface": "#F6F8FF",
    "error": "#FF453A",
    "on_error": "#FFFFFF",
    "error_container": "#FFE2DE",
    "on_error_container": "#5F1410",
    "success": "#248A3D",
    "warning": "#F2A43C",
    "info": "#5C6AC4",
    "shadow": "#020817",
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
            "#0B9BFF",
            "#0071E3",
            "#5E5CE6",
        ],
    )


def elevated_surface_gradient() -> ft.LinearGradient:
    """カードやモーダルのベースに使う柔らかな艶感。"""

    return ft.LinearGradient(
        begin=ft.alignment.top_center,
        end=ft.alignment.bottom_center,
        colors=[
            "#FFFFFF",
            "#F7F9FE",
        ],
    )


def accent_glow_gradient() -> ft.RadialGradient:
    """浮遊感を出すアクセントグロー。"""

    return ft.RadialGradient(
        center=ft.Alignment(0, 0),
        radius=1.25,
        colors=[
            "#5E5CE6EE",
            "#5E5CE608",
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
