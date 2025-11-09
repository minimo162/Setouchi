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

EXTENDED_TYPE_SCALE: dict[str, dict[str, object]] = {
    "display": {"size": 32, "weight": ft.FontWeight.W_600},
    "hero": TYPE_SCALE["hero"],
    "title": TYPE_SCALE["title"],
    "title_sm": {"size": 16, "weight": ft.FontWeight.W_600},
    "subtitle": TYPE_SCALE["subtitle"],
    "body": TYPE_SCALE["body"],
    "body_sm": {"size": 11, "weight": ft.FontWeight.W_500},
    "caption": TYPE_SCALE["caption"],
}

DEPTH_TOKENS: dict[str, dict[str, float]] = {
    "micro": {"y": 2.0, "blur": 8.0, "spread": 0.0, "opacity": 0.06},
    "sm": {"y": 6.0, "blur": 16.0, "spread": 0.0, "opacity": 0.08},
    "md": {"y": 14.0, "blur": 28.0, "spread": 0.0, "opacity": 0.10},
    "lg": {"y": 24.0, "blur": 52.0, "spread": 2.0, "opacity": 0.13},
    "xl": {"y": 40.0, "blur": 84.0, "spread": 6.0, "opacity": 0.16},
}

MOTION_TOKENS: dict[str, dict[str, object]] = {
    "micro": {"duration": 120, "curve": "easeOut"},
    "short": {"duration": 200, "curve": "easeInOut"},
    "medium": {"duration": 320, "curve": "easeInOut"},
    "long": {"duration": 480, "curve": "easeInOut"},
}

AURORA_NOISE_TOKEN: dict[str, object] = {
    "opacity": 0.22,
    "blur": 46,
    "scale": 1.4,
}

AURORA_PARTICLE_TOKEN: dict[str, object] = {
    "count": 14,
    "min_size": 4,
    "max_size": 14,
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


def depth_shadow(level: str = "md") -> ft.BoxShadow:
    """デプススケールを直接参照するシャドウ。"""

    spec = DEPTH_TOKENS.get(level, DEPTH_TOKENS["md"])
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


def motion_token(name: str = "short") -> ft.animation.Animation:
    """共通モーションプリセットを Animation に変換。"""

    spec = MOTION_TOKENS.get(name, MOTION_TOKENS["short"])
    return ft.animation.Animation(spec["duration"], spec["curve"])


__all__ = [
    "EXPRESSIVE_PALETTE",
    "TYPE_SCALE",
    "EXTENDED_TYPE_SCALE",
    "DEPTH_TOKENS",
    "MOTION_TOKENS",
    "AURORA_NOISE_TOKEN",
    "AURORA_PARTICLE_TOKEN",
    "primary_surface_gradient",
    "elevated_surface_gradient",
    "accent_glow_gradient",
    "floating_shadow",
    "depth_shadow",
    "glass_surface",
    "glass_border",
    "motion_token",
]
