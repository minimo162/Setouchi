"""Reusable chat-related UI components for the desktop application."""

import math
from typing import Any, Dict, List, Optional, Union

import flet as ft

from .messages import ResponseType
from .theme import (
    EXPRESSIVE_PALETTE,
    TYPE_SCALE,
    floating_shadow,
    glass_border,
    glass_surface,
    primary_surface_gradient,
)


_MAX_MESSAGE_HEIGHT = 360
_MAX_BUBBLE_WIDTH = 560


class ChatMessage(ft.ResponsiveRow):
    """Material Design inspired chat bubble for rendering responses."""

    def __init__(self, msg_type: Union[ResponseType, str], msg_content: str, metadata: Optional[Dict[str, Any]] = None, animate: bool = True):
        super().__init__()
        self.vertical_alignment = ft.CrossAxisAlignment.START
        self.opacity = 0
        self.animate_opacity = 300
        self.offset = ft.Offset(0, 0.1)
        self.animate_offset = 300
        if not animate:
            self.opacity = 1
            self.animate_opacity = 0
            self.offset = ft.Offset(0, 0)
            self.animate_offset = 0

        metadata = metadata or {}
        display_time = metadata.get("display_time")
        mode_label = metadata.get("mode_label")

        palette = EXPRESSIVE_PALETTE
        body_scale = TYPE_SCALE["body"]
        subtitle_scale = TYPE_SCALE["subtitle"]
        caption_scale = TYPE_SCALE["caption"]

        def _final_answer_gradient() -> ft.LinearGradient:
            return primary_surface_gradient()

        type_map = {
            "user": {
                "bgcolor": glass_surface(0.82),
                "icon": ft.Icons.PERSON_ROUNDED,
                "icon_color": palette["primary"],
                "icon_bgcolor": ft.Colors.with_opacity(0.18, palette["primary"]),
                "border": glass_border(0.34),
                "text_style": {"color": palette["on_surface"], "size": body_scale["size"], "weight": ft.FontWeight.W_500},
            },
            "thought": {
                "bgcolor": ft.Colors.with_opacity(0.85, palette["secondary_container"]),
                "icon": ft.Icons.LIGHTBULB_OUTLINE,
                "icon_color": palette["on_secondary_container"],
                "icon_bgcolor": ft.Colors.with_opacity(0.26, palette["secondary"]),
                "text_style": {"color": palette["on_secondary_container"], "size": body_scale["size"]},
                "title": "AI Thought",
                "border": ft.border.all(1, ft.Colors.with_opacity(0.36, palette["secondary"])),
            },
            "action": {
                "bgcolor": glass_surface(0.68),
                "icon": ft.Icons.CODE,
                "icon_color": palette["on_surface_variant"],
                "icon_bgcolor": ft.Colors.with_opacity(0.22, palette["primary"]),
                "text_style": {"font_family": "monospace", "color": palette["on_surface_variant"], "size": body_scale["size"]},
                "title": "Action",
                "border": glass_border(0.26),
            },
            "observation": {
                "bgcolor": glass_surface(0.76),
                "icon": ft.Icons.FIND_IN_PAGE_OUTLINED,
                "icon_color": palette["on_surface"],
                "icon_bgcolor": ft.Colors.with_opacity(0.18, palette["tertiary"]),
                "text_style": {"color": palette["on_surface"], "size": body_scale["size"]},
                "title": "Observation",
                "border": glass_border(0.28),
            },
            "final_answer": {
                "gradient_factory": _final_answer_gradient,
                "icon": ft.Icons.CHECK_CIRCLE_OUTLINE,
                "icon_color": palette["inverse_on_surface"],
                "icon_bgcolor": ft.Colors.with_opacity(0.26, palette["tertiary"]),
                "text_style": {"color": palette["inverse_on_surface"], "size": body_scale["size"], "weight": ft.FontWeight.W_500},
                "title": "Answer",
                "border": ft.border.all(1, ft.Colors.with_opacity(0.45, ft.Colors.WHITE)),
            },
            "chat_prompt": {
                "bgcolor": ft.Colors.with_opacity(0.82, palette["secondary_container"]),
                "icon": ft.Icons.CONTENT_PASTE,
                "icon_color": palette["on_secondary_container"],
                "icon_bgcolor": ft.Colors.with_opacity(0.22, palette["secondary"]),
                "text_style": {"color": palette["on_secondary_container"], "size": body_scale["size"]},
                "title": "Copilot Prompt",
                "border": ft.border.all(1, ft.Colors.with_opacity(0.32, palette["secondary"])),
            },
            "chat_response": {
                "gradient_factory": _final_answer_gradient,
                "icon": ft.Icons.SMART_TOY_OUTLINED,
                "icon_color": palette["inverse_on_surface"],
                "icon_bgcolor": ft.Colors.with_opacity(0.24, palette["tertiary"]),
                "text_style": {"color": palette["inverse_on_surface"], "size": body_scale["size"], "weight": ft.FontWeight.W_500},
                "title": "Copilot Response",
                "border": ft.border.all(1, ft.Colors.with_opacity(0.45, ft.Colors.WHITE)),
            },
            "info": {
                "text_style": {"color": palette["on_surface_variant"], "size": 12},
                "icon": ft.Icons.INFO_OUTLINE,
                "icon_color": palette["on_surface_variant"],
            },
            "status": {
                "text_style": {"color": palette["on_surface_variant"], "size": 12},
            },
            "error": {
                "icon": ft.Icons.ERROR_OUTLINE_ROUNDED,
                "icon_color": palette["error"],
                "bgcolor": palette["error_container"],
                "icon_bgcolor": ft.Colors.with_opacity(0.22, palette["error"]),
                "text_style": {"color": palette["on_error_container"], "size": body_scale["size"]},
                "title": "Error",
                "border": ft.border.all(1, ft.Colors.with_opacity(0.32, palette["error"])),
            },
        }

        msg_type_value = msg_type.value if isinstance(msg_type, ResponseType) else msg_type
        config = type_map.get(msg_type_value, type_map["info"])

        if msg_type_value in ["info", "status"]:
            controls: List[ft.Control] = []
            header_controls: List[ft.Control] = []
            if mode_label:
                header_controls.append(
                    ft.Text(
                        mode_label,
                        size=caption_scale["size"],
                        weight=caption_scale["weight"],
                        color=palette["primary"],
                    )
                )
            if display_time:
                header_controls.append(
                    ft.Text(
                        display_time,
                        size=caption_scale["size"],
                        weight=caption_scale["weight"],
                        color=palette["on_surface_variant"],
                    )
                )
            if header_controls:
                alignment = ft.MainAxisAlignment.SPACE_BETWEEN if len(header_controls) == 2 else (ft.MainAxisAlignment.START if mode_label else ft.MainAxisAlignment.END)
                controls.append(
                    ft.Row(
                        header_controls,
                        alignment=alignment,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    )
                )
            info_text_style = config.get("text_style", {"color": palette["on_surface_variant"], "size": caption_scale["size"]})
            text_control = ft.Text(msg_content, **info_text_style)
            controls.append(text_control)
            self.controls = [
                ft.Column(
                    controls,
                    col=12,
                    alignment=ft.MainAxisAlignment.START,
                    horizontal_alignment=ft.CrossAxisAlignment.START,
                    spacing=4,
                )
            ]
            return

        content_controls: List[ft.Control] = []
        header_controls: List[ft.Control] = []
        if mode_label:
            header_controls.append(
                ft.Container(
                    ft.Text(
                        mode_label,
                        size=caption_scale["size"],
                        weight=caption_scale["weight"],
                        color=palette["primary"],
                    ),
                    padding=ft.Padding(10, 4, 10, 4),
                    bgcolor=ft.Colors.with_opacity(0.14, palette["primary"]),
                    border_radius=12,
                )
            )
        if display_time:
            header_controls.append(
                ft.Text(
                    display_time,
                    size=caption_scale["size"],
                    weight=caption_scale["weight"],
                    color=palette["on_surface_variant"],
                )
            )
        if header_controls:
            alignment = ft.MainAxisAlignment.SPACE_BETWEEN if len(header_controls) == 2 else (ft.MainAxisAlignment.START if mode_label else ft.MainAxisAlignment.END)
            content_controls.append(
                ft.Row(
                    header_controls,
                    alignment=alignment,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                )
            )

        if config.get("title"):
            content_controls.append(
                ft.Text(
                    config["title"],
                    size=subtitle_scale["size"],
                    weight=subtitle_scale["weight"],
                    color=palette["on_surface_variant"],
                )
            )

        text_style = config.get("text_style", {"color": palette["on_surface"], "size": body_scale["size"], "weight": body_scale["weight"]})
        normalized_content = (msg_content or "").replace("\r\n", "\n")
        line_controls: List[ft.Control] = []
        size = text_style.get("size")
        icon_color = config.get("icon_color", palette["on_surface_variant"])

        for raw_line in normalized_content.split("\n"):
            if raw_line.strip() == "":
                line_controls.append(ft.Container(height=10))
                continue

            stripped = raw_line.strip()
            if stripped.startswith("引用"):
                label, sep, remainder = stripped.partition(":")
                bullet = ft.Text("•", size=size or 13, color=icon_color)
                label_text = ft.Text(
                    label.strip() + (sep if sep else ""),
                    weight=ft.FontWeight.BOLD,
                    size=size or 13,
                    color=icon_color,
                )
                remainder_texts = []
                remainder_value = remainder.strip() if remainder else ""
                if remainder_value:
                    remainder_texts.append(ft.Text(remainder_value, **text_style, selectable=True))
                line_controls.append(
                    ft.Row(
                        [
                            bullet,
                            ft.Column([label_text] + remainder_texts, spacing=2, tight=True),
                        ],
                        alignment=ft.MainAxisAlignment.START,
                        vertical_alignment=ft.CrossAxisAlignment.START,
                        spacing=6,
                    )
                )
            else:
                line_controls.append(ft.Text(raw_line, **text_style, selectable=True))

        content_controls.extend(
            line_controls if line_controls else [ft.Text(msg_content, **text_style, selectable=True)]
        )

        approx_lines = max(
            1,
            normalized_content.count("\n") + 1 if normalized_content else 0,
            math.ceil(len(normalized_content) / 60) if normalized_content else 0,
        )
        estimated_height = 56 + approx_lines * 20
        needs_scroll = estimated_height > _MAX_MESSAGE_HEIGHT
        scroll_mode = ft.ScrollMode.AUTO if needs_scroll else None
        content_column = ft.Column(
            content_controls,
            spacing=8,
            tight=True,
            scroll=scroll_mode,
            auto_scroll=True if needs_scroll else None,
        )

        gradient_factory = config.get("gradient_factory")
        gradient = gradient_factory() if callable(gradient_factory) else None
        icon_gradient_factory = config.get("icon_gradient_factory")
        icon_gradient = icon_gradient_factory() if callable(icon_gradient_factory) else None

        border = config.get("border")
        if border is None:
            border_color = config.get("border_color", ft.Colors.with_opacity(0.22, palette["outline"]))
            border = ft.border.all(1, border_color)

        message_bubble = ft.Container(
            content=content_column,
            bgcolor=config.get("bgcolor", glass_surface(0.84)),
            gradient=gradient if gradient else None,
            border_radius=28,
            padding=ft.Padding(22, 18, 22, 18),
            expand=True,
            height=_MAX_MESSAGE_HEIGHT if needs_scroll else None,
            clip_behavior=ft.ClipBehavior.HARD_EDGE if needs_scroll else ft.ClipBehavior.NONE,
            shadow=floating_shadow("sm"),
            border=border,
            constraints=ft.BoxConstraints(max_width=_MAX_BUBBLE_WIDTH),
        )

        icon_name = config.get("icon", ft.Icons.SMART_BUTTON)
        icon_color = config.get("icon_color", palette["on_surface_variant"])
        icon = ft.Icon(name=icon_name, color=icon_color, size=20)
        icon_container = ft.Container(
            icon,
            width=36,
            height=36,
            gradient=icon_gradient if icon_gradient else None,
            bgcolor=config.get("icon_bgcolor", palette["surface_variant"]),
            alignment=ft.alignment.center,
            border_radius=12,
            margin=ft.margin.only(right=12, left=12, top=6),
            border=ft.border.all(1, ft.Colors.with_opacity(0.14, icon_color)) if icon_gradient or config.get("icon_bgcolor") else None,
            shadow=floating_shadow("sm"),
        )

        row_alignment = ft.MainAxisAlignment.START
        row_children: List[ft.Control]
        if msg_type_value in {"user", "chat_prompt"}:
            row_alignment = ft.MainAxisAlignment.END
            row_children = [message_bubble, icon_container]
        else:
            row_children = [icon_container, message_bubble]

        bubble_and_icon_row = ft.Row(
            row_children,
            vertical_alignment=ft.CrossAxisAlignment.START,
            spacing=16,
            alignment=row_alignment,
        )

        if msg_type_value in {"user", "chat_prompt"}:
            self.controls = [
                ft.Column(col={"sm": 2, "md": 4}),
                ft.Column(col={"sm": 10, "md": 8}, controls=[bubble_and_icon_row]),
            ]
        else:
            self.controls = [
                ft.Column(col={"sm": 10, "md": 8}, controls=[bubble_and_icon_row]),
            ]

    def appear(self):
        self.opacity = 1
        self.offset = ft.Offset(0, 0)
        self.update()


__all__ = ["ChatMessage"]
