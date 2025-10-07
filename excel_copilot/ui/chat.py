"""Reusable chat-related UI components for the desktop application."""

from typing import Union

import flet as ft

from .messages import ResponseType


class ChatMessage(ft.ResponsiveRow):
    """Material Design inspired chat bubble for rendering responses."""

    def __init__(self, msg_type: Union[ResponseType, str], msg_content: str):
        super().__init__()
        self.vertical_alignment = ft.CrossAxisAlignment.START
        self.opacity = 0
        self.animate_opacity = 300
        self.offset = ft.Offset(0, 0.1)
        self.animate_offset = 300

        palette = {
            "primary_container": "#4F378B",
            "on_primary_container": "#EADDFF",
            "secondary_container": "#4A4458",
            "on_secondary_container": "#E8DEF8",
            "tertiary_container": "#633B48",
            "on_tertiary_container": "#FFD8E4",
            "neutral_container": "#332D41",
            "on_neutral_container": "#E8DEF8",
            "surface_variant": "#49454F",
            "on_surface_variant": "#CAC4D0",
            "error_container": "#8C1D18",
            "on_error_container": "#F9DEDC",
        }

        type_map = {
            "user": {
                "bgcolor": palette["primary_container"],
                "icon": ft.Icons.PERSON_ROUNDED,
                "icon_color": palette["on_primary_container"],
                "text_style": {"color": palette["on_primary_container"], "size": 14},
            },
            "thought": {
                "bgcolor": palette["secondary_container"],
                "icon": ft.Icons.LIGHTBULB_OUTLINE,
                "icon_color": palette["on_secondary_container"],
                "text_style": {"color": palette["on_secondary_container"], "size": 13},
                "title": "AI Thought",
            },
            "action": {
                "bgcolor": palette["surface_variant"],
                "icon": ft.Icons.CODE,
                "icon_color": palette["on_surface_variant"],
                "text_style": {"font_family": "monospace", "color": palette["on_surface_variant"], "size": 13},
                "title": "Action",
            },
            "observation": {
                "bgcolor": palette["neutral_container"],
                "icon": ft.Icons.FIND_IN_PAGE_OUTLINED,
                "icon_color": palette["on_neutral_container"],
                "text_style": {"color": palette["on_neutral_container"], "size": 13},
                "title": "Observation",
            },
            "final_answer": {
                "bgcolor": palette["tertiary_container"],
                "icon": ft.Icons.CHECK_CIRCLE_OUTLINE,
                "icon_color": palette["on_tertiary_container"],
                "text_style": {"color": palette["on_tertiary_container"], "size": 14},
                "title": "Answer",
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
                "icon_color": palette["on_error_container"],
                "bgcolor": palette["error_container"],
                "text_style": {"color": palette["on_error_container"], "size": 13},
                "title": "Error",
            },
        }

        msg_type_value = msg_type.value if isinstance(msg_type, ResponseType) else msg_type
        config = type_map.get(msg_type_value, type_map["info"])

        if msg_type_value in ["info", "status"]:
            self.controls = [
                ft.Column(
                    [ft.Text(msg_content, **config.get("text_style", {}))],
                    col=12,
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                )
            ]
            return

        content_controls = []
        if config.get("title"):
            content_controls.append(
                ft.Text(config["title"], weight=ft.FontWeight.BOLD, size=12, color=config.get("icon_color"))
            )

        text_style = dict(config.get("text_style", {}))
        line_controls = []
        icon_color = config.get("icon_color", text_style.get("color"))
        size = text_style.get("size")
        normalized_content = (msg_content or "").replace("\r\n", "\n")
        for raw_line in normalized_content.split("\n"):
            if raw_line.strip() == "":
                line_controls.append(ft.Container(height=6))
                continue

            stripped = raw_line.strip()
            if stripped.startswith("\u5f15\u7528"):
                label, sep, remainder = stripped.partition(":")
                bullet = ft.Text("\u2022", size=size or 13, color=icon_color)
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

        message_bubble = ft.Container(
            content=ft.Column(content_controls, spacing=6, tight=True),
            bgcolor=config.get("bgcolor"),
            border_radius=16,
            padding=16,
            expand=True,
            shadow=ft.BoxShadow(
                spread_radius=1,
                blur_radius=18,
                color="#33000000",
                offset=ft.Offset(2, 4),
            ),
        )

        icon_name = config.get("icon", ft.Icons.SMART_BUTTON)
        icon_color = config.get("icon_color", "#CFD8DC")
        icon = ft.Icon(name=icon_name, color=icon_color, size=20)
        icon_container = ft.Container(icon, margin=ft.margin.only(right=8, left=8, top=3))

        bubble_and_icon_row = ft.Row(
            [icon_container, message_bubble] if msg_type_value != "user" else [message_bubble, icon_container],
            vertical_alignment=ft.CrossAxisAlignment.START,
        )

        if msg_type_value == "user":
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
