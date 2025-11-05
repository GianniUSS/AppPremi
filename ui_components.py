"""Componenti UI riutilizzabili."""
from __future__ import annotations

from typing import Callable, Optional
import tkinter as tk

from config import COLORS, FONTS


def create_button(
    parent: tk.Widget,
    text: str,
    command: Callable[[], None],
    variant: str = "primary",
    *,
    width: Optional[int] = None,
) -> tk.Button:
    """Crea un pulsante Tk con stile coerente per l'applicazione."""
    button = tk.Button(
        parent,
        text=text,
        command=command,
        font=FONTS["button"],
        cursor="hand2",
        bd=0,
        highlightthickness=0,
        padx=14,
        pady=6,
        disabledforeground=COLORS.get("text_light", "#7F8C8D"),
    )

    if width is not None:
        button.configure(width=width)

    primary_bg = COLORS.get("secondary", "#2980B9")
    primary_active = COLORS.get("primary", "#1F3A93")
    secondary_bg = COLORS.get("white", "#FFFFFF")
    secondary_active = COLORS.get("border", "#BDC3C7")
    danger_bg = COLORS.get("danger", "#E74C3C")
    danger_active = "#C0392B"

    variant = variant.lower()
    if variant == "primary":
        button.configure(
            bg=primary_bg,
            fg=COLORS.get("white", "#FFFFFF"),
            activebackground=primary_active,
            activeforeground=COLORS.get("white", "#FFFFFF"),
            highlightbackground=primary_bg,
        )
    elif variant == "secondary":
        button.configure(
            bg=secondary_bg,
            fg=COLORS.get("primary", "#2C3E50"),
            activebackground=secondary_active,
            activeforeground=COLORS.get("primary", "#2C3E50"),
            highlightbackground=secondary_bg,
            bd=1,
            relief="solid",
        )
    elif variant == "danger":
        button.configure(
            bg=danger_bg,
            fg=COLORS.get("white", "#FFFFFF"),
            activebackground=danger_active,
            activeforeground=COLORS.get("white", "#FFFFFF"),
            highlightbackground=danger_bg,
        )
    else:
        raise ValueError(f"Variante pulsante non supportata: {variant}")

    return button
