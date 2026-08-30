"""Palette e stili condivisi dall'interfaccia."""

import tkinter.font as tkfont
from tkinter import ttk

BG = "#EEF2F7"          # sfondo finestra
CARD = "#FFFFFF"        # sfondo schede e tabella
PRIMARY = "#1F4E79"     # blu istituzionale (lo stesso dei file Excel)
PRIMARY_DARK = "#163A5B"
ACCENT = "#2E86C1"
TEXT = "#1B2733"
MUTED = "#6B7A8C"
BORDER = "#D7E0EA"
HEAD_BG = "#E7EEF6"
ROW_ALT = "#F6F9FC"
SELECT = "#CFE4F6"
DANGER = "#B03A2E"
LOAN_BG = "#FBE3E1"     # riga di un dispositivo in prestito
LOAN_FG = "#A93226"
LINK = "#1A6BA8"        # cella-pulsante nell'elenco
IPHONE_ROW = "#E2F4E4"  # riga di un iPhone
IPHONE_ROW_ALT = "#D8EFDB"
TABLET_ROW = "#FCEEDA"  # riga di un tablet Dell
TABLET_ROW_ALT = "#F8E5CB"

ROOM_COLORS = ["#2E86C1", "#16A085", "#8E44AD", "#D68910", "#C0392B"]
IPHONE_COLOR = "#5D6D7E"   # scheda di comodo, non una stanza vera


def pick_family(root):
    """Il font di sistema piu' adatto fra quelli disponibili."""
    available = set(tkfont.families(root))
    for name in ("Segoe UI", "SF Pro Text", "Helvetica Neue", "DejaVu Sans", "Arial"):
        if name in available:
            return name
    return "TkDefaultFont"


def apply(root):
    """Applica il tema alla finestra e ritorna i font usati."""
    family = pick_family(root)
    fonts = {
        "base": (family, 10),
        "small": (family, 9),
        "bold": (family, 10, "bold"),
        "card_title": (family, 11, "bold"),
        "count": (family, 30, "bold"),
        "title": (family, 16, "bold"),
        "subtitle": (family, 9),
        "section": (family, 12, "bold"),
    }
    root.configure(bg=BG)

    style = ttk.Style(root)
    if "clam" in style.theme_names():
        style.theme_use("clam")

    style.configure(".", background=BG, foreground=TEXT, font=fonts["base"])
    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD)
    style.configure("Head.TFrame", background=PRIMARY)
    style.configure("TLabel", background=BG, foreground=TEXT, font=fonts["base"])
    style.configure("Muted.TLabel", background=BG, foreground=MUTED, font=fonts["small"])
    style.configure("Section.TLabel", background=BG, foreground=PRIMARY, font=fonts["section"])
    style.configure("HeadTitle.TLabel", background=PRIMARY, foreground="#FFFFFF",
                    font=fonts["title"])
    style.configure("HeadSub.TLabel", background=PRIMARY, foreground="#BBD4EA",
                    font=fonts["subtitle"])
    style.configure("Status.TLabel", background="#E3E9F0", foreground=MUTED,
                    font=fonts["small"])

    style.configure("TButton", background=CARD, foreground=TEXT, font=fonts["base"],
                    borderwidth=1, relief="flat", padding=(12, 6))
    style.map("TButton",
              background=[("active", HEAD_BG), ("pressed", HEAD_BG)],
              bordercolor=[("!disabled", BORDER)])
    style.configure("Primary.TButton", background=PRIMARY, foreground="#FFFFFF",
                    font=fonts["bold"], borderwidth=0, padding=(14, 6))
    style.map("Primary.TButton", background=[("active", PRIMARY_DARK),
                                             ("pressed", PRIMARY_DARK)])
    style.configure("Ghost.TButton", background=BG, foreground=PRIMARY,
                    font=fonts["bold"], borderwidth=0, padding=(6, 4))
    style.map("Ghost.TButton", background=[("active", BG)], foreground=[("active", ACCENT)])

    style.configure("Inv.Treeview", background=CARD, fieldbackground=CARD,
                    foreground=TEXT, rowheight=32, borderwidth=0, font=fonts["base"])
    style.configure("Inv.Treeview.Heading", background=HEAD_BG, foreground=PRIMARY,
                    font=fonts["bold"], relief="flat", padding=(8, 7))
    style.map("Inv.Treeview.Heading", background=[("active", "#DAE5F0")])
    style.map("Inv.Treeview", background=[("selected", SELECT)],
              foreground=[("selected", TEXT)])
    # Pulsanti veri disegnati sopra le righe dell'elenco.
    style.configure("Row.TButton", background=LINK, foreground="#FFFFFF",
                    font=(family, 9, "bold"), borderwidth=0, relief="flat",
                    padding=(6, 2), focuscolor=LINK)
    style.map("Row.TButton", background=[("active", PRIMARY), ("pressed", PRIMARY)])
    style.configure("RowBack.TButton", background=LOAN_FG, foreground="#FFFFFF",
                    font=(family, 9, "bold"), borderwidth=0, relief="flat",
                    padding=(6, 2), focuscolor=LOAN_FG)
    style.map("RowBack.TButton", background=[("active", "#8E2A21"), ("pressed", "#8E2A21")])

    style.configure("TCombobox", fieldbackground=CARD, background=CARD,
                    bordercolor=BORDER, arrowcolor=PRIMARY)
    style.configure("TEntry", fieldbackground=CARD, bordercolor=BORDER)
    style.configure("TSeparator", background=BORDER)
    return fonts
