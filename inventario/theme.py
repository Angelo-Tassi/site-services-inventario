"""Palette e stili condivisi dall'interfaccia."""

import tkinter.font as tkfont
from tkinter import ttk

BG = "#F3F6F9"          # sfondo finestra
CARD = "#FFFFFF"        # superfici e tabella
CARD_HOVER = "#F0F8F9"  # scheda sotto il puntatore
PRIMARY = "#16344A"     # intestazione e titoli: blu ardesia
PRIMARY_DARK = "#0F283A"
ACCENT = "#087F8C"      # azione principale: verde petrolio
ACCENT_DARK = "#066873"
TEXT = "#223548"
MUTED = "#526779"
BORDER = "#DCE5ED"
HEAD_BG = "#EAF0F5"
HEAD_HOVER = "#DCE8EE"
ROW_ALT = "#F7FAFC"
DISABLED_BG = "#EDF1F5"
DISABLED_FG = "#738394"
# Ambra: e' l'unica tinta che nessuna categoria di riga usa. Con l'azzurro di
# prima una riga selezionata era quasi identica a un tablet Dell, e con venti
# righe accese insieme non si capiva piu' che cosa fosse una spunta e che cosa
# una categoria.
SELECT = "#F6DFAC"
DANGER = "#B03A2E"
LOAN_BG = "#FDEEEC"     # riga di un dispositivo in prestito
LOAN_BG_ALT = "#FAE6E3"
LOAN_FG = "#A93226"
LINK = "#087F8C"        # cella-pulsante nell'elenco
IPHONE_ROW = "#EDF9EF"  # riga di un iPhone
IPHONE_ROW_ALT = "#E5F4E8"
TABLET_ROW = "#EAF4FD"  # riga di un tablet Dell
TABLET_ROW_ALT = "#DFEDFA"
SHIP_ROW = "#F6EFFB"    # riga di un dispositivo rispedito
SHIP_ROW_ALT = "#EFE5F7"
SHIP_FG = "#6C3483"

# Pulsanti della barra, colorati per famiglia: chi sta per premere sa a che
# categoria appartiene il comando prima di leggerlo. Sfondi pastello e testo
# scuro dello stesso tono - il contrasto resta alto, il colore non urla.
# Il cestino: nero pieno, l'unico bottone che non appartiene a nessuna delle tre
# famiglie di colore. Deve staccare, perche' e' l'unico posto da cui si torna
# indietro da un'eliminazione e va trovato senza cercarlo. Bianco su nero: 18:1
# di contrasto, il massimo che il tema puo' dare.
CESTINO_BG = "#1C1C1C"
CESTINO_BG_ON = "#3A3A3A"
CESTINO_FG = "#FFFFFF"

AZIONE_ROSSA_BG = "#FBEAE7"      # ripristino e reset: si tocca l'inventario
AZIONE_ROSSA_BG_ON = "#F6D8D2"
AZIONE_ROSSA_FG = "#96291D"
AZIONE_VERDE_BG = "#E7F4EA"      # esportazioni: i dati escono
AZIONE_VERDE_BG_ON = "#D6EBDC"
AZIONE_VERDE_FG = "#1E6B3A"
AZIONE_ARANCIO_BG = "#FDF0E0"    # importazioni: i dati entrano
AZIONE_ARANCIO_BG_ON = "#F8E2C8"
AZIONE_ARANCIO_FG = "#8A5316"

# Colonne dell'elenco. ttk non permette di colorare una cella o una colonna -
# lo stile vale per tutta la tabella - ma permette di colorare cio' che le
# delimita e le intesta: una riga verticale fra una colonna e l'altra, e una
# barretta nell'intestazione. Bastano a non perdere la colonna mentre si scorre
# un elenco largo, senza toccare la leggibilita' del testo, che resta scuro su
# bianco.
COLORE_COLONNA = {
    "asset_tag": "#1F4E79",       # identificativi: blu istituzionale
    "imei": "#1F4E79",
    "seriale": "#5B8DB8",
    "tipo": "#16A085",            # che cos'e': verde
    "funzione": "#16A085",        # a che cosa serve: e' ancora "che cos'e'"
    "modello": "#16A085",
    "stanza": "#8E44AD",          # dov'e': viola
    "stato": "#C0392B",           # come sta: rosso
    "prestato_a": "#B36B06",      # prestito: ambra scura, per staccare dal bianco
    "prestato_il": "#B36B06",
    "restituito_da": "#6C3483",   # telefoni: lo stesso viola della spedizione
    "spedito_il": "#6C3483",
    "note": "#7F8C8D",            # contorno: grigio
    "modificato_il": "#78888A",
    "modificato_da": "#78888A",
}
COLORE_COLONNA_ALTRO = "#78888A"

ROOM_COLORS = ["#087F8C", "#4567A1", "#8063A5", "#A36B20", "#AC5161"]
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
        "count": (family, 32, "bold"),
        "title": (family, 17, "bold"),
        "subtitle": (family, 9),
        "section": (family, 12, "bold"),
    }
    root.configure(bg=BG)
    # Anche i widget Tk e le tendine dei combobox seguono il font di sistema.
    for nome in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
        tkfont.nametofont(nome, root=root).configure(family=family, size=10)
    root.option_add("*TCombobox*Listbox.font", fonts["base"])
    root.option_add("*TCombobox*Listbox.background", CARD)
    root.option_add("*TCombobox*Listbox.foreground", TEXT)
    root.option_add("*TCombobox*Listbox.selectBackground", PRIMARY)
    root.option_add("*TCombobox*Listbox.selectForeground", "#FFFFFF")

    style = ttk.Style(root)
    if "clam" in style.theme_names():
        style.theme_use("clam")

    style.configure(".", background=BG, foreground=TEXT, font=fonts["base"])
    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD)
    style.configure("Card.TLabel", background=CARD, foreground=MUTED,
                    font=fonts["small"])
    style.configure("Head.TFrame", background=PRIMARY)
    style.configure("TLabel", background=BG, foreground=TEXT, font=fonts["base"])
    style.configure("Muted.TLabel", background=BG, foreground=MUTED, font=fonts["small"])
    style.configure("Section.TLabel", background=BG, foreground=PRIMARY, font=fonts["section"])
    style.configure("HeadTitle.TLabel", background=PRIMARY, foreground="#FFFFFF",
                    font=fonts["title"])
    style.configure("HeadSub.TLabel", background=PRIMARY, foreground="#CBDDE8",
                    font=fonts["subtitle"])
    style.configure("Status.TLabel", background=CARD, foreground=MUTED,
                    font=fonts["small"])

    # Il tema clam permette gli stessi pulsanti piatti su Windows e macOS.
    # Il focus resta visibile per chi usa la tastiera; premuto precede hover.
    style.configure("TButton", background=CARD, foreground=TEXT, font=fonts["base"],
                    borderwidth=0, relief="flat", padding=(12, 8),
                    focuscolor=ACCENT, focusthickness=1)
    style.map("TButton",
              background=[("disabled", DISABLED_BG), ("pressed", HEAD_HOVER),
                          ("active", HEAD_BG)],
              foreground=[("disabled", DISABLED_FG)],
              relief=[("pressed", "flat")])
    style.configure("Primary.TButton", background=ACCENT, foreground="#FFFFFF",
                    font=fonts["bold"], borderwidth=0, padding=(14, 8),
                    focuscolor="#FFFFFF")
    style.map("Primary.TButton",
              background=[("disabled", DISABLED_BG), ("pressed", ACCENT_DARK),
                          ("active", ACCENT_DARK)],
              foreground=[("disabled", DISABLED_FG)])
    for nome, sfondo, acceso, testo in (
        ("Rosso", AZIONE_ROSSA_BG, AZIONE_ROSSA_BG_ON, AZIONE_ROSSA_FG),
        ("Verde", AZIONE_VERDE_BG, AZIONE_VERDE_BG_ON, AZIONE_VERDE_FG),
        ("Arancio", AZIONE_ARANCIO_BG, AZIONE_ARANCIO_BG_ON, AZIONE_ARANCIO_FG),
    ):
        style.configure("%s.TButton" % nome, background=sfondo, foreground=testo,
                        font=fonts["bold"], borderwidth=0, relief="flat",
                        padding=(12, 8))
        style.map("%s.TButton" % nome,
                  background=[("disabled", DISABLED_BG), ("pressed", acceso), ("active", acceso)],
                  foreground=[("disabled", DISABLED_FG), ("active", testo), ("pressed", testo)],
                  bordercolor=[("!disabled", acceso)])

    # Piu' grande e in stampatello: e' un pulsante che si cerca con lo sguardo
    # quando ci si accorge di aver eliminato la riga sbagliata.
    style.configure("Cestino.TButton", background=CESTINO_BG, foreground=CESTINO_FG,
                    font=(family, 11, "bold"), borderwidth=0, relief="flat",
                    padding=(12, 8))
    style.map("Cestino.TButton",
              background=[("disabled", DISABLED_BG), ("pressed", CESTINO_BG_ON), ("active", CESTINO_BG_ON)],
              foreground=[("disabled", DISABLED_FG), ("active", CESTINO_FG), ("pressed", CESTINO_FG)])

    style.configure("Ghost.TButton", background=BG, foreground=PRIMARY,
                    font=fonts["bold"], borderwidth=0, padding=(8, 8))
    style.map("Ghost.TButton",
              background=[("pressed", HEAD_HOVER), ("active", HEAD_BG)],
              foreground=[("disabled", DISABLED_FG), ("active", ACCENT)])

    style.configure("Inv.Treeview", background=CARD, fieldbackground=CARD,
                    foreground=TEXT, rowheight=38, borderwidth=0, font=fonts["base"])
    style.configure("Inv.Treeview.Heading", background=HEAD_BG, foreground=PRIMARY,
                    font=fonts["bold"], relief="flat", padding=(10, 10))
    style.map("Inv.Treeview.Heading", background=[("pressed", HEAD_HOVER), ("active", HEAD_HOVER)])
    # Solo lo sfondo: senza mappare il colore del testo, una riga in prestito
    # selezionata conserva il suo rosso e una spedita il suo viola, cosi' le due
    # informazioni piu' urgenti attraversano la selezione.
    style.map("Inv.Treeview", background=[("selected", SELECT)])
    # Pulsanti veri disegnati sopra le righe dell'elenco.
    style.configure("Row.TButton", background=LINK, foreground="#FFFFFF",
                    font=(family, 9, "bold"), borderwidth=0, relief="flat",
                    padding=(6, 2), focuscolor=LINK)
    style.map("Row.TButton", background=[("pressed", ACCENT_DARK), ("active", ACCENT_DARK)])
    style.configure("RowShip.TButton", background=SHIP_FG, foreground="#FFFFFF",
                    font=(family, 9, "bold"), borderwidth=0, relief="flat",
                    padding=(6, 2), focuscolor=SHIP_FG)
    style.map("RowShip.TButton", background=[("active", "#552D69"), ("pressed", "#552D69")])
    style.configure("RowBack.TButton", background=LOAN_FG, foreground="#FFFFFF",
                    font=(family, 9, "bold"), borderwidth=0, relief="flat",
                    padding=(6, 2), focuscolor=LOAN_FG)
    style.map("RowBack.TButton", background=[("active", "#8E2A21"), ("pressed", "#8E2A21")])

    style.configure("TCombobox", fieldbackground=CARD, background=HEAD_BG,
                    foreground=TEXT, bordercolor=BORDER, arrowcolor=PRIMARY,
                    lightcolor=CARD, darkcolor=CARD, padding=(8, 7))
    style.map("TCombobox", fieldbackground=[("readonly", CARD)],
              bordercolor=[("focus", ACCENT)],
              foreground=[("disabled", DISABLED_FG), ("readonly", TEXT)],
              selectbackground=[("readonly", PRIMARY)],
              selectforeground=[("readonly", "#FFFFFF")])
    style.configure("TEntry", fieldbackground=CARD, foreground=TEXT,
                    bordercolor=BORDER, lightcolor=CARD, darkcolor=CARD,
                    padding=(8, 7))
    style.map("TEntry", bordercolor=[("focus", ACCENT)],
              foreground=[("disabled", DISABLED_FG)])
    style.configure("TCheckbutton", background=BG, foreground=TEXT, padding=(0, 4))
    style.configure("TRadiobutton", background=BG, foreground=TEXT, padding=(0, 4))
    style.configure("TScrollbar", background=HEAD_BG, troughcolor=BG,
                    borderwidth=0, arrowsize=14, arrowcolor=MUTED,
                    relief="flat")
    style.map("TScrollbar", background=[("pressed", BORDER), ("active", BORDER)])
    style.configure("TSeparator", background=BORDER)
    return fonts
