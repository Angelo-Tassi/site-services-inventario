"""Interfaccia grafica dell'inventario (Tkinter).

Due schermate:
  * Home    - schede delle stanze con il numero di dispositivi + inventario completo
  * Stanza  - inventario della singola stanza, aperto cliccando una scheda
"""

import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from . import config, excel_io, theme
from .store import (ALL_FIELDS, DA_RISPEDIRE, HEADERS, InventoryError,
                    MESI_CONSERVAZIONE,
                    InventoryStore, NON_DISPONIBILE, SPEDITO, clean,
                    is_iphone, is_on_loan, is_shipped, new_item, norm_tag,
                    puo_essere_eliminato, rows_from_workbook, testo_spedizione)

NO_ROOM = "(senza stanza)"
CHECK_COLUMN = "_sel"
ACTION_COLUMN = "_azione"
CHECK_ON = "\u25c9"      # cerchio pieno: riga selezionata
CHECK_OFF = "\u25cb"     # cerchio vuoto
COLUMN_WIDTHS = {CHECK_COLUMN: 46, ACTION_COLUMN: 150, "asset_tag": 120, "tipo": 75, "modello": 185,
                 "seriale": 120, "imei": 130, "restituito_da": 135, "stanza": 160,
                 "stato": 105, "prestato_a": 140, "prestato_il": 120, "spedito_il": 120, "note": 180,
                 "modificato_il": 120, "modificato_da": 145}
REFRESH_MS = 15000


# --------------------------------------------------------------- dialoghi


class _Modal(tk.Toplevel):
    def __init__(self, parent, title):
        tk.Toplevel.__init__(self, parent)
        self.title(title)
        self.configure(bg=theme.BG)
        self.resizable(False, False)
        self.transient(parent)
        self.result = None
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.bind("<Escape>", lambda e: self._cancel())

    def _cancel(self):
        self.result = None
        self.destroy()

    def show(self):
        self.update_idletasks()
        parent = self.master
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 3
        self.geometry("+%d+%d" % (max(x, 0), max(y, 0)))
        self.grab_set()
        self.wait_window(self)
        return self.result


class ItemDialog(_Modal):
    """Scheda di inserimento / modifica di un dispositivo.

    Il tipo sta in cima perche' decide la forma del modulo: per un iPhone al
    posto di asset tag e numero di serie si chiedono IMEI e chi lo ha restituito.
    """

    def __init__(self, parent, rooms, types, item=None, iphone_room=None, stati=None):
        _Modal.__init__(self, parent, "Modifica dispositivo" if item else "Nuovo dispositivo")
        item = item or new_item(stanza=rooms[0] if rooms else "")
        self._loan = (item.get("prestato_a", ""), item.get("prestato_il", ""))
        self.rooms = rooms
        self.iphone_room = iphone_room or (rooms[0] if rooms else "")
        self.stati = list(stati or [])

        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)

        self.var_tipo = tk.StringVar(value=item.get("tipo") or (types[0] if types else ""))
        self.var_tag = tk.StringVar(value=item.get("asset_tag", ""))
        self.var_modello = tk.StringVar(value=item.get("modello", ""))
        self.var_seriale = tk.StringVar(value=item.get("seriale", ""))
        self.var_imei = tk.StringVar(value=item.get("imei") or
                                     (item.get("asset_tag", "") if is_iphone(item.get("tipo")) else ""))
        self.var_restituito = tk.StringVar(value=item.get("restituito_da", ""))
        self.var_stanza = tk.StringVar(value=item.get("stanza", ""))
        self.var_stato = tk.StringVar(value=item.get("stato") or
                                      (self.stati[0] if self.stati else ""))

        # --- il tipo, in cima: cambiandolo cambia il modulo
        ttk.Label(body, text="Tipo *").grid(row=0, column=0, sticky="w", pady=5)
        combo_tipo = ttk.Combobox(body, textvariable=self.var_tipo, values=types,
                                  state="readonly", width=32)
        combo_tipo.grid(row=0, column=1, sticky="we", pady=5)
        combo_tipo.bind("<<ComboboxSelected>>", lambda e: self._build_fields())
        ttk.Separator(body, orient="horizontal").grid(
            row=1, column=0, columnspan=2, sticky="we", pady=(6, 10))

        self.fields = ttk.Frame(body)
        self.fields.grid(row=2, column=0, columnspan=2, sticky="nsew")
        self.fields.columnconfigure(1, weight=1)

        ttk.Label(body, style="Muted.TLabel",
                  text="I campi contrassegnati con * sono obbligatori.").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))

        buttons = ttk.Frame(body)
        buttons.grid(row=4, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(buttons, text="Annulla", command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text="Salva", style="Primary.TButton",
                   command=self._ok).pack(side="right")

        self.text_note = None
        self._nota = item.get("note", "")
        self.required = []
        self._build_fields()
        self.bind("<Return>", lambda e: self._ok())

    # ------------------------------------------------------------ modulo

    def is_iphone(self):
        return is_iphone(self.var_tipo.get())

    def _build_fields(self):
        """(Ri)disegna i campi in base al tipo selezionato, conservando i valori."""
        nota = self.text_note.get("1.0", "end").strip() if self.text_note else self._nota_iniziale()
        for child in self.fields.winfo_children():
            child.destroy()

        if self.is_iphone():
            righe = [("IMEI *", self.var_imei),
                     ("Modello *", self.var_modello),
                     ("Restituito da *", self.var_restituito)]
        else:
            righe = [("Asset Tag *", self.var_tag),
                     ("Modello *", self.var_modello),
                     ("Numero di serie *", self.var_seriale)]

        self.required = []
        for riga, (etichetta, var) in enumerate(righe):
            ttk.Label(self.fields, text=etichetta).grid(row=riga, column=0, sticky="w", pady=5)
            entry = ttk.Entry(self.fields, textvariable=var, width=34)
            entry.grid(row=riga, column=1, sticky="we", pady=5)
            self.required.append((etichetta.rstrip(" *"), var, entry))

        riga = len(righe)
        ttk.Label(self.fields, text="Stanza *").grid(row=riga, column=0, sticky="w", pady=5)
        if self.is_iphone():
            # Gli iPhone appartengono sempre alla loro stanza: campo mostrato ma bloccato.
            self.var_stanza.set(self.iphone_room)
            combo = ttk.Combobox(self.fields, textvariable=self.var_stanza,
                                 values=[self.iphone_room], state="disabled", width=32)
        else:
            combo = ttk.Combobox(self.fields, textvariable=self.var_stanza, values=self.rooms,
                                 state="readonly", width=32)
        combo.grid(row=riga, column=1, sticky="we", pady=5)
        self.required.append(("Stanza", self.var_stanza, combo))
        riga += 1
        if self.is_iphone():
            ttk.Label(self.fields, style="Muted.TLabel",
                      text="Gli iPhone restano sempre in %s." % self.iphone_room).grid(
                row=riga, column=1, sticky="w")
            riga += 1

        ttk.Label(self.fields, text="Stato").grid(row=riga, column=0, sticky="w", pady=5)
        if self.is_iphone():
            self.var_stato.set(DA_RISPEDIRE)
            valori, stato_widget = [DA_RISPEDIRE], "disabled"
        elif self._loan[0]:
            self.var_stato.set(NON_DISPONIBILE)
            valori, stato_widget = [NON_DISPONIBILE], "disabled"
        else:
            if self.var_stato.get() not in self.stati and self.stati:
                self.var_stato.set(self.stati[0])
            valori, stato_widget = self.stati, "readonly"
        ttk.Combobox(self.fields, textvariable=self.var_stato, values=valori,
                     state=stato_widget, width=32).grid(row=riga, column=1,
                                                        sticky="we", pady=5)
        riga += 1
        if stato_widget == "disabled":
            motivo = ("Gli iPhone sono sempre \"%s\"." % DA_RISPEDIRE if self.is_iphone()
                      else "In prestito: lo stato torna modificabile dopo il rientro.")
            ttk.Label(self.fields, style="Muted.TLabel", text=motivo).grid(
                row=riga, column=1, sticky="w")
            riga += 1

        ttk.Label(self.fields, text="Note").grid(row=riga, column=0, sticky="nw", pady=5)
        self.text_note = tk.Text(self.fields, width=34, height=4, wrap="word",
                                 relief="solid", borderwidth=1, highlightthickness=0,
                                 bg=theme.CARD, fg=theme.TEXT)
        self.text_note.insert("1.0", nota)
        self.text_note.grid(row=riga, column=1, sticky="we", pady=5)

        self.required[0][2].focus_set()

    def _nota_iniziale(self):
        return getattr(self, "_nota", "")

    # ------------------------------------------------------------ salvataggio

    def missing_fields(self):
        """Campi obbligatori ancora vuoti, nell'ordine in cui compaiono."""
        return [label for label, var, _w in self.required if not var.get().strip()]

    def _ok(self):
        missing = self.missing_fields()
        if missing:
            messagebox.showerror(
                "Dati mancanti",
                "Il dispositivo non e\' stato inserito.\n"
                "Mancano questi dati obbligatori:\n\n%s\n\n"
                "Compilali e premi di nuovo Salva."
                % "\n".join("  \u2022  " + label for label in missing),
                parent=self)
            for label, var, widget in self.required:
                if not var.get().strip():
                    widget.focus_set()
                    break
            return

        comuni = dict(
            tipo=self.var_tipo.get(),
            modello=self.var_modello.get(),
            stanza=self.var_stanza.get(),
            note=self.text_note.get("1.0", "end"),
            stato=self.var_stato.get(),
            prestato_a=self._loan[0],
            prestato_il=self._loan[1],
        )
        if self.is_iphone():
            # per un iPhone l'IMEI e' l'identificativo del dispositivo
            imei = norm_tag(self.var_imei.get())
            self.result = new_item(asset_tag=imei, imei=imei,
                                   restituito_da=self.var_restituito.get(), **comuni)
        else:
            self.result = new_item(asset_tag=norm_tag(self.var_tag.get()),
                                   seriale=self.var_seriale.get(), **comuni)
        self.destroy()


class RoomsDialog(_Modal):
    """Impostazioni condivise: nomi delle stanze e tipi di dispositivo."""

    def __init__(self, parent, rooms, types, loan_rooms, iphone_room=""):
        _Modal.__init__(self, parent, "Impostazioni inventario")
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Stanze (una per riga)").grid(row=0, column=0, sticky="w")
        ttk.Label(body, text="Tipi di dispositivo").grid(row=0, column=1, sticky="w", padx=(14, 0))
        ttk.Label(body, text="Stanze con prestito").grid(row=0, column=2, sticky="w", padx=(14, 0))
        opts = dict(relief="solid", borderwidth=1, highlightthickness=0,
                    bg=theme.CARD, fg=theme.TEXT)
        self.text_rooms = tk.Text(body, width=26, height=8, **opts)
        self.text_rooms.insert("1.0", "\n".join(rooms))
        self.text_rooms.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        self.text_types = tk.Text(body, width=18, height=8, **opts)
        self.text_types.insert("1.0", "\n".join(types))
        self.text_types.grid(row=1, column=1, sticky="nsew", padx=(14, 0), pady=(4, 0))
        self.text_loans = tk.Text(body, width=26, height=8, **opts)
        self.text_loans.insert("1.0", "\n".join(loan_rooms))
        self.text_loans.grid(row=1, column=2, sticky="nsew", padx=(14, 0), pady=(4, 0))
        riga_iphone = ttk.Frame(body)
        riga_iphone.grid(row=2, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Label(riga_iphone, text="Stanza degli iPhone").pack(side="left")
        self.var_iphone_room = tk.StringVar(value=iphone_room or (rooms[0] if rooms else ""))
        ttk.Combobox(riga_iphone, textvariable=self.var_iphone_room, values=rooms,
                     width=30).pack(side="left", padx=(10, 0))

        ttk.Label(
            body, style="Muted.TLabel",
            text="Le impostazioni sono salvate accanto al file dati e valgono per tutti gli utenti.\n"
                 "Nelle stanze con prestito ogni riga dell'elenco ha il pulsante Presta / Registra rientro.\n"
                 "Gli iPhone vengono registrati sempre nella stanza indicata qui sopra e non si spostano.",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
        buttons = ttk.Frame(body)
        buttons.grid(row=4, column=0, columnspan=3, sticky="e", pady=(16, 0))
        ttk.Button(buttons, text="Annulla", command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text="Salva", style="Primary.TButton",
                   command=self._ok).pack(side="right")

    def _ok(self):
        rooms = [r.strip() for r in self.text_rooms.get("1.0", "end").splitlines() if r.strip()]
        types = [t.strip() for t in self.text_types.get("1.0", "end").splitlines() if t.strip()]
        loans = [r.strip() for r in self.text_loans.get("1.0", "end").splitlines() if r.strip()]
        if not rooms:
            messagebox.showwarning("Dato mancante", "Indica almeno una stanza.", parent=self)
            return
        unknown = [r for r in loans if r not in rooms]
        if unknown:
            messagebox.showwarning(
                "Stanza sconosciuta",
                "Queste stanze con prestito non sono nell'elenco delle stanze:\n%s"
                % ", ".join(unknown), parent=self)
            return
        stanza_iphone = self.var_iphone_room.get().strip()
        if stanza_iphone and stanza_iphone not in rooms:
            messagebox.showwarning(
                "Stanza sconosciuta",
                "La stanza degli iPhone (%s) non e' nell'elenco delle stanze."
                % stanza_iphone, parent=self)
            return
        self.result = {"rooms": rooms, "types": types or ["Laptop", "Tablet"],
                       "loan_rooms": loans,
                       "iphone_room": stanza_iphone or rooms[0]}
        self.destroy()


class ImportDialog(_Modal):
    def __init__(self, parent, path, count, skipped):
        _Modal.__init__(self, parent, "Importa inventario")
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=os.path.basename(path),
                  style="Section.TLabel").pack(anchor="w")
        msg = "%d righe valide trovate." % count
        if skipped:
            msg += "  %d righe ignorate (asset tag mancante)." % skipped
        ttk.Label(body, text=msg, style="Muted.TLabel").pack(anchor="w", pady=(2, 12))
        self.var_mode = tk.StringVar(value="merge")
        ttk.Radiobutton(body, variable=self.var_mode, value="merge",
                        text="Unisci: aggiunge i nuovi e aggiorna quelli con lo stesso asset tag"
                        ).pack(anchor="w")
        ttk.Radiobutton(body, variable=self.var_mode, value="replace",
                        text="Sostituisci: cancella l'inventario attuale e carica solo questi dati"
                        ).pack(anchor="w", pady=(4, 0))
        buttons = ttk.Frame(body)
        buttons.pack(anchor="e", pady=(16, 0))
        ttk.Button(buttons, text="Annulla", command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text="Importa", style="Primary.TButton",
                   command=self._ok).pack(side="right")

    def _ok(self):
        self.result = self.var_mode.get()
        self.destroy()


# ------------------------------------------------------------ scheda stanza


class RoomCard(tk.Frame):
    """Riquadro cliccabile con il conteggio dei dispositivi di una stanza."""

    def __init__(self, parent, fonts, name, count, breakdown, color, command,
                 note="", note_color=None):
        tk.Frame.__init__(self, parent, bg=theme.CARD, highlightthickness=1,
                          highlightbackground=theme.BORDER, cursor="hand2")
        self.command = command
        stripe = tk.Frame(self, bg=color, height=4)
        stripe.pack(fill="x")
        inner = tk.Frame(self, bg=theme.CARD, padx=18, pady=14)
        inner.pack(fill="both", expand=True)

        self.labels = []
        title = tk.Label(inner, text=name, bg=theme.CARD, fg=theme.TEXT,
                         font=fonts["card_title"], anchor="w")
        title.pack(fill="x")
        number = tk.Label(inner, text=str(count), bg=theme.CARD, fg=color,
                          font=fonts["count"], anchor="w")
        number.pack(fill="x")
        sub = tk.Label(inner, text=breakdown, bg=theme.CARD, fg=theme.MUTED,
                       font=fonts["small"], anchor="w")
        sub.pack(fill="x")
        self.note = tk.Label(inner, text=note, bg=theme.CARD,
                             fg=note_color or theme.MUTED, font=fonts["small"], anchor="w")
        if note:
            self.note.pack(fill="x", pady=(4, 0))
        link = tk.Label(inner, text="Apri l'inventario  ›", bg=theme.CARD,
                        fg=theme.ACCENT, font=fonts["bold"], anchor="w")
        link.pack(fill="x", pady=(10, 0))

        self.labels = [inner, title, number, sub, self.note, link]
        for widget in [self] + self.labels:
            widget.bind("<Button-1>", lambda e: self.command())
            widget.bind("<Enter>", self._enter)
            widget.bind("<Leave>", self._leave)
            widget.configure(cursor="hand2")

    def _enter(self, _event=None):
        self.configure(highlightbackground=theme.ACCENT)

    def _leave(self, _event=None):
        self.configure(highlightbackground=theme.BORDER)


# ------------------------------------------------------------ finestra principale


class App(tk.Tk):
    def __init__(self, data_path):
        tk.Tk.__init__(self)
        self.title('Site Services : Inventario Iphone, Laptop e Tablet')
        self.geometry("1220x720")
        self.minsize(980, 560)

        self.fonts = theme.apply(self)
        self.cfg = config.load_shared_config(data_path)
        self.store = InventoryStore(data_path,
                                    iphone_room=self.cfg.get("iphone_room"),
                                    stati=self.cfg.get("states"))
        self.sort_field = "asset_tag"
        self.sort_reverse = False
        self.visible = []
        self.view = "home"
        self.tree = None

        self._build_header()
        self._build_toolbar()
        self._build_filters()
        self.body = ttk.Frame(self, padding=(16, 4, 16, 8))
        self.body.pack(fill="both", expand=True)
        self._build_status()
        self._bind_keys()

        self.after(100, self._initial_load)
        self.after(REFRESH_MS, self._auto_refresh)

    # ------------------------------------------------------------ layout

    def _build_header(self):
        head = ttk.Frame(self, style="Head.TFrame", padding=(20, 14))
        head.pack(fill="x")
        left = ttk.Frame(head, style="Head.TFrame")
        left.pack(side="left")
        ttk.Label(left, text='Site Services : Inventario Iphone, Laptop e Tablet',
                  style="HeadTitle.TLabel").pack(anchor="w")
        self.var_subtitle = tk.StringVar(value="Laptop e tablet in nostro possesso")
        ttk.Label(left, textvariable=self.var_subtitle,
                  style="HeadSub.TLabel").pack(anchor="w", pady=(2, 0))
        self.var_head_count = tk.StringVar(value="")
        ttk.Label(head, textvariable=self.var_head_count,
                  style="HeadSub.TLabel").pack(side="right")

    def _build_toolbar(self):
        bar = ttk.Frame(self, padding=(16, 12, 16, 4))
        bar.pack(fill="x")
        self.btn_home = ttk.Button(bar, text="‹  Home", style="Ghost.TButton",
                                   command=self.show_home)
        self.btn_home.pack(side="left", padx=(0, 10))
        ttk.Button(bar, text="Nuovo", style="Primary.TButton",
                   command=self.on_new).pack(side="left", padx=(0, 6))
        for text, command in (
            ("Modifica", self.on_edit),
            ("Elimina", self.on_delete),
            ("Sposta in stanza...", self.on_move),
        ):
            ttk.Button(bar, text=text, command=command).pack(side="left", padx=(0, 6))
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=10)
        for text, command in (
            ("Importa xls...", self.on_import),
            ("Esporta xls...", self.on_export),
            ("Stampa", self.on_print),
        ):
            ttk.Button(bar, text=text, command=command).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="Impostazioni", command=self.on_settings).pack(side="right")
        ttk.Button(bar, text="Aggiorna", command=self.on_refresh).pack(side="right", padx=6)

    def _build_filters(self):
        bar = ttk.Frame(self, padding=(16, 4))
        bar.pack(fill="x")
        ttk.Label(bar, text="Cerca").pack(side="left")
        self.var_search = tk.StringVar()
        self.var_search.trace_add("write", lambda *a: self.refresh_table())
        self.entry_search = ttk.Entry(bar, textvariable=self.var_search, width=34)
        self.entry_search.pack(side="left", padx=(6, 16))

        self.label_room = ttk.Label(bar, text="Stanza")
        self.label_room.pack(side="left")
        self.var_room = tk.StringVar(value="Tutte")
        self.combo_room = ttk.Combobox(bar, textvariable=self.var_room, state="readonly", width=20)
        self.combo_room.pack(side="left", padx=(6, 16))
        self.combo_room.bind("<<ComboboxSelected>>", self._on_room_filter)

        ttk.Label(bar, text="Tipo").pack(side="left")
        self.var_type = tk.StringVar(value="Tutti")
        self.combo_type = ttk.Combobox(bar, textvariable=self.var_type, state="readonly", width=14)
        self.combo_type.pack(side="left", padx=(6, 16))
        self.combo_type.bind("<<ComboboxSelected>>", self._on_type_filter)

        ttk.Button(bar, text="Azzera filtri", command=self.reset_filters).pack(side="left")

    def _build_status(self):
        self.var_status = tk.StringVar()
        ttk.Separator(self).pack(fill="x")
        ttk.Label(self, textvariable=self.var_status, style="Status.TLabel",
                  anchor="w", padding=(16, 6)).pack(fill="x")

    def _bind_keys(self):
        self.bind("<Control-n>", lambda e: self.on_new())
        self.bind("<Control-f>", lambda e: self.entry_search.focus_set())
        self.bind("<Control-p>", lambda e: self.on_print())
        self.bind("<F5>", lambda e: self.on_refresh())
        self.bind("<Escape>", lambda e: self.show_home())
        self.bind("<Delete>", lambda e: self.on_delete())

    # ------------------------------------------------------------ tabella

    @staticmethod
    def is_dell_tablet(item):
        """Tablet di marca Dell: riconosciuto dal modello."""
        return (clean(item.get("tipo")).lower() == "tablet"
                and "dell" in clean(item.get("modello")).lower())

    def row_tag(self, item, dispari):
        """Il colore della riga: prestito, poi tipo, poi la banda alternata."""
        if is_on_loan(item):
            return "loan"
        if is_shipped(item):
            return "spedito_alt" if dispari else "spedito"
        if is_iphone(item.get("tipo")):
            return "iphone_alt" if dispari else "iphone"
        if self.is_dell_tablet(item):
            return "tablet_alt" if dispari else "tablet"
        return "odd" if dispari else ""

    def iphone_type(self):
        """Il tipo iPhone cosi' come e' scritto nelle impostazioni, o None."""
        for tipo in self.cfg.get("types", []):
            if is_iphone(tipo):
                return tipo
        for item in self.store.items:
            if is_iphone(item.get("tipo")):
                return item["tipo"]
        return None

    def iphone_room(self):
        return self.cfg.get("iphone_room") or ""

    def loan_column_visible(self):
        """La colonna Prestito esiste solo dentro una stanza che gestisce prestiti."""
        return (self.view == "room"
                and self.var_room.get() in self.cfg.get("loan_rooms", []))

    def ship_column_visible(self):
        """La colonna Spedizione esiste solo nel contenitore degli iPhone."""
        return self.view == "type" and is_iphone(self.var_type.get())

    def action_column_visible(self):
        return self.loan_column_visible() or self.ship_column_visible()

    def can_lend(self, item):
        return item.get("stanza") in self.cfg.get("loan_rooms", [])

    def _columns(self):
        colonne = [CHECK_COLUMN]
        if self.action_column_visible():
            colonne.append(ACTION_COLUMN)
        return colonne + list(ALL_FIELDS)

    def _make_table(self, parent):
        wrap = tk.Frame(parent, bg=theme.CARD, highlightthickness=1,
                        highlightbackground=theme.BORDER)
        wrap.pack(fill="both", expand=True)
        columns = self._columns()
        tree = ttk.Treeview(wrap, columns=columns, show="headings",
                            selectmode="browse", style="Inv.Treeview")
        for field in columns:
            if field == CHECK_COLUMN:
                tree.heading(field, text="")
                tree.column(field, width=COLUMN_WIDTHS[field], anchor="center",
                            stretch=False)
                continue
            if field == ACTION_COLUMN:
                tree.heading(field, text="Spedizione" if self.ship_column_visible()
                             else "Prestito")
                tree.column(field, width=COLUMN_WIDTHS[field], anchor="center",
                            stretch=False)
                continue
            arrow = ""
            if field == self.sort_field:
                arrow = "  ▾" if self.sort_reverse else "  ▴"
            tree.heading(field, text=HEADERS[field] + arrow,
                         command=lambda f=field: self.sort_by(f))
            tree.column(field, width=COLUMN_WIDTHS[field], anchor="w",
                        stretch=(field in ("modello", "note")))
        scroll = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)

        def on_scroll(primo, ultimo):
            scroll.set(primo, ultimo)
            self._sync_row_buttons()

        tree.configure(yscrollcommand=on_scroll)
        tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        tree.tag_configure("odd", background=theme.ROW_ALT)
        tree.tag_configure("loan", background=theme.LOAN_BG, foreground=theme.LOAN_FG)
        tree.tag_configure("spedito", background=theme.SHIP_ROW, foreground=theme.SHIP_FG)
        tree.tag_configure("spedito_alt", background=theme.SHIP_ROW_ALT,
                           foreground=theme.SHIP_FG)
        tree.tag_configure("iphone", background=theme.IPHONE_ROW)
        tree.tag_configure("iphone_alt", background=theme.IPHONE_ROW_ALT)
        tree.tag_configure("tablet", background=theme.TABLET_ROW)
        tree.tag_configure("tablet_alt", background=theme.TABLET_ROW_ALT)
        tree.bind("<Button-1>", self._on_click)
        tree.bind("<Double-1>", self._on_double_click)
        tree.bind("<<TreeviewSelect>>", self._on_select)
        tree.bind("<Configure>", lambda e: self._sync_row_buttons())
        return tree

    # ------------------------------------------- pulsanti veri sulle righe

    def _clear_row_buttons(self):
        for button in getattr(self, "_row_buttons", {}).values():
            button.destroy()
        self._row_buttons = {}

    def _sync_row_buttons(self):
        """Disegna un vero pulsante sulla cella Prestito delle righe visibili."""
        if self.tree is None or not self.action_column_visible():
            return
        if not hasattr(self, "_row_buttons"):
            self._row_buttons = {}
        colonne = self._columns()
        if ACTION_COLUMN not in colonne:
            return
        indice = colonne.index(ACTION_COLUMN)
        vive = set()
        for tag in self.tree.get_children():
            item = self._item_by_tag(tag)
            if item is None or not self.action_label(item):
                continue
            try:
                box = self.tree.bbox(tag, indice)
            except tk.TclError:
                box = None
            if not box:                       # riga fuori dalla parte visibile
                continue
            vive.add(tag)
            testo = self.action_label(item)
            if self.ship_column_visible():
                stile = "RowShip.TButton"
            else:
                stile = "RowBack.TButton" if is_on_loan(item) else "Row.TButton"
            button = self._row_buttons.get(tag)
            if button is None:
                button = ttk.Button(self.tree, text=testo, style=stile, takefocus=False,
                                    command=lambda t=tag: self._on_row_button(t))
                self._row_buttons[tag] = button
            else:
                button.configure(text=testo, style=stile,
                                 command=lambda t=tag: self._on_row_button(t))
            x, y, larghezza, altezza = box
            button.place(x=x + 6, y=y + 3, width=max(larghezza - 12, 40),
                         height=max(altezza - 6, 18))
        for tag in [t for t in self._row_buttons if t not in vive]:
            self._row_buttons.pop(tag).destroy()

    def _on_row_button(self, tag):
        item = self._item_by_tag(tag)
        if item is None:
            return
        self.tree.selection_set([tag])
        if self.ship_column_visible():
            self.on_ship(tag)
        elif is_on_loan(item):
            self.on_give_back(tag)
        else:
            self.on_lend(tag)

    # -------------------------------------------------- pulsanti nella riga

    def _cell_at(self, event):
        """(item, nome_colonna) della cella sotto il puntatore, o (None, None)."""
        tree = self.tree
        if tree is None or tree.identify_region(event.x, event.y) != "cell":
            return None, None
        tag = tree.identify_row(event.y)
        column = tree.identify_column(event.x)
        if not tag or not column:
            return None, None
        index = int(column[1:]) - 1
        columns = self._columns()
        if not 0 <= index < len(columns):
            return None, None
        return tag, columns[index]

    def action_label(self, item):
        """Testo del pulsante sulla riga, secondo la schermata in cui siamo."""
        if self.ship_column_visible():
            if not is_iphone(item.get("tipo")) or is_shipped(item):
                return ""
            return "SPEDITO"
        if not self.can_lend(item):
            return ""
        return "Registra rientro" if is_on_loan(item) else "Presta"

    def _on_select(self, _event=None):
        """Tiene il segno di spunta allineato alla riga selezionata."""
        if self.tree is None or CHECK_COLUMN not in self._columns():
            return
        scelti = set(self.tree.selection())
        for tag in self.tree.get_children():
            atteso = CHECK_ON if tag in scelti else CHECK_OFF
            if self.tree.set(tag, CHECK_COLUMN) != atteso:
                self.tree.set(tag, CHECK_COLUMN, atteso)

    def _on_click(self, event):
        tag, column = self._cell_at(event)
        if not tag:
            return None
        if column == CHECK_COLUMN:
            if tag in self.tree.selection():
                self.tree.selection_remove(tag)
            else:
                self.tree.selection_set([tag])
            self._on_select()
            return "break"
        if column == ACTION_COLUMN:
            return "break"          # ci pensa il pulsante disegnato sopra
        return None

    def _on_double_click(self, event):
        tag, column = self._cell_at(event)
        if column in (ACTION_COLUMN, CHECK_COLUMN):
            return "break"
        if column == "note" and tag:
            self.edit_note_inline(tag)
            return "break"
        if column == "stato" and tag:
            self.edit_stato_inline(tag)
            return "break"
        self.on_edit()
        return "break"

    def _item_by_tag(self, tag):
        for item in self.store.items:
            if item["asset_tag"] == tag:
                return item
        return None

    def edit_stato_inline(self, tag):
        """Cambia lo stato con una tendina direttamente nell'elenco, senza popup."""
        item = self._item_by_tag(tag)
        if item is None:
            return
        if is_iphone(item.get("tipo")):
            self._segnala("Lo stato degli iPhone e' sempre \"%s\"." % DA_RISPEDIRE)
            return
        if is_on_loan(item):
            self._segnala("%s e' in prestito a %s: registra prima il rientro."
                          % (tag, item["prestato_a"]))
            return
        stati = list(self.cfg.get("states") or [])
        if not stati:
            return
        colonne = self._columns()
        box = self.tree.bbox(tag, colonne.index("stato"))
        if not box:
            return
        var = tk.StringVar(value=item.get("stato") or stati[0])
        combo = ttk.Combobox(self.tree, textvariable=var, values=stati,
                             state="readonly", font=self.fonts["base"])
        combo.place(x=box[0], y=box[1], width=box[2], height=box[3])
        combo.focus_set()
        fatto = {"chiuso": False}

        def chiudi(salva):
            if fatto["chiuso"]:
                return
            fatto["chiuso"] = True
            scelto = var.get()
            combo.destroy()
            if salva and scelto != item.get("stato"):
                self._run(lambda: self.store.set_stato(tag, scelto),
                          "%s: %s." % (tag, scelto))

        combo.bind("<<ComboboxSelected>>", lambda e: chiudi(True))
        combo.bind("<Escape>", lambda e: chiudi(False))
        combo.bind("<FocusOut>", lambda e: chiudi(False))
        combo.event_generate("<Button-1>")      # apre subito la tendina

    def _segnala(self, messaggio):
        """Avviso discreto nella barra di stato, senza aprire finestre."""
        self.var_status.set(messaggio + "     " + self.var_status.get())

    def edit_note_inline(self, tag):
        """Modifica la nota direttamente nell'elenco."""
        item = self._item_by_tag(tag)
        if item is None:
            return
        columns = self._columns()
        box = self.tree.bbox(tag, columns.index("note"))
        if not box:
            return
        entry = tk.Entry(self.tree, relief="solid", borderwidth=1,
                         highlightthickness=1, highlightcolor=theme.ACCENT,
                         bg=theme.CARD, fg=theme.TEXT, font=self.fonts["base"])
        entry.insert(0, item.get("note", ""))
        entry.select_range(0, "end")
        entry.place(x=box[0], y=box[1], width=box[2], height=box[3])
        entry.focus_set()
        state = {"done": False}

        def close(save):
            if state["done"]:
                return
            state["done"] = True
            text = entry.get()
            entry.destroy()
            if save:
                self._run(lambda: self.store.set_note(tag, text),
                          "Nota aggiornata su %s." % tag)

        entry.bind("<Return>", lambda e: close(True))
        entry.bind("<Escape>", lambda e: close(False))
        entry.bind("<FocusOut>", lambda e: close(True))

    # ------------------------------------------------------------ schermate

    def _clear_body(self):
        self._clear_row_buttons()
        for child in self.body.winfo_children():
            child.destroy()
        self.tree = None

    def show_home(self):
        self.view = "home"
        self.var_room.set("Tutte")
        self.var_subtitle.set("Laptop e tablet in nostro possesso")
        self.btn_home.state(["disabled"])
        self._render()

    def show_room(self, room):
        self.view = "room"
        self.var_room.set(room)
        self.var_type.set("Tutti")
        self.var_subtitle.set("Inventario di %s" % room)
        self.btn_home.state(["!disabled"])
        self._render()

    def show_iphones(self):
        """Vista di comodo con i soli iPhone; restano comunque nella loro stanza."""
        tipo = self.iphone_type()
        if not tipo:
            return
        self.view = "type"
        self.var_room.set("Tutte")
        self.var_type.set(tipo)
        self.var_subtitle.set("Telefoni in nostro possesso - registrati in %s"
                              % self.iphone_room())
        self.btn_home.state(["!disabled"])
        self._render()

    def _render(self):
        self._clear_body()
        if self.view == "home":
            self._render_cards()
            header = ttk.Frame(self.body)
            header.pack(fill="x", pady=(18, 8))
            ttk.Label(header, text="Inventario completo",
                      style="Section.TLabel").pack(side="left")
            self.var_section_count = tk.StringVar()
            ttk.Label(header, textvariable=self.var_section_count,
                      style="Muted.TLabel").pack(side="left", padx=(10, 0))
        else:
            header = ttk.Frame(self.body)
            header.pack(fill="x", pady=(10, 8))
            titolo = self.var_type.get() if self.view == "type" else self.var_room.get()
            ttk.Label(header, text=titolo,
                      style="Section.TLabel").pack(side="left")
            self.var_section_count = tk.StringVar()
            ttk.Label(header, textvariable=self.var_section_count,
                      style="Muted.TLabel").pack(side="left", padx=(10, 0))
        self.tree = self._make_table(self.body)
        self.refresh_table(keep_selection=False)

    def _render_cards(self):
        strip = ttk.Frame(self.body)
        strip.pack(fill="x", pady=(12, 0))
        rooms = list(self.cfg["rooms"])
        extra = sorted(set(i.get("stanza", "") for i in self.store.items) - set(rooms))
        for name in extra:
            if name:
                rooms.append(name)
        for column, name in enumerate(rooms):
            subset = [i for i in self.store.items if i.get("stanza") == name]
            counts = {}
            for item in subset:
                key = item.get("tipo") or "altro"
                counts[key] = counts.get(key, 0) + 1
            breakdown = "  ·  ".join(
                "%d %s" % (n, k.lower()) for k, n in sorted(counts.items())) or "nessun dispositivo"
            on_loan = sum(1 for i in subset if is_on_loan(i))
            note = ""
            if name in self.cfg.get("loan_rooms", []):
                note = ("%d in prestito" % on_loan) if on_loan else "nessun prestito in corso"
            card = RoomCard(strip, self.fonts, name, len(subset), breakdown,
                            theme.ROOM_COLORS[column % len(theme.ROOM_COLORS)],
                            lambda r=name: self.show_room(r), note=note,
                            note_color=theme.LOAN_FG if on_loan else theme.MUTED)
            card.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 14, 0))
            strip.columnconfigure(column, weight=1, uniform="cards")

        tipo = self.iphone_type()
        if tipo:
            telefoni = [i for i in self.store.items if is_iphone(i.get("tipo"))]
            column = len(rooms)
            card = RoomCard(strip, self.fonts, tipo, len(telefoni),
                            "tutti i telefoni, ovunque siano registrati",
                            theme.IPHONE_COLOR, self.show_iphones,
                            note="anche in %s" % self.iphone_room())
            card.grid(row=0, column=column, sticky="nsew", padx=(14, 0))
            strip.columnconfigure(column, weight=1, uniform="cards")

    # ------------------------------------------------------------ dati

    def _initial_load(self):
        try:
            self.store.create_if_missing()
            self.store.load()
        except InventoryError as exc:
            messagebox.showerror("Errore", str(exc), parent=self)
            self.destroy()
            return
        self._sync_filter_values()
        self.show_home()

    def _sync_filter_values(self):
        rooms = list(self.cfg["rooms"])
        for item in self.store.items:
            room = item.get("stanza", "")
            if room and room not in rooms:
                rooms.append(room)
        self.combo_room["values"] = ["Tutte"] + rooms + [NO_ROOM]
        if self.var_room.get() not in self.combo_room["values"]:
            self.var_room.set("Tutte")
        self.combo_type["values"] = ["Tutti"] + list(self.cfg["types"])
        if self.var_type.get() not in self.combo_type["values"]:
            self.var_type.set("Tutti")

    def _on_room_filter(self, _event=None):
        room = self.var_room.get()
        if room in ("Tutte", NO_ROOM):
            if self.view in ("room", "type"):
                self.show_home()
                return
            self.refresh_table()
        else:
            self.show_room(room)

    def _on_type_filter(self, _event=None):
        tipo = self.var_type.get()
        if self.view == "type" and not is_iphone(tipo):
            self.show_home()
            return
        self.refresh_table()

    def filtered_items(self):
        text = self.var_search.get().strip().lower()
        room = self.var_room.get()
        tipo = self.var_type.get()
        result = []
        for item in self.store.items:
            if room == NO_ROOM:
                if item.get("stanza"):
                    continue
            elif room != "Tutte" and item.get("stanza") != room:
                continue
            if tipo != "Tutti" and item.get("tipo") != tipo:
                continue
            if text and not any(
                text in str(item.get(f, "")).lower()
                for f in ("asset_tag", "modello", "seriale", "imei",
                          "restituito_da", "note", "tipo", "stanza",
                          "stato", "prestato_a")
            ):
                continue
            result.append(item)
        result.sort(key=lambda it: str(it.get(self.sort_field, "")).lower(),
                    reverse=self.sort_reverse)
        return result

    def refresh_table(self, keep_selection=True):
        self.visible = self.filtered_items()
        if self.tree is None:
            self._update_status()
            return
        selected = set(self.selected_tags()) if keep_selection else set()
        self.tree.delete(*self.tree.get_children())
        columns = self._columns()
        scelti = set(selected)
        for i, item in enumerate(self.visible):
            values = []
            for field in columns:
                if field == CHECK_COLUMN:
                    values.append(CHECK_ON if item["asset_tag"] in scelti else CHECK_OFF)
                elif field == ACTION_COLUMN:
                    values.append("")          # il pulsante e' disegnato sopra
                else:
                    values.append(item.get(field, ""))
            tag = self.row_tag(item, i % 2)
            self.tree.insert("", "end", iid=item["asset_tag"], values=values,
                             tags=(tag,) if tag else ())
        restore = [t for t in selected if self.tree.exists(t)]
        if restore:
            self.tree.selection_set(restore[:1])
            self.tree.see(restore[0])
        self._update_status()
        self._sync_row_buttons()

    def _update_status(self):
        total = len(self.store.items)
        parts = []
        for room in self.cfg["rooms"]:
            parts.append("%s: %d" % (room, sum(
                1 for i in self.store.items if i.get("stanza") == room)))
        others = sum(1 for i in self.store.items if i.get("stanza") not in self.cfg["rooms"])
        if others:
            parts.append("altre/nessuna: %d" % others)
        self.var_head_count.set("%d dispositivi     %s" % (total, "     ".join(parts)))
        if getattr(self, "var_section_count", None) is not None:
            label = "%d dispositivi" % len(self.visible)
            if len(self.visible) != total:
                label += " di %d" % total
            self.var_section_count.set(label)
        shown = "" if len(self.visible) == total else "  |  visualizzati: %d" % len(self.visible)
        self.var_status.set(
            "%d dispositivi  (%s)%s     File: %s"
            % (total, ", ".join(parts), shown, self.store.path)
        )

    def sort_by(self, field):
        if self.sort_field == field:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_field, self.sort_reverse = field, False
        if self.tree is not None:
            for name in ALL_FIELDS:
                arrow = ""
                if name == self.sort_field:
                    arrow = "  ▾" if self.sort_reverse else "  ▴"
                self.tree.heading(name, text=HEADERS[name] + arrow)
        self.refresh_table()

    def reset_filters(self):
        self.var_search.set("")
        if self.view == "room":
            self.var_type.set("Tutti")
        elif self.view != "type":          # nella vista iPhone il tipo e' la vista
            self.var_type.set("Tutti")
            self.var_room.set("Tutte")
        self.refresh_table()

    def selected_tags(self):
        return list(self.tree.selection()) if self.tree is not None else []

    def selected_items(self):
        tags = set(self.selected_tags())
        return [i for i in self.store.items if i["asset_tag"] in tags]

    def _reload(self, message=None):
        try:
            self.store.load()
        except InventoryError as exc:
            messagebox.showerror("Errore", str(exc), parent=self)
            return
        self.cfg = config.load_shared_config(self.store.path)
        self.store.iphone_room = self.cfg.get("iphone_room")
        self.store.stati = list(self.cfg.get("states") or [])
        self._sync_filter_values()
        self._render()
        if message:
            self.var_status.set(message + "     " + self.var_status.get())

    def _auto_refresh(self):
        try:
            if self.store.changed_on_disk():
                self._reload("Inventario aggiornato da un altro utente.")
        except Exception:
            pass
        self.after(REFRESH_MS, self._auto_refresh)

    def _run(self, action, success=None):
        """Esegue un'operazione sull'archivio gestendo gli errori."""
        try:
            result = action()
        except InventoryError as exc:
            messagebox.showerror("Operazione non riuscita", str(exc), parent=self)
            self._reload()
            return None
        self._sync_filter_values()
        if self.view == "home":
            self._render()          # i conteggi delle schede cambiano
        else:
            self.refresh_table()
        if success:
            self.var_status.set(success + "     " + self.var_status.get())
        return result

    # ------------------------------------------------------------ azioni

    def on_refresh(self):
        self._reload("Elenco ricaricato.")

    def on_new(self):
        rooms = self.cfg["rooms"]
        preset = None
        if self.view == "room":
            preset = new_item(stanza=self.var_room.get())
        elif self.view == "type":
            preset = new_item(tipo=self.var_type.get(), stanza=self.iphone_room())
        item = ItemDialog(self, rooms, self.cfg["types"], preset,
                          iphone_room=self.iphone_room(),
                          stati=self.cfg.get("states")).show()
        if item:
            self._run(lambda: self.store.add(item), "Aggiunto %s." % item["asset_tag"])

    def on_edit(self):
        items = self.selected_items()
        if len(items) != 1:
            messagebox.showinfo("Modifica",
                                "Spunta il dispositivo da modificare, oppure fai doppio clic sulla riga.",
                                parent=self)
            return
        old = items[0]
        edited = ItemDialog(self, self.cfg["rooms"], self.cfg["types"], old,
                            iphone_room=self.iphone_room(),
                            stati=self.cfg.get("states")).show()
        if edited:
            self._run(lambda: self.store.update(old["asset_tag"], edited),
                      "Salvato %s." % edited["asset_tag"])

    def on_delete(self):
        tags = self.selected_tags()
        if not tags:
            return
        item = self._item_by_tag(tags[0])
        if item is not None:
            libero, sblocco = puo_essere_eliminato(item)
            if not libero:
                messagebox.showwarning(
                    "Eliminazione non consentita",
                    "%s\n\n%s\n\n"
                    "Il dispositivo e' stato rispedito al servizio telefonia il %s e\n"
                    "va conservato in inventario per consultazione.\n\n"
                    "Potrai eliminarlo a partire dal %s."
                    % (tags[0], item.get("modello", ""), item["spedito_il"],
                       sblocco.strftime("%d/%m/%Y")),
                    parent=self)
                return
        question = "Eliminare %s dall'inventario?" % tags[0]
        if item and item.get("modello"):
            question = "Eliminare %s - %s dall'inventario?" % (tags[0], item["modello"])
        if not messagebox.askyesno("Conferma eliminazione", question, parent=self):
            return
        self._run(lambda: self.store.delete(tags), "Eliminato %s." % tags[0])

    def on_move(self):
        tags = self.selected_tags()
        if not tags:
            messagebox.showinfo("Sposta", "Spunta il dispositivo da spostare.", parent=self)
            return
        telefoni = [t for t in tags if is_iphone((self._item_by_tag(t) or {}).get("tipo"))]
        if telefoni and len(telefoni) == len(tags):
            messagebox.showinfo(
                "Sposta",
                "Gli iPhone restano sempre in %s e non possono essere spostati."
                % self.iphone_room(), parent=self)
            return
        room = self._ask_room("Sposta %s in:" % tags[0])
        if not room:
            return
        esito = self._run(lambda: self.store.move_to_room(tags, room))
        if esito:
            spostati, bloccati = esito
            messaggio = "Spostati %d dispositivi in %s." % (spostati, room) if spostati \
                else "Nessuno spostamento."
            if bloccati:
                messaggio += "  %d iPhone lasciati in %s." % (bloccati, self.iphone_room())
            self.var_status.set(messaggio + "     " + self.var_status.get())

    def on_lend(self, tag=None):
        """Registra il prestito del dispositivo a una persona."""
        tag = tag or self._single_selection("Presta")
        if not tag:
            return
        item = self._item_by_tag(tag)
        if item is None:
            return
        if not self.can_lend(item):
            messagebox.showinfo(
                "Prestito",
                "La gestione dei prestiti e' attiva solo per: %s."
                % ", ".join(self.cfg.get("loan_rooms") or ["nessuna stanza"]), parent=self)
            return
        person = self._ask_person(item)
        if not person:
            return
        when = self._run(lambda: self.store.lend(tag, person))
        if when:
            self.var_status.set("%s prestato a %s il %s.     %s"
                                % (tag, person, when, self.var_status.get()))

    def on_ship(self, tag=None):
        """Registra la spedizione dell'iPhone al servizio telefonia."""
        tag = tag or self._single_selection("Spedizione")
        if not tag:
            return
        item = self._item_by_tag(tag)
        if item is None:
            return
        if not is_iphone(item.get("tipo")):
            messagebox.showinfo("Spedizione",
                                "La spedizione al servizio telefonia riguarda solo gli iPhone.",
                                parent=self)
            return
        if is_shipped(item):
            messagebox.showinfo("Spedizione",
                                "%s risulta gia' spedito il %s." % (tag, item["spedito_il"]),
                                parent=self)
            return
        if not messagebox.askyesno(
            "Conferma spedizione",
            "%s - %s\n\nRegistrare la spedizione al servizio telefonia?\n\n"
            "Data e ora vengono registrate adesso. Il dispositivo resta in\n"
            "inventario per consultazione per %d mesi, poi potra' essere eliminato."
            % (tag, item.get("modello", ""), MESI_CONSERVAZIONE), parent=self
        ):
            return
        testo = self._run(lambda: self.store.ship(tag))
        if testo:
            messagebox.showinfo("Spedizione registrata", testo, parent=self)

    def on_give_back(self, tag=None):
        """Chiude il prestito: il dispositivo torna disponibile."""
        tag = tag or self._single_selection("Rientro")
        if not tag:
            return
        item = self._item_by_tag(tag)
        if item is None or not is_on_loan(item):
            messagebox.showinfo("Rientro", "Il dispositivo non risulta in prestito.", parent=self)
            return
        if not messagebox.askyesno(
            "Registra rientro",
            "%s - %s\n\nIn prestito a %s dal %s.\nRegistrare il rientro?"
            % (tag, item.get("modello", ""), item["prestato_a"], item["prestato_il"]),
            parent=self
        ):
            return
        person = self._run(lambda: self.store.give_back(tag))
        if person:
            self.var_status.set("%s rientrato da %s.     %s"
                                % (tag, person, self.var_status.get()))

    def _single_selection(self, action):
        tags = self.selected_tags()
        if len(tags) != 1:
            messagebox.showinfo(action, "Spunta il dispositivo su cui vuoi agire.",
                                parent=self)
            return None
        return tags[0]

    def _ask_person(self, item):
        dialog = _Modal(self, "Presta dispositivo")
        body = ttk.Frame(dialog, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="%s - %s" % (item["asset_tag"], item.get("modello", "")),
                  style="Section.TLabel").pack(anchor="w")
        ttk.Label(body, text="Data e ora del prestito vengono registrate in automatico.",
                  style="Muted.TLabel").pack(anchor="w", pady=(2, 12))
        ttk.Label(body, text="Nome della persona").pack(anchor="w")
        var = tk.StringVar()
        entry = ttk.Entry(body, textvariable=var, width=34)
        entry.pack(fill="x", pady=(4, 0))
        buttons = ttk.Frame(body)
        buttons.pack(anchor="e", pady=(16, 0))

        def ok():
            if not var.get().strip():
                messagebox.showwarning("Campo mancante",
                                       "Indica il nome della persona.", parent=dialog)
                return
            dialog.result = var.get().strip()
            dialog.destroy()

        ttk.Button(buttons, text="Annulla", command=dialog._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text="Registra prestito", style="Primary.TButton",
                   command=ok).pack(side="right")
        dialog.bind("<Return>", lambda e: ok())
        entry.focus_set()
        return dialog.show()

    def _ask_room(self, prompt):
        dialog = _Modal(self, "Scegli stanza")
        body = ttk.Frame(dialog, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=prompt).pack(anchor="w", pady=(0, 8))
        var = tk.StringVar(value=self.cfg["rooms"][0])
        ttk.Combobox(body, textvariable=var, values=self.cfg["rooms"],
                     state="readonly", width=28).pack(fill="x")
        buttons = ttk.Frame(body)
        buttons.pack(anchor="e", pady=(16, 0))

        def ok():
            dialog.result = var.get()
            dialog.destroy()

        ttk.Button(buttons, text="Annulla", command=dialog._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text="Sposta", style="Primary.TButton",
                   command=ok).pack(side="right")
        return dialog.show()

    def on_import(self):
        path = filedialog.askopenfilename(
            parent=self, title="Seleziona il file da importare",
            filetypes=[("File Excel", "*.xlsx *.xlsm"), ("Tutti i file", "*.*")])
        if not path:
            return
        try:
            items, skipped = rows_from_workbook(path)
        except InventoryError as exc:
            messagebox.showerror("Importazione non riuscita", str(exc), parent=self)
            return
        if not items:
            messagebox.showwarning("Importazione", "Nessuna riga valida trovata nel file.",
                                   parent=self)
            return
        mode = ImportDialog(self, path, len(items), skipped).show()
        if not mode:
            return
        if mode == "replace" and not messagebox.askyesno(
            "Conferma sostituzione",
            "L'inventario attuale (%d dispositivi) verra' sostituito dai %d dal file.\n"
            "Procedere?" % (len(self.store.items), len(items)), parent=self
        ):
            return
        result = self._run(lambda: self.store.import_items(items, mode))
        if result:
            added, updated = result
            messagebox.showinfo("Importazione completata",
                                "Aggiunti: %d\nAggiornati: %d" % (added, updated), parent=self)

    def on_export(self):
        items = self.filtered_items()
        if not items and not messagebox.askyesno(
            "Esporta", "Nessun dispositivo nella vista corrente. Esportare comunque?",
            parent=self
        ):
            return
        default = "Inventario_%s.xlsx" % datetime.now().strftime("%Y%m%d")
        path = filedialog.asksaveasfilename(
            parent=self, title="Esporta inventario", defaultextension=".xlsx",
            initialfile=default, filetypes=[("File Excel", "*.xlsx")])
        if not path:
            return
        per_room = self.view == "home" and messagebox.askyesno(
            "Esporta", "Creare un foglio separato per ogni stanza?\n\n"
            "No = un unico elenco.", parent=self)
        try:
            excel_io.export(items, path, group_by_room=per_room, rooms=self.cfg["rooms"])
        except InventoryError as exc:
            messagebox.showerror("Esportazione non riuscita", str(exc), parent=self)
            return
        if messagebox.askyesno("Esportazione completata",
                               "File creato:\n%s\n\nAprirlo ora?" % path, parent=self):
            excel_io.open_file(path)

    def on_print(self):
        items = self.filtered_items()
        if not items:
            messagebox.showinfo("Stampa", "Non c'e' nulla da stampare nella vista corrente.",
                                parent=self)
            return
        per_room = self.view == "home" and messagebox.askyesno(
            "Stampa", "Stampare una pagina separata per ogni stanza?\n\n"
            "No = un unico elenco.", parent=self)
        try:
            path = excel_io.build_print_file(items, group_by_room=per_room,
                                             rooms=self.cfg["rooms"])
            printed = excel_io.send_to_printer(path)
        except InventoryError as exc:
            messagebox.showerror("Stampa non riuscita", str(exc), parent=self)
            return
        if printed:
            self.var_status.set("Stampa inviata alla stampante predefinita.     "
                                + self.var_status.get())
        else:
            messagebox.showinfo(
                "Stampa", "Il documento e' stato aperto in Excel.\n"
                "Usa File > Stampa per inviarlo alla stampante.", parent=self)

    def on_settings(self):
        result = RoomsDialog(self, self.cfg["rooms"], self.cfg["types"],
                             self.cfg.get("loan_rooms", []),
                             self.cfg.get("iphone_room", "")).show()
        if not result:
            return
        try:
            config.save_shared_config(self.store.path, result)
        except OSError as exc:
            messagebox.showerror("Errore", "Impossibile salvare le impostazioni:\n%s" % exc,
                                 parent=self)
            return
        self.cfg = config.load_shared_config(self.store.path)
        self.store.iphone_room = self.cfg.get("iphone_room")
        self.store.stati = list(self.cfg.get("states") or [])
        self._sync_filter_values()
        self.show_home()


# ------------------------------------------------------------ avvio


def choose_data_file(root):
    """Prima configurazione: dove si trova, o dove creare, il file inventario.

    Non viene mostrata quando accanto al programma c'e' gia' Inventario.xlsx,
    che e' il caso normale con eseguibile e dati nella stessa cartella di rete.
    """
    cartella = config.app_dir()
    answer = messagebox.askyesnocancel(
        'Site Services : Inventario Iphone, Laptop e Tablet',
        "Non ho ancora trovato il file dell'inventario.\n\n"
        "Di norma si chiama Inventario.xlsx e sta nella stessa cartella del\n"
        "programma:\n%s\n\n"
        "Si'  = apri un file inventario gia' esistente\n"
        "No   = crea qui un nuovo inventario vuoto\n"
        "Annulla = esci" % cartella,
        parent=root)
    if answer is None:
        return None
    if answer:
        return filedialog.askopenfilename(
            parent=root, title="Seleziona il file inventario",
            initialdir=cartella,
            filetypes=[("File Excel", "*.xlsx")]) or None
    return filedialog.asksaveasfilename(
        parent=root, title="Crea il file inventario",
        defaultextension=".xlsx", initialdir=cartella,
        initialfile=config.DATA_FILE_NAME,
        filetypes=[("File Excel", "*.xlsx")]) or None


def main():
    path = config.load_data_path()
    if not path or not os.path.isdir(os.path.dirname(os.path.abspath(path))):
        root = tk.Tk()
        root.withdraw()
        path = choose_data_file(root)
        root.destroy()
        if not path:
            return 1
        # Se il file scelto e' quello accanto al programma non serve ricordarlo:
        # al prossimo avvio viene ritrovato da solo.
        if os.path.abspath(path) != os.path.abspath(config.default_data_path()):
            config.save_data_path(path)
    App(path).mainloop()
    return 0
