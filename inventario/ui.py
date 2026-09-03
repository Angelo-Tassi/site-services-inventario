"""Interfaccia grafica dell'inventario (Tkinter).

Due schermate:
  * Home    - schede delle stanze con il numero di dispositivi + inventario completo
  * Stanza  - inventario della singola stanza, aperto cliccando una scheda
"""

import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, font as tkfont, messagebox, ttk

from . import __version__, config, excel_io, theme
from . import lingua as lang
from .lingua import T, intestazione, stato as traduci_stato
from .store import (ALL_FIELDS, COPIE_DA_TENERE, DA_RISPEDIRE,
                    ELIMINATI_GIORNI, ELIMINATI_MASSIMO, HEADERS,
                    InventoryError, MESI_CONSERVAZIONE,
                    InventoryStore, NON_DISPONIBILE, SPEDITO, clean,
                    is_iphone, is_on_loan, is_shipped, new_item, norm_tag,
                    puo_essere_eliminato, righe_separatore,
                    rinomina_tocca_gli_iphone, rinomine_stanze,
                    rinomine_tipi, rows_from_workbook,
                    sembra_un_foglio_da_importare, testo_spedizione,
                    valore_visibile)

MODIFICHE_PER_PROMEMORIA = 20   # record toccati prima di ricordare la copia
ALTEZZA_MINIMA_TABELLA = 160   # pixel: circa cinque righe
SPAZIO_CELLA = 22              # margini della cella, perche' il testo non tocchi i bordi
SPAZIO_INTESTAZIONE = 34       # margini piu' la barretta colorata e la freccia
LARGHEZZA_MINIMA = 46

# L'ordine in cui si leggono le colonne: prima quello che si cerca a colpo
# d'occhio - qual e' l'oggetto, dov'e', come sta, che cosa c'e' da sapere - e
# per ultimi i campi lunghi, che si leggono solo quando servono davvero. Il
# contenitore Iphone non segue questo ordine: li' l'IMEI e' l'identificativo e
# viene per primo.
ORDINE_COLONNE = ["asset_tag", "tipo", "stanza", "note", "stato",
                  "modello", "seriale",
                  "imei", "restituito_da", "prestato_a", "prestato_il",
                  "spedito_il", "modificato_il", "modificato_da"]

NO_ROOM_IT = "(senza stanza)"


def NO_ROOM():
    return T(NO_ROOM_IT)


def TUTTE():
    return T("Tutte")


def TUTTI():
    return T("Tutti")


CAMPI_DATA = ("modificato_il", "prestato_il", "spedito_il")


def chiave_ordinamento(item, campo):
    """Le colonne con una data si ordinano per data, non per testo."""
    valore = item.get(campo, "")
    if campo in CAMPI_DATA:
        for formato in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
            try:
                return datetime.strptime(str(valore), formato)
            except ValueError:
                continue
        return datetime.min          # senza data finisce in fondo
    return str(valore).lower()


def seleziona_per_stanza(items, esito, stanza):
    """Che cosa importare quando l'importazione riguarda una sola stanza.

    Se il foglio dichiara le stanze con le righe-separatore, comandano quelle:
    si prende la sezione della stanza scelta e si scarta il resto. Se invece il
    foglio non dichiara nessuna stanza, e' la scelta dell'utente a valere e
    tutte le righe finiscono li'.

    Ritorna (righe, scartate, regola), dove regola vale "separatori" o "tutte".
    Ritorna (None, 0, "mancante") se il foglio dichiara delle stanze ma non
    quella richiesta: in quel caso non si deve importare niente.
    """
    trovate = list(esito.get("stanze_trovate") or [])
    if not trovate:
        return list(items), 0, "tutte"
    if stanza not in trovate:
        return None, 0, "mancante"
    miei = [i for i in items if i.get("stanza") == stanza]
    return miei, len(items) - len(miei), "separatori"


def nome_file(testo):
    """Un nome di stanza utilizzabile come nome di file."""
    pulito = "".join(c if c.isalnum() or c in " -_" else "-" for c in str(testo))
    return "_".join(pulito.split()) or "stanza"


def stato_canonico(scelto, stati):
    """Dalla voce mostrata nella tendina al valore italiano da salvare."""
    for valore in stati:
        if traduci_stato(valore) == scelto:
            return valore
    return scelto


def item_stato_iphone(spedito):
    return SPEDITO if spedito else DA_RISPEDIRE

CHECK_COLUMN = "_sel"
ACTION_COLUMN = "_azione"
# Un iPhone non ha numero di serie e non si presta: nel suo contenitore quelle
# colonne sarebbero sempre vuote.
COLONNE_NON_IPHONE = ("asset_tag", "seriale", "prestato_a", "prestato_il")
# Bit dei modificatori dentro `event.state`. Mod1 si legge solo dove Command
# esiste davvero: su Windows quello stesso bit e' il Bloc Num, e leggerlo li'
# vorrebbe dire scambiare ogni clic per un Ctrl+clic a chi tiene acceso il
# tastierino numerico.
STATO_SHIFT = 0x0001
STATO_CONTROL = 0x0004
STATO_COMANDO = 0x0008


def con_modificatore(stato, comando=False):
    """True se il clic porta un tasto che significa "piu' righe insieme"."""
    maschera = STATO_SHIFT | STATO_CONTROL | (STATO_COMANDO if comando else 0)
    return bool(stato & maschera)


CHECK_ON = "\u25c9"      # cerchio pieno: riga selezionata
CHECK_OFF = "\u25cb"     # cerchio vuoto
COLUMN_WIDTHS = {CHECK_COLUMN: 46, ACTION_COLUMN: 175, "asset_tag": 120, "tipo": 75, "modello": 185,
                 "seriale": 120, "imei": 130, "restituito_da": 135, "stanza": 160,
                 "stato": 185, "prestato_a": 140, "prestato_il": 120, "spedito_il": 120, "note": 180,
                 "modificato_il": 120, "modificato_da": 145}
REFRESH_MS = 15000


# --------------------------------------------------------------- dialoghi


def installa_copia_incolla(root):
    """Copia, taglia e incolla in tutti i campi dell'applicazione.

    Tk ha le sue scorciatoie predefinite, ma non sono le stesse su tutti i
    sistemi ne' con tutte le disposizioni di tastiera, e in qualche caso non
    funzionano affatto: chi incolla un codice preso da Excel non deve chiedersi
    se in quel campo si puo' fare. Qui si collegano a mano - Control e Command -
    e si aggiunge il menu del tasto destro.

    Si collega alle classi di widget, non ai singoli campi, cosi' vale anche per
    le finestre costruite piu' tardi.

    Vale anche per i campi in sola lettura: da li' non si scrive, ma copiare
    quello che c'e' scritto e' sempre lecito.
    """

    def scrivibile(widget):
        try:
            return str(widget.cget("state")) not in ("readonly", "disabled")
        except tk.TclError:
            return True

    def testo_selezionato(widget):
        try:
            if widget.selection_present():
                return widget.selection_get()
        except (tk.TclError, AttributeError):
            pass
        try:
            return widget.get("sel.first", "sel.last")          # Text
        except (tk.TclError, TypeError):
            return None

    def tutto_il_testo(widget):
        try:
            return widget.get()
        except (TypeError, tk.TclError):
            try:
                return widget.get("1.0", "end").rstrip("\n")    # Text
            except tk.TclError:
                return ""

    def copia(evento):
        widget = evento.widget
        testo = testo_selezionato(widget)
        if testo is None:
            testo = tutto_il_testo(widget)      # niente selezionato: tutto
        if testo:
            widget.clipboard_clear()
            widget.clipboard_append(testo)
        return "break"

    def taglia(evento):
        copia(evento)
        widget = evento.widget
        if not scrivibile(widget):
            return "break"
        try:
            if widget.selection_present():
                widget.delete("sel.first", "sel.last")
            else:
                widget.delete(0, "end")
        except (tk.TclError, AttributeError):
            try:
                widget.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
        return "break"

    def incolla(evento):
        widget = evento.widget
        if not scrivibile(widget):
            return "break"
        try:
            arrivato = widget.clipboard_get()
        except tk.TclError:
            return "break"
        try:
            if widget.selection_present():
                widget.delete("sel.first", "sel.last")
        except (tk.TclError, AttributeError):
            try:
                widget.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
        try:
            widget.insert("insert", arrivato)
        except tk.TclError:
            pass
        return "break"

    def seleziona_tutto(evento):
        widget = evento.widget
        try:
            widget.select_range(0, "end")
            widget.icursor("end")
        except (tk.TclError, AttributeError):
            try:
                widget.tag_add("sel", "1.0", "end")
            except tk.TclError:
                pass
        return "break"

    menu = tk.Menu(root, tearoff=0)

    def apri_menu(evento):
        widget = evento.widget
        finto = type("Evento", (), {"widget": widget})()
        menu.delete(0, "end")
        menu.add_command(label=T("Copia"), command=lambda: copia(finto))
        menu.add_command(label=T("Taglia"), command=lambda: taglia(finto),
                         state="normal" if scrivibile(widget) else "disabled")
        menu.add_command(label=T("Incolla"), command=lambda: incolla(finto),
                         state="normal" if scrivibile(widget) else "disabled")
        menu.add_separator()
        menu.add_command(label=T("Seleziona tutto"),
                         command=lambda: seleziona_tutto(finto))
        try:
            widget.focus_set()
            menu.tk_popup(evento.x_root, evento.y_root)
        finally:
            menu.grab_release()
        return "break"

    scorciatoie = (("c", copia), ("x", taglia), ("v", incolla), ("a", seleziona_tutto))
    root._appunti = {"copia": copia, "taglia": taglia, "incolla": incolla,
                     "seleziona_tutto": seleziona_tutto, "menu": apri_menu}
    # Si sostituiscono i collegamenti predefiniti invece di aggiungersi: due
    # gestori sulla stessa combinazione incollerebbero due volte.
    virtuali = {"c": "<<Copy>>", "x": "<<Cut>>", "v": "<<Paste>>",
                "a": "<<SelectAll>>"}
    for classe in ("Entry", "TEntry", "TCombobox", "Text"):
        for tasto, azione in scorciatoie:
            for modificatore in ("Control", "Command"):
                for lettera in (tasto, tasto.upper()):
                    try:
                        root.bind_class(classe, "<%s-%s>" % (modificatore, lettera),
                                        azione)
                    except tk.TclError:
                        pass      # combinazione non disponibile su questo sistema
            try:
                root.bind_class(classe, virtuali[tasto], azione)
            except tk.TclError:
                pass
        for tasto_destro in ("<Button-3>", "<Button-2>", "<Control-Button-1>"):
            try:
                root.bind_class(classe, tasto_destro, apri_menu, add="+")
            except tk.TclError:
                pass


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
        _Modal.__init__(self, parent, T("Modifica dispositivo") if item
                        else T("Nuovo dispositivo"))
        item = item or new_item(stanza=rooms[0] if rooms else "")
        # Gli iPhone non si prestano: nessun dato di prestito da conservare.
        self._loan = ("", "") if is_iphone(item.get("tipo")) \
            else (item.get("prestato_a", ""), item.get("prestato_il", ""))
        self.rooms = rooms
        self.iphone_room = iphone_room or (rooms[0] if rooms else "")
        self.stati = list(stati or [])
        self._item_spedito = bool(item.get("spedito_il"))

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
        ttk.Label(body, text=T("Tipo *")).grid(row=0, column=0, sticky="w", pady=5)
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
                  text=T("I campi contrassegnati con * sono obbligatori.")).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))

        buttons = ttk.Frame(body)
        buttons.grid(row=4, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(buttons, text=T("Annulla"), command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text=T("Salva"), style="Primary.TButton",
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

        # Obbligatorio e' solo l'identificativo: e' l'unica cosa senza la quale
        # il dispositivo non esiste in inventario. Il resto spesso non si ha
        # sottomano nel momento in cui si registra un arrivo, e pretenderlo
        # significa solo far rimandare l'inserimento - cioe' perdere la riga.
        if self.is_iphone():
            righe = [(T("IMEI"), self.var_imei, True),
                     (T("Modello"), self.var_modello, False),
                     (T("Restituito da"), self.var_restituito, False)]
        else:
            righe = [(T("Asset Tag"), self.var_tag, True),
                     (T("Modello"), self.var_modello, False),
                     (T("Numero di serie"), self.var_seriale, False)]

        self.required = []
        for riga, (etichetta, var, obbligatorio) in enumerate(righe):
            ttk.Label(self.fields, text=etichetta + (" *" if obbligatorio else "")).grid(
                row=riga, column=0, sticky="w", pady=5)
            entry = ttk.Entry(self.fields, textvariable=var, width=34)
            entry.grid(row=riga, column=1, sticky="we", pady=5)
            if obbligatorio:
                self.required.append((etichetta, var, entry))

        riga = len(righe)
        # La stanza non blocca il salvataggio: e' una tendina, e se non e' stata
        # scelta si parte dalla prima invece di lasciare un dispositivo senza
        # stanza, che non comparirebbe in nessuna scheda.
        if not self.var_stanza.get() and self.rooms:
            self.var_stanza.set(self.rooms[0])
        ttk.Label(self.fields, text=T("Stanza")).grid(row=riga, column=0, sticky="w", pady=5)
        if self.is_iphone():
            # Gli iPhone appartengono sempre alla loro stanza: campo mostrato ma bloccato.
            self.var_stanza.set(self.iphone_room)
            combo = ttk.Combobox(self.fields, textvariable=self.var_stanza,
                                 values=[self.iphone_room], state="disabled", width=32)
        else:
            combo = ttk.Combobox(self.fields, textvariable=self.var_stanza, values=self.rooms,
                                 state="readonly", width=32)
        combo.grid(row=riga, column=1, sticky="we", pady=5)
        riga += 1
        if self.is_iphone():
            ttk.Label(self.fields, style="Muted.TLabel",
                      text=T("Gli iPhone restano sempre in %s.") % self.iphone_room).grid(
                row=riga, column=1, sticky="w")
            riga += 1

        ttk.Label(self.fields, text=T("Stato")).grid(row=riga, column=0, sticky="w", pady=5)
        if self.is_iphone():
            self.var_stato.set(traduci_stato(item_stato_iphone(self._item_spedito)))
            valori, stato_widget = [self.var_stato.get()], "disabled"
        elif self._loan[0]:
            self.var_stato.set(traduci_stato(NON_DISPONIBILE))
            valori, stato_widget = [self.var_stato.get()], "disabled"
        else:
            if self.var_stato.get() not in self.stati and self.stati:
                self.var_stato.set(self.stati[0])
            valori, stato_widget = [traduci_stato(v) for v in self.stati], "readonly"
            self.var_stato.set(traduci_stato(self.var_stato.get()))
        ttk.Combobox(self.fields, textvariable=self.var_stato, values=valori,
                     state=stato_widget, width=32).grid(row=riga, column=1,
                                                        sticky="we", pady=5)
        riga += 1
        if stato_widget == "disabled":
            motivo = (T("Gli iPhone non si prestano: lo stato lo decide la spedizione.")
                      if self.is_iphone()
                      else T("In prestito: lo stato torna modificabile dopo il rientro."))
            ttk.Label(self.fields, style="Muted.TLabel", text=motivo).grid(
                row=riga, column=1, sticky="w")
            riga += 1

        ttk.Label(self.fields, text=T("Note")).grid(row=riga, column=0, sticky="nw", pady=5)
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
                T("Dati mancanti"),
                T("Il dispositivo non e\' stato inserito.\n"
                "Mancano questi dati obbligatori:\n\n%s\n\n"
                "Compilali e premi di nuovo Salva.")
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
            stato=stato_canonico(self.var_stato.get(), self.stati),
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

    def __init__(self, parent, rooms, types, loan_rooms, iphone_room="", lingua=None):
        _Modal.__init__(self, parent, T("Impostazioni inventario"))
        self.stanze_di_partenza = list(rooms)
        self.tipi_di_partenza = list(types)
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=T("Stanze (una per riga)")).grid(row=0, column=0, sticky="w")
        ttk.Label(body, text=T("Tipi di dispositivo")).grid(row=0, column=1, sticky="w", padx=(14, 0))
        ttk.Label(body, text=T("Stanze con prestito")).grid(row=0, column=2, sticky="w", padx=(14, 0))
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
        riga_lingua = ttk.Frame(body)
        riga_lingua.grid(row=2, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Label(riga_lingua, text=T("Lingua")).pack(side="left")
        self.var_lingua = tk.StringVar(value=lang.nome_lingua(lingua or lang.corrente()))
        ttk.Combobox(riga_lingua, textvariable=self.var_lingua,
                     values=[nome for nome, _ in lang.LINGUE], state="readonly",
                     width=18).pack(side="left", padx=(10, 0))
        ttk.Label(riga_lingua, style="Muted.TLabel",
                  text=T("Vale solo per questo computer.")).pack(side="left", padx=(10, 0))

        riga_iphone = ttk.Frame(body)
        riga_iphone.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Label(riga_iphone, text=T("Stanza degli iPhone")).pack(side="left")
        self.var_iphone_room = tk.StringVar(value=iphone_room or (rooms[0] if rooms else ""))
        ttk.Combobox(riga_iphone, textvariable=self.var_iphone_room, values=rooms,
                     width=30).pack(side="left", padx=(10, 0))

        ttk.Label(
            body, style="Muted.TLabel",
            text=T("Le impostazioni sono salvate accanto al file dati e valgono per tutti gli utenti.\n"
                 "Nelle stanze con prestito ogni riga dell'elenco ha il pulsante Presta / Registra rientro.\n"
                 "Gli iPhone vengono registrati sempre nella stanza indicata qui sopra e non si spostano."),
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))
        # I comandi che riscrivono l'inventario di tutti stanno qui, non nella
        # barra: si usano di rado, e nella barra rubavano lo spazio ai comandi
        # di tutti i giorni fino a farli sparire dallo schermo.
        comandi = ttk.LabelFrame(body, text=T("Copie e ripristino"), padding=10)
        comandi.grid(row=5, column=0, columnspan=3, sticky="we", pady=(14, 0))
        ttk.Button(comandi, text=T("Salva copia in locale..."), style="Verde.TButton",
                   command=self._copia_locale).pack(side="left", padx=(0, 8))
        ttk.Button(comandi, text=T("Ripristina da una copia..."), style="Rosso.TButton",
                   command=self._ripristina).pack(side="left", padx=(0, 8))
        ttk.Button(comandi, text=T("Ripristina da un file locale..."),
                   style="Rosso.TButton",
                   command=self._ripristina_locale).pack(side="left", padx=(0, 8))
        ttk.Button(comandi, text=T("Reset inventario"), style="Rosso.TButton",
                   command=self._reset).pack(side="left")
        ttk.Label(comandi, style="Muted.TLabel",
                  text=T("Delle copie automatiche se ne tengono %d: quando ne arriva\n"
                         "una nuova, la piu' vecchia viene cancellata.\n"
                         "Stanno sulla rete accanto ai dati: se sparisce quella\n"
                         "cartella, si riparte da un file locale.")
                  % COPIE_DA_TENERE).pack(side="left", padx=(14, 0))

        buttons = ttk.Frame(body)
        buttons.grid(row=6, column=0, columnspan=3, sticky="we", pady=(16, 0))
        ttk.Button(buttons, text=T("Collega inventario condiviso..."),
                   command=self._collega).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text=T("Annulla"), command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text=T("Salva"), style="Primary.TButton",
                   command=self._ok).pack(side="right")

    def _collega(self):
        self.result = {"collega": True}
        self.destroy()

    def _ripristina(self):
        self.result = {"ripristina": True}
        self.destroy()

    def _copia_locale(self):
        self.result = {"copia_locale": True}
        self.destroy()

    def _ripristina_locale(self):
        self.result = {"ripristina_locale": True}
        self.destroy()

    def _reset(self):
        self.result = {"reset": True}
        self.destroy()

    def _ok(self):
        rooms = [r.strip() for r in self.text_rooms.get("1.0", "end").splitlines() if r.strip()]
        types = [t.strip() for t in self.text_types.get("1.0", "end").splitlines() if t.strip()]
        loans = [r.strip() for r in self.text_loans.get("1.0", "end").splitlines() if r.strip()]
        if not rooms:
            messagebox.showwarning(T("Dato mancante"), T("Indica almeno una stanza."), parent=self)
            return
        doppie = [r for r in rooms if rooms.count(r) > 1]
        if doppie:
            messagebox.showwarning(
                T("Stanza ripetuta"),
                T("Questa stanza compare due volte nell'elenco:\n%s\n\n"
                  "Due stanze non possono chiamarsi allo stesso modo.")
                % doppie[0], parent=self)
            return
        doppi = [t for t in types if types.count(t) > 1]
        if doppi:
            messagebox.showwarning(
                T("Tipo ripetuto"),
                T("Questo tipo compare due volte nell'elenco:\n%s\n\n"
                  "Due tipi non possono chiamarsi allo stesso modo.")
                % doppi[0], parent=self)
            return

        # Il tipo iPhone e' la parola con cui il programma riconosce i telefoni:
        # si dice di no qui, prima di salvare, invece di lasciare che i telefoni
        # smettano di essere telefoni.
        cambi_tipo = rinomine_tipi(self.tipi_di_partenza, types)
        tocca = rinomina_tocca_gli_iphone(cambi_tipo)
        if tocca:
            messagebox.showwarning(
                T("Il tipo iPhone non si rinomina"),
                T("Stai rinominando \"%s\" in \"%s\".\n\n"
                  "E' la parola con cui il programma riconosce i telefoni: da li'\n"
                  "vengono l'IMEI al posto dell'asset tag, la stanza obbligata, la\n"
                  "spedizione al servizio telefonia e il fatto che non si eliminino.\n\n"
                  "Rimetti \"%s\" e salva.") % (tocca[0], tocca[1], tocca[0]),
                parent=self)
            return

        # Rinominata una stanza, chi la nominava deve seguirla: le stanze con
        # prestito e la stanza degli iPhone si aggiornano da sole, o il
        # controllo qui sotto fermerebbe il salvataggio per un nome che
        # l'utente ha appena cambiato di proposito.
        rinomine = dict(rinomine_stanze(self.stanze_di_partenza, rooms))
        loans = [rinomine.get(r, r) for r in loans]
        stanza_iphone_scelta = rinomine.get(self.var_iphone_room.get().strip(),
                                            self.var_iphone_room.get().strip())

        unknown = [r for r in loans if r not in rooms]
        if unknown:
            messagebox.showwarning(
                T("Stanza sconosciuta"),
                T("Queste stanze con prestito non sono nell'elenco delle stanze:\n%s")
                % ", ".join(unknown), parent=self)
            return
        stanza_iphone = stanza_iphone_scelta
        if stanza_iphone and stanza_iphone not in rooms:
            messagebox.showwarning(
                T("Stanza sconosciuta"),
                T("La stanza degli iPhone (%s) non e' nell'elenco delle stanze.")
                % stanza_iphone, parent=self)
            return
        scelta = dict((nome, codice) for nome, codice in lang.LINGUE)
        self.result = {"rooms": rooms, "types": types or ["Laptop", "Tablet"],
                       "rinomine": list(rinomine.items()),
                       "rinomine_tipi": cambi_tipo,
                       "loan_rooms": loans,
                       "iphone_room": stanza_iphone or rooms[0],
                       "lingua": scelta.get(self.var_lingua.get(), lang.ITALIANO)}
        self.destroy()


class TypeChoiceDialog(_Modal):
    """Primo passo di Aggiungi: che cosa si sta inserendo."""

    def __init__(self, parent, types, predefinito=""):
        _Modal.__init__(self, parent, T("Aggiungi dispositivo"))
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=T("Che cosa vuoi aggiungere?"),
                  style="Section.TLabel").pack(anchor="w")
        ttk.Label(body, style="Muted.TLabel",
                  text=T("Il tipo decide i campi da compilare e cosa si puo'\n"
                       "leggere con il lettore di codici.")).pack(anchor="w", pady=(4, 12))
        self.var_tipo = tk.StringVar(value=predefinito or (types[0] if types else ""))
        combo = ttk.Combobox(body, textvariable=self.var_tipo, values=types,
                             state="readonly", width=30)
        combo.pack(fill="x")
        buttons = ttk.Frame(body)
        buttons.pack(anchor="e", pady=(16, 0))
        ttk.Button(buttons, text=T("Annulla"), command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text=T("Avanti"), style="Primary.TButton",
                   command=self._ok).pack(side="right")
        self.bind("<Return>", lambda e: self._ok())
        combo.focus_set()

    def _ok(self):
        if not self.var_tipo.get().strip():
            return
        self.result = self.var_tipo.get()
        self.destroy()


class AddChoiceDialog(_Modal):
    """Come aggiungere un dispositivo: a mano o leggendo i codici a barre."""

    def __init__(self, parent, iphone=False):
        _Modal.__init__(self, parent, T("Aggiungi iPhone") if iphone
                        else T("Aggiungi dispositivo"))
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=T("Come vuoi aggiungerlo?"),
                  style="Section.TLabel").pack(anchor="w")
        if iphone:
            aiuto = (T("La scansione legge l'IMEI dal codice a barre: un iPhone non\n"
                     "ha asset tag ne' numero di serie."))
            etichetta = T("Scansiona l'IMEI con il lettore di codici")
        else:
            aiuto = (T("La scansione compila asset tag e numero di serie con il\n"
                     "lettore di codici a barre."))
            etichetta = T("Scansiona con il lettore di codici")
        ttk.Label(body, style="Muted.TLabel", text=aiuto).pack(anchor="w", pady=(4, 14))

        def scegli(modo):
            self.result = modo
            self.destroy()

        ttk.Button(body, text=etichetta, style="Primary.TButton",
                   command=lambda: scegli("barcode")).pack(fill="x", pady=(0, 8))
        ttk.Button(body, text=T("Inserimento manuale"),
                   command=lambda: scegli("manuale")).pack(fill="x")
        ttk.Button(body, text=T("Annulla"), command=self._cancel).pack(anchor="e", pady=(14, 0))


class ScanDialog(_Modal):
    """Un passo della procedura con lettore: si legge il codice, oppure si scrive.

    Il lettore di codici si comporta come una tastiera: scrive nel campo e
    conferma da solo. Se non riesce a leggere, l'operatore passa alla scrittura
    a mano con il pulsante dedicato, senza perdere quello che ha gia' inserito.
    """

    def __init__(self, parent, titolo, campo, passo, totale, manuale=False,
                 valore=""):
        _Modal.__init__(self, parent, titolo)
        self.campo = campo
        self.manuale = manuale
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text=T("Passo %d di %d") % (passo, totale),
                  style="Muted.TLabel").pack(anchor="w")
        self.var_titolo = tk.StringVar()
        ttk.Label(body, textvariable=self.var_titolo,
                  style="Section.TLabel").pack(anchor="w", pady=(2, 2))
        self.var_aiuto = tk.StringVar()
        ttk.Label(body, textvariable=self.var_aiuto, style="Muted.TLabel",
                  justify="left").pack(anchor="w", pady=(0, 10))

        self.var_valore = tk.StringVar(value=valore)
        self.entry = ttk.Entry(body, textvariable=self.var_valore, width=36,
                               font=(theme.pick_family(self), 13))
        self.entry.pack(fill="x")

        self.btn_manuale = ttk.Button(
            body, text=T("Non riesco a scansionare - inserisci a mano"),
            command=self.passa_a_manuale)
        self.btn_manuale.pack(anchor="w", pady=(10, 0))

        buttons = ttk.Frame(body)
        buttons.pack(anchor="e", pady=(14, 0))
        ttk.Button(buttons, text=T("Annulla"), command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text=T("Avanti"), style="Primary.TButton",
                   command=self._ok).pack(side="right")

        self.bind("<Return>", lambda e: self._ok())
        self._aggiorna()
        self.entry.focus_set()

    def _aggiorna(self):
        if self.manuale:
            self.var_titolo.set(T("Scrivi %s") % self.campo)
            self.var_aiuto.set(T("Digita il valore e premi Invio."))
            self.btn_manuale.pack_forget()
        else:
            self.var_titolo.set(T("Scansiona %s") % self.campo)
            self.var_aiuto.set(T("Inquadra il codice: il lettore compila il campo e\n"
                               "conferma da solo. Puoi anche digitarlo."))

    def passa_a_manuale(self):
        self.manuale = True
        self._aggiorna()
        self.entry.focus_set()

    def _ok(self):
        valore = self.var_valore.get().strip()
        if not valore:
            messagebox.showwarning(
                T("Campo vuoto"),
                T("%s non puo' restare vuoto.\n\n"
                "Riprova la scansione, oppure usa il pulsante per inserirlo a mano.")
                % self.campo.capitalize(), parent=self)
            self.entry.focus_set()
            return
        self.result = valore
        self.destroy()


PAROLA_RESET = "ELIMINA TUTTO"
PAROLA_ELIMINA = "ACCETTO"


class ResetDialog(_Modal):
    """Conferma dello svuotamento dell'inventario.

    Non basta un si': va scritta una frase, perche' l'operazione riguarda i dati
    condivisi di tutti e non si annulla con un tasto.
    """

    def __init__(self, parent, da_eliminare, protetti):
        _Modal.__init__(self, parent, T("Reset dell'inventario"))
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)

        tk.Label(body, text=T("Stai per svuotare l'inventario condiviso"),
                 bg=theme.LOAN_FG, fg="#FFFFFF", font=self.master.fonts["card_title"],
                 padx=12, pady=9, anchor="w").pack(fill="x")

        righe = [T("Verranno eliminati %d dispositivi, per tutti gli utenti.") % da_eliminare,
                 T("L'operazione non si annulla dal programma.")]
        if protetti:
            righe.append("")
            righe.append(T("Restano dentro %d iPhone: il reset non li elimina mai,\n"
                           "perche' non potrebbero essere ricaricati da un file.")
                         % protetti)
        tk.Label(body, text=T("\n").join(righe), justify="left", anchor="w",
                 bg=theme.LOAN_BG, fg=theme.LOAN_FG, padx=12, pady=10).pack(
            fill="x", pady=(0, 10))

        ttk.Label(body, style="Muted.TLabel", justify="left",
                  text=T("Prima di procedere il programma salva una copia del file dati\n"
                       "nella cartella Backup, dentro quella del programma: se qualcosa\n"
                       "va storto, l'inventario si recupera da li'.")).pack(anchor="w")

        ttk.Label(body, text=T("Per confermare, scrivi   %s") % PAROLA_RESET).pack(
            anchor="w", pady=(14, 4))
        self.var_conferma = tk.StringVar()
        entry = ttk.Entry(body, textvariable=self.var_conferma, width=34)
        entry.pack(fill="x")

        buttons = ttk.Frame(body)
        buttons.pack(anchor="e", pady=(16, 0))
        ttk.Button(buttons, text=T("Annulla"), command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text=T("Svuota l'inventario"),
                   command=self._ok).pack(side="right")
        self.bind("<Return>", lambda e: self._ok())
        entry.focus_set()

    def parola_giusta(self):
        return self.var_conferma.get().strip().upper() == PAROLA_RESET

    def _ok(self):
        if not self.parola_giusta():
            messagebox.showwarning(
                T("Conferma non valida"),
                T("Per svuotare l'inventario devi scrivere esattamente:\n\n%s")
                % PAROLA_RESET, parent=self)
            return
        self.result = True
        self.destroy()


class EsportazioneFattaDialog(_Modal):
    """Cosa fare del file appena esportato: mandarlo, aprirlo, o niente."""

    def __init__(self, parent, descrizione, percorsi):
        _Modal.__init__(self, parent, T("Esportazione completata"))
        self.percorsi = list(percorsi)
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=descrizione, style="Section.TLabel",
                  wraplength=420, justify="left").pack(anchor="w")

        # Percorsi interi in un campo di testo, non in un'etichetta: da qui si
        # copiano - con Ctrl+C o con il tasto destro - e si incollano dove
        # servono, che e' la prima cosa che si vuole fare con un file appena
        # prodotto.
        elenco = "\n".join(self.percorsi)
        cornice = tk.Frame(body, bg=theme.CARD, highlightthickness=1,
                           highlightbackground=theme.BORDER)
        cornice.pack(fill="both", expand=True, pady=(6, 14))
        percorsi = tk.Text(cornice, width=52, height=min(6, max(1, len(self.percorsi))),
                           wrap="none", relief="flat", highlightthickness=0,
                           bg=theme.CARD, fg=theme.TEXT, font=self.master.fonts["base"])
        percorsi.insert("1.0", elenco)
        percorsi.configure(state="disabled")
        barra = ttk.Scrollbar(cornice, orient="horizontal", command=percorsi.xview)
        percorsi.configure(xscrollcommand=barra.set)
        percorsi.pack(side="top", fill="both", expand=True, padx=8, pady=(8, 0))
        barra.pack(side="bottom", fill="x")

        ttk.Button(body, text=T("Invia per e-mail con Outlook"),
                   style="Primary.TButton",
                   command=lambda: self._scegli("email")).pack(fill="x")
        ttk.Label(body, style="Muted.TLabel", justify="left",
                  text=T("Apre un messaggio nuovo con il file gia' allegato:\n"
                         "destinatario e testo li scrivi tu, l'invio resta a te.")).pack(
            anchor="w", padx=(6, 0), pady=(4, 10))
        ttk.Button(body, text=T("Apri il file"),
                   command=lambda: self._scegli("apri")).pack(fill="x")
        ttk.Button(body, text=T("Ho finito"),
                   command=self._cancel).pack(fill="x", pady=(8, 0))

    def _scegli(self, azione):
        self.result = azione
        self.destroy()


class RestoreDialog(_Modal):
    """Scelta della copia di sicurezza da cui ripartire."""

    def __init__(self, parent, copie, quanti_ora):
        _Modal.__init__(self, parent, T("Ripristina da una copia"))
        self.copie = copie
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=T("Da quale copia vuoi ripartire?"),
                  style="Section.TLabel").pack(anchor="w")
        ttk.Label(body, style="Muted.TLabel", justify="left",
                  text=T("In inventario ci sono adesso %d dispositivi. La copia scelta\n"
                         "prendera' il loro posto; lo stato attuale viene salvato prima,\n"
                         "cosi' puoi tornare indietro anche da qui.") % quanti_ora).pack(
            anchor="w", pady=(4, 12))

        cornice = tk.Frame(body, bg=theme.CARD, highlightthickness=1,
                           highlightbackground=theme.BORDER)
        cornice.pack(fill="both", expand=True)
        self.elenco = ttk.Treeview(cornice, columns=("quando", "quanti", "file"),
                                   show="headings", selectmode="browse", height=10,
                                   style="Inv.Treeview")
        for campo, testo, larghezza in (("quando", T("Salvata il"), 170),
                                        ("quanti", T("Dispositivi"), 100),
                                        ("file", T("File"), 260)):
            self.elenco.heading(campo, text=testo)
            self.elenco.column(campo, width=larghezza, anchor="w")
        scorri = ttk.Scrollbar(cornice, orient="vertical", command=self.elenco.yview)
        self.elenco.configure(yscrollcommand=scorri.set)
        self.elenco.pack(side="left", fill="both", expand=True)
        scorri.pack(side="right", fill="y")
        for indice, (percorso, quando, quanti) in enumerate(copie):
            self.elenco.insert("", "end", iid=str(indice),
                               values=(quando.strftime("%d/%m/%Y  %H:%M:%S"), quanti,
                                       os.path.basename(percorso)))
        if copie:
            self.elenco.selection_set("0")
        self.elenco.bind("<Double-1>", lambda e: self._ok())

        buttons = ttk.Frame(body)
        buttons.pack(anchor="e", pady=(16, 0))
        ttk.Button(buttons, text=T("Annulla"), command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text=T("Ripristina"), style="Primary.TButton",
                   command=self._ok).pack(side="right")

    def _ok(self):
        scelta = self.elenco.selection()
        if not scelta:
            messagebox.showwarning(T("Nessuna copia"), T("Scegli una copia dall'elenco."),
                                   parent=self)
            return
        self.result = self.copie[int(scelta[0])][0]
        self.destroy()


class ExportOptionsDialog(_Modal):
    """Primo passo dell'esportazione: che cosa si esporta, e in che forma."""

    def __init__(self, parent, rooms):
        _Modal.__init__(self, parent, T("Esporta in Excel"))
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text=T("Che cosa vuoi esportare"),
                  style="Section.TLabel").pack(anchor="w")
        self.var_ambito = tk.StringVar(value="tutto")
        ttk.Radiobutton(body, variable=self.var_ambito, value="tutto",
                        text=T("Tutto l'inventario"),
                        command=self._aggiorna).pack(anchor="w", pady=(6, 0))
        ttk.Radiobutton(body, variable=self.var_ambito, value="stanza",
                        text=T("Una sola stanza"),
                        command=self._aggiorna).pack(anchor="w", pady=(2, 0))
        riga = ttk.Frame(body)
        riga.pack(anchor="w", fill="x", padx=(24, 0), pady=(4, 0))
        self.var_stanza = tk.StringVar(value=rooms[0] if rooms else "")
        self.combo = ttk.Combobox(riga, textvariable=self.var_stanza, values=rooms,
                                  state="readonly", width=30)
        self.combo.pack(side="left")

        ttk.Separator(body, orient="horizontal").pack(fill="x", pady=12)

        ttk.Label(body, text=T("In che forma"), style="Section.TLabel").pack(anchor="w")
        self.var_forma = tk.StringVar(value="unico")
        self.scelte_forma = []
        for valore, testo in (
            ("unico", "Un unico elenco, in un solo foglio"),
            ("fogli", "Un foglio per ogni stanza, nello stesso file"),
            ("file", "Un file separato per ogni stanza"),
        ):
            b = ttk.Radiobutton(body, variable=self.var_forma, value=valore, text=T(testo))
            b.pack(anchor="w", pady=(6 if valore == "unico" else 2, 0))
            self.scelte_forma.append(b)
        self.nota = ttk.Label(body, style="Muted.TLabel", justify="left",
                              text=T("Ogni foglio porta in testa il nome della stanza,\n"
                                   "la data e il numero di dispositivi."))
        self.nota.pack(anchor="w", padx=(24, 0), pady=(6, 0))

        ttk.Separator(body, orient="horizontal").pack(fill="x", pady=12)
        self.var_inglese = tk.BooleanVar(value=lang.corrente() == lang.INGLESE)
        ttk.Checkbutton(body, variable=self.var_inglese,
                        text=T("Esporta i file in inglese")).pack(anchor="w")
        ttk.Label(body, style="Muted.TLabel", justify="left",
                  text=T("Intestazioni e stati in inglese. Nomi delle stanze e dei\n"
                         "tipi restano come li hai scritti tu.")).pack(
            anchor="w", padx=(24, 0), pady=(4, 0))

        buttons = ttk.Frame(body)
        buttons.pack(anchor="e", pady=(16, 0))
        ttk.Button(buttons, text=T("Annulla"), command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text=T("Avanti"), style="Primary.TButton",
                   command=self._ok).pack(side="right")
        self._aggiorna()

    def _lingua(self):
        return lang.INGLESE if self.var_inglese.get() else lang.ITALIANO

    def _aggiorna(self):
        singola = self.var_ambito.get() == "stanza"
        self.combo.configure(state="readonly" if singola else "disabled")
        for b in self.scelte_forma:
            b.configure(state="disabled" if singola else "normal")
        self.nota.configure(
            text=T("Una stanza sola sta in un file solo, con il suo nome in testa.")
            if singola else T("Ogni foglio porta in testa il nome della stanza,\n"
                              "la data e il numero di dispositivi."))

    def _ok(self):
        if self.var_ambito.get() == "stanza":
            if not self.var_stanza.get():
                messagebox.showwarning(T("Stanza mancante"), T("Scegli la stanza."), parent=self)
                return
            self.result = {"stanza": self.var_stanza.get(), "forma": "unico",
                           "lingua": self._lingua()}
        else:
            self.result = {"stanza": None, "forma": self.var_forma.get(),
                           "lingua": self._lingua()}
        self.destroy()


class ImportOptionsDialog(_Modal):
    """Primo passo dell'importazione: che cosa si carica, e come."""

    def __init__(self, parent, rooms, stanza_fissa=None):
        _Modal.__init__(self, parent,
                        T("Importa in %s") % stanza_fissa if stanza_fissa
                        else T("Importa da Excel"))
        self.stanza_fissa = stanza_fissa
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)

        if stanza_fissa:
            ttk.Label(body, text=T("Importa in %s") % stanza_fissa,
                      style="Section.TLabel").pack(anchor="w")
            ttk.Label(body, style="Muted.TLabel", justify="left",
                      text=T("Se il foglio dichiara le stanze, viene caricata solo la\n"
                           "sezione di questa stanza e il resto si scarta. Se non le\n"
                           "dichiara, tutte le righe finiscono qui.")).pack(
                anchor="w", pady=(4, 0))
            ttk.Separator(body, orient="horizontal").pack(fill="x", pady=12)
            self.var_ambito = tk.StringVar(value="stanza")
            self.var_stanza = tk.StringVar(value=stanza_fissa)
            self.combo = None
        else:
            self._scelta_ambito(body, rooms)
        self._scelta_modo(body)

    def _scelta_ambito(self, body, rooms):
        ttk.Label(body, text=T("Che cosa vuoi caricare"),
                  style="Section.TLabel").pack(anchor="w")
        self.var_ambito = tk.StringVar(value="tutto")
        ttk.Radiobutton(body, variable=self.var_ambito, value="tutto",
                        text=T("Tutto l'inventario"),
                        command=self._aggiorna).pack(anchor="w", pady=(6, 0))
        ttk.Radiobutton(body, variable=self.var_ambito, value="stanza",
                        text=T("Una sola stanza"),
                        command=self._aggiorna).pack(anchor="w", pady=(2, 0))

        riga = ttk.Frame(body)
        riga.pack(anchor="w", fill="x", padx=(24, 0), pady=(4, 0))
        self.var_stanza = tk.StringVar(value=rooms[0] if rooms else "")
        self.combo = ttk.Combobox(riga, textvariable=self.var_stanza, values=rooms,
                                  state="readonly", width=30)
        self.combo.pack(side="left")
        self.nota_stanza = ttk.Label(
            body, style="Muted.TLabel", justify="left",
            text=T("Se il foglio dichiara le stanze con le righe-separatore, viene\n"
                 "caricata solo la sezione della stanza scelta e il resto si scarta.\n"
                 "Se non le dichiara, tutte le righe finiscono nella stanza scelta."))
        self.nota_stanza.pack(anchor="w", padx=(24, 0), pady=(4, 0))

        ttk.Separator(body, orient="horizontal").pack(fill="x", pady=12)

    def _scelta_modo(self, body):
        ttk.Label(body, text=T("Come"), style="Section.TLabel").pack(anchor="w")
        self.var_mode = tk.StringVar(value="merge")
        ttk.Radiobutton(body, variable=self.var_mode, value="merge",
                        text=T("Unisci: aggiunge i nuovi e aggiorna quelli gia' presenti")
                        ).pack(anchor="w", pady=(6, 0))
        ttk.Radiobutton(body, variable=self.var_mode, value="replace",
                        text=T("Sostituisci: svuota prima, poi carica solo il file")
                        ).pack(anchor="w", pady=(2, 0))
        ttk.Label(body, style="Muted.TLabel", justify="left",
                  text=T("Prima di una sostituzione una copia del file dati va in Backup.\n"
                       "Gli iPhone non vengono mai toccati: si inseriscono solo a mano.")
                  ).pack(anchor="w", padx=(24, 0), pady=(4, 0))

        buttons = ttk.Frame(body)
        buttons.pack(anchor="e", pady=(16, 0))
        ttk.Button(buttons, text=T("Annulla"), command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text=T("Scegli il file"), style="Primary.TButton",
                   command=self._ok).pack(side="right")
        self._aggiorna()

    def _aggiorna(self):
        if self.combo is None:
            return
        stato = "readonly" if self.var_ambito.get() == "stanza" else "disabled"
        self.combo.configure(state=stato)

    def _ok(self):
        stanza = self.var_stanza.get() if self.var_ambito.get() == "stanza" else None
        if self.var_ambito.get() == "stanza" and not stanza:
            messagebox.showwarning(T("Stanza mancante"), T("Scegli la stanza."), parent=self)
            return
        self.result = {"stanza": stanza, "mode": self.var_mode.get()}
        self.destroy()


class ImportDialog(_Modal):
    """Riepilogo di quello che si sta per caricare, e conferma finale."""

    def __init__(self, parent, path, count, esito=None, opzioni=None,
                 da_eliminare=0, anteprima=None):
        _Modal.__init__(self, parent, T("Importa inventario"))
        opzioni = opzioni or {"stanza": None, "mode": "merge"}
        self.opzioni = opzioni
        self.sostituzione_totale = (opzioni["mode"] == "replace"
                                    and opzioni["stanza"] is None)
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=os.path.basename(path),
                  style="Section.TLabel").pack(anchor="w")
        esito = esito or {}
        righe = [T("%d righe valide trovate.") % count]
        if esito.get("scartate"):
            righe.append(T("%d righe ignorate: manca l'asset tag.") % esito["scartate"])
        if esito.get("da_tag"):
            righe.append(T("%d righe hanno preso la stanza dai separatori nel foglio.")
                         % esito["da_tag"])
        if esito.get("iphone"):
            righe.append(T("%d iPhone ignorati: si inseriscono solo a mano.") % esito["iphone"])
        doppioni = esito.get("doppioni") or []
        if doppioni:
            righe.append(T("%d asset tag compaiono piu' volte nel foglio (%s): "
                           "vale l'ultima riga.")
                         % (len(doppioni), ", ".join(doppioni[:3])
                            + (T(" e altri") if len(doppioni) > 3 else "")))
        if esito.get("altre_stanze"):
            righe.append(T("%d righe di altre stanze scartate.") % esito["altre_stanze"])
        if esito.get("regola") == "tutte" and opzioni.get("stanza"):
            righe.append(T("Il foglio non dichiara stanze: tutte le righe finiranno "
                         "in %s.") % opzioni["stanza"])
        ttk.Label(body, text=T("\n").join(righe), style="Muted.TLabel",
                  justify="left").pack(anchor="w", pady=(2, 8))

        if anteprima:
            self._dettaglio(body, anteprima, opzioni)

        for testo in self._avvertenze(esito, opzioni, count):
            avviso = tk.Label(body, text=testo, justify="left", anchor="w",
                              bg=theme.LOAN_BG, fg=theme.LOAN_FG,
                              padx=10, pady=8, wraplength=430)
            avviso.pack(fill="x", pady=(0, 8))
        dove = opzioni["stanza"] or T("tutto l'inventario")
        come = T("Sostituzione") if opzioni["mode"] == "replace" else T("Unione")
        tk.Label(body, text=T("%s  \u2192  %s") % (come, dove), anchor="w",
                 bg=theme.HEAD_BG, fg=theme.PRIMARY, padx=10, pady=7,
                 font=self.master.fonts["bold"]).pack(fill="x", pady=(0, 8))

        self.var_conferma = tk.StringVar()
        if opzioni["mode"] == "replace":
            testo = [T("Verranno prima eliminati %d dispositivi gia' in inventario%s.")
                     % (da_eliminare, "" if opzioni["stanza"] is None
                        else T(" in %s") % opzioni["stanza"])]
            testo.append(T("Gli iPhone non vengono toccati."))
            testo.append(T("Una copia del file dati viene salvata prima di procedere."))
            tk.Label(body, text=T("\n").join(testo), justify="left", anchor="w",
                     bg=theme.LOAN_BG, fg=theme.LOAN_FG, padx=10, pady=8,
                     wraplength=430).pack(fill="x", pady=(0, 8))
        if self.sostituzione_totale:
            ttk.Label(body, text=T("Per confermare, scrivi   %s") % PAROLA_RESET).pack(
                anchor="w", pady=(2, 4))
            ttk.Entry(body, textvariable=self.var_conferma, width=34).pack(fill="x")

        buttons = ttk.Frame(body)
        buttons.pack(anchor="e", pady=(16, 0))
        ttk.Button(buttons, text=T("Annulla"), command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text=T("Importa"), style="Primary.TButton",
                   command=self._ok).pack(side="right")

    def _dettaglio(self, body, anteprima, opzioni):
        """Che cosa entra, dove, e quanti dispositivi ci saranno alla fine.

        Un riepilogo con due numeri - "12 aggiunti, 3 aggiornati" - non dice in
        quale stanza finiscono ne' se l'inventario raddoppia. Chi sta per
        riscrivere l'inventario di tutti ha il diritto di vederlo prima.
        """
        cornice = tk.Frame(body, bg=theme.CARD, highlightthickness=1,
                           highlightbackground=theme.BORDER)
        cornice.pack(fill="both", expand=True, pady=(0, 10))
        righe = []
        per_stanza = anteprima.get("per_stanza") or {}
        for stanza in sorted(per_stanza):
            conto = per_stanza[stanza]
            pezzi = []
            if conto["aggiunti"]:
                pezzi.append(T("%d nuovi") % conto["aggiunti"])
            if conto["aggiornati"]:
                pezzi.append(T("%d aggiornati") % conto["aggiornati"])
            righe.append("  %-30s %s" % (stanza, ", ".join(pezzi) or T("niente")))
        if righe:
            righe.insert(0, T("Dove finiscono:"))
        saltate = []
        if anteprima.get("senza_identificativo"):
            saltate.append(T("  %d senza asset tag")
                           % anteprima["senza_identificativo"])
        if saltate:
            righe.append("")
            righe.append(T("Saltate:"))
            righe.extend(saltate)
        righe.append("")
        if anteprima.get("eliminati"):
            righe.append(T("Prima eliminati: %d") % anteprima["eliminati"])
        righe.append(T("In inventario adesso: %d") % anteprima.get("prima", 0))
        righe.append(T("Dopo l'importazione: %d") % anteprima.get("dopo", 0))
        testo = tk.Text(cornice, width=46,
                        height=min(14, max(4, len(righe))), wrap="none",
                        relief="flat", highlightthickness=0,
                        bg=theme.CARD, fg=theme.TEXT, font=self.master.fonts["base"])
        testo.insert("1.0", "\n".join(righe))
        testo.configure(state="disabled")
        testo.pack(fill="both", expand=True, padx=10, pady=8)

    @staticmethod
    def _avvertenze(esito, opzioni=None, count=0):
        """Cosa il file conteneva e il programma non ha potuto usare."""
        messaggi = []
        opzioni = opzioni or {}
        if count and opzioni.get("stanza") is None \
                and not esito.get("stanze_trovate"):
            # e' il difetto che si nota solo dopo: le schede delle stanze
            # restano a zero e tutto finisce nell'elenco completo
            messaggi.append(
                T("Nel foglio non c'e' nessuna riga che dichiari una stanza.\n"
                  "I %d dispositivi verranno importati SENZA STANZA: le schede\n"
                  "delle stanze resteranno vuote.\n\n"
                  "Una riga separatore e' una riga con scritto solo il nome della\n"
                  "stanza, per esempio  Site Services BAU  (vanno bene anche BAU,\n"
                  "KIOSK, DISASTER). Vale per tutte le righe che la seguono.")
                % count)
        ignorate = esito.get("colonne_ignorate") or []
        if ignorate:
            elenco = ", ".join(ignorate[:6])
            if len(ignorate) > 6:
                elenco += T(" e altre %d") % (len(ignorate) - 6)
            messaggi.append(
                T("Colonne non riconosciute, il cui contenuto non verra' importato:\n%s\n"
                  "Se una di queste e' un dato che ti serve, rinominala come la colonna\n"
                  "corrispondente dell'inventario e riprova.") % elenco)
        senza = esito.get("senza_modello")
        if senza:
            messaggi.append(
                T("%d righe non hanno il modello del dispositivo: verranno importate\n"
                  "con quel campo vuoto, da completare a mano.") % senza)
        return messaggi

    def _ok(self):
        if self.sostituzione_totale and \
                self.var_conferma.get().strip().upper() != PAROLA_RESET:
            messagebox.showwarning(
                T("Conferma non valida"),
                T("Stai per svuotare l'inventario di tutti.\n\n"
                "Per procedere scrivi esattamente:\n%s") % PAROLA_RESET, parent=self)
            return
        self.result = self.opzioni
        self.destroy()


# ------------------------------------------------------------ scheda stanza


class EliminaPlusDialog(_Modal):
    """Eliminazione in blocco: si incollano i codici, si legge, si conferma.

    Il senso e' togliere trenta dispositivi senza cercarli uno per uno, ma senza
    rinunciare a sapere che cosa sparisce: prima di eliminare si vede l'elenco
    completo, diviso per stanza, con quello che verra' saltato e perche'.
    """

    def __init__(self, parent, store):
        _Modal.__init__(self, parent, T("Elimina piu' dispositivi"))
        self.store = store
        self.esito = None
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text=T("Incolla qui i dispositivi da eliminare"),
                  style="Section.TLabel").pack(anchor="w")
        ttk.Label(body, style="Muted.TLabel", justify="left",
                  text=T("Uno per riga. Va bene incollare una colonna di asset tag da\n"
                         "Excel, o righe intere: viene letto il primo codice che\n"
                         "corrisponde a un dispositivo in inventario.")).pack(
            anchor="w", pady=(2, 8))

        cornice = tk.Frame(body, bg=theme.CARD, highlightthickness=1,
                           highlightbackground=theme.BORDER)
        cornice.pack(fill="both", expand=True)
        self.testo = tk.Text(cornice, width=46, height=8, wrap="none", relief="flat",
                             highlightthickness=0, bg=theme.CARD, fg=theme.TEXT)
        barra = ttk.Scrollbar(cornice, orient="vertical", command=self.testo.yview)
        self.testo.configure(yscrollcommand=barra.set)
        self.testo.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        self.testo.focus_set()

        ttk.Button(body, text=T("Controlla"), style="Primary.TButton",
                   command=self._controlla).pack(anchor="w", pady=(10, 0))

        self.riepilogo = tk.Text(body, width=46, height=12, wrap="word", relief="solid",
                                 borderwidth=1, highlightthickness=0, bg=theme.CARD,
                                 fg=theme.TEXT, state="disabled")
        self.riepilogo.pack(fill="both", expand=True, pady=(12, 0))

        self.conferma = ttk.Frame(body)
        self.conferma.pack(fill="x", pady=(10, 0))
        self.var_conferma = tk.StringVar()

        buttons = ttk.Frame(body)
        buttons.pack(anchor="e", pady=(14, 0))
        ttk.Button(buttons, text=T("Annulla"), command=self._cancel).pack(side="right", padx=6)
        self.btn_elimina = ttk.Button(buttons, text=T("Elimina"), style="Rosso.TButton",
                                      command=self._ok, state="disabled")
        self.btn_elimina.pack(side="right")

    # ------------------------------------------------------------ controllo

    def righe_incollate(self):
        return [r for r in self.testo.get("1.0", "end").splitlines() if r.strip()]

    def _controlla(self):
        righe = self.righe_incollate()
        if not righe:
            self._scrivi([T("Non hai incollato niente.")])
            return
        da_eliminare, non_trovati, bloccati = self.store.anteprima_eliminazione(righe)
        self.esito = (da_eliminare, non_trovati, bloccati)
        self._scrivi(self._testo_riepilogo(righe, da_eliminare, non_trovati, bloccati))
        for figlio in self.conferma.winfo_children():
            figlio.destroy()
        if not da_eliminare:
            self.btn_elimina.state(["disabled"])
            return
        ttk.Label(self.conferma,
                  text=T("Per eliminare questi %d dispositivi, scrivi   %s")
                  % (len(da_eliminare), PAROLA_ELIMINA)).pack(anchor="w")
        self.var_conferma.set("")
        ttk.Entry(self.conferma, textvariable=self.var_conferma,
                  width=34).pack(fill="x", pady=(4, 0))
        self.btn_elimina.state(["!disabled"])

    def _testo_riepilogo(self, righe, da_eliminare, non_trovati, bloccati):
        fuori = [T("Righe incollate: %d") % len(righe), ""]
        if da_eliminare:
            per_stanza = {}
            for item in da_eliminare:
                per_stanza.setdefault(item.get("stanza") or T("(senza stanza)"),
                                      []).append(item)
            fuori.append(T("VERRANNO ELIMINATI: %d") % len(da_eliminare))
            for stanza in sorted(per_stanza):
                elenco = per_stanza[stanza]
                fuori.append("")
                fuori.append(T("  %s - %d dispositivi") % (stanza, len(elenco)))
                for item in elenco:
                    riga = "    %s  %s" % (valore_visibile(item, "asset_tag"),
                                           item.get("modello") or "")
                    if item.get("prestato_a"):
                        riga += T("   [in prestito a %s]") % item["prestato_a"]
                    fuori.append(riga.rstrip())
        else:
            fuori.append(T("Non c'e' niente da eliminare."))
        if bloccati:
            fuori.append("")
            fuori.append(T("SALTATI perche' non si possono eliminare: %d") % len(bloccati))
            for item, motivo in bloccati:
                fuori.append("    %s  -  %s"
                             % (valore_visibile(item, "asset_tag"), motivo))
        if non_trovati:
            fuori.append("")
            fuori.append(T("SALTATI perche' non sono in inventario: %d") % len(non_trovati))
            for codice in non_trovati[:20]:
                fuori.append("    %s" % codice)
            if len(non_trovati) > 20:
                fuori.append(T("    e altri %d") % (len(non_trovati) - 20))
        return fuori

    def _scrivi(self, righe):
        self.riepilogo.configure(state="normal")
        self.riepilogo.delete("1.0", "end")
        self.riepilogo.insert("1.0", "\n".join(righe))
        self.riepilogo.configure(state="disabled")

    def _ok(self):
        if not self.esito or not self.esito[0]:
            return
        if self.var_conferma.get().strip().upper() != PAROLA_ELIMINA:
            messagebox.showwarning(
                T("Conferma non valida"),
                T("Stai per eliminare %d dispositivi dall'inventario di tutti.\n\n"
                  "Per procedere scrivi esattamente:\n%s")
                % (len(self.esito[0]), PAROLA_ELIMINA), parent=self)
            return
        self.result = [i["asset_tag"] for i in self.esito[0]]
        self.destroy()


def _per_stanza(items):
    """I dispositivi raggruppati per stanza, in ordine di nome."""
    gruppi = {}
    for item in items:
        gruppi.setdefault(item.get("stanza") or T("(senza stanza)"), []).append(item)
    return sorted(gruppi.items())


def _riga_dispositivo(item, con_prestito=True):
    riga = "    %s  %s" % (valore_visibile(item, "asset_tag") or
                           valore_visibile(item, "imei"),
                           (item.get("modello") or "")[:34])
    if con_prestito and item.get("prestato_a"):
        riga += T("   [in prestito a %s]") % item["prestato_a"]
    return riga.rstrip()


def riepilogo_eliminazione(da_eliminare, bloccati, restano):
    """Le righe da leggere prima di eliminare.

    Un'operazione che tocca venti dispositivi non si conferma con un si': chi la
    lancia deve poter vedere quali spariscono, da quale stanza, e che cosa il
    programma salta e perche'.
    """
    righe = []
    if da_eliminare:
        righe.append(T("VERRANNO ELIMINATI: %d") % len(da_eliminare))
        for stanza, elenco in _per_stanza(da_eliminare):
            righe.append("")
            righe.append(T("  %s - %d dispositivi") % (stanza, len(elenco)))
            righe.extend(_riga_dispositivo(i) for i in elenco)
    else:
        righe.append(T("Non c'e' niente da eliminare."))
    if bloccati:
        righe.append("")
        righe.append(T("SALTATI perche' non si possono eliminare: %d") % len(bloccati))
        for item, motivo in bloccati:
            righe.append("    %s  -  %s" % (valore_visibile(item, "asset_tag") or
                                            valore_visibile(item, "imei"), motivo))
    righe.append("")
    righe.append(T("In inventario resteranno %d dispositivi.") % restano)
    return righe


def riepilogo_spostamento(spostabili, telefoni, prestati, destinazione,
                          stanza_iphone, conteggi):
    """Le righe da leggere prima di spostare fra stanze."""
    righe = []
    restano_dove_sono = [i for i in spostabili if i.get("stanza") == destinazione]
    da_spostare = [i for i in spostabili if i.get("stanza") != destinazione]
    if da_spostare:
        righe.append(T("SPOSTATI IN %s: %d") % (destinazione, len(da_spostare)))
        for stanza, elenco in _per_stanza(da_spostare):
            righe.append("")
            righe.append(T("  da %s - %d dispositivi") % (stanza, len(elenco)))
            righe.extend(_riga_dispositivo(i) for i in elenco)
    else:
        righe.append(T("Nessun cambiamento: sono gia' tutti in %s.") % destinazione)
    if restano_dove_sono and da_spostare:
        righe.append("")
        righe.append(T("Gia' in %s, restano dove sono: %d")
                     % (destinazione, len(restano_dove_sono)))
    if telefoni:
        righe.append("")
        righe.append(T("RESTANO FERMI perche' sono iPhone: %d") % len(telefoni))
        righe.append(T("Gli iPhone sono sempre registrati in %s.") % stanza_iphone)
    if prestati:
        righe.append("")
        righe.append(T("RESTANO FERMI perche' sono in prestito: %d") % len(prestati))
        righe.append(T("Registra prima il rientro, poi si potranno spostare."))
        righe.extend(_riga_dispositivo(i) for i in prestati)
    if conteggi and da_spostare:
        righe.append("")
        righe.append(T("Come restano le stanze:"))
        for stanza, (prima, dopo) in conteggi:
            if prima != dopo:
                righe.append("  %-28s %d  ->  %d" % (stanza, prima, dopo))
    return righe


class CestinoDialog(_Modal):
    """Gli eliminati di recente, con il ripristino riga per riga.

    Fa vedere due sole colonne - asset tag e tipo - perche' serve a ritrovare
    un dispositivo, non a consultarne la scheda: quella torna intera con il
    ripristino. La data e la stanza di provenienza si leggono in fondo, sulla
    riga scelta.
    """

    PER_PAGINA = 10

    def __init__(self, parent, store, stanze):
        _Modal.__init__(self, parent, T("Eliminati di recente"))
        self.store = store
        self.stanze = list(stanze)
        self.app = parent
        self.pagina = 0
        self.ripristinati = 0
        self._pulsanti = {}

        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, style="Muted.TLabel", wraplength=520, justify="left",
                  text=T("Restano qui %d giorni dall'eliminazione, e al massimo %d.\n"
                         "Non compaiono in nessuna ricerca, esportazione o stampa.")
                  % (ELIMINATI_GIORNI, ELIMINATI_MASSIMO)).pack(anchor="w",
                                                                pady=(0, 10))

        riga = ttk.Frame(body)
        riga.pack(fill="x")
        ttk.Label(riga, text=T("Cerca")).pack(side="left")
        self.var_cerca = tk.StringVar()
        self.var_cerca.trace_add("write", lambda *a: self._cerca_cambiata())
        entry = ttk.Entry(riga, textvariable=self.var_cerca, width=28)
        entry.pack(side="left", padx=(6, 14))
        self.btn_ripristina = ttk.Button(riga, text=T("Ripristina i selezionati"),
                                         style="Primary.TButton",
                                         command=self._ripristina_selezionati)
        self.btn_ripristina.pack(side="left")

        cornice = ttk.Frame(body)
        cornice.pack(fill="both", expand=True, pady=(12, 0))
        colonne = ("asset_tag", "tipo", ACTION_COLUMN)
        self.elenco = ttk.Treeview(cornice, columns=colonne, show="headings",
                                   height=self.PER_PAGINA, selectmode="extended",
                                   style="Inv.Treeview")
        self.elenco.heading("asset_tag", text=intestazione(HEADERS["asset_tag"]))
        self.elenco.heading("tipo", text=intestazione(HEADERS["tipo"]))
        self.elenco.heading(ACTION_COLUMN, text="")
        self.elenco.column("asset_tag", width=200, anchor="w", stretch=False)
        self.elenco.column("tipo", width=150, anchor="w", stretch=False)
        self.elenco.column(ACTION_COLUMN, width=140, anchor="w", stretch=False)
        self.elenco.pack(fill="both", expand=True)
        self.elenco.bind("<<TreeviewSelect>>", lambda e: self._mostra_dettaglio())
        self.elenco.bind("<Configure>", self._disegna_pulsanti)
        # Copia e incolla come nell'elenco principale: la scorciatoia e il
        # menu del tasto destro. installa_copia_incolla() copre i campi di
        # testo, ma una tabella non e' un campo di testo.
        self.elenco.bind("<Control-c>", self._copia)
        self.elenco.bind("<Command-c>", self._copia)
        for tasto in ("<Button-3>", "<Button-2>"):
            self.elenco.bind(tasto, self._menu)

        self.var_dettaglio = tk.StringVar()
        ttk.Label(body, textvariable=self.var_dettaglio, style="Muted.TLabel",
                  wraplength=520, justify="left").pack(anchor="w", pady=(8, 0))

        piede = ttk.Frame(body)
        piede.pack(fill="x", pady=(12, 0))
        self.btn_prima = ttk.Button(piede, text=T("‹  Precedenti"),
                                    command=lambda: self._vai(-1))
        self.btn_prima.pack(side="left")
        self.var_pagina = tk.StringVar()
        ttk.Label(piede, textvariable=self.var_pagina,
                  style="Muted.TLabel").pack(side="left", padx=10)
        self.btn_dopo = ttk.Button(piede, text=T("Successivi  ›"),
                                   command=lambda: self._vai(1))
        self.btn_dopo.pack(side="left")
        ttk.Button(piede, text=T("Chiudi"), command=self._chiudi).pack(side="right")

        self._ricarica()
        entry.focus_set()

    # ---------------------------------------------------------- contenuto

    def _ricarica(self):
        self.voci = self.store.eliminati(self.var_cerca.get())
        pagine = max(1, (len(self.voci) + self.PER_PAGINA - 1) // self.PER_PAGINA)
        self.pagina = min(self.pagina, pagine - 1)
        inizio = self.pagina * self.PER_PAGINA
        self.in_pagina = self.voci[inizio:inizio + self.PER_PAGINA]
        self.elenco.delete(*self.elenco.get_children())
        for voce in self.in_pagina:
            self.elenco.insert("", "end", iid=voce["asset_tag"],
                               values=(voce.get("asset_tag", ""),
                                       voce.get("tipo", ""), ""),
                               tags=("orfano",) if voce.get("orfano") else ())
        self.elenco.tag_configure("orfano", background=theme.LOAN_BG)
        if not self.voci:
            self.var_pagina.set(T("Nessun dispositivo eliminato di recente."))
        else:
            self.var_pagina.set(T("%d-%d di %d")
                                % (inizio + 1, inizio + len(self.in_pagina),
                                   len(self.voci)))
        self.btn_prima.configure(state="normal" if self.pagina > 0 else "disabled")
        self.btn_dopo.configure(
            state="normal" if inizio + self.PER_PAGINA < len(self.voci) else "disabled")
        self.btn_ripristina.configure(state="normal" if self.voci else "disabled")
        self._mostra_dettaglio()
        self.after_idle(self._disegna_pulsanti)

    def _tabella_disposta(self, righe):
        """Vero quando ogni riga ha davvero il suo posto sullo schermo.

        Appena costruita, la tabella non e' ancora stata disposta: bbox()
        risponde per la prima riga con le coordinate della riga delle
        intestazioni - y a zero - e per le altre non risponde affatto. Un
        pulsante messo li' finisce accanto ai titoli delle colonne invece che
        accanto al suo dispositivo, e le altre righe restano senza. Si aspetta
        che le coordinate siano vere.
        """
        for tag in righe:
            try:
                box = self.elenco.bbox(tag, 2)
            except tk.TclError:
                return False
            if not box:
                return False
        return self.elenco.bbox(righe[0], 2)[1] > 0

    def _disegna_pulsanti(self, _evento=None, tentativi=0):
        """Un pulsante Ripristina vero, sopra la cella di ogni riga."""
        if not self.elenco.winfo_exists():
            return
        righe = self.elenco.get_children()
        if righe and not self._tabella_disposta(righe):
            if tentativi < 25:      # ~mezzo secondo, poi si disegna comunque
                self.after(20, lambda: self._disegna_pulsanti(tentativi=tentativi + 1))
                return
        indice = 2
        vivi = set()
        for tag in righe:
            try:
                box = self.elenco.bbox(tag, indice)
            except tk.TclError:
                box = None
            if not box:
                continue
            vivi.add(tag)
            pulsante = self._pulsanti.get(tag)
            if pulsante is None:
                pulsante = ttk.Button(self.elenco, text=T("Ripristina"),
                                      style="Row.TButton", takefocus=False,
                                      command=lambda t=tag: self._ripristina([t]))
                self._pulsanti[tag] = pulsante
            x, y, larghezza, altezza = box
            pulsante.place(x=x + 6, y=y + 3, width=max(larghezza - 12, 40),
                           height=max(altezza - 6, 18))
        for tag in [t for t in self._pulsanti if t not in vivi]:
            self._pulsanti.pop(tag).destroy()

    def _cerca_cambiata(self):
        self.pagina = 0
        self._ricarica()

    def _vai(self, passo):
        self.pagina = max(0, self.pagina + passo)
        self._ricarica()

    def _voce(self, tag):
        return next((v for v in self.voci if v.get("asset_tag") == tag), None)

    def _mostra_dettaglio(self):
        scelti = self.elenco.selection()
        if len(scelti) != 1:
            self.var_dettaglio.set(
                T("%d selezionati") % len(scelti) if scelti else "")
            return
        voce = self._voce(scelti[0])
        if voce is None:
            self.var_dettaglio.set("")
            return
        if voce.get("orfano"):
            dove = T("nessuna stanza: la sua e' stata tolta, te la chiedera'")
        else:
            dove = T("torna in %s") % (voce.get("stanza") or T("(senza stanza)"))
        self.var_dettaglio.set(T("Eliminato il %s da %s  -  %s")
                               % (voce.get("eliminato_il", "?"),
                                  voce.get("eliminato_da", "?"), dove))

    def _copia(self, _evento=None, solo_identificativo=False):
        scelti = self.elenco.selection()
        if not scelti:
            return "break"
        righe = []
        for tag in scelti:
            valori = self.elenco.item(tag, "values")
            righe.append(valori[0] if solo_identificativo
                         else "%s\t%s" % (valori[0], valori[1]))
        self.clipboard_clear()
        self.clipboard_append("\n".join(righe))
        return "break"

    def _menu(self, evento):
        riga = self.elenco.identify_row(evento.y)
        if riga and riga not in self.elenco.selection():
            self.elenco.selection_set([riga])
        quante = len(self.elenco.selection())
        if not quante:
            return "break"
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label=T("Ripristina (%d)") % quante,
                         command=lambda: self._ripristina(list(self.elenco.selection())))
        menu.add_separator()
        menu.add_command(label=T("Copia le righe (%d)") % quante,
                         command=self._copia)
        menu.add_command(label=T("Copia l'identificativo"),
                         command=lambda: self._copia(solo_identificativo=True))
        try:
            menu.tk_popup(evento.x_root, evento.y_root)
        finally:
            menu.grab_release()
        return "break"

    # --------------------------------------------------------- ripristino

    def _ripristina_selezionati(self):
        scelti = list(self.elenco.selection())
        if not scelti:
            messagebox.showinfo(T("Ripristina"),
                                T("Scegli i dispositivi da ripristinare."), parent=self)
            return
        self._ripristina(scelti)

    def _ripristina(self, tags):
        voci = [self._voce(t) for t in tags]
        voci = [v for v in voci if v is not None]
        if not voci:
            return
        stanza = None
        orfani = [v for v in voci if v.get("orfano") or not clean(v.get("stanza"))]
        if orfani:
            stanza = self._chiedi_stanza(orfani)
            if not stanza:
                return
        try:
            rimessi, saltati = self.store.ripristina_eliminati(
                [v["asset_tag"] for v in voci], stanza)
        except InventoryError as exc:
            messagebox.showerror(T("Ripristino non riuscito"), str(exc), parent=self)
            return
        self.ripristinati += len(rimessi)
        self._ricarica()
        righe = []
        if rimessi:
            righe.append(T("Ripristinati %d dispositivi.") % len(rimessi))
        if saltati:
            righe.append("")
            righe.append(T("SALTATI: %d") % len(saltati))
            righe.extend("    %s  -  %s" % (tag, motivo) for tag, motivo in saltati)
        messagebox.showinfo(T("Ripristino"), "\n".join(righe), parent=self)

    def _chiedi_stanza(self, orfani):
        """Un orfano non ha una stanza a cui tornare: si chiede quale."""
        elenco = ", ".join(v["asset_tag"] for v in orfani[:5])
        if len(orfani) > 5:
            elenco += "..."
        return self.app._ask_room(
            T("Questi dispositivi erano in una stanza che non esiste piu':\n%s\n\n"
              "In che stanza rimetterli?") % elenco)

    def _chiudi(self):
        self.result = self.ripristinati
        self.destroy()

    def _cancel(self):
        self._chiudi()


def riepilogo_copia_locale(rapporto, adesso, stanze_adesso, cfg_adesso):
    """Che cosa tornerebbe indietro ripristinando una copia locale.

    Il ripristino piu' pesante che il programma sappia fare: riscrive
    l'inventario di tutti e anche le stanze. Va letto prima, non dopo.
    """
    righe = [T("La copia contiene %d dispositivi; adesso in inventario ce ne sono %d.")
             % (rapporto["dispositivi"], adesso)]
    if rapporto.get("quando"):
        righe.append(T("Risale al %s.")
                     % rapporto["quando"].strftime("%d/%m/%Y %H:%M"))
    righe.append("")
    righe.append(T("COME RESTANO LE STANZE:"))
    dalla_copia = rapporto.get("per_stanza") or {}
    for stanza in sorted(set(dalla_copia) | set(stanze_adesso)):
        righe.append("  %-28s %d  ->  %d"
                     % (stanza, stanze_adesso.get(stanza, 0), dalla_copia.get(stanza, 0)))

    impostazioni = rapporto.get("impostazioni")
    righe.append("")
    if not impostazioni:
        righe.append(T("La copia non porta le impostazioni: stanze, tipi e stati\n"
                       "restano quelli di adesso."))
        return righe
    righe.append(T("TORNANO ANCHE LE IMPOSTAZIONI:"))
    for chiave, etichetta in (("rooms", T("Stanze")),
                              ("types", T("Tipi di dispositivo")),
                              ("loan_rooms", T("Stanze con prestito")),
                              ("iphone_room", T("Stanza degli iPhone"))):
        valore = impostazioni.get(chiave)
        if isinstance(valore, list):
            valore = ", ".join(valore)
        attuale = cfg_adesso.get(chiave)
        if isinstance(attuale, list):
            attuale = ", ".join(attuale)
        segno = "  " if clean(str(valore)) == clean(str(attuale)) else "* "
        righe.append("%s%s: %s" % (segno, etichetta, valore or T("(niente)")))
    if any(r.startswith("* ") for r in righe):
        righe.append("")
        righe.append(T("Le righe con * cambiano rispetto a com'e' adesso."))
    return righe


class ConfermaOperazioneDialog(_Modal):
    """Che cosa sta per succedere, riga per riga, prima di farlo succedere."""

    def __init__(self, parent, titolo, intestazione, righe, testo_ok,
                 stile_ok="Primary.TButton", avvertenza=None):
        _Modal.__init__(self, parent, titolo)
        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=intestazione, style="Section.TLabel",
                  wraplength=430, justify="left").pack(anchor="w", pady=(0, 10))

        cornice = tk.Frame(body, bg=theme.CARD, highlightthickness=1,
                           highlightbackground=theme.BORDER)
        cornice.pack(fill="both", expand=True)
        testo = tk.Text(cornice, width=52, height=min(16, max(4, len(righe))),
                        wrap="none", relief="flat", highlightthickness=0,
                        bg=theme.CARD, fg=theme.TEXT, font=self.master.fonts["base"])
        testo.insert("1.0", "\n".join(righe))
        testo.configure(state="disabled")
        barra = ttk.Scrollbar(cornice, orient="vertical", command=testo.yview)
        testo.configure(yscrollcommand=barra.set)
        testo.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        barra.pack(side="right", fill="y")

        if avvertenza:
            tk.Label(body, text=avvertenza, justify="left", anchor="w",
                     bg=theme.LOAN_BG, fg=theme.LOAN_FG, padx=10, pady=8,
                     wraplength=430).pack(fill="x", pady=(10, 0))

        buttons = ttk.Frame(body)
        buttons.pack(anchor="e", pady=(14, 0))
        ttk.Button(buttons, text=T("Annulla"),
                   command=self._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text=testo_ok, style=stile_ok,
                   command=self._ok).pack(side="right")

    def _ok(self):
        self.result = True
        self.destroy()


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
        link = tk.Label(inner, text=T("Apri l'inventario  ›"), bg=theme.CARD,
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
        self.title(T("Site Services : Inventario Iphone, Laptop e Tablet")
                   + "  -  v" + __version__)
        self.geometry("1220x720")
        self.minsize(980, 560)

        lang.imposta(config.load_language())
        self.fonts = theme.apply(self)
        installa_copia_incolla(self)
        # Command esiste solo su macOS: dove non c'e', quel bit di stato
        # significa altro e non va letto.
        self._ha_comando = self.tk.call("tk", "windowingsystem") == "aqua"
        self.cfg = config.load_shared_config(data_path)
        self.store = InventoryStore(data_path,
                                    iphone_room=self.cfg.get("iphone_room"),
                                    stati=self.cfg.get("states"))
        self.sort_field = "modificato_il"
        self.sort_reverse = True        # il piu' recente in cima
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

    def _cambia_lingua(self, _event=None):
        """La tendina nell'intestazione: cambia lingua e ridisegna la finestra."""
        scelte = dict((nome, codice) for nome, codice in lang.LINGUE)
        nuova = scelte.get(self.var_lingua.get(), lang.ITALIANO)
        if nuova == lang.corrente():
            return
        lang.imposta(nuova)
        config.save_language(nuova)
        self.ricostruisci()

    def ricostruisci(self):
        """Ridisegna tutta la finestra: serve dopo un cambio di lingua."""
        self._clear_row_buttons()
        for figlio in self.winfo_children():
            figlio.destroy()
        self.title(T("Site Services : Inventario Iphone, Laptop e Tablet")
                   + "  -  v" + __version__)
        self.cfg = config.load_shared_config(self.store.path)
        self.store.iphone_room = self.cfg.get("iphone_room")
        self.store.stati = list(self.cfg.get("states") or [])
        self.tree = None
        self._build_header()
        self._build_toolbar()
        self._build_filters()
        self.body = ttk.Frame(self, padding=(16, 4, 16, 8))
        self.body.pack(fill="both", expand=True)
        self._build_status()
        self._sync_filter_values()
        self.show_home()

    def _build_header(self):
        head = ttk.Frame(self, style="Head.TFrame", padding=(20, 14))
        head.pack(fill="x")
        left = ttk.Frame(head, style="Head.TFrame")
        left.pack(side="left")
        ttk.Label(left, text=T("Site Services : Inventario Iphone, Laptop e Tablet"),
                  style="HeadTitle.TLabel").pack(anchor="w")
        ttk.Label(left, text="v" + __version__,
                  style="HeadSub.TLabel").pack(anchor="w")
        self.var_subtitle = tk.StringVar(value=T("Laptop e tablet in nostro possesso"))
        ttk.Label(left, textvariable=self.var_subtitle,
                  style="HeadSub.TLabel").pack(anchor="w", pady=(2, 0))
        destra = ttk.Frame(head, style="Head.TFrame")
        destra.pack(side="right")
        scelta = ttk.Frame(destra, style="Head.TFrame")
        scelta.pack(anchor="e")
        ttk.Label(scelta, text=T("Lingua"), style="HeadSub.TLabel").pack(side="left",
                                                                        padx=(0, 6))
        self.var_lingua = tk.StringVar(value=lang.nome_lingua(lang.corrente()))
        combo = ttk.Combobox(scelta, textvariable=self.var_lingua,
                             values=[nome for nome, _ in lang.LINGUE],
                             state="readonly", width=11)
        combo.pack(side="left")
        combo.bind("<<ComboboxSelected>>", self._cambia_lingua)
        self.var_head_count = tk.StringVar(value="")
        ttk.Label(destra, textvariable=self.var_head_count,
                  style="HeadSub.TLabel").pack(anchor="e", pady=(6, 0))

    def _build_toolbar(self):
        """La barra dei comandi, in una riga o in due a seconda dello spazio.

        In una riga sola i pulsanti che non ci stavano non venivano stretti:
        Tk semplicemente non li disponeva, e sparivano senza dire niente -
        Impostazioni compresa, che e' la porta di tutto il resto. Qui i comandi
        stanno in tre gruppi e i gruppi vanno a capo quando serve, cosi' non
        se ne perde mai nessuno.
        """
        self.barra = ttk.Frame(self, padding=(16, 12, 16, 4))
        self.barra.pack(fill="x")
        self.barra.columnconfigure(1, weight=1)

        # 1. i comandi di tutti i giorni
        self.gruppo_lavoro = ttk.Frame(self.barra)
        self.btn_home = ttk.Button(self.gruppo_lavoro, text=T("‹  Home"),
                                   style="Ghost.TButton", command=self.show_home)
        self.btn_home.pack(side="left", padx=(0, 10))
        ttk.Button(self.gruppo_lavoro, text=T("Aggiungi"), style="Primary.TButton",
                   command=self.on_add).pack(side="left", padx=(0, 6))
        for text, command in (("Modifica", self.on_edit), ("Elimina", self.on_delete)):
            ttk.Button(self.gruppo_lavoro, text=T(text),
                       command=command).pack(side="left", padx=(0, 6))
        ttk.Button(self.gruppo_lavoro, text=T("Elimina +"), style="Rosso.TButton",
                   command=self.on_delete_many).pack(side="left", padx=(0, 6))
        ttk.Button(self.gruppo_lavoro, text=T("Sposta in stanza..."),
                   command=self.on_move).pack(side="left", padx=(0, 6))

        # 2. i file che entrano ed escono. Il colore dice a quale famiglia
        # appartiene il comando: arancione i dati che entrano, verde quelli che
        # escono, rosso quello che riscrive l'inventario di tutti.
        self.gruppo_file = ttk.Frame(self.barra)
        self.sep_barra = ttk.Separator(self.gruppo_file, orient="vertical")
        self.sep_barra.pack(side="left", fill="y", padx=10)
        for text, command, stile in (
            ("Importa xls...", self.on_import, "Arancio.TButton"),
            ("Esporta xls...", self.on_export, "Verde.TButton"),
            ("Stampa", self.on_print, "TButton"),
        ):
            ttk.Button(self.gruppo_file, text=T(text), style=stile,
                       command=command).pack(side="left", padx=(0, 6))

        # 3. quello che non riguarda i dispositivi
        self.gruppo_lato = ttk.Frame(self.barra)
        ttk.Button(self.gruppo_lato, text=T("Aggiorna"),
                   command=self.on_refresh).pack(side="left", padx=(0, 6))
        ttk.Button(self.gruppo_lato, text=T("Impostazioni"),
                   command=self.on_settings).pack(side="left")

        self._righe_barra = None
        self.barra.bind("<Configure>", self._sistema_barra)
        self._sistema_barra()

    def _sistema_barra(self, event=None):
        """Decide se i comandi stanno in una riga o in due.

        Cambia la disposizione solo quando cambia davvero: <Configure> arriva a
        ogni pixel di ridimensionamento, e rifare la griglia ogni volta
        sarebbe lavoro buttato. Andare a capo cambia l'altezza, non la
        larghezza, quindi la scelta non puo' rimbalzare fra i due stati.
        """
        larghezza = event.width if event is not None else self.barra.winfo_width()
        if larghezza <= 1:
            return
        lavoro = self.gruppo_lavoro.winfo_reqwidth()
        file_ = self.gruppo_file.winfo_reqwidth()
        lato = self.gruppo_lato.winfo_reqwidth()
        if lavoro + file_ + lato <= larghezza:
            stato = "una"
        elif lavoro + lato <= larghezza:
            stato = "due"
        else:
            stato = "due_basso"
        if stato == self._righe_barra:
            return
        self._righe_barra = stato
        for gruppo in (self.gruppo_lavoro, self.gruppo_file, self.gruppo_lato):
            gruppo.grid_forget()
        self.sep_barra.pack_forget()
        if stato == "una":
            self.sep_barra.pack(side="left", fill="y", padx=10, before=self.sep_dopo())
            self.gruppo_lavoro.grid(row=0, column=0, sticky="w")
            self.gruppo_file.grid(row=0, column=1, sticky="w")
            self.gruppo_lato.grid(row=0, column=2, sticky="e")
        elif stato == "due":
            self.gruppo_lavoro.grid(row=0, column=0, sticky="w")
            self.gruppo_lato.grid(row=0, column=2, sticky="e")
            self.gruppo_file.grid(row=1, column=0, columnspan=3, sticky="w",
                                  pady=(6, 0))
        else:
            self.gruppo_lavoro.grid(row=0, column=0, columnspan=3, sticky="w")
            self.gruppo_file.grid(row=1, column=0, sticky="w", pady=(6, 0))
            self.gruppo_lato.grid(row=1, column=2, sticky="e", pady=(6, 0))

    def sep_dopo(self):
        """Il primo pulsante del gruppo file: il separatore gli sta davanti."""
        figli = [c for c in self.gruppo_file.winfo_children() if c is not self.sep_barra]
        return figli[0] if figli else None

    def _build_filters(self):
        bar = ttk.Frame(self, padding=(16, 4))
        bar.pack(fill="x")
        ttk.Label(bar, text=T("Cerca")).pack(side="left")
        self.var_search = tk.StringVar()
        self.var_search.trace_add("write", lambda *a: self.refresh_table())
        self.entry_search = ttk.Entry(bar, textvariable=self.var_search, width=34)
        self.entry_search.pack(side="left", padx=(6, 16))

        self.label_room = ttk.Label(bar, text=T("Stanza"))
        self.label_room.pack(side="left")
        self.var_room = tk.StringVar(value=TUTTE())
        self.combo_room = ttk.Combobox(bar, textvariable=self.var_room, state="readonly", width=20)
        self.combo_room.pack(side="left", padx=(6, 16))
        self.combo_room.bind("<<ComboboxSelected>>", self._on_room_filter)

        ttk.Label(bar, text=T("Tipo")).pack(side="left")
        self.var_type = tk.StringVar(value=TUTTI())
        self.combo_type = ttk.Combobox(bar, textvariable=self.var_type, state="readonly", width=14)
        self.combo_type.pack(side="left", padx=(6, 16))
        self.combo_type.bind("<<ComboboxSelected>>", self._on_type_filter)

        ttk.Button(bar, text=T("Azzera filtri"), command=self.reset_filters).pack(side="left")

    def _build_status(self):
        self.var_status = tk.StringVar()
        ttk.Separator(self).pack(fill="x")
        ttk.Label(self, textvariable=self.var_status, style="Status.TLabel",
                  anchor="w", padding=(16, 6)).pack(fill="x")

    def _bind_keys(self):
        self.bind("<Control-n>", lambda e: self.on_add())
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
            return "loan_alt" if dispari else "loan"
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

    def _campi_visibili(self):
        """I campi che hanno senso nella vista aperta.

        Una colonna vuota per costruzione - il prestito dove i prestiti non
        esistono, l'IMEI dove non ci sono telefoni, la stanza dentro una stanza
        sola - non porta informazione: toglie spazio a quello che si deve
        leggere davvero.

        Si decide da come e' configurato l'inventario, non dai dati presenti in
        quel momento: cosi' le colonne non compaiono e spariscono mentre si
        lavora, e una stanza vuota mostra le stesse colonne di quando sara'
        piena.

        L'esportazione e la stampa non passano di qui: portano via tutto.
        """
        campi = list(ALL_FIELDS)

        def togli(*nomi):
            for nome in nomi:
                if nome in campi:
                    campi.remove(nome)

        contenitore_iphone = self.ship_column_visible()
        stanza = self.var_room.get() if self.view == "room" else None

        # Quello che e' uguale su tutte le righe sta nel titolo, non in colonna.
        if stanza is not None or contenitore_iphone:
            togli("stanza")
        if self.view == "type":
            togli("tipo")

        # I campi dei telefoni si vedono dove i telefoni possono esserci.
        if contenitore_iphone:
            ci_sono_iphone = True
        elif not self.iphone_type() or self.view == "type":
            ci_sono_iphone = False
        elif stanza is not None:
            ci_sono_iphone = (stanza == self.iphone_room())
        else:
            ci_sono_iphone = True
        if not ci_sono_iphone:
            togli("imei", "restituito_da", "spedito_il")
        if contenitore_iphone:
            togli(*COLONNE_NON_IPHONE)      # un iPhone non ha asset tag ne' seriale

        # Il prestito si vede dove i prestiti sono previsti.
        stanze_prestito = self.cfg.get("loan_rooms", [])
        if contenitore_iphone:
            si_presta = False
        elif stanza is not None:
            si_presta = stanza in stanze_prestito
        else:
            si_presta = bool(stanze_prestito)
        if not si_presta:
            togli("prestato_a", "prestato_il")

        if not contenitore_iphone:      # la vista degli iPhone ha un ordine suo
            campi.sort(key=lambda c: ORDINE_COLONNE.index(c)
                       if c in ORDINE_COLONNE else len(ORDINE_COLONNE))

        # L'inventario completo e' una panoramica, non una scheda: deve dire in
        # una riga che cos'e' un dispositivo, dov'e' e come sta, e da quando -
        # chi ha toccato la riga per ultimo e quando e' la domanda che ci si fa
        # guardando tutto insieme. Il modello e il numero di serie servono
        # dentro la stanza, davanti all'oggetto che si ha in mano; il prestito e
        # la spedizione lo stato in panoramica li riassume gia': "In prestito",
        # "Spedito al servizio telefonia".
        if self.view == "home":
            togli("imei", "prestato_a", "prestato_il", "restituito_da",
                  "spedito_il", "modello", "seriale")
        return campi

    def _columns(self):
        colonne = [CHECK_COLUMN]
        if self.action_column_visible():
            colonne.append(ACTION_COLUMN)
        return colonne + self._campi_visibili()

    def _larghezze_colonne(self, columns):
        """Quanto deve essere larga ogni colonna perche' il testo si legga tutto.

        Una colonna che taglia il testo costringe ad allargarla a mano ogni
        volta, e intanto quello che nasconde non si sa. Si misura il contenuto
        con il font vero e si prende il piu' largo, mai meno dell'intestazione:
        una colonna vuota deve comunque poter mostrare il proprio nome, o non si
        capisce nemmeno che cosa contiene.
        """
        # i font del tema sono tuple (famiglia, corpo): per misurare servono
        # oggetti Font veri, che si costruiscono una volta sola
        if not hasattr(self, "_font_misura"):
            self._font_misura = {
                "base": tkfont.Font(root=self, font=self.fonts["base"]),
                "bold": tkfont.Font(root=self, font=self.fonts["bold"]),
            }
        base = self._font_misura["base"]
        grassetto = self._font_misura["bold"]
        larghezze = {}
        for campo in columns:
            if campo in (CHECK_COLUMN, ACTION_COLUMN):
                larghezze[campo] = COLUMN_WIDTHS[campo]
                continue
            # l'intestazione porta anche la freccia dell'ordinamento e la
            # barretta colorata: vanno contate, o il nome viene tagliato
            titolo = intestazione(HEADERS[campo])
            minimo = grassetto.measure(titolo + "  \u25be") + SPAZIO_INTESTAZIONE
            valori = [valore_visibile(i, campo) for i in self.visible]
            # misurare tutto e' inutile: bastano i piu' lunghi, che con un font
            # proporzionale non sono sempre quelli con piu' caratteri
            candidati = sorted(set(v for v in valori if v), key=len, reverse=True)[:5]
            largo = max([base.measure(v) for v in candidati] or [0]) + SPAZIO_CELLA
            larghezze[campo] = max(minimo, largo, LARGHEZZA_MINIMA)
        return larghezze

    def _applica_larghezze(self):
        """Adatta le colonne al contenuto che si sta mostrando."""
        if getattr(self, "tree", None) is None or not self.tree.winfo_exists():
            return
        colonne = self._columns()
        try:
            for campo, largo in self._larghezze_colonne(colonne).items():
                self.tree.column(campo, width=largo, minwidth=largo)
        except tk.TclError:
            return          # tabella in ricostruzione
        self._sync_righelli()

    def _control_e_della_selezione(self):
        """True se qui Control+clic serve gia' alla selezione multipla."""
        try:
            gesti = str(self.tk.call("event", "info", "<<ToggleSelection>>"))
        except tk.TclError:
            return False
        return "Control" in gesti

    def _make_table(self, parent):
        wrap = tk.Frame(parent, bg=theme.CARD, highlightthickness=1,
                        highlightbackground=theme.BORDER)
        wrap.pack(fill="both", expand=True)
        columns = self._columns()
        # "extended" e' la selezione multipla di ttk: Ctrl per righe sparse,
        # Shift per un intervallo. La fa Tk, quindi non c'e' nostro codice sugli
        # eventi del mouse e non c'e' niente che possa rallentare l'elenco.
        tree = ttk.Treeview(wrap, columns=columns, show="headings",
                            selectmode="extended", style="Inv.Treeview")
        for field in columns:
            if field == CHECK_COLUMN:
                tree.heading(field, text=T(""))
                tree.column(field, width=COLUMN_WIDTHS[field], anchor="center",
                            stretch=False)
                continue
            if field == ACTION_COLUMN:
                tree.heading(field, text=T("Spedizione") if self.ship_column_visible()
                             else T("Prestito"))
                tree.column(field, width=COLUMN_WIDTHS[field], anchor="center",
                            stretch=False)
                continue
            arrow = ""
            if field == self.sort_field:
                arrow = "  ▾" if self.sort_reverse else "  ▴"
            tree.heading(field, text=intestazione(HEADERS[field]) + arrow,
                         image=self._segno_colonna(field),
                         command=lambda f=field: self.sort_by(f))
            tree.column(field, width=COLUMN_WIDTHS[field], anchor="w",
                        stretch=False)
        scroll = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        scroll_x = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview)

        def on_scroll(primo, ultimo):
            scroll.set(primo, ultimo)
            self._sync_row_buttons()
            self._sync_righelli()

        def on_scroll_x(primo, ultimo):
            scroll_x.set(primo, ultimo)
            self._sync_row_buttons()
            self._sync_righelli()

        tree.configure(yscrollcommand=on_scroll, xscrollcommand=on_scroll_x)
        # con la griglia le due barre restano ai bordi giusti anche quando la
        # tabella e' piu' larga della finestra: le colonne in fondo - note,
        # modificato il, modificato da - si raggiungono scorrendo a destra
        tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        # minsize e' la protezione contro lo schermo bianco: senza, quando lo
        # spazio verticale non basta - font piu' grandi, ingrandimento di
        # Windows al 125%, finestra bassa - la griglia riduce la tabella fino
        # ad altezza zero, e i dispositivi ci sono ma non si vedono.
        wrap.rowconfigure(0, weight=1, minsize=ALTEZZA_MINIMA_TABELLA)
        wrap.columnconfigure(0, weight=1)

        def rotella_orizzontale(event):
            passo = event.delta
            if passo:                      # Windows e macOS
                tree.xview_scroll(-1 if passo > 0 else 1, "units")
            self._sync_row_buttons()
            return "break"

        def scorri_di_lato(passi):
            tree.xview_scroll(passi, "units")
            self._sync_row_buttons()

        # I pulsanti 6 e 7 sono la rotella orizzontale di X11: su Windows non
        # esistono e Tk rifiuta il collegamento con un errore. Ogni scorciatoia
        # si aggiunge per conto suo, perche' una non disponibile non deve
        # impedire alla tabella di comparire.
        for evento, azione in (("<Shift-MouseWheel>", rotella_orizzontale),
                               ("<Button-6>", lambda e: scorri_di_lato(-1)),
                               ("<Button-7>", lambda e: scorri_di_lato(1))):
            try:
                tree.bind(evento, azione)
            except tk.TclError:
                continue          # scorciatoia non disponibile su questo sistema
        tree.tag_configure("odd", background=theme.ROW_ALT)
        tree.tag_configure("loan", background=theme.LOAN_BG, foreground=theme.LOAN_FG)
        tree.tag_configure("loan_alt", background=theme.LOAN_BG_ALT, foreground=theme.LOAN_FG)
        tree.tag_configure("spedito", background=theme.SHIP_ROW, foreground=theme.SHIP_FG)
        tree.tag_configure("spedito_alt", background=theme.SHIP_ROW_ALT,
                           foreground=theme.SHIP_FG)
        tree.tag_configure("iphone", background=theme.IPHONE_ROW)
        tree.tag_configure("iphone_alt", background=theme.IPHONE_ROW_ALT)
        tree.tag_configure("tablet", background=theme.TABLET_ROW)
        tree.tag_configure("tablet_alt", background=theme.TABLET_ROW_ALT)
        self._righelli = []
        self._rientro = None            # si rimisura per ogni tabella nuova
        tree.bind("<Button-1>", self._on_click)
        tree.bind("<Double-1>", self._on_double_click)
        tree.bind("<<TreeviewSelect>>", self._on_select)
        for combinazione in ("<Control-c>", "<Control-C>", "<Command-c>",
                             "<Command-C>", "<<Copy>>"):
            try:
                tree.bind(combinazione, self._copia_selezione)
            except tk.TclError:
                pass
        # Su Windows Control+clic aggiunge una riga alla selezione; su macOS
        # quel gesto e' il tasto destro. Invece di dedurlo dal sistema operativo
        # lo si chiede a Tk, che lo sa gia'.
        tasti_destri = ["<Button-3>", "<Button-2>"]
        if not self._control_e_della_selezione():
            tasti_destri.append("<Control-Button-1>")
        for tasto_destro in tasti_destri:
            tree.bind(tasto_destro, self._menu_riga)
        tree.bind("<Configure>", lambda e: (self._sync_row_buttons(),
                                            self._sync_righelli()))
        return tree

    def _copia_selezione(self, _evento=None, solo_identificativo=False):
        """Copia negli appunti la riga selezionata, pronta per Excel.

        Serve soprattutto all'incontrario: si copiano gli asset tag da qui e si
        incollano in Elimina +, o in un foglio.
        """
        tags = self.selected_tags() or list(self.tree.selection())
        if not tags:
            self._segnala(T("Non c'e' nessuna riga selezionata."))
            return "break"
        colonne = [c for c in self._columns()
                   if c not in (CHECK_COLUMN, ACTION_COLUMN)]
        righe = []
        for tag in tags:
            item = self._item_by_tag(tag)
            if item is None:
                continue
            if solo_identificativo:
                righe.append(valore_visibile(item, "asset_tag")
                             or valore_visibile(item, "imei"))
            else:
                righe.append("\t".join(valore_visibile(item, c) for c in colonne))
        if not righe:
            return "break"
        self.clipboard_clear()
        self.clipboard_append("\n".join(righe))
        self._segnala(T("Copiato negli appunti: %s.") % ", ".join(tags[:3]))
        return "break"

    def _menu_riga(self, evento):
        """Menu del tasto destro sull'elenco."""
        if self.tree is None:
            return "break"
        riga = self.tree.identify_row(evento.y)
        if riga and riga not in self.tree.selection():
            # su una riga gia' scelta il menu agisce su tutte: azzerare la
            # selezione per aprirlo vorrebbe dire perdere venti spunte
            self.tree.selection_set(riga)
            self.tree.focus(riga)
        item = self._item_by_tag(riga) if riga else None
        etichetta = (T("Copia l'IMEI") if item is not None and is_iphone(item.get("tipo"))
                     else T("Copia l'asset tag"))
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label=etichetta,
                         command=lambda: self._copia_selezione(solo_identificativo=True))
        menu.add_command(label=T("Copia la riga"), command=self._copia_selezione)
        menu.add_separator()
        menu.add_command(label=T("Modifica..."), command=self.on_edit)
        menu.add_command(label=T("Elimina"), command=self.on_delete)
        try:
            menu.tk_popup(evento.x_root, evento.y_root)
        finally:
            menu.grab_release()
        return "break"

    # --------------------------------------------- colori delle colonne

    def _segno_colonna(self, field):
        """Barretta colorata da mettere nell'intestazione della colonna.

        ttk non sa colorare una colonna: lo stile della tabella vale per tutte.
        Un'immaginetta nell'intestazione invece si puo', ed e' quanto basta per
        dare a ogni colonna un colore suo senza toccare il testo, che resta
        scuro su bianco.
        """
        if not hasattr(self, "_segni"):
            self._segni = {}
        if field not in self._segni:
            colore = theme.COLORE_COLONNA.get(field, theme.COLORE_COLONNA_ALTRO)
            immagine = tk.PhotoImage(width=4, height=14)
            immagine.put(colore, to=(0, 0, 4, 14))
            self._segni[field] = immagine       # va tenuta, o Tk la butta via
        return self._segni[field]

    def _sync_righelli(self):
        """Righe verticali colorate fra una colonna e l'altra.

        Servono a non perdere la colonna scorrendo un elenco largo. Si disegnano
        sopra la tabella, come i pulsanti delle righe, perche' ttk non ha le
        righe di griglia verticali.
        """
        try:
            self._disegna_righelli()
        except tk.TclError:
            pass          # tabella in ricostruzione: al prossimo giro ci sara'

    def _rientro_tabella(self):
        """Di quanto la prima colonna e' rientrata rispetto al bordo del widget.

        Il conto delle larghezze parte da zero, le coordinate di Tk dal bordo
        esterno della tabella: senza questo scarto - due pixel di cornice - le
        righe cadrebbero sempre un po' a sinistra del punto giusto. Si misura
        una volta, a tabella ferma, e si tiene.
        """
        if getattr(self, "_rientro", None) is None:
            figli = self.tree.get_children()
            riquadro = (self.tree.bbox(figli[0], self._columns()[0])
                        if figli else None)
            if riquadro and self.tree.xview()[0] == 0:
                self._rientro = riquadro[0]
            else:
                return 2                # elenco vuoto o gia' scorso: si riprova
        return self._rientro

    def _disegna_righelli(self):
        if getattr(self, "tree", None) is None or not self.tree.winfo_exists():
            return
        colonne = self._columns()
        larghezze = [int(self.tree.column(c, "width")) for c in colonne]
        totale = sum(larghezze) or 1
        # xview() dice a che punto del totale siamo. Il caso "non si scorre" e'
        # gia' compreso: li' il primo estremo vale zero. Trattare a parte
        # l'estremo destro - com'era prima - azzerava lo scostamento proprio
        # quando la tabella era spostata al massimo, e le righe finivano in
        # mezzo al testo.
        scostamento = self.tree.xview()[0] * totale
        altezza = self.tree.winfo_height()
        larghezza = self.tree.winfo_width()
        for riga in self._righelli:
            riga.place_forget()
        x = self._rientro_tabella() - scostamento
        for indice, campo in enumerate(colonne[:-1]):
            x += larghezze[indice]
            if not (0 < x < larghezza):
                continue
            colore = theme.COLORE_COLONNA.get(colonne[indice + 1],
                                              theme.COLORE_COLONNA_ALTRO)
            if indice >= len(self._righelli):
                self._righelli.append(tk.Frame(self.tree, width=2))
            riga = self._righelli[indice]
            riga.configure(bg=colore, width=2)
            riga.place(x=int(x) - 1, y=0, height=altezza)

    # ------------------------------------------- pulsanti veri sulle righe

    def _clear_row_buttons(self):
        for button in getattr(self, "_row_buttons", {}).values():
            button.destroy()
        self._row_buttons = {}

    def _sync_row_buttons(self):
        """Disegna un vero pulsante sulla cella Prestito delle righe visibili.

        Viene chiamata anche dalle barre di scorrimento, cioe' mentre Tk sta
        ancora disegnando: se qui scappasse un'eccezione, quella finirebbe
        dentro il disegno della tabella. Per questo non solleva mai niente.
        """
        try:
            self._disegna_pulsanti_riga()
        except tk.TclError:
            pass          # tabella in ricostruzione: al prossimo giro ci sara'

    def _disegna_pulsanti_riga(self):
        if getattr(self, "tree", None) is None or not self.tree.winfo_exists() \
                or not self.action_column_visible():
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
            return T("Conferma spedizione")
        if not self.can_lend(item):
            return ""
        return T("Registra rientro") if is_on_loan(item) else T("Presta")

    def _on_select(self, _event=None):
        """Tiene il segno di spunta allineato alla riga selezionata."""
        if self.tree is None or CHECK_COLUMN not in self._columns():
            return
        scelti = set(self.tree.selection())
        for tag in self.tree.get_children():
            atteso = CHECK_ON if tag in scelti else CHECK_OFF
            if self.tree.set(tag, CHECK_COLUMN) != atteso:
                self.tree.set(tag, CHECK_COLUMN, atteso)
        self._update_status()

    def _on_click(self, event):
        tag, column = self._cell_at(event)
        if not tag:
            return None
        if con_modificatore(event.state, self._ha_comando):
            return None          # comanda ttk: aggiunge, toglie o estende
        if column == CHECK_COLUMN:
            if self.tree.selection() == (tag,):
                self.tree.selection_remove(tag)   # era l'unica: si toglie
            else:
                self.tree.selection_set([tag])    # senza tasti si riparte da qui
            # la riga diventa l'ancora da cui Shift estendera'
            self.tree.focus(tag)
            self._on_select()
            return "break"
        if column == ACTION_COLUMN:
            return "break"          # ci pensa il pulsante disegnato sopra
        return None

    def _on_double_click(self, event):
        if con_modificatore(event.state, self._ha_comando):
            return "break"       # un doppio Ctrl+clic non apre la scheda
        tag, column = self._cell_at(event)
        if column in (ACTION_COLUMN, CHECK_COLUMN):
            return "break"
        if tag and column in InventoryStore.CAMPI_AL_VOLO:
            self.edit_testo_inline(tag, column)
            return "break"
        if column == "stato" and tag:
            self.edit_stato_inline(tag)
            return "break"
        if column == "tipo" and tag:
            self.edit_tipo_inline(tag)
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
            self._segnala(T("Lo stato degli iPhone e' sempre \"%s\".") % traduci_stato(DA_RISPEDIRE))
            return
        if is_on_loan(item):
            self._segnala(T("%s e' in prestito a %s: registra prima il rientro.")
                          % (tag, item["prestato_a"]))
            return
        stati = list(self.cfg.get("states") or [])
        if not stati:
            return
        colonne = self._columns()
        box = self.tree.bbox(tag, colonne.index("stato"))
        if not box:
            return
        var = tk.StringVar(value=traduci_stato(item.get("stato") or stati[0]))
        combo = ttk.Combobox(self.tree, textvariable=var,
                             values=[traduci_stato(v) for v in stati],
                             state="readonly", font=self.fonts["base"])
        combo.place(x=box[0], y=box[1], width=box[2], height=box[3])
        combo.focus_set()
        fatto = {"chiuso": False}

        def chiudi(salva):
            if fatto["chiuso"]:
                return
            fatto["chiuso"] = True
            scelto = stato_canonico(var.get(), stati)
            combo.destroy()
            if salva and scelto != item.get("stato"):
                self._run(lambda: self.store.set_stato(tag, scelto),
                          T("%s: %s.") % (tag, scelto))

        combo.bind("<<ComboboxSelected>>", lambda e: chiudi(True))
        combo.bind("<Escape>", lambda e: chiudi(False))
        combo.bind("<FocusOut>", lambda e: chiudi(False))
        combo.event_generate("<Button-1>")      # apre subito la tendina

    def edit_tipo_inline(self, tag):
        """Cambia il tipo con una tendina direttamente nell'elenco."""
        item = self._item_by_tag(tag)
        if item is None:
            return
        tipi = [t for t in (self.cfg.get("types") or [])
                if is_iphone(t) == is_iphone(item.get("tipo"))]
        if len(tipi) < 2:
            self._segnala(T("Non c'e' un altro tipo in cui trasformarlo."))
            return
        colonne = self._columns()
        if "tipo" not in colonne:
            return
        box = self.tree.bbox(tag, colonne.index("tipo"))
        if not box:
            return
        var = tk.StringVar(value=item.get("tipo") or tipi[0])
        # scrivibile, non solo a scelta: un tipo si incolla da un foglio o si
        # scrive per esteso quando quelli configurati non bastano
        combo = ttk.Combobox(self.tree, textvariable=var, values=tipi,
                             state="normal", font=self.fonts["base"])
        combo.select_range(0, "end")
        combo.place(x=box[0], y=box[1], width=box[2], height=box[3])
        combo.focus_set()
        fatto = {"chiuso": False}

        def chiudi(salva):
            if fatto["chiuso"]:
                return
            fatto["chiuso"] = True
            scelto = var.get()
            combo.destroy()
            if salva and scelto != item.get("tipo"):
                self._run(lambda: self.store.set_tipo(tag, scelto),
                          T("%s: %s.") % (tag, scelto))

        combo.bind("<<ComboboxSelected>>", lambda e: chiudi(True))
        combo.bind("<Return>", lambda e: chiudi(True))
        combo.bind("<Escape>", lambda e: chiudi(False))
        combo.bind("<FocusOut>", lambda e: chiudi(True))

    def _segnala(self, messaggio):
        """Avviso discreto nella barra di stato, senza aprire finestre."""
        self.var_status.set(messaggio + "     " + self.var_status.get())

    def edit_note_inline(self, tag):
        """Modifica la nota direttamente nell'elenco."""
        return self.edit_testo_inline(tag, "note")

    def edit_testo_inline(self, tag, campo):
        """Modifica un campo di testo direttamente nell'elenco, senza popup."""
        item = self._item_by_tag(tag)
        if item is None:
            return
        columns = self._columns()
        if campo not in columns:
            return
        box = self.tree.bbox(tag, columns.index(campo))
        if not box:
            return
        entry = tk.Entry(self.tree, relief="solid", borderwidth=1,
                         highlightthickness=1, highlightcolor=theme.ACCENT,
                         bg=theme.CARD, fg=theme.TEXT, font=self.fonts["base"])
        entry.insert(0, item.get(campo, ""))
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
                self._run(lambda: self.store.set_campo(tag, campo, text),
                          T("%s aggiornato su %s.")
                          % (intestazione(HEADERS[campo]), tag))

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
        self.var_room.set(TUTTE())
        self.var_type.set(TUTTI())      # il contenitore iPhone lascia il filtro impostato
        self.var_subtitle.set(T("Laptop e tablet in nostro possesso"))
        self.btn_home.state(["disabled"])
        self._render()

    def show_room(self, room):
        self.view = "room"
        self.var_room.set(room)
        self.var_type.set(TUTTI())
        self.var_subtitle.set(T("Inventario di %s") % room)
        self.btn_home.state(["!disabled"])
        self._render()

    def show_iphones(self):
        """Vista di comodo con i soli iPhone; restano comunque nella loro stanza."""
        tipo = self.iphone_type()
        if not tipo:
            return
        self.view = "type"
        self.var_room.set(TUTTE())
        self.var_type.set(tipo)
        self.var_subtitle.set(T("Telefoni in nostro possesso - registrati in %s")
                              % self.iphone_room())
        self.btn_home.state(["!disabled"])
        self._render()

    def _render(self):
        self._clear_body()
        if self.view == "home":
            self._render_cards()
            header = ttk.Frame(self.body)
            header.pack(fill="x", pady=(18, 8))
            ttk.Label(header, text=T("Inventario completo"),
                      style="Section.TLabel").pack(side="left")
            self.var_section_count = tk.StringVar()
            ttk.Label(header, textvariable=self.var_section_count,
                      style="Muted.TLabel").pack(side="left", padx=(10, 0))
            ttk.Button(header, text=T("Scarica il modello di importazione"),
                       command=self.on_template).pack(side="right")
            ttk.Button(header, text=T("Controllo generale duplicati"),
                       style="Rosso.TButton",
                       command=self.on_duplicati).pack(side="right", padx=(0, 6))
            self.btn_cestino = ttk.Button(header, style="Cestino.TButton",
                                          text=self._etichetta_cestino(),
                                          command=self.on_cestino)
            self.btn_cestino.pack(side="right", padx=(0, 10))
        else:
            header = ttk.Frame(self.body)
            header.pack(fill="x", pady=(10, 8))
            titolo = self.var_type.get() if self.view == "type" else self.var_room.get()
            ttk.Label(header, text=titolo,
                      style="Section.TLabel").pack(side="left")
            if self.view == "room":
                ttk.Button(header, text=T("Esporta questa stanza in xls"),
                           style="Verde.TButton",
                           command=self.on_export_room).pack(side="right")
                ttk.Button(header, text=T("Importa i dati di questa stanza"),
                           style="Arancio.TButton",
                           command=self.on_import_room).pack(side="right", padx=(0, 6))
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
                "%d %s" % (n, k.lower()) for k, n in sorted(counts.items())) or T("nessun dispositivo")
            on_loan = sum(1 for i in subset if is_on_loan(i))
            note = ""
            if name in self.cfg.get("loan_rooms", []):
                note = (T("%d in prestito") % on_loan) if on_loan else T("nessun prestito in corso")
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
                            T("tutti i telefoni, ovunque siano registrati"),
                            theme.IPHONE_COLOR, self.show_iphones,
                            note=T("anche in %s") % self.iphone_room())
            card.grid(row=0, column=column, sticky="nsew", padx=(14, 0))
            strip.columnconfigure(column, weight=1, uniform="cards")

    # ------------------------------------------------------------ dati

    def _initial_load(self):
        try:
            self.store.create_if_missing()
            self.store.load()
        except InventoryError as exc:
            messagebox.showerror(T("Errore"), str(exc), parent=self)
            self.destroy()
            return
        self._sync_filter_values()
        self.show_home()
        self._avvisa_se_foglio_da_importare()

    def _avvisa_se_foglio_da_importare(self):
        """Dice all'utente se il file aperto e' un foglio da importare.

        Chi ci finisce dentro vede un inventario che sembra caricato ma con le
        stanze vuote e i nomi delle stanze in elenco come dispositivi, e non ha
        nessun modo di capire perche'. Vale la pena dirlo subito.
        """
        separatori = righe_separatore(self.store.items, self.cfg.get("rooms"))
        if not separatori:
            return
        messagebox.showwarning(
            T("Questo file non e' un inventario"),
            T("Il file aperto e' un foglio da IMPORTARE, non un inventario:\n%s\n\n"
              "Contiene le righe che dividono i dispositivi per stanza (%s),\n"
              "che qui compaiono in elenco come se fossero dispositivi. Per\n"
              "questo le stanze restano vuote.\n\n"
              "Come sistemare:\n"
              "1. chiudi il programma;\n"
              "2. cancella il file inventario_percorso.json accanto al programma,\n"
              "   se c'e': e' li' che resta memorizzata la scelta sbagliata;\n"
              "3. riapri: l'inventario vuoto viene creato da solo in Produzione;\n"
              "4. carica questo foglio con  Importa xls...  da dentro il programma.")
            % (self.store.path, ", ".join(separatori[:3])), parent=self)

    def _sync_filter_values(self):
        rooms = list(self.cfg["rooms"])
        for item in self.store.items:
            room = item.get("stanza", "")
            if room and room not in rooms:
                rooms.append(room)
        self.combo_room["values"] = [TUTTE()] + rooms + [NO_ROOM()]
        if self.var_room.get() not in self.combo_room["values"]:
            self.var_room.set(TUTTE())
        tipi = list(self.cfg["types"])
        for item in self.store.items:
            tipo = clean(item.get("tipo"))
            if tipo and tipo not in tipi:
                tipi.append(tipo)      # scritto a mano: si filtra lo stesso
        self.combo_type["values"] = [TUTTI()] + tipi
        if self.var_type.get() not in self.combo_type["values"]:
            self.var_type.set(TUTTI())

    def _on_room_filter(self, _event=None):
        room = self.var_room.get()
        if room in (TUTTE(), NO_ROOM()):
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
            if room == NO_ROOM():
                if item.get("stanza"):
                    continue
            elif room != TUTTE() and item.get("stanza") != room:
                continue
            if tipo != TUTTI() and item.get("tipo") != tipo:
                continue
            if text and not any(
                text in str(item.get(f, "")).lower()
                for f in ("asset_tag", "modello", "seriale", "imei",
                          "restituito_da", "note", "tipo", "stanza",
                          "stato", "prestato_a")
            ):
                continue
            result.append(item)
        result.sort(key=lambda it: chiave_ordinamento(it, self.sort_field),
                    reverse=self.sort_reverse)
        return result

    def refresh_table(self, keep_selection=True):
        self.visible = self.filtered_items()
        if self.tree is None:
            self._update_status()
            return
        selected = self.selected_tags() if keep_selection else []
        ancora = self.tree.focus() if keep_selection else ""
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
                elif field == "stato":
                    values.append(traduci_stato(valore_visibile(item, field)))
                else:
                    values.append(valore_visibile(item, field))
            tag = self.row_tag(item, i % 2)
            self.tree.insert("", "end", iid=item["asset_tag"], values=values,
                             tags=(tag,) if tag else ())
        restore = [t for t in selected if self.tree.exists(t)]
        if restore:
            self.tree.selection_set(restore)
            if ancora in restore:
                self.tree.focus(ancora)
            self.tree.see(restore[0])
        self._update_status()
        self._applica_larghezze()
        self._sync_row_buttons()

    def _update_status(self):
        total = len(self.store.items)
        parts = []
        for room in self.cfg["rooms"]:
            parts.append("%s: %d" % (room, sum(
                1 for i in self.store.items if i.get("stanza") == room)))
        others = sum(1 for i in self.store.items if i.get("stanza") not in self.cfg["rooms"])
        if others:
            parts.append(T("altre/nessuna: %d") % others)
        self.var_head_count.set(T("%d dispositivi     %s") % (total, "     ".join(parts)))
        if getattr(self, "var_section_count", None) is not None:
            label = T("%d dispositivi") % len(self.visible)
            if len(self.visible) != total:
                label += T(" di %d") % total
            scelti = len(self.tree.selection()) if self.tree is not None else 0
            if scelti > 1:
                label += T("  -  %d selezionati") % scelti
            self.var_section_count.set(label)
        shown = "" if len(self.visible) == total else T("  |  visualizzati: %d") % len(self.visible)
        self.var_status.set(
            T("%d dispositivi  (%s)%s     File: %s")
            % (total, ", ".join(parts), shown, self.store.path)
        )

    def sort_by(self, field):
        if self.sort_field == field:
            self.sort_reverse = not self.sort_reverse
        else:
            # le date partono dalla piu' recente, il testo dalla A
            self.sort_field = field
            self.sort_reverse = field in CAMPI_DATA
        if self.tree is not None:
            # solo le colonne di questa vista: le altre nella tabella non ci sono
            for name in self._campi_visibili():
                arrow = ""
                if name == self.sort_field:
                    arrow = "  ▾" if self.sort_reverse else "  ▴"
                self.tree.heading(name, text=intestazione(HEADERS[name]) + arrow)
        self.refresh_table()

    def reset_filters(self):
        self.var_search.set("")
        if self.view == "room":
            self.var_type.set(TUTTI())
        elif self.view != "type":          # nella vista iPhone il tipo e' la vista
            self.var_type.set(TUTTI())
            self.var_room.set(TUTTE())
        self.refresh_table()

    def selected_tags(self):
        """Gli asset tag selezionati, nell'ordine in cui si vedono nell'elenco.

        Tk li restituisce nell'ordine in cui la selezione e' stata costruita:
        chi sceglie dal basso verso l'alto se li ritroverebbe capovolti nel
        riepilogo e negli appunti, dove ci si aspetta l'ordine dell'elenco.
        """
        if self.tree is None:
            return []
        scelti = set(self.tree.selection())
        return [t for t in self.tree.get_children() if t in scelti]

    def selected_items(self):
        tags = set(self.selected_tags())
        return [i for i in self.store.items if i["asset_tag"] in tags]

    def _items_by_tag(self, tags):
        """I dispositivi corrispondenti ai codici, in un giro solo.

        Cercarli uno per uno con _item_by_tag costerebbe un giro dell'inventario
        per ogni riga selezionata: con venti righe su trecento dispositivi si
        sente.
        """
        voluti = set(tags)
        trovati = dict((i["asset_tag"], i) for i in self.store.items
                       if i["asset_tag"] in voluti)
        return [trovati[t] for t in tags if t in trovati]

    def _reload(self, message=None):
        try:
            self.store.load()
        except InventoryError as exc:
            messagebox.showerror(T("Errore"), str(exc), parent=self)
            return
        self.cfg = config.load_shared_config(self.store.path)
        self.store.iphone_room = self.cfg.get("iphone_room")
        self.store.stati = list(self.cfg.get("states") or [])
        self._sync_filter_values()
        scelti, ancora = self.selected_tags(), (self.tree.focus() if self.tree else "")
        self._render()
        if scelti and self.tree is not None:
            # l'elenco si ricarica da solo ogni quindici secondi: senza questo,
            # una selezione costruita riga per riga sparirebbe da sola
            restano = [t for t in scelti if self.tree.exists(t)]
            if restano:
                self.tree.selection_set(restano)
                if ancora in restano:
                    self.tree.focus(ancora)
        if message:
            self.var_status.set(message + "     " + self.var_status.get())

    def _auto_refresh(self):
        try:
            if self.store.changed_on_disk():
                self._reload(T("Inventario aggiornato da un altro utente."))
        except Exception:
            pass
        self.after(REFRESH_MS, self._auto_refresh)

    def _run(self, action, success=None):
        """Esegue un'operazione sull'archivio gestendo gli errori."""
        try:
            result = action()
        except InventoryError as exc:
            messagebox.showerror(T("Operazione non riuscita"), str(exc), parent=self)
            self._reload()
            return None
        self._sync_filter_values()
        if self.view == "home":
            self._render()          # i conteggi delle schede cambiano
        else:
            self.refresh_table()
        self._aggiorna_cestino()
        if success:
            self.var_status.set(success + "     " + self.var_status.get())
        self._forse_ricorda_copia()
        return result

    def _forse_ricorda_copia(self):
        """Ogni tanto ricorda di portarsi via una copia dell'inventario.

        Le copie automatiche stanno sulla share, accanto ai dati: coprono
        l'errore umano ma non la cartella di rete che sparisce. Quella copia la
        deve volere qualcuno, e chi sta lavorando non ci pensa: glielo si
        ricorda dopo un po' di modifiche, non a ogni singola riga.
        """
        if getattr(self, "_copia_chiesta", False):
            return                      # gia' in corso: non si annida
        fatte = self.store.modifiche - getattr(self, "_modifiche_alla_copia", 0)
        if fatte < MODIFICHE_PER_PROMEMORIA:
            return
        self._copia_chiesta = True
        try:
            salva = messagebox.askyesno(
                T("Conviene una copia"),
                T("Hai modificato %d dispositivi dall'ultima copia locale.\n\n"
                  "Le copie automatiche stanno sulla cartella di rete, accanto\n"
                  "ai dati: se sparisce quella, spariscono anche loro.\n\n"
                  "Vuoi salvare adesso una copia dell'inventario sul tuo\n"
                  "computer? Ci vogliono cinque secondi.") % fatte, parent=self)
            self._modifiche_alla_copia = self.store.modifiche
            if salva:
                self.on_copia_locale()
        finally:
            self._copia_chiesta = False

    # ------------------------------------------------------------ azioni

    def on_refresh(self):
        self._reload(T("Elenco ricaricato."))

    def on_add(self):
        """Prima cosa si aggiunge, poi come: a mano o con il lettore di codici."""
        tipi = self.cfg.get("types") or []
        if not tipi:
            messagebox.showinfo(T("Aggiungi"),
                                T("Non ci sono tipi di dispositivo configurati."), parent=self)
            return
        tipo = TypeChoiceDialog(self, tipi, self.tipo_predefinito()).show()
        if not tipo:
            return
        iphone = is_iphone(tipo)
        modo = AddChoiceDialog(self, iphone=iphone).show()
        if modo == "manuale":
            self.on_new(tipo)
        elif modo == "barcode":
            if iphone:
                self.on_new_barcode_iphone(tipo)
            else:
                self.on_new_barcode(tipo)

    def tipo_predefinito(self):
        """Nel contenitore iPhone si aggiunge quasi sempre un iPhone."""
        tipi = self.cfg.get("types") or []
        if self.view == "type":
            return self.iphone_type() or (tipi[0] if tipi else "")
        return next((t for t in tipi if not is_iphone(t)), tipi[0] if tipi else "")

    def on_new_barcode_iphone(self, tipo=None):
        """Un solo codice da leggere: l'IMEI. Il resto si scrive nella scheda."""
        imei = ScanDialog(self, T("IMEI"), T("l'IMEI del telefono"), 1, 1).show()
        if not imei:
            return
        preset = new_item(tipo=tipo or self.iphone_type(), imei=imei,
                          stanza=self.iphone_room())
        item = ItemDialog(self, self.cfg["rooms"], self.cfg["types"], preset,
                          iphone_room=self.iphone_room(),
                          stati=self.cfg.get("states")).show()
        if item:
            self._run(lambda: self.store.add(item), T("Aggiunto %s.") % item["asset_tag"])

    def stanza_predefinita(self):
        if self.view == "room":
            return self.var_room.get()
        if self.view == "type":
            return self.iphone_room()
        rooms = self.cfg.get("rooms") or []
        return rooms[0] if rooms else ""

    def on_new_barcode(self, tipo=None):
        """Asset tag e seriale con il lettore, poi il modello a mano."""
        tag = ScanDialog(self, T("Asset tag"), T("l'asset tag"), 1, 3).show()
        if not tag:
            return
        seriale = ScanDialog(self, T("Numero di serie"), T("il numero di serie"), 2, 3).show()
        if not seriale:
            return
        modello = ScanDialog(self, T("Modello"), T("il modello del dispositivo"), 3, 3,
                             manuale=True).show()
        if not modello:
            return
        preset = new_item(
            asset_tag=tag, seriale=seriale, modello=modello,
            tipo=tipo or self.tipo_predefinito(),
            stanza=self.stanza_predefinita())
        item = ItemDialog(self, self.cfg["rooms"], self.cfg["types"], preset,
                          iphone_room=self.iphone_room(),
                          stati=self.cfg.get("states")).show()
        if item:
            self._run(lambda: self.store.add(item), T("Aggiunto %s.") % item["asset_tag"])

    def on_new(self, tipo=None):
        rooms = self.cfg["rooms"]
        tipo = tipo or self.tipo_predefinito()
        stanza = self.iphone_room() if is_iphone(tipo) else self.stanza_predefinita()
        preset = new_item(tipo=tipo, stanza=stanza)
        item = ItemDialog(self, rooms, self.cfg["types"], preset,
                          iphone_room=self.iphone_room(),
                          stati=self.cfg.get("states")).show()
        if item:
            self._run(lambda: self.store.add(item), T("Aggiunto %s.") % item["asset_tag"])

    def on_edit(self):
        items = self.selected_items()
        if len(items) != 1:
            messagebox.showinfo(T("Modifica"),
                                T("Spunta il dispositivo da modificare, oppure fai doppio clic sulla riga."),
                                parent=self)
            return
        old = items[0]
        edited = ItemDialog(self, self.cfg["rooms"], self.cfg["types"], old,
                            iphone_room=self.iphone_room(),
                            stati=self.cfg.get("states")).show()
        if edited:
            self._run(lambda: self.store.update(old["asset_tag"], edited),
                      T("Salvato %s.") % edited["asset_tag"])

    def on_delete(self):
        """Elimina i dispositivi selezionati, dopo averli mostrati uno per uno."""
        tags = self.selected_tags()
        if not tags:
            messagebox.showinfo(T("Elimina"),
                                T("Spunta i dispositivi da eliminare."), parent=self)
            return

        da_eliminare, _non_trovati, bloccati = self.store.anteprima_eliminazione(tags)
        # l'anteprima ha riletto il file: la tabella va riallineata a quello che
        # il riepilogo sta per raccontare
        self.refresh_table()

        # Un solo dispositivo bloccato: si dice perche' e cosa fare per
        # sbloccarlo, che e' piu' utile di un elenco con dentro una riga sola.
        if not da_eliminare and len(bloccati) == 1:
            item = bloccati[0][0]
            libero, sblocco = puo_essere_eliminato(item)
            if is_on_loan(item):
                messagebox.showwarning(
                    T("Eliminazione non consentita"),
                    T("%s - %s\n\n"
                    "Il dispositivo e' in prestito a %s dal %s.\n\n"
                    "Un dispositivo in prestito non si elimina: l'inventario e'\n"
                    "l'unica traccia di chi ce l'ha. Registra prima il rientro\n"
                    "con il pulsante Registra rientro, poi potrai eliminarlo.")
                    % (valore_visibile(item, "asset_tag") or valore_visibile(item, "imei"),
                       item.get("modello", ""), item.get("prestato_a", ""),
                       item.get("prestato_il", "")),
                    parent=self)
            elif sblocco is None:
                messagebox.showwarning(
                    T("Eliminazione non consentita"),
                    T("%s - %s\n\n"
                    "Questo iPhone non e' ancora stato rispedito al servizio\n"
                    "telefonia, quindi non puo' essere eliminato dall'inventario.\n\n"
                    "Registra prima la spedizione con il pulsante Conferma\n"
                    "spedizione, nel contenitore Iphone. Da quel momento restera'\n"
                    "consultabile per %d mesi, e poi potra' essere eliminato.")
                    % (valore_visibile(item, "asset_tag") or valore_visibile(item, "imei"),
                       item.get("modello", ""), MESI_CONSERVAZIONE),
                    parent=self)
            else:
                messagebox.showwarning(
                    T("Eliminazione non consentita"),
                    T("%s - %s\n\n"
                    "Il dispositivo e' stato rispedito al servizio telefonia il %s e\n"
                    "va conservato in inventario per consultazione.\n\n"
                    "Potrai eliminarlo a partire dal %s.")
                    % (valore_visibile(item, "asset_tag") or valore_visibile(item, "imei"),
                       item.get("modello", ""), item["spedito_il"],
                       sblocco.strftime("%d/%m/%Y")),
                    parent=self)
            return

        restano = len(self.store.items) - len(da_eliminare)
        righe = riepilogo_eliminazione(da_eliminare, bloccati, restano)
        if not da_eliminare:
            messagebox.showwarning(T("Elimina"), "\n".join(righe), parent=self)
            return

        avvertenza = T("Stai per eliminare %d dispositivi dall'inventario di tutti.\n"
                       "Una copia del file dati viene salvata prima di procedere.")
        if len(da_eliminare) == 1:
            avvertenza = T("Il dispositivo sparisce dall'inventario di tutti.\n"
                           "Una copia del file dati viene salvata prima di procedere.")
        if not ConfermaOperazioneDialog(
            self, T("Conferma eliminazione"),
            T("Eliminare %d dispositivi?") % len(da_eliminare) if len(da_eliminare) > 1
            else T("Eliminare questo dispositivo?"),
            righe, T("Elimina"), "Rosso.TButton", avvertenza
        ).show():
            return

        try:
            copia = self.store.copia_di_sicurezza()
        except InventoryError as exc:
            messagebox.showerror(T("Eliminazione annullata"), str(exc), parent=self)
            return

        # solo quelli davvero eliminabili: store.delete e' tutto-o-niente, e un
        # iPhone protetto in mezzo bloccherebbe anche gli altri
        eliminabili = [i["asset_tag"] for i in da_eliminare]
        fatto = self._run(lambda: self.store.delete(eliminabili))
        if fatto is None:
            return
        messagebox.showinfo(
            T("Eliminazione completata"),
            T("Eliminati %d dispositivi.\n\nIn inventario ne restano %d.\n\n"
              "Copia di sicurezza del file precedente:\n%s")
            % (len(eliminabili), len(self.store.items), copia), parent=self)

    def on_delete_many(self):
        """Eliminazione in blocco, incollando i codici da un foglio."""
        tags = EliminaPlusDialog(self, self.store).show()
        if not tags:
            return
        copia = None
        try:
            copia = self.store.copia_di_sicurezza()
        except InventoryError as exc:
            messagebox.showerror(T("Eliminazione annullata"), str(exc), parent=self)
            return
        if not self._run(lambda: self.store.delete(tags),
                         T("Eliminati %d dispositivi.") % len(tags)):
            return
        self.refresh_table()
        self.show_home() if self.view == "home" else self._render()
        messagebox.showinfo(
            T("Eliminazione completata"),
            T("%d dispositivi eliminati.\n\nIn inventario ne restano %d.\n\n"
              "Copia di sicurezza del file precedente:\n%s")
            % (len(tags), len(self.store.items), copia), parent=self)

    def on_move(self):
        """Sposta in un'altra stanza i dispositivi selezionati."""
        tags = self.selected_tags()
        if not tags:
            messagebox.showinfo(T("Sposta"),
                                T("Spunta i dispositivi da spostare."), parent=self)
            return
        items = self._items_by_tag(tags)
        telefoni = [i for i in items if is_iphone(i.get("tipo"))]
        prestati = [i for i in items
                    if not is_iphone(i.get("tipo")) and is_on_loan(i)]
        spostabili = [i for i in items
                      if not is_iphone(i.get("tipo")) and not is_on_loan(i)]
        if not spostabili:
            if prestati:
                messagebox.showinfo(
                    T("Sposta"),
                    T("%d dispositivi sono in prestito e non si spostano.\n\n"
                      "Registra prima il rientro, poi si potranno spostare.")
                    % len(prestati) if len(prestati) > 1 else
                    T("%s e' in prestito a %s: registra prima il rientro,\n"
                      "poi si potra' spostare.")
                    % (valore_visibile(prestati[0], "asset_tag"),
                       prestati[0].get("prestato_a", "")), parent=self)
                return
            messagebox.showinfo(
                T("Sposta"),
                T("Gli iPhone restano sempre in %s e non possono essere spostati.")
                % self.iphone_room(), parent=self)
            return

        domanda = (T("Sposta %s in:") % tags[0] if len(spostabili) == 1
                   else T("Sposta %d dispositivi in:") % len(spostabili))
        room = self._ask_room(domanda)
        if not room:
            return

        in_partenza = sum(1 for i in spostabili if i.get("stanza") != room)
        conteggi = []
        for stanza in self.cfg["rooms"]:
            prima = sum(1 for i in self.store.items if i.get("stanza") == stanza)
            escono = sum(1 for i in spostabili
                         if i.get("stanza") == stanza and stanza != room)
            dopo = prima - escono + (in_partenza if stanza == room else 0)
            conteggi.append((stanza, (prima, dopo)))
        righe = riepilogo_spostamento(spostabili, telefoni, prestati, room,
                                      self.iphone_room(), conteggi)
        if not in_partenza:
            messagebox.showinfo(T("Sposta"), "\n".join(righe), parent=self)
            return
        if not ConfermaOperazioneDialog(
            self, T("Conferma spostamento"),
            T("Spostare %d dispositivi in %s?") % (in_partenza, room)
            if in_partenza > 1 else T("Spostare questo dispositivo in %s?") % room,
            righe, T("Sposta")
        ).show():
            return

        esito = self._run(lambda: self.store.move_to_room(tags, room))
        if esito:
            spostati, fermi_telefoni, fermi_prestiti = esito
            messaggio = T("Spostati %d dispositivi in %s.") % (spostati, room) if spostati \
                else T("Nessuno spostamento.")
            if fermi_telefoni:
                messaggio += T("  %d iPhone lasciati in %s.") % (fermi_telefoni,
                                                                 self.iphone_room())
            if fermi_prestiti:
                messaggio += T("  %d in prestito lasciati dove sono.") % fermi_prestiti
            self.var_status.set(messaggio + "     " + self.var_status.get())

    def on_lend(self, tag=None):
        """Registra il prestito del dispositivo a una persona."""
        tag = tag or self._single_selection(T("Presta"))
        if not tag:
            return
        item = self._item_by_tag(tag)
        if item is None:
            return
        if not self.can_lend(item):
            messagebox.showinfo(
                T("Prestito"),
                T("La gestione dei prestiti e' attiva solo per: %s.")
                % ", ".join(self.cfg.get("loan_rooms") or ["nessuna stanza"]), parent=self)
            return
        person = self._ask_person(item)
        if not person:
            return
        when = self._run(lambda: self.store.lend(tag, person))
        if when:
            self.var_status.set(T("%s prestato a %s il %s.     %s")
                                % (tag, person, when, self.var_status.get()))

    def on_ship(self, tag=None):
        """Registra la spedizione dell'iPhone al servizio telefonia."""
        tag = tag or self._single_selection(T("Spedizione"))
        if not tag:
            return
        item = self._item_by_tag(tag)
        if item is None:
            return
        if not is_iphone(item.get("tipo")):
            messagebox.showinfo(T("Spedizione"),
                                T("La spedizione al servizio telefonia riguarda solo gli iPhone."),
                                parent=self)
            return
        if is_shipped(item):
            messagebox.showinfo(T("Spedizione"),
                                T("%s risulta gia' spedito il %s.") % (tag, item["spedito_il"]),
                                parent=self)
            return
        if not messagebox.askyesno(
            T("Conferma spedizione"),
            T("%s - %s\n\nRegistrare la spedizione al servizio telefonia?\n\n"
            "Data e ora vengono registrate adesso. Il dispositivo resta in\n"
            "inventario per consultazione per %d mesi, poi potra' essere eliminato.")
            % (tag, item.get("modello", ""), MESI_CONSERVAZIONE), parent=self
        ):
            return
        testo = self._run(lambda: self.store.ship(tag))
        if testo:
            messagebox.showinfo(T("Spedizione registrata"), testo, parent=self)

    def on_give_back(self, tag=None):
        """Chiude il prestito: il dispositivo torna disponibile."""
        tag = tag or self._single_selection(T("Rientro"))
        if not tag:
            return
        item = self._item_by_tag(tag)
        if item is None or not is_on_loan(item):
            messagebox.showinfo(T("Rientro"), T("Il dispositivo non risulta in prestito."), parent=self)
            return
        if not messagebox.askyesno(
            T("Registra rientro"),
            T("%s - %s\n\nIn prestito a %s dal %s.\nRegistrare il rientro?")
            % (tag, item.get("modello", ""), item["prestato_a"], item["prestato_il"]),
            parent=self
        ):
            return
        person = self._run(lambda: self.store.give_back(tag))
        if person:
            self.var_status.set(T("%s rientrato da %s.     %s")
                                % (tag, person, self.var_status.get()))

    def _single_selection(self, action):
        tags = self.selected_tags()
        if len(tags) != 1:
            messagebox.showinfo(action, T("Spunta il dispositivo su cui vuoi agire."),
                                parent=self)
            return None
        return tags[0]

    def _ask_person(self, item):
        dialog = _Modal(self, T("Presta dispositivo"))
        body = ttk.Frame(dialog, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=T("%s - %s") % (item["asset_tag"], item.get("modello", "")),
                  style="Section.TLabel").pack(anchor="w")
        ttk.Label(body, text=T("Data e ora del prestito vengono registrate in automatico."),
                  style="Muted.TLabel").pack(anchor="w", pady=(2, 12))
        ttk.Label(body, text=T("Nome della persona")).pack(anchor="w")
        var = tk.StringVar()
        entry = ttk.Entry(body, textvariable=var, width=34)
        entry.pack(fill="x", pady=(4, 0))
        buttons = ttk.Frame(body)
        buttons.pack(anchor="e", pady=(16, 0))

        def ok():
            if not var.get().strip():
                messagebox.showwarning(T("Campo mancante"),
                                       T("Indica il nome della persona."), parent=dialog)
                return
            dialog.result = var.get().strip()
            dialog.destroy()

        ttk.Button(buttons, text=T("Annulla"), command=dialog._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text=T("Registra prestito"), style="Primary.TButton",
                   command=ok).pack(side="right")
        dialog.bind("<Return>", lambda e: ok())
        entry.focus_set()
        return dialog.show()

    def _ask_room(self, prompt):
        dialog = _Modal(self, T("Scegli stanza"))
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

        ttk.Button(buttons, text=T("Annulla"), command=dialog._cancel).pack(side="right", padx=6)
        ttk.Button(buttons, text=T("Sposta"), style="Primary.TButton",
                   command=ok).pack(side="right")
        return dialog.show()

    def _etichetta_cestino(self):
        """ELIMINATI DI RECENTE (n): il numero e' la ragione per aprirlo."""
        try:
            quanti = len(self.store.eliminati())
        except Exception:
            quanti = 0
        return T("ELIMINATI DI RECENTE (%d)") % quanti

    def _aggiorna_cestino(self):
        pulsante = getattr(self, "btn_cestino", None)
        if pulsante is not None and pulsante.winfo_exists():
            pulsante.configure(text=self._etichetta_cestino())

    def on_cestino(self):
        """Gli eliminati di recente, da cui si ripesca quello tolto per sbaglio."""
        ripristinati = CestinoDialog(self, self.store, self.cfg["rooms"]).show()
        self._aggiorna_cestino()
        if ripristinati:
            self.store.load()
            self._sync_filter_values()
            self.show_home()
            self.var_status.set(
                T("Ripristinati %d dispositivi dagli eliminati di recente.")
                % ripristinati + "     " + self.var_status.get())

    def on_duplicati(self):
        """Cerca in tutto l'inventario i dispositivi registrati due volte.

        Il file dati e' un .xlsx che si puo' aprire e correggere a mano: e' da
        li' che i doppioni entrano. Il programma non ne crea, ma deve saperli
        togliere.
        """
        gruppi, seriali = self.store.trova_duplicati()
        if not gruppi and not seriali:
            messagebox.showinfo(
                T("Nessun duplicato"),
                T("Ho confrontato l'ASSET TAG di %d dispositivi: ognuno compare una\n"
                  "volta sola. E' quello l'identificativo, e su quello si controlla.\n\n"
                  "Nemmeno un numero di serie risulta ripetuto, ma quella e' solo una\n"
                  "verifica in piu': il seriale non identifica il dispositivo.")
                % len(self.store.items), parent=self)
            return

        righe = [T("Confrontato l'asset tag di %d dispositivi.")
                 % len(self.store.items), ""]
        if gruppi:
            quanti = sum(len(e) - 1 for _t, e in gruppi)
            righe.append(T("DOPPIONI TROVATI: %d asset tag registrati piu' volte, "
                           "%d righe in piu'.") % (len(gruppi), quanti))
            for tag, elenco in gruppi[:12]:
                tenuto = self.store._piu_recente(elenco)
                righe.append("")
                righe.append("  %s  -  %d registrazioni" % (tag, len(elenco)))
                for item in elenco:
                    segno = T("TIENE") if item is tenuto else T("elimina")
                    righe.append("    [%s] %s  %s  %s"
                                 % (segno, item.get("stanza") or "-",
                                    (item.get("modello") or "")[:28],
                                    item.get("modificato_il") or ""))
            if len(gruppi) > 12:
                righe.append("")
                righe.append(T("  ... e altri %d asset tag") % (len(gruppi) - 12))
            righe.append("")
            righe.append(T("Si tiene la registrazione modificata piu' di recente."))
        if seriali:
            righe.append("")
            righe.append(T("Nota a margine - numeri di serie ripetuti su asset tag\n"
                           "diversi: %d") % len(seriali))
            for seriale, elenco in seriali[:8]:
                righe.append("    %s  ->  %s" % (seriale, ", ".join(
                    valore_visibile(i, "asset_tag") for i in elenco)))
            righe.append(T("Il seriale non identifica il dispositivo, quindi questi\n"
                           "NON sono duplicati e non vengono toccati. Di solito pero'\n"
                           "sono un errore di battitura, e conviene guardarli."))

        if not gruppi:
            messagebox.showwarning(T("Controllo duplicati"), "\n".join(righe),
                                   parent=self)
            return
        if not messagebox.askyesno(
            T("Controllo duplicati"),
            "\n".join(righe) + T("\n\nProcedo a eliminare le righe in piu'?\n"
                                 "Una copia di sicurezza viene salvata prima."),
            parent=self
        ):
            return

        rapporto = self._run(lambda: self.store.rimuovi_duplicati(),
                             T("Duplicati rimossi."))
        if not rapporto:
            return
        self._render()
        fine = [T("Eliminate %d righe in piu'.") % len(rapporto["eliminati"]), ""]
        for item in rapporto["eliminati"][:15]:
            fine.append("    %s  %s  %s" % (valore_visibile(item, "asset_tag"),
                                            item.get("stanza") or "-",
                                            (item.get("modello") or "")[:24]))
        if len(rapporto["eliminati"]) > 15:
            fine.append(T("    ... e altre %d") % (len(rapporto["eliminati"]) - 15))
        if rapporto["protetti"]:
            fine.append("")
            fine.append(T("NON eliminate, protette: %d") % len(rapporto["protetti"]))
            for item, motivo in rapporto["protetti"]:
                fine.append("    %s  -  %s"
                            % (valore_visibile(item, "asset_tag"), motivo))
        fine.append("")
        fine.append(T("Dispositivi prima: %d") % rapporto["prima"])
        fine.append(T("Dispositivi adesso: %d") % rapporto["dopo"])
        if rapporto["copia"]:
            fine.append("")
            fine.append(T("Copia di sicurezza del file precedente:"))
            fine.append(rapporto["copia"])
        messagebox.showinfo(T("Duplicati rimossi"), "\n".join(fine), parent=self)

    def on_template(self):
        """Genera il modello vuoto da compilare e reimportare."""
        percorso = filedialog.asksaveasfilename(
            parent=self, title=T("Salva il modello di inventario"),
            defaultextension=".xlsx", initialfile=T("Modello_inventario.xlsx"),
            filetypes=[(T("File Excel"), "*.xlsx")])
        if not percorso:
            return
        try:
            excel_io.build_template(percorso, self.cfg.get("rooms", []),
                                    self.cfg.get("states"), lingua=lang.corrente())
        except InventoryError as exc:
            messagebox.showerror(T("Modello non creato"), str(exc), parent=self)
            return
        if messagebox.askyesno(
            T("Modello creato"),
            T("%s\n\nCompila il foglio \"Inventario\" e reimportalo con\n"
            "Importa xls...  Le righe con il nome della stanza dividono\n"
            "l'elenco: quello che scrivi sotto finisce in quella stanza.\n\n"
            "Aprirlo ora?") % percorso, parent=self
        ):
            excel_io.open_file(percorso)

    def on_import(self):
        """Prima cosa si carica e come, poi il file, poi la conferma."""
        opzioni = ImportOptionsDialog(self, self.cfg.get("rooms") or []).show()
        if not opzioni:
            return
        path = filedialog.askopenfilename(
            parent=self, title=T("Seleziona il file da importare"),
            filetypes=[(T("File Excel"), "*.xlsx *.xlsm"), (T("Tutti i file"), "*.*")])
        if not path:
            return
        try:
            items, esito = rows_from_workbook(path, self.cfg.get("rooms"))
        except InventoryError as exc:
            messagebox.showerror(T("Importazione non riuscita"), str(exc), parent=self)
            return
        if not items:
            messagebox.showwarning(T("Importazione"), T("Nessuna riga valida trovata nel file."),
                                   parent=self)
            return
        stanza = opzioni["stanza"]
        scartati = 0
        if stanza is not None:
            items, scartati, regola = seleziona_per_stanza(items, esito, stanza)
            if regola == "mancante":
                self._avviso_stanza_mancante(stanza, esito)
                return
            if not items:
                messagebox.showwarning(
                    T("Nessun dispositivo"),
                    T("La riga %s c'e', ma sotto non ci sono dispositivi validi.\n"
                    "Non e' stato importato niente.") % stanza, parent=self)
                return
            esito = dict(esito)
            esito["altre_stanze"] = scartati
            esito["regola"] = regola
        conferma = ImportDialog(
            self, path, len(items), esito, opzioni,
            self.contati_in_eliminazione(opzioni),
            self.store.anteprima_importazione(items, opzioni["mode"], stanza)).show()
        if not conferma:
            return
        risultato = self._run(lambda: self.store.import_items(
            items, opzioni["mode"], stanza))
        if not risultato:
            return
        messagebox.showinfo(T("Importazione completata"),
                            "\n".join(self._resoconto_importazione(
                                risultato, esito, opzioni, scartati)),
                            parent=self)

    def _resoconto_importazione(self, risultato, esito, opzioni, scartati=0):
        """Che cosa e' stato fatto davvero, riga per riga.

        A operazione avvenuta i numeri non bastano: chi ha appena riscritto
        l'inventario di tutti deve poter controllare che sia successo quello che
        si aspettava, e ritrovare la copia di sicurezza se non lo e'.
        """
        esito = esito or {}
        righe = [T("%s in %s.")
                 % (T("Sostituzione") if opzioni["mode"] == "replace" else T("Unione"),
                    opzioni["stanza"] or T("tutto l'inventario")), ""]
        if risultato.get("eliminati"):
            righe.append(T("Eliminati prima del caricamento: %d")
                         % risultato["eliminati"])
        righe.append(T("Aggiunti: %d") % risultato["aggiunti"])
        righe.append(T("Aggiornati: %d") % risultato["aggiornati"])

        saltate = []
        if esito.get("scartate"):
            saltate.append(T("  %d senza asset tag") % esito["scartate"])
        if esito.get("iphone"):
            saltate.append(T("  %d iPhone: si inseriscono solo a mano") % esito["iphone"])
        if scartati:
            saltate.append(T("  %d di altre stanze") % scartati)
        if esito.get("doppioni"):
            saltate.append(T("  %d doppioni nel foglio: tenuta l'ultima riga")
                           % len(esito["doppioni"]))
        if esito.get("senza_modello"):
            saltate.append(T("  %d senza modello: caricate lo stesso")
                           % esito["senza_modello"])
        if saltate:
            righe.append("")
            righe.append(T("Righe non caricate, o caricate con riserva:"))
            righe.extend(saltate)
        if esito.get("colonne_ignorate"):
            righe.append("")
            righe.append(T("Colonne del foglio non riconosciute: %s")
                         % ", ".join(esito["colonne_ignorate"][:6]))

        righe.append("")
        righe.append(T("In inventario adesso: %d dispositivi.") % len(self.store.items))
        per_stanza = []
        for stanza in self.cfg["rooms"]:
            quanti = sum(1 for i in self.store.items if i.get("stanza") == stanza)
            per_stanza.append("  %-28s %d" % (stanza, quanti))
        righe.extend(per_stanza)
        if risultato.get("copia"):
            righe.append("")
            righe.append(T("Copia di sicurezza del file precedente:"))
            righe.append(risultato["copia"])
        return righe

    def on_import_room(self):
        """Carica dal file solo la sezione che riguarda la stanza aperta."""
        stanza = self.var_room.get()
        opzioni = ImportOptionsDialog(self, self.cfg.get("rooms") or [],
                                      stanza_fissa=stanza).show()
        if not opzioni:
            return
        path = filedialog.askopenfilename(
            parent=self, title=T("File da importare in %s") % stanza,
            filetypes=[(T("File Excel"), "*.xlsx *.xlsm"), (T("Tutti i file"), "*.*")])
        if not path:
            return
        try:
            items, esito = rows_from_workbook(path, self.cfg.get("rooms"))
        except InventoryError as exc:
            messagebox.showerror(T("Importazione non riuscita"), str(exc), parent=self)
            return

        miei, scartati, regola = seleziona_per_stanza(items, esito, stanza)
        if regola == "mancante":
            self._avviso_stanza_mancante(stanza, esito)
            return
        if not miei:
            messagebox.showwarning(
                T("Nessun dispositivo"),
                T("La riga %s c'e', ma sotto non ci sono dispositivi validi.\n"
                "Non e' stato importato niente.") % stanza, parent=self)
            return
        esito_stanza = dict(esito)
        esito_stanza["altre_stanze"] = scartati
        esito_stanza["regola"] = regola
        conferma = ImportDialog(
            self, path, len(miei), esito_stanza, opzioni,
            self.contati_in_eliminazione(opzioni),
            self.store.anteprima_importazione(miei, opzioni["mode"], stanza)).show()
        if not conferma:
            return
        risultato = self._run(lambda: self.store.import_items(
            miei, opzioni["mode"], stanza))
        if not risultato:
            return
        messagebox.showinfo(T("Importazione completata"),
                            "\n".join(self._resoconto_importazione(
                                risultato, esito_stanza, opzioni, scartati)),
                            parent=self)

    def _avviso_stanza_mancante(self, stanza, esito):
        trovate = esito.get("stanze_trovate") or []
        messagebox.showwarning(
            T("Manca la riga della stanza"),
            T("Il foglio dichiara le stanze con le righe-separatore, ma nessuna\n"
            "indica %s: non so quali dispositivi siano suoi.\n"
            "Non e' stato importato niente.\n\n"
            "Aggiungi al file una riga vuota con scritto soltanto\n\n"
            "        %s\n\n"
            "nella prima cella, e sotto elenca i dispositivi della stanza.\n\n"
            "Nel file ho trovato invece: %s.")
            % (stanza, stanza.upper(), ", ".join(trovate)), parent=self)

    def contati_in_eliminazione(self, opzioni):
        """Quanti dispositivi verrebbero rimossi da una sostituzione."""
        if opzioni["mode"] != "replace":
            return 0
        stanza = opzioni["stanza"]
        return sum(1 for i in self.store.items
                   if not is_iphone(i.get("tipo"))
                   and (stanza is None or i.get("stanza") == stanza))

    def fine_esportazione(self, descrizione, percorsi, cartella=None):
        """Chiude ogni esportazione offrendo invio, apertura, o niente."""
        azione = EsportazioneFattaDialog(self, descrizione, percorsi).show()
        if azione == "email":
            try:
                allegato = excel_io.allega_a_outlook(percorsi)
            except InventoryError as exc:
                messagebox.showwarning(T("Invio per e-mail"), str(exc), parent=self)
                return
            self.var_status.set(
                T("Messaggio aperto in Outlook con %s allegato.     ")
                % os.path.basename(allegato) + self.var_status.get())
        elif azione == "apri":
            excel_io.open_file(cartella or percorsi[0])

    def on_export_room(self):
        """Esporta il contenuto della stanza aperta, filtri esclusi."""
        stanza = self.var_room.get()
        items = [i for i in self.store.items if i.get("stanza") == stanza]
        if not items:
            messagebox.showinfo(T("Esporta"), T("%s non contiene dispositivi.") % stanza,
                                parent=self)
            return
        percorso = filedialog.asksaveasfilename(
            parent=self, title=T("Esporta %s") % stanza, defaultextension=".xlsx",
            initialfile="Inventario_%s_%s.xlsx" % (
                nome_file(stanza), datetime.now().strftime("%Y%m%d")),
            filetypes=[(T("File Excel"), "*.xlsx")])
        if not percorso:
            return
        try:
            excel_io.export(items, percorso, rooms=[stanza], titolo=stanza,
                            lingua=lang.corrente())
        except InventoryError as exc:
            messagebox.showerror(T("Esportazione non riuscita"), str(exc), parent=self)
            return
        self.fine_esportazione(
            T("%d dispositivi di %s esportati in:\n%s") % (len(items), stanza, percorso),
            [percorso])

    def on_export(self):
        """Prima cosa esportare e in che forma, poi dove salvarlo."""
        rooms = self.cfg.get("rooms") or []
        opzioni = ExportOptionsDialog(self, rooms).show()
        if not opzioni:
            return
        stanza, forma = opzioni["stanza"], opzioni["forma"]
        lingua_file = opzioni.get("lingua")
        items = [i for i in self.store.items
                 if stanza is None or i.get("stanza") == stanza]
        if not items:
            messagebox.showinfo(
                T("Esporta"),
                T("%s non contiene dispositivi da esportare.")
                % (stanza or T("L'inventario")), parent=self)
            return

        if forma == "file":
            cartella = filedialog.askdirectory(
                parent=self, title=T("Cartella in cui salvare un file per stanza"))
            if not cartella:
                return
            try:
                scritti = excel_io.export_per_stanza(items, cartella, rooms,
                                                     lingua=lingua_file)
            except InventoryError as exc:
                messagebox.showerror(T("Esportazione non riuscita"), str(exc), parent=self)
                return
            if not scritti:
                messagebox.showinfo(T("Esporta"), T("Nessuna stanza contiene dispositivi."),
                                    parent=self)
                return
            self.fine_esportazione(
                T("%d file scritti in:\n%s") % (len(scritti), cartella),
                scritti, cartella=cartella)
            return

        if stanza:
            proposto = "Inventario_%s_%s.xlsx" % (nome_file(stanza),
                                                  datetime.now().strftime("%Y%m%d"))
        else:
            proposto = "Inventario_%s.xlsx" % datetime.now().strftime("%Y%m%d")
        percorso = filedialog.asksaveasfilename(
            parent=self, title=T("Esporta inventario"), defaultextension=".xlsx",
            initialfile=proposto, filetypes=[(T("File Excel"), "*.xlsx")])
        if not percorso:
            return
        try:
            excel_io.export(items, percorso, group_by_room=(forma == "fogli"),
                            rooms=[stanza] if stanza else rooms,
                            titolo=stanza, lingua=lingua_file)
        except InventoryError as exc:
            messagebox.showerror(T("Esportazione non riuscita"), str(exc), parent=self)
            return
        descrizione = (T("%d dispositivi di %s") % (len(items), stanza) if stanza
                       else T("%d dispositivi") % len(items))
        if forma == "fogli":
            descrizione += T(", un foglio per stanza")
        self.fine_esportazione(
            T("%s esportati in:\n%s") % (descrizione, percorso), [percorso])

    def on_print(self):
        items = self.filtered_items()
        if not items:
            messagebox.showinfo(T("Stampa"), T("Non c'e' nulla da stampare nella vista corrente."),
                                parent=self)
            return
        per_room = self.view == "home" and messagebox.askyesno(
            T("Stampa"), T("Stampare una pagina separata per ogni stanza?\n\n"
            "No = un unico elenco."), parent=self)
        try:
            path = excel_io.build_print_file(items, group_by_room=per_room,
                                             rooms=self.cfg["rooms"])
            printed = excel_io.send_to_printer(path)
        except InventoryError as exc:
            messagebox.showerror(T("Stampa non riuscita"), str(exc), parent=self)
            return
        if printed:
            self.var_status.set(T("Stampa inviata alla stampante predefinita.     ")
                                + self.var_status.get())
        else:
            messagebox.showinfo(
                T("Stampa"), T("Il documento e' stato aperto in Excel.\n"
                "Usa File > Stampa per inviarlo alla stampante."), parent=self)

    def on_reset(self):
        """Svuota l'inventario, per poi ricaricarlo da un'importazione."""
        if not self.store.items:
            messagebox.showinfo(T("Reset"), T("L'inventario e' gia' vuoto."), parent=self)
            return
        protetti = [i for i in self.store.items
                    if is_iphone(i.get("tipo")) or not puo_essere_eliminato(i)[0]]
        da_eliminare = len(self.store.items) - len(protetti)
        if not da_eliminare:
            messagebox.showinfo(
                T("Reset"),
                T("Non c'e' niente da eliminare: tutti i %d dispositivi in inventario\n"
                "sono iPhone protetti dalla conservazione.") % len(protetti), parent=self)
            return
        if not ResetDialog(self, da_eliminare, len(protetti)).show():
            return
        esito = self._run(lambda: self.store.reset())
        if not esito:
            return
        eliminati, tenuti, copia = esito
        messaggio = T("Eliminati %d dispositivi.") % eliminati
        if tenuti:
            messaggio += T("\nMantenuti %d iPhone.") % tenuti
        messaggio += (T("\n\nCopia di sicurezza del file precedente:\n%s\n\n"
                      "Ora puoi ricaricare l'inventario con Importa xls...") % copia)
        messagebox.showinfo(T("Inventario svuotato"), messaggio, parent=self)

    def on_restore(self, scegli=False):
        """Riporta l'inventario a una copia di sicurezza.

        Senza argomenti propone l'ultima copia, che e' il caso di gran lunga piu'
        frequente: un'importazione andata storta si annulla cosi'. Con
        scegli=True apre l'elenco completo.
        """
        copie = self.store.copie_disponibili()
        if not copie:
            messagebox.showinfo(
                T("Ripristina"),
                T("Non c'e' ancora nessuna copia di sicurezza.\n\n"
                  "Ne viene salvata una a ogni reset e a ogni importazione\n"
                  "che sostituisce i dati."), parent=self)
            return
        if scegli:
            percorso = RestoreDialog(self, copie, len(self.store.items)).show()
            if not percorso:
                return
            quando, quanti = next((q, n) for p, q, n in copie if p == percorso)
        else:
            percorso, quando, quanti = copie[0]
            if not messagebox.askyesno(
                T("Ripristina l'ultima copia"),
                T("Ultima copia salvata: %s\n"
                  "Contiene %d dispositivi; adesso in inventario ce ne sono %d.\n\n"
                  "L'inventario di tutti tornera' com'era in quel momento.\n"
                  "Lo stato attuale viene salvato prima, cosi' puoi tornare indietro.\n\n"
                  "Procedere?")
                % (quando.strftime("%d/%m/%Y %H:%M:%S"), quanti, len(self.store.items)),
                parent=self
            ):
                return
        esito = self._run(lambda: self.store.restore(percorso))
        if not esito:
            return
        ripristinati, precedente = esito
        self._sync_filter_values()
        self.show_home()
        messagebox.showinfo(
            T("Inventario ripristinato"),
            T("Ripristinati %d dispositivi dalla copia del %s.\n\n"
              "Lo stato precedente e' stato salvato in:\n%s")
            % (ripristinati, quando.strftime("%d/%m/%Y %H:%M:%S"), precedente),
            parent=self)

    def on_ripristino_locale(self):
        """Rimette tutto - inventario e stanze - da una copia salvata in locale.

        E' la via d'uscita dal caso peggiore: la cartella di rete sparita, e con
        lei i backup automatici che ci stavano dentro. Per questo legge anche le
        impostazioni: dal solo elenco dei dispositivi le stanze non tornano.
        """
        percorso = filedialog.askopenfilename(
            parent=self, title=T("Scegli la copia da cui ripristinare"),
            initialdir=os.path.expanduser("~"),
            filetypes=[(T("Copia dell'inventario"), "*.zip *.xlsx"),
                       (T("Copia completa (zip)"), "*.zip"),
                       (T("File Excel"), "*.xlsx")])
        if not percorso:
            return
        try:
            rapporto = self.store.anteprima_copia_locale(percorso)
        except InventoryError as exc:
            messagebox.showerror(T("Ripristino non riuscito"), str(exc), parent=self)
            return

        stanze_adesso = {}
        for item in self.store.items:
            stanza = clean(item.get("stanza")) or NO_ROOM()
            stanze_adesso[stanza] = stanze_adesso.get(stanza, 0) + 1
        righe = riepilogo_copia_locale(rapporto, len(self.store.items),
                                       stanze_adesso, self.cfg)
        avvertenza = T("L'inventario di tutti torna com'era nella copia.\n"
                       "Lo stato attuale viene salvato prima, cosi' puoi tornare indietro.")
        if not ConfermaOperazioneDialog(
            self, T("Ripristina da un file locale"),
            T("Ripristinare tutto da %s?") % os.path.basename(percorso),
            righe, T("Ripristina"), "Rosso.TButton", avvertenza
        ).show():
            return

        esito = self._run(lambda: self.store.ripristina_da_copia_locale(percorso))
        if not esito:
            return
        quanti, precedente, con_impostazioni = esito
        self.cfg = config.load_shared_config(self.store.path)
        self.store.iphone_room = self.cfg.get("iphone_room")
        self.store.stati = list(self.cfg.get("states") or [])
        self._sync_filter_values()
        self.show_home()
        coda = T("Sono tornate anche le impostazioni: stanze, tipi e stati.") \
            if con_impostazioni else \
            T("La copia non portava le impostazioni: stanze e tipi sono rimasti\n"
              "quelli di prima.")
        messagebox.showinfo(
            T("Inventario ripristinato"),
            T("Ripristinati %d dispositivi da %s.\n\n%s\n\n"
              "Lo stato precedente e' stato salvato in:\n%s")
            % (quanti, os.path.basename(percorso), coda, precedente), parent=self)

    def on_copia_locale(self):
        """Salva una copia dell'inventario dove decide l'utente.

        Serve contro il caso che nessun backup automatico copre: la cartella di
        rete che sparisce, o qualcuno che ci cancella dentro. La copia e' un
        inventario completo e apribile - non un estratto - quindi da sola basta
        a ripartire.
        """
        adesso = datetime.now()
        proposto = "Inventario_%s.zip" % adesso.strftime("%Y-%m-%d_%H-%M")
        percorso = filedialog.asksaveasfilename(
            parent=self, title=T("Salva una copia dell'inventario"),
            defaultextension=".zip", initialfile=proposto,
            initialdir=os.path.expanduser("~"),
            filetypes=[(T("Copia completa (zip)"), "*.zip"),
                       (T("Solo i dispositivi (xlsx)"), "*.xlsx")])
        if not percorso:
            return
        try:
            salvato, impostazioni, quanti = self.store.copia_in(percorso)
        except InventoryError as exc:
            messagebox.showerror(T("Copia non riuscita"), str(exc), parent=self)
            return
        e_uno_zip = salvato.lower().endswith(".zip")
        righe = [T("%d dispositivi, come sono in questo momento.") % quanti,
                 "",
                 salvato]
        if impostazioni and e_uno_zip:
            righe.append("")
            righe.append(T("Dentro ci sono l'inventario e le impostazioni: stanze,\n"
                           "tipi e stati per rimetterlo com'era."))
        elif impostazioni:
            righe.append(os.path.basename(impostazioni))
            righe.append("")
            righe.append(T("Accanto ai dati e' stato salvato anche il file delle\n"
                           "impostazioni: stanze, tipi e stati per rimetterlo\n"
                           "com'era."))
        righe.append("")
        if e_uno_zip:
            righe.append(T("Da qui si riparte anche se la cartella di rete sparisce:\n"
                           "Impostazioni > Ripristina da un file locale...\n\n"
                           "Per guardare i dati in Excel, apri lo zip con un doppio\n"
                           "clic: l'inventario dentro e' un .xlsx normale."))
        else:
            righe.append(T("E' un inventario completo: si apre in Excel, e in caso di\n"
                           "guaio si ricarica con Ripristina da un file locale...\n"
                           "o con Importa xls... in modalita' Sostituisci."))
        self._modifiche_alla_copia = self.store.modifiche
        messagebox.showinfo(T("Copia salvata"), "\n".join(righe), parent=self)

    def on_collega(self):
        """Sceglie la cartella condivisa in cui sta l'inventario di tutti."""
        from . import configura
        attuale = os.path.dirname(os.path.dirname(os.path.abspath(self.store.path)))
        cartella = filedialog.askdirectory(
            parent=self, title=T("Scegli la cartella condivisa dell'inventario"),
            initialdir=attuale if os.path.isdir(attuale) else None)
        if not cartella:
            return
        percorso = configura.percorso_inventario(cartella)
        if os.path.abspath(percorso) == os.path.abspath(self.store.path):
            messagebox.showinfo(
                T("Gia' collegato"),
                T("E' gia' questo l'inventario aperto:\n\n%s") % percorso,
                parent=self)
            return
        esiste = os.path.exists(percorso)
        if not messagebox.askyesno(
            T("Collega inventario condiviso"),
            T("%s\n\n%s\n\nDa adesso questa postazione lavorera' su quel file, e\n"
              "l'inventario aperto ora non verra' piu' usato ne' modificato.\n\n"
              "Il programma va chiuso e riaperto. Procedo?")
            % (percorso, T("L'inventario e' gia' li' e non verra' toccato.") if esiste
               else T("Li' non c'e' ancora nessun inventario: ne verra' creato uno vuoto.")),
            parent=self
        ):
            return
        try:
            percorso, _ = configura.collega(cartella)
        except InventoryError as exc:
            messagebox.showerror(T("Collegamento non riuscito"), str(exc), parent=self)
            return
        messagebox.showinfo(
            T("Collegato"),
            T("Questa postazione ora apre:\n\n%s\n\nRiapri il programma per lavorarci.")
            % percorso, parent=self)
        self.destroy()

    def on_settings(self):
        result = RoomsDialog(self, self.cfg["rooms"], self.cfg["types"],
                             self.cfg.get("loan_rooms", []),
                             self.cfg.get("iphone_room", ""),
                             lang.corrente()).show()
        if not result:
            return
        if result.get("collega"):
            self.on_collega()
            return
        if result.get("ripristina"):
            self.on_restore(scegli=True)
            return
        if result.get("copia_locale"):
            self.on_copia_locale()
            return
        if result.get("reset"):
            self.on_reset()
            return
        if result.get("ripristina_locale"):
            self.on_ripristino_locale()
            return
        nuova_lingua = result.pop("lingua", lang.corrente())
        rinomine = result.pop("rinomine", [])
        cambi_tipo = result.pop("rinomine_tipi", [])

        # Una stanza tolta lascerebbe i suoi dispositivi con il nome di una
        # stanza che non esiste piu': non si trovano con i filtri e non si sa
        # che fine abbiano fatto. Vanno negli eliminati di recente, ma non di
        # nascosto: si chiede, e se si dice di no non si salva niente.
        rinominate = set(v for v, _n in rinomine)
        tolte = [r for r in self.cfg.get("rooms", [])
                 if r not in result.get("rooms", []) and r not in rinominate]
        conteggio = {k: v for k, v in self.store.quanti_nelle_stanze(tolte).items() if v}
        if conteggio:
            righe = "\n".join("  %s: %d dispositivi" % (stanza, quanti)
                               for stanza, quanti in sorted(conteggio.items()))
            if not messagebox.askyesno(
                T("Stanza tolta con dispositivi dentro"),
                T("Stai togliendo dall'elenco stanze che non sono vuote:\n\n%s\n\n"
                  "Questi dispositivi finiscono in Eliminati di recente, da dove\n"
                  "si ripristinano scegliendo una stanza. Restano li' %d giorni.\n\n"
                  "Procedere?") % (righe, ELIMINATI_GIORNI), parent=self):
                return
            if self._run(lambda: self.store.porta_via_gli_orfani(tolte)) is None:
                return
        spostati = {}
        cambiati = {}
        if cambi_tipo:
            esito = self._run(lambda: self.store.rinomina_tipi(cambi_tipo))
            if esito is None:
                return
            cambiati = esito
        if rinomine:
            # Prima i dispositivi, poi le impostazioni: se lo spostamento non
            # riesce - la share che non risponde, un altro tecnico che sta
            # salvando - non si e' ancora scritto niente, e l'inventario resta
            # coerente con l'elenco delle stanze.
            esito = self._run(lambda: self.store.rinomina_stanze(rinomine))
            if esito is None:
                return
            spostati = esito
        try:
            config.save_shared_config(self.store.path, result)
        except OSError as exc:
            messagebox.showerror(T("Errore"), T("Impossibile salvare le impostazioni:\n%s") % exc,
                                 parent=self)
            return
        if nuova_lingua != lang.corrente():
            lang.imposta(nuova_lingua)
            config.save_language(nuova_lingua)
            self.ricostruisci()
            return
        self.cfg = config.load_shared_config(self.store.path)
        self.store.iphone_room = self.cfg.get("iphone_room")
        self.store.stati = list(self.cfg.get("states") or [])
        self._sync_filter_values()
        self.show_home()
        if rinomine or cambi_tipo:
            righe = []
            if rinomine:
                righe.append(T("STANZE:"))
                righe.extend(T("%s  ->  %s   (%d dispositivi)")
                             % (vecchio, nuovo, spostati.get(nuovo, 0))
                             for vecchio, nuovo in rinomine)
            if cambi_tipo:
                if righe:
                    righe.append("")
                righe.append(T("TIPI DI DISPOSITIVO:"))
                righe.extend(T("%s  ->  %s   (%d dispositivi)")
                             % (vecchio, nuovo, cambiati.get(nuovo, 0))
                             for vecchio, nuovo in cambi_tipo)
            messagebox.showinfo(
                T("Rinomina completata"),
                T("%s\n\nI dispositivi sono stati aggiornati con il nome nuovo.\n"
                  "Non hanno perso niente: cambia solo l'etichetta.")
                % "\n".join(righe), parent=self)


# ------------------------------------------------------------ avvio


def choose_data_file(root):
    """Prima configurazione: dove si trova, o dove creare, il file inventario.

    Non viene mostrata quando accanto al programma c'e' gia' Inventario.xlsx,
    che e' il caso normale con eseguibile e dati nella stessa cartella di rete.
    """
    cartella = config.app_dir()
    answer = messagebox.askyesnocancel(
        T('Site Services : Inventario Iphone, Laptop e Tablet'),
        T("Non ho ancora trovato il file dell'inventario.\n\n"
        "Di norma si chiama Inventario.xlsx e sta nella stessa cartella del\n"
        "programma:\n%s\n\n"
        "Si'  = apri un file inventario gia' esistente\n"
        "No   = crea qui un nuovo inventario vuoto\n"
        "Annulla = esci") % cartella,
        parent=root)
    if answer is None:
        return None
    if answer:
        while True:
            scelto = filedialog.askopenfilename(
                parent=root, title=T("Seleziona il file inventario"),
                initialdir=config.production_dir() if os.path.isdir(
                    config.production_dir()) else cartella,
                filetypes=[(T("File Excel"), "*.xlsx")]) or None
            if not scelto:
                return None
            da_importare, motivo = sembra_un_foglio_da_importare(
                scelto, config.load_shared_config(scelto).get("rooms"))
            if not da_importare:
                return scelto
            messagebox.showwarning(
                T("Questo non e' un inventario"),
                T("%s\n\n%s\n\nUn foglio del genere si CARICA in un inventario con\n"
                  "Importa xls..., non si apre come inventario: aprendolo, le\n"
                  "righe separatore diventerebbero dispositivi e nessun\n"
                  "dispositivo avrebbe una stanza.\n\n"
                  "Scegli Annulla, lascia creare l'inventario vuoto, e importa\n"
                  "questo file da dentro il programma.")
                % (os.path.basename(scelto), motivo), parent=root)
    return filedialog.asksaveasfilename(
        parent=root, title=T("Crea il file inventario"),
        defaultextension=".xlsx", initialdir=cartella,
        initialfile=config.DATA_FILE_NAME,
        filetypes=[(T("File Excel"), "*.xlsx")]) or None


def main():
    path = config.load_data_path()
    if not path:
        # Se questa installazione ha gia' un inventario assegnato - il caso del
        # programma sulle postazioni e dei dati sulla share - e non lo si
        # raggiunge, ci si ferma. Crearne uno nuovo in locale farebbe lavorare
        # il tecnico su una copia che nessun altro vede, ed e' il modo piu'
        # silenzioso di perdere il lavoro di una giornata.
        atteso, sorgente = config.configured_data_path()
        if atteso:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                T("Inventario non raggiungibile"),
                T("Questo programma deve aprire l'inventario condiviso:\n\n%s\n\n"
                  "In questo momento non si raggiunge. Di solito e' la cartella\n"
                  "di rete che non risponde, o la connessione.\n\n"
                  "Controlla di vedere quella cartella da Esplora risorse, poi\n"
                  "riapri il programma. Non viene creato nessun inventario\n"
                  "locale: si lavora tutti sullo stesso file.\n\n"
                  "Il percorso e' scritto in:\n%s") % (atteso, sorgente),
                parent=root)
            root.destroy()
            return 2
    if not path:
        # Al primo avvio l'inventario si crea da solo, in Produzione accanto al
        # programma. Chiedere all'utente dove metterlo era il modo piu' rapido
        # per ritrovarsi come inventario un file di prova scelto per sbaglio:
        # la domanda si fa solo se qui non si puo' scrivere davvero.
        candidato = config.default_data_path()
        try:
            cartella = os.path.dirname(os.path.abspath(candidato))
            if not os.path.isdir(cartella):
                os.makedirs(cartella)
            InventoryStore(candidato).create_if_missing()
            path = candidato
        except (OSError, InventoryError):
            path = None
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
