"""Ogni finestra con un campo di testo riceve i tasti, tutti, e li restituisce.

Il caso vero, su Windows: si preme Presta su una riga, si apre la finestra del
nome, e si riesce a scrivere due caratteri. Oppure la tastiera non risponde
piu' del tutto. Qui ogni finestra modale viene aperta PER DAVVERO - con show(),
la presa e il fuoco - e le si mandano venti tasti: devono arrivare tutti e
venti. Alla chiusura, la presa deve essere sparita e il fuoco deve essere
tornato alla finestra principale: una finestra in cui nessun widget ha il fuoco
e' esattamente "la tastiera non risponde".
"""
import os, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
import tkinter as tk
from tkinter import ttk, messagebox
from inventario import ui
from inventario.store import InventoryStore, new_item
from inventario.ui import (App, CestinoDialog, EliminaDaExcelDialog,
                           EliminaPlusDialog, ItemDialog, StanzeDaAssegnareDialog,
                           _Modal, e_un_campo_di_testo)

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
VENTI = "abcdefghijklmnopqrst"

app = App(fixture.build())
app._initial_load()
app.update()
scadenza = time.time() + 0.4
while time.time() < scadenza:
    app.update()
avvisi = []
for nome in ("showinfo", "showwarning", "showerror"):
    setattr(messagebox, nome, lambda t, m, **k: avvisi.append((t, m)))
esplosioni = []
app.report_callback_exception = lambda *a: esplosioni.append(a)
problemi = []

def respira(n=3):
    for _ in range(n):
        app.update_idletasks(); app.update()

def finestra_aperta():
    """L'ultima finestra modale viva."""
    vive = [w for w in app.winfo_children()
            if isinstance(w, _Modal) and w.winfo_exists()]
    return vive[-1] if vive else None

def quando_pronta(azione, scadenza=None):
    """Esegue `azione` appena la finestra modale e' davvero pronta.

    Pronta vuol dire quello che vede un utente: a schermo, con la presa e con
    il fuoco dentro. Un timer fisso non basta - sotto carico scatta mentre la
    finestra sta ancora comparendo, e la prova osserverebbe uno stato che
    nessuno puo' vedere, perche' la finestra non e' ancora a schermo.
    """
    scadenza = scadenza or time.time() + 10
    d = finestra_aperta()
    pronta = False
    if d is not None:
        try:
            fuoco = app.focus_get()
        except KeyError:
            fuoco = None
        pronta = (d.winfo_viewable() and app.grab_current() is d
                  and fuoco is not None and str(fuoco).startswith(str(d)))
    if pronta or time.time() > scadenza:
        azione()
    else:
        app.after(20, lambda: quando_pronta(azione, scadenza))

def scrivi(campo, testo):
    for lettera in testo:
        campo.event_generate("<KeyPress>", keysym="space" if lettera == " " else lettera)
    respira()

def scrivibile(campo):
    """Una tendina a scelta fissa riceve i tasti ma non ci scrive: si salta."""
    try:
        return str(campo.cget("state")) not in ("readonly", "disabled")
    except tk.TclError:
        return True

def contenuto(campo):
    if isinstance(campo, tk.Text):
        return campo.get("1.0", "end").rstrip("\n")
    return campo.get()

def prova(nome, apri, campo_atteso=None, quanti=len(VENTI)):
    """Apre la finestra per davvero e, mentre e' aperta, le manda i tasti.

    `apri()` blocca (e' show()): quello che va fatto dentro si programma
    prima con after, e gira nel ciclo di eventi annidato della finestra.
    """
    esito = {}

    def dentro():
        d = finestra_aperta()
        esito["finestra"] = d
        if d is None:
            esito["errore"] = "nessuna finestra aperta"
            return
        try:
            esito["presa"] = app.grab_current()
            try:
                fuoco = app.focus_get()
            except KeyError:
                fuoco = "(interno di Tk)"
            esito["fuoco"] = fuoco
            campo = campo_atteso(d) if campo_atteso else fuoco
            esito["campo"] = campo
            esito["scrivibile"] = e_un_campo_di_testo(campo) and scrivibile(campo)
            if esito["scrivibile"]:
                prima = contenuto(campo)
                scrivi(campo, VENTI)
                esito["scritto"] = contenuto(campo)[len(prima):]
        finally:
            if d.winfo_exists():
                d.destroy()

    app.after(20, lambda: quando_pronta(dentro))
    apri()
    respira()
    d = esito.get("finestra")
    ok = True
    def verifica(c, cosa):
        nonlocal ok
        if not c:
            ok = False
            problemi.append("%s: %s" % (nome, cosa))
    verifica(d is not None, "la finestra non si e' aperta")
    if d is not None:
        verifica(esito.get("presa") is d, "la presa non era sulla finestra: %s" % esito.get("presa"))
        fuoco = esito.get("fuoco")
        verifica(fuoco is not None and str(fuoco).startswith(str(d)),
                 "il fuoco non era dentro la finestra: %s" % fuoco)
        if esito.get("scrivibile"):
            verifica(esito.get("scritto") == VENTI[:quanti],
                     "arrivati %r invece di %d tasti" % (esito.get("scritto"), quanti))
    verifica(app.grab_current() is None, "la presa e' rimasta: %s" % app.grab_current())
    try:
        dopo = app.focus_get()
    except KeyError:
        dopo = None
    verifica(dopo is not None and str(dopo).startswith(str(app)),
             "alla chiusura il fuoco non e' tornato alla finestra principale: %s" % dopo)
    verifica(not esplosioni, "eccezioni nascoste: %s" % (esplosioni[:1],))
    print("  %-34s %s" % (nome, "ok" if ok else "NO"))
    if not ok:
        print("      " + problemi[-1])

# ============ tutte le finestre con un campo di testo ============
app.show_room(KIOSK); respira(6)
libero = [i for i in app.store.items if i["stanza"] == KIOSK and not i.get("prestato_a")][0]

print("== finestre ==")
prova("Presta dispositivo", lambda: app._ask_person(libero))
prova("Scegli stanza", lambda: app._ask_room("dove?"))
prova("Nuovo dispositivo",
      lambda: ItemDialog(app, app.cfg["rooms"], app.cfg["types"],
                         iphone_room=BAU, stati=app.cfg["states"]).show())
prova("Modifica dispositivo",
      lambda: ItemDialog(app, app.cfg["rooms"], app.cfg["types"], item=dict(libero),
                         iphone_room=BAU, stati=app.cfg["states"]).show())
prova("Elimina +", lambda: EliminaPlusDialog(app, app.store).show())
prova("Elimina da Excel", lambda: EliminaDaExcelDialog(app, app.store).show())
prova("Stanze da assegnare",
      lambda: StanzeDaAssegnareDialog(app, 3, app.cfg["rooms"]).show())
prova("Eliminati di recente",
      lambda: CestinoDialog(app, app.store, app.cfg["rooms"]).show())

# ============ il prestito dal pulsante di riga, con la ricarica in mezzo ============
print("== il prestito dal pulsante di riga ==")
app.show_room(KIOSK); respira(6)
tag = libero["asset_tag"]
altro = InventoryStore(app.store.path, iphone_room=BAU)
altro.stanze = list(app.cfg["rooms"])
esito = {}

def dentro_il_prestito():
    d = finestra_aperta()
    esito["finestra"] = d
    if d is None:
        return
    try:
        campo = app.focus_get()
        esito["campo"] = campo
        scrivi(campo, "Ma")                 # due caratteri...
        # ...e in quel momento un altro tecnico salva e scatta la ricarica
        altro.load()
        altro.add(new_item("IT-7778", "Laptop", "T14", "PF8", BAU))
        assert app.store.changed_on_disk()
        app._auto_refresh(); respira()
        esito["viva"] = d.winfo_exists()
        esito["fuoco_dopo"] = app.focus_get()
        scrivi(campo, "rio Rossi")          # ...e si continua a scrivere
        esito["nome"] = campo.get()
        esito["ricaricato"] = bool([i for i in app.store.items
                                    if i["asset_tag"] == "IT-7778"])
    finally:
        if d.winfo_exists():
            d.destroy()

app.after(20, lambda: quando_pronta(dentro_il_prestito))
app._on_row_button(tag)          # rimandato con after_idle: parte dall'update
respira()
assert esito.get("finestra") is not None, "la finestra del prestito non si e' aperta"
assert esito.get("viva"), "la ricarica ha chiuso la finestra del prestito"
assert esito.get("nome") == "Mario Rossi", \
    "i tasti dopo il secondo si sono persi: %r" % esito.get("nome")
assert esito.get("fuoco_dopo") is esito.get("campo"), \
    "la ricarica ha spostato il fuoco: %s" % esito.get("fuoco_dopo")
assert not esito.get("ricaricato"), "l'elenco non si ricarica mentre si scrive"
assert app.grab_current() is None
assert not esplosioni, esplosioni
print("  prestito dal pulsante di riga    ok")

# a finestra chiusa, la ricarica passa
app.tree.focus_set(); respira()
app._auto_refresh(); respira()
assert [i for i in app.store.items if i["asset_tag"] == "IT-7778"], "poi si ricarica"

# ============ Canc nella ricerca cancella un carattere, non un dispositivo ============
print("== i tasti della finestra dentro i campi ==")
chiamate = []
app.on_delete = lambda *a: chiamate.append("on_delete")
app.entry_search.focus_force(); respira()
app.var_search.set("abc"); app.entry_search.icursor(0); respira()
app.entry_search.event_generate("<Delete>"); respira()
assert not chiamate, "Canc nella ricerca ha aperto l'eliminazione"
assert app.var_search.get() == "bc", app.var_search.get()
app.entry_search.event_generate("<Escape>"); respira()
assert app.view == "room", "Esc nella ricerca non riporta alla home"
app.tree.focus_force(); respira()
app.tree.event_generate("<Delete>"); respira()
assert chiamate == ["on_delete"], "dall'elenco, Canc elimina come sempre"
app.var_search.set("")
print("  Canc ed Esc nei campi              ok")

app.destroy()
if problemi:
    print("\nPROBLEMI:")
    for p in problemi:
        print("  -", p)
    raise SystemExit(1)
print("TASTIERA FINESTRE OK")
