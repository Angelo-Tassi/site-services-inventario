"""La stanza dei prestiti si sceglie da una tendina, come quella degli iPhone.

Sono la stessa cosa - una stanza a cui e' affidato un ruolo - ma i prestiti si
scrivevano a mano in un riquadro di testo: bisognava azzeccare il nome, e
sbagliandolo il salvataggio veniva rifiutato.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import messagebox
from inventario import config, ui
from inventario.ui import App, RoomsDialog, SENZA_PRESTITI

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR

app = App(fixture.build())
app._initial_load()
app.update()
avvisi = []
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))

def finestra():
    return RoomsDialog(app, app.cfg["rooms"], app.cfg["types"],
                       app.cfg.get("loan_rooms", []), app.cfg.get("iphone_room", ""))

def salva(scelta):
    d = finestra()
    if scelta is not None:
        d.var_loan_room.set(scelta)
    d._ok()
    esito = d.result
    if d.winfo_exists():
        d.destroy()
    return esito

# ---- la tendina c'e', sotto quella degli iPhone, e il riquadro a mano non c'e' piu'
d = finestra()
app.update(); d.update()
assert not hasattr(d, "text_loans"), "il riquadro da scrivere a mano e' stato sostituito"
assert str(d.combo_prestiti.cget("state")) == "readonly", "si sceglie, non si scrive"
assert list(d.combo_prestiti.cget("values")) == [SENZA_PRESTITI()] + app.cfg["rooms"]
assert d.var_loan_room.get() == KIOSK, d.var_loan_room.get()
# sta sotto la stanza degli iPhone, non sopra
riga_prestiti = d.combo_prestiti.master
riga_iphone = None
for figlio in riga_prestiti.master.winfo_children():
    for nipote in figlio.winfo_children() if figlio.winfo_children() else []:
        if str(getattr(nipote, "cget", lambda k: "")("text")) == "Stanza degli iPhone":
            riga_iphone = figlio
assert riga_iphone is not None, "non trovo la riga della stanza degli iPhone"
assert riga_iphone.winfo_y() < riga_prestiti.winfo_y(), \
    (riga_iphone.winfo_y(), riga_prestiti.winfo_y())
d.destroy()

# ---- scegliendo un'altra stanza, e' quella la stanza dei prestiti
assert salva(DR)["loan_rooms"] == [DR]
# ---- e con (nessuna) i prestiti si spengono
assert salva(SENZA_PRESTITI())["loan_rooms"] == []
# ---- senza toccarla resta quella di prima
assert salva(None)["loan_rooms"] == [KIOSK]

# ---- il giro vero: la colonna Prestito segue la tendina
ui.RoomsDialog.show = lambda self: salva(DR)
app.on_settings()
assert app.cfg["loan_rooms"] == [DR], app.cfg["loan_rooms"]
app.show_room(DR); app.update()
assert app.loan_column_visible(), "nella stanza scelta compaiono i pulsanti Presta"
app.show_room(KIOSK); app.update()
assert not app.loan_column_visible(), "e spariscono da quella di prima"

# ---- spegnendoli non compaiono da nessuna parte
ui.RoomsDialog.show = lambda self: salva(SENZA_PRESTITI())
app.on_settings()
assert app.cfg["loan_rooms"] == [], app.cfg["loan_rooms"]
for stanza in (BAU, KIOSK, DR):
    app.show_room(stanza); app.update()
    assert not app.loan_column_visible(), stanza

# ---- una configurazione con piu' stanze non perde le altre di nascosto
config.save_shared_config(app.store.path, {"rooms": [BAU, KIOSK, DR],
                                           "types": app.cfg["types"],
                                           "loan_rooms": [KIOSK, DR],
                                           "iphone_room": BAU})
app.cfg = config.load_shared_config(app.store.path)
d = finestra()
assert d.var_loan_room.get() == KIOSK, d.var_loan_room.get()
assert d.altre_stanze_prestito == [DR], d.altre_stanze_prestito
d._ok()
assert d.result["loan_rooms"] == [KIOSK, DR], d.result["loan_rooms"]
d.destroy()

app.destroy()
print("STANZA PRESTITI OK")
