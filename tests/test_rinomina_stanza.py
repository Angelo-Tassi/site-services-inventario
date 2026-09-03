"""Rinominare una stanza porta con se' i dispositivi che ci stavano dentro.

Prima restavano etichettati con il vecchio nome e comparivano in una scheda a
parte, con una stanza che non esisteva piu': bisognava spostarli a mano, uno per
uno. Il difficile non e' spostarli, e' capire quando si tratta davvero di una
rinomina e non di una stanza aggiunta, tolta o solo riordinata.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import messagebox
from inventario import config, ui
from inventario.store import InventoryStore, rinomine_stanze
from inventario.ui import App, RoomsDialog

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR

# ---- che cos'e' una rinomina, e che cosa non lo e'
assert rinomine_stanze([BAU, KIOSK, DR], [BAU, "Kiosk 2", DR]) == [(KIOSK, "Kiosk 2")]
assert rinomine_stanze(["A", "B", "C"], ["A1", "B", "C1"]) == [("A", "A1"), ("C", "C1")]
# riordinare non e' rinominare: i nomi sono gli stessi, cambia solo la fila
assert rinomine_stanze(["A", "B", "C"], ["C", "B", "A"]) == []
# aggiungere e togliere nemmeno
assert rinomine_stanze(["A", "B"], ["A", "B", "C"]) == []
assert rinomine_stanze(["A", "B", "C"], ["A", "C"]) == []
assert rinomine_stanze(["A", "B"], ["A", "B"]) == []
assert rinomine_stanze([], ["A"]) == [] and rinomine_stanze(["A"], []) == []
# gli spazi di troppo non fanno una rinomina
assert rinomine_stanze(["A", "B"], ["A", "  B  "]) == []

# ---- l'archivio sposta davvero i dispositivi
percorso = fixture.build()
store = InventoryStore(percorso, iphone_room=BAU)
store.load()
prima = len([i for i in store.items if i.get("stanza") == KIOSK])
assert prima == 5, prima
spostati = store.rinomina_stanze([(KIOSK, "Kiosk rinnovato")])
assert spostati == {"Kiosk rinnovato": 5}, spostati
store.load()
assert not [i for i in store.items if i.get("stanza") == KIOSK], "nessuno resta indietro"
assert len([i for i in store.items if i.get("stanza") == "Kiosk rinnovato"]) == 5
# i prestiti in corso si spostano con il dispositivo
prestati = [i for i in store.items if i.get("prestato_a")]
assert all(i["stanza"] == "Kiosk rinnovato" for i in prestati), prestati

# ---- una stanza vuota si rinomina lo stesso, e lo dice
vuoto = store.rinomina_stanze([("Stanza che non esiste", "Nuova")])
assert vuoto == {"Nuova": 0}, vuoto
assert store.rinomina_stanze([]) == {}

# ============================ dalle impostazioni ============================
app = App(fixture.build())
app._initial_load()
app.update()
avvisi = []
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))

def impostazioni(stanze, prestiti=None, iphone=None):
    """Apre la finestra, scrive le stanze e preme Salva, come farebbe l'utente."""
    dlg = RoomsDialog(app, app.cfg["rooms"], app.cfg["types"],
                      app.cfg.get("loan_rooms", []), app.cfg.get("iphone_room", ""))
    dlg.text_rooms.delete("1.0", "end")
    dlg.text_rooms.insert("1.0", "\n".join(stanze))
    if prestiti is not None:
        dlg.text_loans.delete("1.0", "end")
        dlg.text_loans.insert("1.0", "\n".join(prestiti))
    if iphone is not None:
        dlg.var_iphone_room.set(iphone)
    dlg._ok()
    risultato = dlg.result
    if dlg.winfo_exists():
        dlg.destroy()
    return risultato

# ---- rinominando il Digital Kiosk, prestiti e dispositivi lo seguono
risultato = impostazioni([BAU, "Kiosk 2", DR])
assert risultato is not None, avvisi
assert risultato["rinomine"] == [(KIOSK, "Kiosk 2")], risultato["rinomine"]
assert risultato["loan_rooms"] == ["Kiosk 2"], \
    "la stanza con prestito segue il nome nuovo, o il salvataggio si bloccherebbe"

# ---- e rinominando la stanza degli iPhone, la segue anche quella
risultato = impostazioni(["BAU 2", KIOSK, DR])
assert risultato["iphone_room"] == "BAU 2", risultato
assert risultato["rinomine"] == [(BAU, "BAU 2")]

# ---- due stanze con lo stesso nome vengono rifiutate
avvisi.clear()
assert impostazioni([BAU, BAU, DR]) is None
assert "compare due volte" in avvisi[-1][1], avvisi[-1]

# ---- il giro intero: si salva e i dispositivi sono nella stanza nuova
avvisi.clear()
ui.RoomsDialog.show = lambda self: impostazioni([BAU, "Kiosk 2", DR])
app.on_settings()
app.store.load()
assert app.cfg["rooms"] == [BAU, "Kiosk 2", DR], app.cfg["rooms"]
assert app.cfg["loan_rooms"] == ["Kiosk 2"], app.cfg
assert len([i for i in app.store.items if i.get("stanza") == "Kiosk 2"]) == 5
assert not [i for i in app.store.items if i.get("stanza") == KIOSK]
titolo, corpo = avvisi[-1]
assert titolo == "Rinomina completata", avvisi[-1]
assert "STANZE:" in corpo, corpo
assert "Digital Kiosk  ->  Kiosk 2   (5 dispositivi)" in corpo, corpo

# ---- e la scheda della stanza vecchia non esiste piu': niente orfani
app.show_home(); app.update()
assert "Kiosk 2" in [n for n in app.cfg["rooms"]]
senza_stanza = [i for i in app.store.items
                if i.get("stanza") not in app.cfg["rooms"]]
assert not senza_stanza, senza_stanza

# ---- la stanza degli iPhone si puo' rinominare e resta la stanza degli iPhone
# E' il caso piu' delicato: i telefoni ci stanno per forza, e se il vincolo
# continuasse a puntare al nome vecchio finirebbero in una stanza inesistente.
from inventario.store import new_item
app._run(lambda: app.store.add(new_item(tipo=fixture.TIPO_IPHONE,
                                        modello="Apple iPhone 14",
                                        imei="356938035643809",
                                        restituito_da="M. B.")), "ok")
assert app.cfg["iphone_room"] == BAU and app.store.iphone_room == BAU

avvisi.clear()
ui.RoomsDialog.show = lambda self: impostazioni(["BAU rinnovato", "Kiosk 2", DR])
app.on_settings()
app.store.load()
assert app.cfg["iphone_room"] == "BAU rinnovato", app.cfg["iphone_room"]
assert app.store.iphone_room == "BAU rinnovato", app.store.iphone_room
telefoni = [i for i in app.store.items if i["tipo"] == fixture.TIPO_IPHONE]
assert telefoni and all(i["stanza"] == "BAU rinnovato" for i in telefoni), telefoni
# e un telefono inserito dopo va nella stanza con il nome nuovo
app._run(lambda: app.store.add(new_item(tipo=fixture.TIPO_IPHONE,
                                        modello="Apple iPhone 15",
                                        imei="351111111111111",
                                        restituito_da="E. R.")), "ok")
app.store.load()
aggiunto = [i for i in app.store.items if i["asset_tag"] == "351111111111111"][0]
assert aggiunto["stanza"] == "BAU rinnovato", aggiunto["stanza"]

# ---- salvare senza toccare le stanze non sposta niente
avvisi.clear()
ui.RoomsDialog.show = lambda self: impostazioni(["BAU rinnovato", "Kiosk 2", DR])
app.on_settings()
assert not [t for t, _m in avvisi if t == "Rinomina completata"], avvisi

app.destroy()
print("RINOMINA STANZA OK")
