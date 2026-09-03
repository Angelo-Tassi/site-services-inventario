"""Togliere la stanza degli iPhone o quella dei prestiti e' un trasloco.

Prima il salvataggio veniva rifiutato con "la stanza non e' nell'elenco", e non
c'era modo di andare avanti se non rimettendo il nome. Ma togliere quella stanza
non e' un errore di battitura: e' un trasloco, e il programma deve chiedere dove
va quello che c'era dentro.

Si portano dietro tutto, compresi i dispositivi in prestito: la regola che li
tiene fermi serve a non perderne la traccia, e lasciarli in una stanza che non
esiste piu' sarebbe esattamente perderla.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import messagebox
from inventario import config, ui
from inventario.store import InventoryStore, is_iphone, new_item
from inventario.ui import App

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
IPHONE = fixture.TIPO_IPHONE

# ---- l'archivio: si porta via tutto, prestiti compresi
percorso = fixture.build()
store = InventoryStore(percorso, iphone_room=BAU)
store.load()
store.add(new_item(tipo=IPHONE, modello="Apple iPhone 14", imei="356938035643809",
                   restituito_da="M. B."))
store.load()
conto = store.trasloca_stanza(KIOSK, DR)
assert conto == {"totale": 5, "prestiti": 2, "iphone": 0}, conto
store.load()
assert not [i for i in store.items if i["stanza"] == KIOSK]
prestati = [i for i in store.items if i.get("prestato_a")]
assert len(prestati) == 2 and all(i["stanza"] == DR for i in prestati), prestati
assert prestati[0]["prestato_a"], "il prestito resta aperto"

# ---- e gli iPhone, rispediti e non
store.ship("356938035643809")
store.add(new_item(tipo=IPHONE, modello="Apple iPhone 15", imei="351111111111111",
                   restituito_da="E. R."))
store.load()
telefoni = [i for i in store.items if is_iphone(i["tipo"])]
assert len(telefoni) == 2 and all(i["stanza"] == BAU for i in telefoni)
assert len([i for i in telefoni if i.get("spedito_il")]) == 1, "uno spedito, uno no"
conto = store.trasloca_stanza(BAU, DR)
store.load()
telefoni = [i for i in store.items if is_iphone(i["tipo"])]
assert conto["iphone"] == 2, conto
assert all(i["stanza"] == DR for i in telefoni), telefoni

# ---- una stanza vuota o uguale a se stessa non fa niente
assert store.trasloca_stanza(BAU, DR) == {"totale": 0, "prestiti": 0, "iphone": 0}
assert store.trasloca_stanza(DR, DR) == {"totale": 0, "prestiti": 0, "iphone": 0}

# ============================ dalle impostazioni ============================
app = App(fixture.build())
app._initial_load()
app.update()
avvisi = []
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))
messagebox.askyesno = lambda t, m, **k: True
app._run(lambda: app.store.add(new_item(tipo=IPHONE, modello="Apple iPhone 14",
                                        imei="356938035643809",
                                        restituito_da="M. B.")), "ok")

domande = []
def rispondi(dove):
    def finto(titolo, prompt, stanze):
        domande.append((titolo, prompt, list(stanze)))
        return dove
    return finto

# ---- togliendo la stanza degli iPhone, il programma chiede dove vanno
app._scegli_stanza = rispondi(DR)
ui.RoomsDialog.show = lambda self: {
    "rooms": [KIOSK, DR], "types": app.cfg["types"],
    "loan_rooms": [KIOSK], "prestiti_spariti": [],
    "iphone_room": "", "iphone_sparita": BAU,
    "rinomine": [], "rinomine_tipi": [], "lingua": "it"}
app.on_settings()
app.store.load()
assert domande, "doveva chiedere dove spostarli"
titolo, prompt, scelte = domande[-1]
assert titolo == "Dove vanno gli iPhone?", titolo
assert BAU in prompt and "rispediti e non" in prompt, prompt
assert scelte == [KIOSK, DR], "si sceglie fra le stanze che restano"
assert app.cfg["iphone_room"] == DR, app.cfg["iphone_room"]
assert app.store.iphone_room == DR
telefoni = [i for i in app.store.items if is_iphone(i["tipo"])]
assert telefoni and all(i["stanza"] == DR for i in telefoni), telefoni
assert not [i for i in app.store.items if i["stanza"] == BAU], "la stanza si e' svuotata"
assert [t for t, _m in avvisi if t == "Stanza traslocata"], avvisi
corpo = [m for t, m in avvisi if t == "Stanza traslocata"][-1]
assert ("%s  ->  %s" % (BAU, DR)) in corpo, corpo
assert "iPhone, rispediti e non" in corpo, corpo

# ---- e i dispositivi non sono finiti nel cestino: sono stati spostati
assert not [v for v in app.store.eliminati()], app.store.eliminati()

# ---- togliendo la stanza dei prestiti, stessa cosa, e la nuova la sostituisce
avvisi.clear(); domande.clear()
app._scegli_stanza = rispondi(DR)
ui.RoomsDialog.show = lambda self: {
    "rooms": [DR], "types": app.cfg["types"],
    "loan_rooms": [], "prestiti_spariti": [KIOSK],
    "iphone_room": DR, "iphone_sparita": "",
    "rinomine": [], "rinomine_tipi": [], "lingua": "it"}
app.on_settings()
app.store.load()
titolo, prompt, _scelte = domande[-1]
assert titolo == "Dove vanno i prestiti?", titolo
assert KIOSK in prompt, prompt
assert app.cfg["loan_rooms"] == [DR], app.cfg["loan_rooms"]
assert not [i for i in app.store.items if i["stanza"] == KIOSK]
prestati = [i for i in app.store.items if i.get("prestato_a")]
assert prestati and all(i["stanza"] == DR for i in prestati), prestati
assert all(i["prestato_a"] for i in prestati), "i prestiti restano aperti"
corpo = [m for t, m in avvisi if t == "Stanza traslocata"][-1]
assert "erano in prestito" in corpo, corpo

# ---- annullando la scelta non si salva niente
avvisi.clear()
prima = list(app.cfg["rooms"])
quanti = len(app.store.items)
app._scegli_stanza = lambda titolo, prompt, stanze: None
ui.RoomsDialog.show = lambda self: {
    "rooms": ["Stanza nuova"], "types": app.cfg["types"],
    "loan_rooms": [], "prestiti_spariti": [DR],
    "iphone_room": "", "iphone_sparita": DR,
    "rinomine": [], "rinomine_tipi": [], "lingua": "it"}
app.on_settings()
app.store.load()
assert app.cfg["rooms"] == prima, app.cfg["rooms"]
assert len(app.store.items) == quanti, "nessun dispositivo doveva muoversi"

app.destroy()
print("TRASLOCO STANZA OK")
