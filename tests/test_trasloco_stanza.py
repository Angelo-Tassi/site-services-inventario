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

# ---- la finestra vera: e' lei che deve accorgersi del trasloco
# Nelle prove qui sotto la finestra viene simulata, quindi qui si controlla che
# quella vera produca davvero le segnalazioni su cui il resto si appoggia.
vera = App(fixture.build())
vera._initial_load()
vera.update()

def salva_stanze(stanze):
    """Come farebbe l'utente: riscrive il riquadro delle stanze e preme Salva."""
    d = ui.RoomsDialog(vera, vera.cfg["rooms"], vera.cfg["types"],
                       vera.cfg.get("loan_rooms", []), vera.cfg.get("iphone_room", ""))
    d.text_rooms.delete("1.0", "end")
    d.text_rooms.insert("1.0", "\n".join(stanze))
    d._ok()
    esito = d.result
    if d.winfo_exists():
        d.destroy()
    return esito

# togliendo la stanza con prestito, la finestra lo dichiara invece di rifiutare
esito = salva_stanze([BAU, DR])
assert esito is not None, "il salvataggio non deve piu' essere rifiutato"
assert esito["prestiti_spariti"] == [KIOSK], esito["prestiti_spariti"]
assert esito["loan_rooms"] == [], esito["loan_rooms"]
assert esito["iphone_sparita"] == "", esito

# togliendo quella degli iPhone, lo stesso
esito = salva_stanze([KIOSK, DR])
assert esito["iphone_sparita"] == BAU, esito
assert esito["iphone_room"] == "", "la destinazione la chiede on_settings"
assert esito["prestiti_spariti"] == [], esito

# tolte tutte e due, si segnalano tutte e due
esito = salva_stanze([DR])
assert esito["iphone_sparita"] == BAU and esito["prestiti_spariti"] == [KIOSK], esito
vera.destroy()

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

# ---- il buco chiuso: togliendo la stanza dai DUE riquadri in un colpo solo,
# i dispositivi in prestito sarebbero finiti nel cestino, cioe' eliminati.
# Adesso il programma chiede lo stesso dove spostarli.
terza = App(fixture.build())
terza._initial_load()
terza.update()
domande_terza = []
def rispondi_terza(titolo, prompt, stanze):
    domande_terza.append((titolo, prompt, list(stanze)))
    return DR
terza._scegli_stanza = rispondi_terza
# la stanza sparisce dall'elenco E dalle stanze con prestito insieme
ui.RoomsDialog.show = lambda self: {
    "rooms": [BAU, DR], "types": terza.cfg["types"],
    "loan_rooms": [], "prestiti_spariti": [],
    "iphone_room": BAU, "iphone_sparita": "",
    "rinomine": [], "rinomine_tipi": [], "lingua": "it"}
terza.on_settings()
terza.store.load()
assert domande_terza, "un prestito aperto non puo' finire nel cestino in silenzio"
assert domande_terza[-1][0] == "Dove vanno i dispositivi in prestito?", domande_terza[-1][0]
prestati = [i for i in terza.store.items if i.get("prestato_a")]
assert prestati and all(i["stanza"] == DR for i in prestati), prestati
assert not [v for v in terza.store.eliminati() if v.get("scheda", {}).get("prestato_a")], \
    "nessun dispositivo in prestito deve essere finito fra gli eliminati"

# ---- e l'archivio si rifiuta comunque, chiunque glielo chieda
from inventario.store import InventoryError
quarta = InventoryStore(fixture.build(), iphone_room=BAU)
quarta.load()
try:
    quarta.porta_via_gli_orfani([KIOSK])
    raise SystemExit("ha eliminato una stanza con un prestito dentro")
except InventoryError as exc:
    assert "in prestito" in str(exc), str(exc)
quarta.load()
assert len([i for i in quarta.items if i["stanza"] == KIOSK]) == 5, "niente e' stato toccato"
# senza prestiti aperti, invece, la stanza si svuota come prima
quarta.give_back("IT-0107"); quarta.give_back("IT-0110")
portati = quarta.porta_via_gli_orfani([KIOSK])
assert len(portati) == 5, portati
terza.destroy()

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
