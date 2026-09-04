"""Un dispositivo senza stanza non entra in inventario. Mai.

Un dispositivo senza stanza non comparirebbe in nessuna stanza, non uscirebbe
da nessuna esportazione per stanza e da nessuna stampa: per ritrovarlo
bisognerebbe gia' sapere che c'e'. Quindi non si importa.

Quando il foglio non lo dice - non ha separatori, o ha righe sopra il primo -
la stanza si chiede prima di importare, in due modi: tutti nella stessa stanza,
oppure uno per uno. Chi non ne assegna nessuna vede quelle righe restare fuori,
dichiarate nel riepilogo.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import messagebox
from inventario import ui
from inventario.store import InventoryStore, new_item
from inventario.ui import App, StanzaDelDispositivoDialog, StanzeDaAssegnareDialog

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR

def senza_stanza(tag, modello="T14"):
    riga = new_item(tag, "Laptop", modello, "PF" + tag[-3:], "", "")
    riga["stanza"] = ""
    return riga

# ======================= l'archivio non li fa entrare =======================
store = InventoryStore(fixture.build(), iphone_room=BAU)
store.load()
quanti = len(store.items)

esito = store.import_items([senza_stanza("IT-8001"),
                            new_item("IT-8002", "Laptop", "T14", "PF2", DR)], "merge")
assert esito["aggiunti"] == 1, esito
assert esito["senza_stanza"] == ["IT-8001"], esito["senza_stanza"]
store.load()
assert len(store.items) == quanti + 1
assert not [i for i in store.items if i["asset_tag"] == "IT-8001"], "non doveva entrare"
assert not [i for i in store.items if not i.get("stanza")], "nessuno senza stanza"

# ---- nemmeno in sostituzione, dove il resto viene cancellato prima
esito = store.import_items([senza_stanza("IT-8003")], "replace")
assert esito["aggiunti"] == 0 and esito["senza_stanza"] == ["IT-8003"], esito
store.load()
assert not [i for i in store.items if not i.get("stanza")]

# ---- importando dentro una stanza, invece, la stanza ce l'hanno tutti
store = InventoryStore(fixture.build(), iphone_room=BAU)
store.load()
esito = store.import_items([senza_stanza("IT-8004"), senza_stanza("IT-8005")],
                           "merge", KIOSK)
assert esito["aggiunti"] == 2 and esito["senza_stanza"] == [], esito
store.load()
assert all(i["stanza"] == KIOSK
           for i in store.items if i["asset_tag"] in ("IT-8004", "IT-8005"))

# ---- e l'anteprima lo dice prima che sia scritto qualcosa
a = store.anteprima_importazione(
    [senza_stanza("IT-8006"), new_item("IT-8007", "Laptop", "T14", "PF7", DR)], "merge")
assert a["senza_stanza"] == 1, a
assert a["aggiunti"] == 1 and a["dopo"] == a["prima"] + 1, a
assert list(a["per_stanza"]) == [DR], a["per_stanza"]

# ============ una stanza che non esiste vale come nessuna stanza ============
# "Cantina", o "Digital  Kiosk" con due spazi, non e' una stanza: il
# dispositivo non comparirebbe in nessuna scheda, e per ritrovarlo bisognerebbe
# gia' sapere che c'e'.
store = InventoryStore(fixture.build(), iphone_room=BAU)
store.stanze = [BAU, KIOSK, DR]
store.load()
assert store.stanza_ammessa(KIOSK) and not store.stanza_ammessa("Cantina")
assert not store.stanza_ammessa("")
# scritta male ma riconoscibile e' la stessa stanza, e entra col nome ufficiale
assert store.stanza_canonica("digital  kiosk") == KIOSK, store.stanza_canonica("digital  kiosk")
assert store.stanza_canonica("Cantina") == ""

fantasma = new_item("IT-8400", "Laptop", "T14", "PF4", "Cantina di Zio Bob")
buono = new_item("IT-8401", "Laptop", "T14", "PF1", DR)
a = store.anteprima_importazione([fantasma, buono], "merge")
assert a["senza_stanza"] == 1 and a["aggiunti"] == 1, a
esito = store.import_items([fantasma, buono], "merge")
assert esito["aggiunti"] == 1 and esito["senza_stanza"] == ["IT-8400"], esito
store.load()
assert not [i for i in store.items if i["asset_tag"] == "IT-8400"], "non doveva entrare"
assert not [i for i in store.items if i["stanza"] not in (BAU, KIOSK, DR)], \
    "nessuno in una stanza che non esiste"

# ---- e chi la scrive male entra lo stesso, ma col nome ufficiale
store.import_items([new_item("IT-8402", "Laptop", "T14", "PF2", "digital  kiosk")],
                   "merge")
store.load()
messo = [i for i in store.items if i["asset_tag"] == "IT-8402"][0]
assert messo["stanza"] == KIOSK, messo["stanza"]

# ---- senza l'elenco delle stanze l'archivio non puo' saperlo, e non lo pretende
libero = InventoryStore(fixture.build(), iphone_room=BAU)
libero.load()
assert libero.stanze is None and libero.stanza_ammessa("Cantina")

# ======================= le due finestre, quelle vere =======================
app = App(fixture.build())
app._initial_load()
app.update()
avvisi = []
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))

d = StanzeDaAssegnareDialog(app, 3, [BAU, KIOSK, DR])
d.var_come.set("una"); d.var_stanza.set(DR); d._aggiorna(); d._ok()
assert d.result == {"come": "una", "stanza": DR}, d.result

d = StanzeDaAssegnareDialog(app, 3, [BAU, KIOSK, DR])
d.var_come.set("uno"); d._aggiorna()
assert str(d.combo.cget("state")) == "disabled", "la tendina non serve piu'"
d._ok()
assert d.result["come"] == "uno", d.result

riga = senza_stanza("IT-8100", "Lenovo ThinkPad T14")
riga["note"] = "trovato in magazzino"
d = StanzaDelDispositivoDialog(app, riga, [BAU, KIOSK, DR], 2, 5)
d.var_stanza.set(KIOSK); d._ok()
assert d.result == {"stanza": KIOSK}, d.result

d = StanzaDelDispositivoDialog(app, riga, [BAU, KIOSK, DR], 2, 5)
d._salta()
assert d.result == {"salta": True}, d.result

d = StanzaDelDispositivoDialog(app, riga, [BAU, KIOSK, DR], 2, 5)
d._cancel()
assert d.result is None, "annulla tutto vuol dire non importare niente"

# ======================= la domanda dentro l'importazione =======================
def rispondi(prima, per_uno=()):
    """Sostituisce le due finestre con le risposte gia' pronte."""
    risposte = list(per_uno)
    class FintaPrima:
        def __init__(self, *a, **k): pass
        def show(self): return prima
    class FintaSingola:
        def __init__(self, *a, **k): pass
        def show(self): return risposte.pop(0) if risposte else None
    ui.StanzeDaAssegnareDialog = FintaPrima
    ui.StanzaDelDispositivoDialog = FintaSingola

def righe():
    return [senza_stanza("IT-8201"), senza_stanza("IT-8202"),
            new_item("IT-8203", "Laptop", "T14", "PF3", DR)]

def con_fantasma():
    return [new_item("IT-8210", "Laptop", "T14", "PF0", "Cantina di Zio Bob"),
            new_item("IT-8211", "Laptop", "T14", "PF1", DR)]

# ---- niente da chiedere se la stanza ce l'hanno tutti
rispondi(None)
intatte = [new_item("IT-8300", "Laptop", "T14", "PF0", DR)]
uscite, fuori = app._chiedi_le_stanze_mancanti(intatte)
assert uscite is intatte and fuori == 0, "non c'era niente da chiedere"

# ---- tutti nella stessa stanza
rispondi({"come": "una", "stanza": KIOSK})
uscite, fuori = app._chiedi_le_stanze_mancanti(righe())
assert fuori == 0 and len(uscite) == 3, uscite
assert [i["stanza"] for i in uscite] == [KIOSK, KIOSK, DR], \
    "solo gli orfani cambiano stanza"

# ---- uno per uno, ciascuno dove dico io
rispondi({"come": "uno"}, [{"stanza": BAU}, {"stanza": DR}])
uscite, fuori = app._chiedi_le_stanze_mancanti(righe())
assert fuori == 0 and [i["stanza"] for i in uscite] == [BAU, DR, DR], uscite

# ---- uno per uno, e uno lo lascio fuori: non entra, e si conta
rispondi({"come": "uno"}, [{"salta": True}, {"stanza": DR}])
uscite, fuori = app._chiedi_le_stanze_mancanti(righe())
assert fuori == 1, fuori
assert [i["asset_tag"] for i in uscite] == ["IT-8202", "IT-8203"], uscite
assert all(i["stanza"] for i in uscite), "chi resta ha la stanza"

# ---- e una stanza che non esiste si chiede come se non ci fosse
rispondi({"come": "una", "stanza": KIOSK})
uscite, fuori = app._chiedi_le_stanze_mancanti(con_fantasma())
assert fuori == 0 and [i["stanza"] for i in uscite] == [KIOSK, DR], uscite

# ---- annullando non si importa niente, nemmeno le righe gia' a posto
rispondi(None)
assert app._chiedi_le_stanze_mancanti(righe()) == (None, 0)
rispondi({"come": "uno"}, [{"stanza": BAU}, None])
assert app._chiedi_le_stanze_mancanti(righe()) == (None, 0), \
    "annullare a meta' strada annulla tutto"

# ---- senza nemmeno una stanza in cui metterli, lo dice e non importa
avvisi.clear()
senza_stanze = App(fixture.build())
senza_stanze._initial_load()
senza_stanze.cfg["rooms"] = []
rispondi({"come": "una", "stanza": KIOSK})
assert senza_stanze._chiedi_le_stanze_mancanti(righe()) == (None, 0)
assert avvisi and avvisi[-1][0] == "Nessuna stanza", avvisi
assert "2 dispositivi senza stanza" in avvisi[-1][1], avvisi[-1][1]
senza_stanze.destroy()

# ======================= e il riepilogo finale lo racconta =======================
risultato = {"aggiunti": 1, "eliminati": 0, "copia": None, "gia_presenti": [],
             "senza_stanza": ["IT-8201", "IT-8202"]}
testo = "\n".join(app._resoconto_importazione(
    risultato, {"non_assegnati": 3}, {"stanza": None, "mode": "merge"}, 0, []))
assert "NON IMPORTATI, senza stanza: 2" in testo, testo
assert "IT-8201" in testo and "IT-8202" in testo, testo
assert "3 senza stanza: non ne hai assegnata nessuna" in testo, testo

app.destroy()
print("STANZA OBBLIGATORIA OK")
