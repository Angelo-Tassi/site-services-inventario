"""Un dispositivo in prestito non si sposta e non si elimina.

Finche' e' nelle mani di una persona, l'inventario e' l'unica traccia di dove
sia finito e di chi ce l'abbia: spostarlo in un'altra stanza o cancellarlo
significherebbe perderlo. Prima si registra il rientro.

Nella configurazione i prestiti sono attivi sul Digital Kiosk, quindi e' li'
che la regola si vede all'opera.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import messagebox
from inventario import lingua as lang
from inventario import ui
from inventario.store import InventoryError, InventoryStore, is_on_loan, new_item
from inventario.ui import App, riepilogo_spostamento

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
percorso = fixture.build()
store = InventoryStore(percorso, iphone_room=BAU)
store.load()

# IT-0107 e IT-0110 sono in prestito nel Digital Kiosk, IT-0106 no
assert is_on_loan(store.items and [i for i in store.items
                                   if i["asset_tag"] == "IT-0107"][0])

# ---- l'archivio non si sposta e non si elimina, comunque lo si chiami
assert store.move_to_room(["IT-0107"], DR) == (0, 0, 1)
store.load()
assert [i for i in store.items if i["asset_tag"] == "IT-0107"][0]["stanza"] == KIOSK

try:
    store.delete(["IT-0107"])
    raise SystemExit("un prestito aperto e' stato eliminato")
except InventoryError as exc:
    assert "in prestito a Marco Bianchi" in str(exc), str(exc)
    assert "registra prima il rientro" in str(exc).lower(), str(exc)

# ---- store.delete e' tutto-o-niente: un prestito in mezzo non deve far
# sparire niente, nemmeno i liberi
prima = len(store.items)
try:
    store.delete(["IT-0106", "IT-0107"])
    raise SystemExit("eliminati insieme a un prestito")
except InventoryError:
    pass
store.load()
assert len(store.items) == prima, "nessuno doveva sparire"

# ---- l'anteprima lo salta e dice perche'
da_eliminare, _non, bloccati = store.anteprima_eliminazione(["IT-0106", "IT-0107"])
assert [i["asset_tag"] for i in da_eliminare] == ["IT-0106"], da_eliminare
assert [i["asset_tag"] for i, _m in bloccati] == ["IT-0107"], bloccati
assert bloccati[0][1] == "in prestito a Marco Bianchi: registra prima il rientro"

# ---- la scheda non e' la porta di servizio: cambiare stanza li' e' spostare
scheda = dict([i for i in store.items if i["asset_tag"] == "IT-0107"][0])
scheda["stanza"] = DR
try:
    store.update("IT-0107", scheda)
    raise SystemExit("prestito spostato dalla scheda")
except InventoryError as exc:
    assert "in prestito" in str(exc), str(exc)
# le altre modifiche restano possibili: si cambia solo la nota
scheda["stanza"] = KIOSK
scheda["note"] = "Consegnato con l'alimentatore"
store.update("IT-0107", scheda)
store.load()
assert [i for i in store.items
        if i["asset_tag"] == "IT-0107"][0]["note"] == "Consegnato con l'alimentatore"

# ---- dopo il rientro tutto torna possibile
store.give_back("IT-0107")
assert store.move_to_room(["IT-0107"], DR) == (1, 0, 0)
store.load()
assert [i for i in store.items if i["asset_tag"] == "IT-0107"][0]["stanza"] == DR

# ---- il motivo si legge anche in inglese: e' tradotto dove viene costruito,
# non dove viene mostrato
lang.imposta(lang.INGLESE)
_da, _non, bloccati = store.anteprima_eliminazione(["IT-0110"])
assert bloccati[0][1] == "on loan to Elena Rossi: register the return first", bloccati
lang.imposta(lang.ITALIANO)

# ============================ dall'interfaccia ============================
app = App(fixture.build())
app._initial_load()
app.update()
avvisi = []
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))
ui.ConfermaOperazioneDialog.show = lambda self: True
app._ask_room = lambda domanda: DR

app.show_home()

# ---- Sposta con dentro solo prestiti: non chiede nemmeno la stanza
app.tree.selection_set(["IT-0107", "IT-0110"]); app._on_select()
app.on_move()
assert "sono in prestito e non si spostano" in avvisi[-1][1], avvisi[-1]
app.store.load()
assert app._item_by_tag("IT-0107")["stanza"] == KIOSK

# ---- selezione mista: partono i liberi, i prestiti restano e si dice quanti
app.tree.selection_set(["IT-0106", "IT-0107"]); app._on_select()
app.on_move()
app.store.load()
assert app._item_by_tag("IT-0106")["stanza"] == DR
assert app._item_by_tag("IT-0107")["stanza"] == KIOSK
assert "in prestito lasciati dove sono" in app.var_status.get(), app.var_status.get()

# ---- Elimina su un solo prestito: il messaggio dice cosa fare
avvisi.clear()
app.tree.selection_set(["IT-0107"]); app._on_select()
app.on_delete()
titolo, corpo = avvisi[-1]
assert titolo == "Eliminazione non consentita", avvisi[-1]
assert "in prestito a Marco Bianchi" in corpo, corpo
assert "Registra rientro" in corpo, corpo
app.store.load()
assert app._item_by_tag("IT-0107") is not None

# ---- Elimina misto: sparisce il libero, il prestito resta
prima = len(app.store.items)
app.tree.selection_set(["IT-0108", "IT-0107"]); app._on_select()
app.on_delete()
app.store.load()
assert len(app.store.items) == prima - 1, len(app.store.items)
assert app._item_by_tag("IT-0107") is not None, "il prestito resta"
assert app._item_by_tag("IT-0108") is None, "il libero se ne va"

# ---- il riepilogo dello spostamento lo mette in chiaro
items = [app._item_by_tag("IT-0107")]
testo = "\n".join(riepilogo_spostamento([], [], items, DR, BAU, []))
assert "RESTANO FERMI perche' sono in prestito: 1" in testo, testo
assert "Registra prima il rientro" in testo, testo
assert "IT-0107" in testo, testo

app.destroy()
print("PRESTITI PROTETTI OK")
