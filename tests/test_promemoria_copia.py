"""Dopo un po' di modifiche il programma ricorda di portarsi via una copia.

Le copie automatiche stanno sulla share accanto ai dati: coprono l'errore umano
ma non la cartella di rete che sparisce. Quella copia la deve volere qualcuno, e
chi sta lavorando non ci pensa.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario import ui
from inventario.store import InventoryError, InventoryStore, new_item
from inventario.ui import MODIFICHE_PER_PROMEMORIA, App

app = App(fixture.build())
app._initial_load()
store = app.store

chieste = []
def finta_domanda(titolo, testo, **kwargs):
    chieste.append(testo)
    return False                      # l'utente dice di no
ui.messagebox.askyesno = finta_domanda
errori = []
ui.messagebox.showerror = lambda t, m, **k: errori.append((t, m))

# ---- il contatore segue i record, non le operazioni
assert store.modifiche == 0
app._run(lambda: store.add(new_item("IT-9001", "Laptop", "T14", "PF1", fixture.BAU, "")))
assert store.modifiche == 1, store.modifiche
assert not chieste, "troppo presto per chiedere"

app._run(lambda: store.set_campo("IT-9001", "note", "provata"))
assert store.modifiche == 2
app._run(lambda: store.set_campo("IT-9001", "note", "provata"))
assert store.modifiche == 2, "una modifica a vuoto non conta"

# ---- fino alla soglia non chiede niente
mancanti = MODIFICHE_PER_PROMEMORIA - store.modifiche
da_togliere = [i["asset_tag"] for i in store.items
               if i["asset_tag"] not in ("IT-9001",)
               and not i["asset_tag"].startswith("35")][:mancanti - 1]
app._run(lambda: store.delete(da_togliere))
assert store.modifiche == MODIFICHE_PER_PROMEMORIA - 1, store.modifiche
assert not chieste, "un record prima della soglia non si chiede ancora"

# ---- al record numero MODIFICHE_PER_PROMEMORIA lo chiede
app._run(lambda: store.delete(["IT-9001"]))
assert store.modifiche == MODIFICHE_PER_PROMEMORIA, store.modifiche
assert len(chieste) == 1, chieste
assert "%d dispositivi" % MODIFICHE_PER_PROMEMORIA in chieste[0], chieste[0]

# ---- e non lo richiede a ogni modifica successiva: riparte da capo
app._run(lambda: store.set_campo(store.items[0]["asset_tag"], "note", "una"))
assert len(chieste) == 1, "non deve insistere"

# ---- eliminandone parecchi in blocco lo chiede subito: contano i record, non
# le operazioni, e un'eliminazione da dieci li conta tutti e dieci
nuovi = ["IT-95%02d" % n for n in range(MODIFICHE_PER_PROMEMORIA)]
for tag in nuovi:
    store.add(new_item(tag, "Laptop", "T14", "PF" + tag, fixture.DR, ""))
store.load()
chieste[:] = []
app._modifiche_alla_copia = store.modifiche      # si riparte da zero
app._run(lambda: store.delete(nuovi))
assert len(chieste) == 1, chieste
assert "%d dispositivi" % MODIFICHE_PER_PROMEMORIA in chieste[0], chieste[0]

# ---- un asset tag gia' presente non entra, e il messaggio dice chi ce l'ha
esistente = store.items[0]
app._run(lambda: store.add(new_item(esistente["asset_tag"], "Laptop", "T14",
                                    "PFX", fixture.DR, "")))
assert errori, "doveva comparire l'errore"
titolo, messaggio = errori[-1]
assert esistente["asset_tag"] in messaggio
assert "Non e' stato inserito niente" in messaggio, messaggio
assert esistente["stanza"] in messaggio, messaggio
assert sum(1 for i in store.items if i["asset_tag"] == esistente["asset_tag"]) == 1

app.destroy()
print("PROMEMORIA COPIA OK")
