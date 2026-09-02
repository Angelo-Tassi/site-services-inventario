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

# ---- al quinto record il programma lo chiede
app._run(lambda: store.delete(["IT-9001", "IT-0101", "IT-0102"]))
assert store.modifiche == 5, store.modifiche
assert len(chieste) == 1, chieste
assert "5 dispositivi" in chieste[0], chieste[0]

# ---- e non lo richiede a ogni modifica successiva: riparte da capo
app._run(lambda: store.set_campo("IT-0103", "note", "una"))
assert len(chieste) == 1, "non deve insistere"
app._run(lambda: store.delete(["IT-0104", "IT-0105", "IT-0106", "IT-0107"]))
assert len(chieste) == 2, chieste
assert "5 dispositivi" in chieste[1], chieste[1]

# ---- eliminandone parecchi in blocco lo chiede subito: contano i record
prima = len(chieste)
molti = [i["asset_tag"] for i in store.items
         if not i["asset_tag"].startswith("35")][:MODIFICHE_PER_PROMEMORIA]
assert len(molti) == MODIFICHE_PER_PROMEMORIA, molti
app._run(lambda: store.delete(molti))
assert len(chieste) == prima + 1, chieste

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
