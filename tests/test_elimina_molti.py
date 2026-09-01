"""Eliminazione in blocco: si legge prima che cosa sparisce, poi si conferma.

Il rischio di un'eliminazione di massa non e' tecnico, e' che chi la lancia non
sappia esattamente che cosa sta togliendo. Per questo l'anteprima dice quali
dispositivi, da quale stanza, e che cosa viene saltato e perche'.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario.store import InventoryStore, new_item
from inventario.ui import PAROLA_ELIMINA, App, EliminaPlusDialog

percorso = fixture.build()
app = App(percorso)
app._initial_load()
store = app.store

# un iPhone mai spedito: non si elimina, e va detto
app._run(lambda: store.add(new_item(tipo=fixture.TIPO_IPHONE, modello="Apple iPhone 14",
                                    imei="356938035643809", restituito_da="M. B.")), "ok")

# ---- l'anteprima divide fra quello che si elimina, quello che non c'e' e
# quello che non si puo' toccare
codici = ["IT-0101",                       # esiste
          "IT-0107\tLaptop\tThinkPad",     # riga intera incollata da Excel
          "IT-0101",                       # doppione: conta una volta sola
          "IT-9999",                       # non esiste
          "356938035643809",               # iPhone mai spedito
          "   ",                           # riga vuota: ignorata
          ]
da_eliminare, non_trovati, bloccati = store.anteprima_eliminazione(codici)
assert [i["asset_tag"] for i in da_eliminare] == ["IT-0101", "IT-0107"], da_eliminare
assert non_trovati == ["IT-9999"], non_trovati
assert [i["asset_tag"] for i, _m in bloccati] == ["356938035643809"], bloccati
assert "non ancora rispedito" in bloccati[0][1], bloccati[0][1]

# ---- il riepilogo mostrato dice le stanze, i prestiti e i motivi
dlg = EliminaPlusDialog(app, store)
dlg.testo.insert("1.0", "\n".join(codici))
dlg._controlla()
testo = dlg.riepilogo.get("1.0", "end")
assert "VERRANNO ELIMINATI: 2" in testo, testo
assert fixture.BAU in testo and fixture.KIOSK in testo, testo
assert "IT-0101" in testo and "IT-0107" in testo
assert "in prestito a Marco Bianchi" in testo, "un prestito in corso va segnalato"
assert "non ancora rispedito" in testo
assert "IT-9999" in testo

# ---- senza la parola di conferma non si elimina niente
dlg.var_conferma.set("si")
dlg._ok()
assert dlg.result is None and dlg.winfo_exists()
dlg.var_conferma.set(PAROLA_ELIMINA.lower())     # anche minuscolo va bene
dlg._ok()
assert dlg.result == ["IT-0101", "IT-0107"], dlg.result

# ---- e se non c'e' niente da eliminare non si puo' nemmeno provare
vuoto = EliminaPlusDialog(app, store)
vuoto.testo.insert("1.0", "IT-0000\nIT-0001")
vuoto._controlla()
assert "Non c'e' niente da eliminare." in vuoto.riepilogo.get("1.0", "end")
assert str(vuoto.btn_elimina.cget("state")) == "disabled"
vuoto._ok()
assert vuoto.result is None
vuoto.destroy()

# ---- l'eliminazione vera toglie solo quelli scelti
prima = len(store.items)
store.delete(["IT-0101", "IT-0107"])
store.load()
assert len(store.items) == prima - 2
assert not [i for i in store.items if i["asset_tag"] in ("IT-0101", "IT-0107")]
assert [i for i in store.items if i["asset_tag"] == "356938035643809"], \
    "l'iPhone non si tocca"

app.destroy()
print("ELIMINA MOLTI OK")
