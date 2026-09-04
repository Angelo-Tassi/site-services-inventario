"""Prima di importare si vede dove finiscono i dispositivi e quanti saranno.

Due numeri - "12 aggiunti, 3 saltati" - non dicono in quale stanza finiscono
ne' se l'inventario raddoppia. Chi sta per riscrivere l'inventario di tutti ha
il diritto di vederlo prima di premere.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario.store import InventoryStore, new_item
from inventario.ui import App, ImportDialog

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
app = App(fixture.build())
app._initial_load()
store = app.store

nuovi = [new_item("IT-0101", "Laptop", "T14 Gen 5", "PF1", BAU, ""),      # esiste
         new_item("IT-7777", "Tablet", "Dell 7320", "8H7", KIOSK, ""),   # nuovo
         new_item("IT-7778", "Laptop", "T14 Gen 4", "PF8", DR, ""),      # nuovo
         new_item("IT-7779", "Laptop", "T14 Gen 4", "PF9", DR, "")]      # nuovo

# ---- unione: si contano le aggiunte e i salti stanza per stanza. Un asset tag
# gia' in inventario non si importa: la riga si salta e si dice dov'e' quello
# che c'e' gia', invece di riscrivergli la scheda sotto senza dirlo
a = store.anteprima_importazione(nuovi, "merge")
assert a["per_stanza"][BAU] == {"aggiunti": 0, "saltati": 1}, a["per_stanza"]
assert a["per_stanza"][KIOSK] == {"aggiunti": 1, "saltati": 0}
assert a["per_stanza"][DR] == {"aggiunti": 2, "saltati": 0}
assert (a["aggiunti"], a["saltati"]) == (3, 1)
assert a["gia_presenti"] == [{"asset_tag": "IT-0101", "stanza": BAU}], a["gia_presenti"]
assert a["prima"] == 13 and a["dopo"] == 16, a
assert a["eliminati"] == 0

# ---- sostituzione: si dice anche quanti spariscono prima.
# In sostituzione il confronto si fa con quello che sopravvive alla pulizia:
# IT-0101 viene prima cancellato, quindi la sua riga entra come nuova. Prima
# l'anteprima lo contava come aggiornamento e prometteva un dispositivo in meno
# di quelli che poi si trovavano davvero in inventario.
a = store.anteprima_importazione(nuovi, "replace")
assert a["eliminati"] == 13, a["eliminati"]
assert a["saltati"] == 0, a["saltati"]
assert a["dopo"] == 4, a          # 13 - 13 + 4 righe, IT-0101 compresa

# ---- gli iPhone non si contano fra gli eliminati: non si toccano mai
app._run(lambda: store.add(new_item(tipo=fixture.TIPO_IPHONE, modello="Apple iPhone 14",
                                    imei="356938035643809", restituito_da="M. B.")), "ok")
a = store.anteprima_importazione(nuovi, "replace")
assert a["prima"] == 14 and a["eliminati"] == 13, a
assert a["dopo"] == 5, a          # 14 - 13 + 4: l'iPhone resta, piu' le 4 righe

# ---- importando in una sola stanza, tutto finisce li'
a = store.anteprima_importazione(nuovi, "merge", KIOSK)
assert list(a["per_stanza"]) == [KIOSK], a["per_stanza"]
assert a["per_stanza"][KIOSK] == {"aggiunti": 3, "saltati": 1}

# ---- le righe senza identificativo si contano a parte
a = store.anteprima_importazione(nuovi + [new_item("", "Laptop", "T14", "PFX", DR, "")],
                                 "merge")
assert a["senza_identificativo"] == 1, a

# ---- e il riepilogo mostrato dice tutte queste cose
anteprima = store.anteprima_importazione(nuovi, "merge")
dlg = ImportDialog(app, "prova.xlsx", len(nuovi), {"da_tag": 4},
                   {"stanza": None, "mode": "merge"}, 0, anteprima)
testi = []
def raccogli(w):
    if w.winfo_class() == "Text":
        testi.append(w.get("1.0", "end"))
    for c in w.winfo_children():
        raccogli(c)
raccogli(dlg)
dettaglio = "\n".join(testi)
assert "Dove finiscono:" in dettaglio, dettaglio
assert BAU in dettaglio and KIOSK in dettaglio and DR in dettaglio
assert "2 nuovi" in dettaglio, dettaglio
assert "1 gia' in inventario" in dettaglio, dettaglio
assert "non vengono importate" in dettaglio, dettaglio
assert "In inventario adesso: 14" in dettaglio, dettaglio
assert "Dopo l'importazione: 17" in dettaglio, dettaglio
dlg.destroy()

app.destroy()
print("RIEPILOGO IMPORTAZIONE OK")
