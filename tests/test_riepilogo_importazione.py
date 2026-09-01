"""Prima di importare si vede dove finiscono i dispositivi e quanti saranno.

Due numeri - "12 aggiunti, 3 aggiornati" - non dicono in quale stanza finiscono
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

# ---- unione: si contano aggiunte e aggiornamenti stanza per stanza
a = store.anteprima_importazione(nuovi, "merge")
assert a["per_stanza"][BAU] == {"aggiunti": 0, "aggiornati": 1}, a["per_stanza"]
assert a["per_stanza"][KIOSK] == {"aggiunti": 1, "aggiornati": 0}
assert a["per_stanza"][DR] == {"aggiunti": 2, "aggiornati": 0}
assert (a["aggiunti"], a["aggiornati"]) == (3, 1)
assert a["prima"] == 13 and a["dopo"] == 16, a
assert a["eliminati"] == 0

# ---- sostituzione: si dice anche quanti spariscono prima
a = store.anteprima_importazione(nuovi, "replace")
assert a["eliminati"] == 13, a["eliminati"]
assert a["dopo"] == 3, a          # 13 - 13 + 3 nuovi

# ---- gli iPhone non si contano fra gli eliminati: non si toccano mai
app._run(lambda: store.add(new_item(tipo=fixture.TIPO_IPHONE, modello="Apple iPhone 14",
                                    imei="356938035643809", restituito_da="M. B.")), "ok")
a = store.anteprima_importazione(nuovi, "replace")
assert a["prima"] == 14 and a["eliminati"] == 13, a
assert a["dopo"] == 4, a          # 14 - 13 + 3

# ---- importando in una sola stanza, tutto finisce li'
a = store.anteprima_importazione(nuovi, "merge", KIOSK)
assert list(a["per_stanza"]) == [KIOSK], a["per_stanza"]
assert a["per_stanza"][KIOSK] == {"aggiunti": 3, "aggiornati": 1}

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
assert "1 aggiornati" in dettaglio and "2 nuovi" in dettaglio, dettaglio
assert "In inventario adesso: 14" in dettaglio, dettaglio
assert "Dopo l'importazione: 17" in dettaglio, dettaglio
dlg.destroy()

app.destroy()
print("RIEPILOGO IMPORTAZIONE OK")
