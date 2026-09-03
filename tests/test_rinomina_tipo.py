"""Rinominare un tipo di dispositivo aggiorna i dispositivi di quel tipo.

Cambia l'etichetta e basta: asset tag, numero di serie, stanza, stato, note,
prestito e spedizione restano quelli che erano.

Il tipo iPhone fa eccezione e non si rinomina. Non e' un'etichetta come le
altre: e' la parola con cui il programma riconosce i telefoni, e da li' vengono
l'IMEI al posto dell'asset tag, la stanza obbligata, la spedizione al servizio
telefonia e il fatto che non si eliminino.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import messagebox
from inventario import config, ui
from inventario.store import (InventoryError, InventoryStore, is_iphone, new_item,
                              rinomina_tocca_gli_iphone, rinomine_tipi)
from inventario.ui import App, RoomsDialog

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
IPHONE = fixture.TIPO_IPHONE

# ---- si riconosce una rinomina come per le stanze
assert rinomine_tipi(["Laptop", "Tablet"], ["Notebook", "Tablet"]) == [("Laptop", "Notebook")]
assert rinomine_tipi(["Laptop", "Tablet"], ["Tablet", "Laptop"]) == []
assert rinomine_tipi(["Laptop"], ["Laptop", "Tablet"]) == []

# ---- e si riconosce il tipo che non si tocca, da tutti e due i lati
assert rinomina_tocca_gli_iphone([("Iphone", "Telefono")]) == ("Iphone", "Telefono")
assert rinomina_tocca_gli_iphone([("Telefono", "iphone")]) == ("Telefono", "iphone")
assert rinomina_tocca_gli_iphone([("Laptop", "Notebook")]) is None

percorso = fixture.build()
store = InventoryStore(percorso, iphone_room=BAU)
store.load()
store.add(new_item(tipo=IPHONE, modello="Apple iPhone 14", imei="356938035643809",
                   restituito_da="M. B."))
store.load()

# ---- niente si perde: si guarda un dispositivo prima e dopo
prima = dict([i for i in store.items if i["asset_tag"] == "IT-0107"][0])
assert prima["tipo"] == "Laptop" and prima["prestato_a"] == "Marco Bianchi"

cambiati = store.rinomina_tipi([("Laptop", "Notebook")])
assert cambiati == {"Notebook": 8}, cambiati
store.load()
dopo = [i for i in store.items if i["asset_tag"] == "IT-0107"][0]
assert dopo["tipo"] == "Notebook", dopo["tipo"]
for campo in ("asset_tag", "seriale", "stanza", "stato", "note",
              "prestato_a", "prestato_il", "modello", "imei", "spedito_il"):
    assert dopo[campo] == prima[campo], (campo, prima[campo], dopo[campo])
assert not [i for i in store.items if i["tipo"] == "Laptop"], "nessuno resta indietro"

# ---- i telefoni non sono stati toccati
telefoni = [i for i in store.items if is_iphone(i["tipo"])]
assert len(telefoni) == 1 and telefoni[0]["imei"] == "356938035643809"

# ---- il tipo iPhone non si rinomina, nemmeno chiamando l'archivio
try:
    store.rinomina_tipi([(IPHONE, "Telefono")])
    raise SystemExit("ha rinominato il tipo iPhone")
except InventoryError as exc:
    assert "non si rinomina" in str(exc), str(exc)
store.load()
assert [i for i in store.items if is_iphone(i["tipo"])], "i telefoni sono ancora telefoni"

# ---- e nemmeno al contrario: un altro tipo non diventa "Iphone"
try:
    store.rinomina_tipi([("Tablet", "Iphone")])
    raise SystemExit("ha creato iPhone per rinomina")
except InventoryError:
    pass

# ============================ dalle impostazioni ============================
app = App(fixture.build())
app._initial_load()
app.update()
avvisi = []
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))

def impostazioni(tipi, stanze=None):
    dlg = RoomsDialog(app, app.cfg["rooms"], app.cfg["types"],
                      app.cfg.get("loan_rooms", []), app.cfg.get("iphone_room", ""))
    dlg.text_types.delete("1.0", "end")
    dlg.text_types.insert("1.0", "\n".join(tipi))
    if stanze is not None:
        dlg.text_rooms.delete("1.0", "end")
        dlg.text_rooms.insert("1.0", "\n".join(stanze))
    dlg._ok()
    risultato = dlg.result
    if dlg.winfo_exists():
        dlg.destroy()
    return risultato

# ---- rinominando Laptop, i dispositivi seguono
avvisi.clear()
ui.RoomsDialog.show = lambda self: impostazioni(["Notebook", "Tablet", IPHONE])
app.on_settings()
app.store.load()
assert app.cfg["types"] == ["Notebook", "Tablet", IPHONE], app.cfg["types"]
assert len([i for i in app.store.items if i["tipo"] == "Notebook"]) == 8
assert not [i for i in app.store.items if i["tipo"] == "Laptop"]
titolo, corpo = avvisi[-1]
assert titolo == "Rinomina completata", avvisi[-1]
assert "TIPI DI DISPOSITIVO:" in corpo and "Laptop  ->  Notebook   (8 dispositivi)" in corpo, corpo

# ---- rinominare il tipo iPhone viene rifiutato, e non si salva niente
avvisi.clear()
tipi_prima = list(app.cfg["types"])
assert impostazioni(["Notebook", "Tablet", "Telefono"]) is None
assert "non si rinomina" in avvisi[-1][0] or "non si rinomina" in avvisi[-1][1], avvisi[-1]
assert app.cfg["types"] == tipi_prima, "le impostazioni non devono cambiare"

# ---- due tipi con lo stesso nome vengono rifiutati
avvisi.clear()
assert impostazioni(["Notebook", "Notebook", IPHONE]) is None
assert "compare due volte" in avvisi[-1][1], avvisi[-1]

# ---- stanza e tipo rinominati insieme: un solo riepilogo, due sezioni
avvisi.clear()
ui.RoomsDialog.show = lambda self: impostazioni(["Portatile", "Tablet", IPHONE],
                                                stanze=[BAU, "Kiosk 2", DR])
app.on_settings()
app.store.load()
corpo = avvisi[-1][1]
assert "STANZE:" in corpo and "TIPI DI DISPOSITIVO:" in corpo, corpo
assert "Digital Kiosk  ->  Kiosk 2   (5 dispositivi)" in corpo, corpo
assert "Notebook  ->  Portatile   (8 dispositivi)" in corpo, corpo
assert len([i for i in app.store.items if i["tipo"] == "Portatile"]) == 8
assert len([i for i in app.store.items if i["stanza"] == "Kiosk 2"]) == 5

app.destroy()
print("RINOMINA TIPO OK")
