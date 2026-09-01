"""Nessuna colonna taglia quello che contiene.

Una colonna che nasconde il testo costringe ad allargarla a mano ogni volta, e
intanto quello che nasconde non si sa. Una colonna vuota deve comunque mostrare
il proprio nome, altrimenti non si capisce nemmeno che cosa conterrebbe.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario.lingua import intestazione
from inventario.store import HEADERS, InventoryStore, new_item, valore_visibile
from inventario.ui import ACTION_COLUMN, CHECK_COLUMN, App

app = App(fixture.build())
app._initial_load()
# un modello lunghissimo, per vedere se la colonna lo segue
app._run(lambda: app.store.add(new_item(
    "IT-8000", "Laptop",
    "Lenovo ThinkPad T14 Gen 5 - configurazione speciale reparto grafica",
    "PF8000", fixture.BAU, "")), "ok")
app.geometry("1250x640")
app.update()

def verifica(dove):
    base = app._font_misura["base"]
    grassetto = app._font_misura["bold"]
    for campo in app._columns():
        largo = int(app.tree.column(campo, "width"))
        if campo in (CHECK_COLUMN, ACTION_COLUMN):
            continue
        titolo = intestazione(HEADERS[campo])
        assert largo >= grassetto.measure(titolo), \
            "%s / %s: %d px per un titolo di %d" % (dove, campo, largo,
                                                    grassetto.measure(titolo))
        for item in app.visible:
            testo = valore_visibile(item, campo)
            assert largo >= base.measure(testo), \
                "%s / %s: %d px per %r" % (dove, campo, largo, testo)

for dove, vai in (("home", app.show_home),
                  (fixture.BAU, lambda: app.show_room(fixture.BAU)),
                  (fixture.KIOSK, lambda: app.show_room(fixture.KIOSK)),
                  ("iphone", app.show_iphones)):
    vai(); app.update(); verifica(dove)

# ---- la colonna si riadatta quando il contenuto cambia
app.show_home(); app.update()
prima = int(app.tree.column("note", "width"))
app._run(lambda: app.store.set_campo(
    "IT-8000", "note", "Nota molto lunga scritta apposta per allargare la colonna"), "ok")
app.refresh_table(); app.update()
assert int(app.tree.column("note", "width")) > prima, (prima,
                                                       app.tree.column("note", "width"))

# ---- il modello si modifica dall'elenco, come le note
assert "modello" in InventoryStore.CAMPI_AL_VOLO
app.store.set_campo("IT-8000", "modello", "Dell Latitude 5450")
assert app._item_by_tag("IT-8000") is not None
app.on_refresh()
assert [i for i in app.store.items if i["asset_tag"] == "IT-8000"][0]["modello"] \
    == "Dell Latitude 5450"

# ---- e non si modificano dall'elenco campi che non lo prevedono
try:
    app.store.set_campo("IT-8000", "asset_tag", "IT-9999")
    raise AssertionError("doveva rifiutare")
except Exception as exc:
    assert "non si modifica" in str(exc), exc

# ---- anche il tipo si cambia dall'elenco, con la tendina dei tipi configurati
app.show_home(); app.update()
prima = app._item_by_tag("IT-8000")["tipo"]
assert prima == "Laptop", prima
app._run(lambda: app.store.set_tipo("IT-8000", "Tablet"), "ok")
assert app._item_by_tag("IT-8000")["tipo"] == "Tablet"

# ---- ma un iPhone non diventa un laptop: si perderebbe l'identificativo
telefono = new_item(tipo=fixture.TIPO_IPHONE, modello="Apple iPhone 14",
                    imei="356938035643809", restituito_da="M. B.")
app._run(lambda: app.store.add(telefono), "ok")
try:
    app.store.set_tipo("356938035643809", "Laptop")
    raise AssertionError("doveva rifiutare")
except Exception as exc:
    assert "iPhone" in str(exc), exc
assert app._item_by_tag("356938035643809")["imei"] == "356938035643809"
# e nemmeno il contrario
try:
    app.store.set_tipo("IT-8000", fixture.TIPO_IPHONE)
    raise AssertionError("doveva rifiutare")
except Exception as exc:
    assert "iPhone" in str(exc), exc
assert app._item_by_tag("IT-8000")["asset_tag"] == "IT-8000"

# ---- l'intestazione del modello dice anche che ci si scrive una descrizione
assert HEADERS["modello"] == "Modello/Descrizione"
assert intestazione(HEADERS["modello"]) == "Modello/Descrizione"

app.destroy()
print("LARGHEZZA COLONNE OK")
