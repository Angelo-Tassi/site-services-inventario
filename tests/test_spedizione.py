import os, sys
from datetime import datetime, timedelta
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import messagebox
from inventario import theme
from inventario.store import (BloccoConservazione, DA_RISPEDIRE, InventoryError,
                              MESI_CONSERVAZIONE, SPEDITO, eliminabile_dal, new_item,
                              puo_essere_eliminato, testo_spedizione)
from inventario.ui import ACTION_COLUMN, App

BAU, KIOSK = fixture.BAU, fixture.KIOSK
app = App(fixture.build()); app._initial_load()
TIPO = app.iphone_type()
avvisi = []
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))

for imei, chi in (("356938035643809", "M. Bianchi"), ("351234567890123", "E. Rossi")):
    app._run(lambda i=imei, c=chi: app.store.add(
        new_item(tipo=TIPO, modello="Apple iPhone 14", imei=i, restituito_da=c)), "ok")

# --- il pulsante SPEDITO esiste solo nel contenitore iPhone
app.show_home()
assert ACTION_COLUMN not in app._columns()
app.show_room(KIOSK)
assert app.tree.heading(ACTION_COLUMN)["text"] == "Prestito"
assert app.action_label(app._item_by_tag("IT-0106")) == "Presta"
app.show_iphones()
assert app.ship_column_visible()
assert app.tree.heading(ACTION_COLUMN)["text"] == "Spedizione", app.tree.heading(ACTION_COLUMN)
assert app.action_label(app._item_by_tag("356938035643809")) == "SPEDITO"

# --- prima della spedizione: verde, Da Rispedire, eliminabile
tel = app._item_by_tag("356938035643809")
assert tel["stato"] == DA_RISPEDIRE and tel["spedito_il"] == ""
assert app.tree.item("356938035643809", "tags")[0].replace("_alt", "") == "iphone"
assert puo_essere_eliminato(tel)[0] is True

# --- spedizione
testo = app._run(lambda: app.store.ship("356938035643809"), "ok")
tel = app._item_by_tag("356938035643809")
assert tel["stato"] == SPEDITO, tel["stato"]
assert tel["spedito_il"], tel
assert "servizio telefonia" in testo and "consultazione" in testo, testo
sblocco = eliminabile_dal(tel)
atteso = MESI_CONSERVAZIONE
assert sblocco.strftime("%d/%m/%Y") in testo, testo
print("frase:", testo)

# --- riga viola, e niente piu' pulsante su quella riga
assert app.tree.item("356938035643809", "tags")[0].replace("_alt", "") == "spedito", \
    app.tree.item("356938035643809", "tags")
assert app.action_label(tel) == ""
assert app.action_label(app._item_by_tag("351234567890123")) == "SPEDITO"   # l'altro no
assert str(app.tree.tag_configure("spedito", "background")) == theme.SHIP_ROW

# --- resta visibile ovunque
app.show_home()
assert "356938035643809" in app.tree.get_children()
app.show_room(BAU)
assert "356938035643809" in app.tree.get_children()
app.show_iphones()
assert "356938035643809" in app.tree.get_children()

# --- non si spedisce due volte
try:
    app.store.ship("356938035643809"); raise SystemExit("doppia spedizione accettata")
except InventoryError as e:
    print("doppia spedizione:", e)
try:
    app.store.ship("IT-0101"); raise SystemExit("laptop spedito")
except InventoryError as e:
    print("solo iPhone:", e)

# --- eliminazione bloccata per tre mesi, con la data nel popup
try:
    app.store.delete(["356938035643809"]); raise SystemExit("eliminato prima del tempo")
except BloccoConservazione as e:
    assert e.sblocco == sblocco
    print("blocco:", str(e).replace(chr(10), " "))
avvisi.clear()
app.tree.selection_set(["356938035643809"])
app.on_delete()
assert avvisi and avvisi[-1][0] == "Eliminazione non consentita", avvisi
assert sblocco.strftime("%d/%m/%Y") in avvisi[-1][1], avvisi[-1][1]
assert app._item_by_tag("356938035643809") is not None      # non e' stato toccato

# --- passati i tre mesi si sblocca
scaduto = dict(app._item_by_tag("356938035643809"))
scaduto["spedito_il"] = (datetime.now() - timedelta(days=95)).strftime("%d/%m/%Y %H:%M")
libero, quando = puo_essere_eliminato(scaduto)
assert libero is True, (libero, quando)
app._run(lambda: app.store.update("356938035643809", scaduto), "ok")
assert app._run(lambda: app.store.delete(["356938035643809"]), "ok") == 1
assert app._item_by_tag("356938035643809") is None

# --- un dispositivo mai spedito si elimina sempre
assert puo_essere_eliminato(app._item_by_tag("IT-0101"))[0] is True
app.destroy()
print("SPEDIZIONE OK")
