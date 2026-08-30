"""I quattro punti segnalati: tipo prima di tutto, ritorno alla home,
ordine cronologico, iPhone senza seriale ne' prestiti."""
import os, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario.ui import (ACTION_COLUMN, AddChoiceDialog, App, COLONNE_NON_IPHONE,
                           TypeChoiceDialog, chiave_ordinamento)
from inventario.store import new_item

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
app = App(fixture.build()); app._initial_load()
TIPO = app.iphone_type()

# ---------------------------------------------- 1. prima si sceglie il tipo
scelta = TypeChoiceDialog(app, app.cfg["types"], app.tipo_predefinito())
assert list(scelta.winfo_children()[0].winfo_children()[2].cget("values")) == app.cfg["types"]
assert scelta.var_tipo.get() == "Laptop", scelta.var_tipo.get()
scelta.var_tipo.set(TIPO); scelta._ok()
assert scelta.result == TIPO
# nel contenitore iPhone la tendina propone gia' l'iPhone
app.show_iphones()
assert app.tipo_predefinito() == TIPO
app.show_home()
assert app.tipo_predefinito() == "Laptop"
# solo dopo si sceglie come inserirlo, con i campi del tipo scelto
assert AddChoiceDialog(app, iphone=True).title() == "Aggiungi iPhone"
assert AddChoiceDialog(app, iphone=False).title() == "Aggiungi dispositivo"

# ---------------------------------------------- 2. Home torna davvero alla home
app.show_iphones()
assert app.view == "type" and app.var_type.get() == TIPO
app.show_home()
assert app.view == "home", app.view
assert app.var_type.get() == "Tutti", app.var_type.get()
assert app.var_room.get() == "Tutte"
assert len(app.visible) == len(app.store.items), (len(app.visible), len(app.store.items))
# e lo stesso vale passando da una stanza
app.show_room(KIOSK); app.show_home()
assert len(app.visible) == len(app.store.items)

# ---------------------------------------------- 3. il piu' recente in cima
assert app.sort_field == "modificato_il" and app.sort_reverse is True
time.sleep(1.1)
app._run(lambda: app.store.add(new_item("IT-0900", "Laptop", "T14 Gen 5", "PFNEW1", DR)), "ok")
assert app.visible[0]["asset_tag"] == "IT-0900", [i["asset_tag"] for i in app.visible[:3]]
time.sleep(1.1)   # secondi diversi: l ordine deve essere certo
app._run(lambda: app.store.add(new_item("IT-0901", "Tablet",
                                        "Dell Latitude 7320 Detachable", "8HNEW", BAU)), "ok")
assert app.visible[0]["asset_tag"] == "IT-0901", [i["asset_tag"] for i in app.visible[:3]]
# anche dentro una stanza
app.show_room(BAU)
assert app.visible[0]["asset_tag"] == "IT-0901"
app.show_home()

# le date si ordinano per data, non per testo
a = {"modificato_il": "09/12/2026 08:00"}
b = {"modificato_il": "10/01/2027 08:00"}
assert chiave_ordinamento(a, "modificato_il") < chiave_ordinamento(b, "modificato_il")
assert chiave_ordinamento({"modificato_il": ""}, "modificato_il") \
    < chiave_ordinamento(a, "modificato_il")
# il testo resta alfabetico e parte dalla A
app.sort_by("modello")
assert app.sort_reverse is False
app.sort_by("spedito_il")
assert app.sort_reverse is True, "le date partono dalla piu' recente"

# ---------------------------------------------- 4. iPhone senza seriale ne' prestiti
app._run(lambda: app.store.add(new_item(tipo=TIPO, modello="Apple iPhone 14",
                                        imei="356938035643809", restituito_da="M. B.",
                                        seriale="RESIDUO", prestato_a="Tizio",
                                        prestato_il="01/01/2026 10:00")), "ok")
tel = app._item_by_tag("356938035643809")
assert tel["seriale"] == "" and tel["prestato_a"] == "" and tel["prestato_il"] == ""
# e nel loro contenitore quelle colonne non compaiono proprio
app.show_iphones()
colonne = app._columns()
for campo in COLONNE_NON_IPHONE:
    assert campo not in colonne, (campo, colonne)
assert "imei" in colonne and "restituito_da" in colonne
assert ACTION_COLUMN in colonne
# nelle altre schermate restano, servono a laptop e tablet
app.show_room(KIOSK)
for campo in COLONNE_NON_IPHONE:
    assert campo in app._columns(), campo
app.destroy()
print("CORREZIONI OK")
