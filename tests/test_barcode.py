"""Aggiunta con lettore di codici, e iPhone senza seriale ne' prestiti."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import messagebox
from inventario.store import (DA_RISPEDIRE, InventoryError, is_on_loan, new_item,
                              normalize_iphone)
from inventario.ui import AddChoiceDialog, App, ItemDialog, ScanDialog

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
app = App(fixture.build()); app._initial_load()
TIPO = app.iphone_type()
avvisi = []
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))

# ---------------------------------------------- iPhone: niente seriale, niente prestiti
tel = new_item(tipo=TIPO, modello="Apple iPhone 14", seriale="NONDEVERESTARE",
               imei="356938035643809", restituito_da="M. B.",
               prestato_a="Tizio", prestato_il="01/01/2026 10:00")
assert tel["seriale"] == "" and tel["prestato_a"] == "" and tel["prestato_il"] == ""
assert tel["stato"] == DA_RISPEDIRE and not is_on_loan(tel)
app._run(lambda: app.store.add(tel), "ok")
salvato = app._item_by_tag("356938035643809")
assert salvato["seriale"] == "" and salvato["prestato_a"] == ""

try:
    app.store.lend("356938035643809", "Tizio"); raise SystemExit("iPhone prestato")
except InventoryError as e:
    print("prestito rifiutato:", e)

# anche forzando i campi da una modifica, vengono ripuliti
forzato = dict(salvato)
forzato.update(seriale="XYZ", prestato_a="Caio", prestato_il="02/01/2026 09:00")
app._run(lambda: app.store.update("356938035643809", forzato), "ok")
def etichette_scheda(dialogo):
    return [str(w.cget("text")).rstrip(" *") for w in dialogo.fields.winfo_children()
            if w.winfo_class() == "TLabel" and w.grid_info().get("column") == 0]

pulito = app._item_by_tag("356938035643809")
assert pulito["seriale"] == "" and pulito["prestato_a"] == "" and pulito["prestato_il"] == ""

# la scheda di un iPhone non mostra il numero di serie
dlg = ItemDialog(app, app.cfg["rooms"], app.cfg["types"], pulito,
                 iphone_room=app.iphone_room(), stati=app.cfg["states"])
# per un iPhone l'unico campo che non si puo' non avere e' l'IMEI
assert [l for l, _v, _w in dlg.required] == ["IMEI"]
assert etichette_scheda(dlg) == ["IMEI", "Modello", "Restituito da", "Stanza",
                                 "Stato", "Note"], etichette_scheda(dlg)
assert dlg._loan == ("", "")
dlg._ok()
assert dlg.result["seriale"] == "" and dlg.result["prestato_a"] == ""

# ---------------------------------------------- scelta fra manuale e scansione
def voci(dialogo):
    return [str(b.cget("text")) for b in dialogo.winfo_children()[0].winfo_children()
            if b.winfo_class() == "TButton"]

app.show_home()
scelta = AddChoiceDialog(app)
assert voci(scelta) == ["Scansiona con il lettore di codici",
                        "Inserimento manuale", "Annulla"], voci(scelta)
scelta._cancel()

# nel contenitore iPhone si scansiona l'IMEI, non l'asset tag
app.show_iphones()
assert app.view == "type"
scelta = AddChoiceDialog(app, iphone=True)
assert voci(scelta) == ["Scansiona l'IMEI con il lettore di codici",
                        "Inserimento manuale", "Annulla"], voci(scelta)
assert scelta.title() == "Aggiungi iPhone"
scelta._cancel()

# un solo passo, e la scheda si apre gia' in modalita' iPhone
passo_imei = ScanDialog(app, "IMEI", "l'IMEI del telefono", 1, 1)
assert passo_imei.var_titolo.get() == "Scansiona l'IMEI del telefono"
passo_imei.var_valore.set("351234567890123")
passo_imei._ok()
preset_tel = new_item(tipo=app.iphone_type(), imei=passo_imei.result,
                      stanza=app.iphone_room())
scheda = ItemDialog(app, app.cfg["rooms"], app.cfg["types"], preset_tel,
                    iphone_room=app.iphone_room(), stati=app.cfg["states"])
assert scheda.is_iphone()
assert [l for l, _v, _w in scheda.required] == ["IMEI"]
assert scheda.var_imei.get() == "351234567890123"
assert scheda.missing_fields() == [], "con l'IMEI si salva anche senza il resto"
scheda.var_modello.set("Apple iPhone 13"); scheda.var_restituito.set("E. Rossi")
scheda._ok()
assert scheda.result["asset_tag"] == "351234567890123"
assert scheda.result["seriale"] == "" and scheda.result["stanza"] == BAU
app._run(lambda: app.store.add(scheda.result), "ok")
assert app._item_by_tag("351234567890123")["stato"] == DA_RISPEDIRE
app.show_home()

# ---------------------------------------------- i passi della scansione
passo = ScanDialog(app, "Asset tag", "l'asset tag", 1, 3)
assert passo.manuale is False
assert passo.var_titolo.get() == "Scansiona l'asset tag"
assert "lettore compila" in passo.var_aiuto.get()
# il lettore scrive nel campo e conferma
passo.var_valore.set("IT-0800")
passo._ok()
assert passo.result == "IT-0800"

# campo vuoto: non si va avanti
passo = ScanDialog(app, "Numero di serie", "il numero di serie", 2, 3)
avvisi.clear(); passo._ok()
assert passo.result is None and avvisi[-1][0] == "Campo vuoto"
assert passo.winfo_exists()

# l'ancora: si passa alla scrittura a mano senza perdere il passo
assert str(passo.btn_manuale.cget("text")) == "Non riesco a scansionare - inserisci a mano"
passo.passa_a_manuale()
assert passo.manuale is True
assert passo.var_titolo.get() == "Scrivi il numero di serie"
assert "Digita" in passo.var_aiuto.get()
assert not passo.btn_manuale.winfo_ismapped(), "sparito il pulsante, siamo gia' a mano"
passo.var_valore.set("  PF5NEW9  ")
passo._ok()
assert passo.result == "PF5NEW9", passo.result

# il modello parte gia' in modalita' manuale
passo = ScanDialog(app, "Modello", "il modello del dispositivo", 3, 3, manuale=True)
assert passo.manuale is True and passo.var_titolo.get() == "Scrivi il modello del dispositivo"
passo._cancel()

# ---------------------------------------------- la scheda si apre precompilata
app.show_room(DR)
tipi = app.cfg["types"]
preset = new_item(asset_tag="IT-0800", seriale="PF5NEW9",
                  modello="Lenovo ThinkPad T14 Gen 5",
                  tipo=next(t for t in tipi if t.lower() != "iphone"),
                  stanza=app.stanza_predefinita())
assert preset["stanza"] == DR, "nella vista stanza propone quella stanza"
dlg = ItemDialog(app, app.cfg["rooms"], tipi, preset,
                 iphone_room=app.iphone_room(), stati=app.cfg["states"])
assert dlg.var_tag.get() == "IT-0800" and dlg.var_seriale.get() == "PF5NEW9"
assert dlg.var_modello.get() == "Lenovo ThinkPad T14 Gen 5"
assert dlg.missing_fields() == [], "tutto gia' pronto: basta confermare"
dlg._ok()
app._run(lambda: app.store.add(dlg.result), "ok")
aggiunto = app._item_by_tag("IT-0800")
assert aggiunto["stanza"] == DR and aggiunto["seriale"] == "PF5NEW9"

app.show_home()
assert app.stanza_predefinita() == app.cfg["rooms"][0]
app.show_iphones()
assert app.stanza_predefinita() == BAU
app.destroy()
print("BARCODE OK")
