"""Scheda di inserimento: campi obbligatori, variante iPhone, stati, contenitore."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import messagebox
from inventario.store import DA_RISPEDIRE, DISPONIBILE, InventoryError, new_item
from inventario.ui import ACTION_COLUMN, App, ItemDialog, RoomCard

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
app = App(fixture.build()); app._initial_load()
TIPO = app.iphone_type()
STATI = app.cfg["states"]
errori, avvisi = [], []
messagebox.showerror = lambda t, m, **k: errori.append((t, m))
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))

def apri(item=None):
    return ItemDialog(app, app.cfg["rooms"], app.cfg["types"], item,
                      iphone_room=app.iphone_room(), stati=STATI)

def etichette(d):
    """Le etichette dei campi del modulo, senza l'asterisco."""
    return [str(w.cget("text")).rstrip(" *") for w in d.fields.winfo_children()
            if w.winfo_class() == "TLabel" and w.grid_info().get("column") == 0]


def obbligatori(d):
    return [lbl for lbl, _v, _w in d.required]

# ---- il tipo e' il primo campo
dlg = apri()
corpo = dlg.winfo_children()[0]
prima = [w.cget("text") for w in corpo.winfo_children()
         if w.grid_info().get("row") == 0 and w.grid_info().get("column") == 0]
assert prima == ["Tipo *"], prima
assert etichette(dlg) == ["Asset Tag", "Modello", "Numero di serie", "Stanza",
                          "Stato", "Note"], etichette(dlg)
assert obbligatori(dlg) == ["Asset Tag"], obbligatori(dlg)
assert dlg.var_stato.get() == DISPONIBILE

# ---- obbligatorio e' solo l'identificativo: senza quello il dispositivo non
# esiste, il resto si completa quando lo si ha sottomano
assert dlg.missing_fields() == ["Asset Tag"]
dlg._ok()
assert dlg.result is None and errori[-1][0] == "Dati mancanti"
assert dlg.winfo_exists(), "il popup resta aperto"
dlg.var_tag.set("   ")
assert dlg.missing_fields() == ["Asset Tag"], "gli spazi non contano"
dlg._ok(); assert dlg.result is None
# con il solo asset tag si salva: modello e seriale restano da completare
dlg.var_tag.set("IT-0700")
assert dlg.missing_fields() == []
# la stanza non blocca: se non e' stata scelta si parte dalla prima
assert dlg.var_stanza.get() in dlg.rooms, dlg.var_stanza.get()

# ---- inserimento completo, con stato scelto
dlg.var_modello.set("Lenovo ThinkPad T14 Gen 5"); dlg.var_seriale.set("PF5NEW9")
dlg.var_stanza.set(DR); dlg.var_stato.set("Guasto in attesa tecnico")
dlg._ok()
assert dlg.result["stato"] == "Guasto in attesa tecnico"
assert not dlg.winfo_exists()
app._run(lambda: app.store.add(dlg.result), "ok")
assert app._item_by_tag("IT-0700")["stato"] == "Guasto in attesa tecnico"
assert app.store.set_stato("IT-0700", "Da rebuildare") is True

# ---- variante iPhone
dlg = apri()
dlg.var_tipo.set(TIPO); dlg._build_fields()
assert dlg.is_iphone()
assert etichette(dlg) == ["IMEI", "Modello", "Restituito da", "Stanza",
                          "Stato", "Note"], etichette(dlg)
assert obbligatori(dlg) == ["IMEI"], obbligatori(dlg)
combo_stanza = [w for w in dlg.fields.winfo_children()
                if w.winfo_class() == "TCombobox"][0]
assert str(combo_stanza.cget("state")) == "disabled"
assert dlg.var_stanza.get() == BAU and list(combo_stanza.cget("values")) == [BAU]
assert dlg.var_stato.get() == DA_RISPEDIRE
assert dlg.missing_fields() == ["IMEI"]
dlg._ok(); assert dlg.result is None

# la nota e i valori sopravvivono al cambio di tipo
dlg.var_imei.set("356938035643809"); dlg.var_modello.set("Apple iPhone 14")
dlg.text_note.insert("1.0", "consegnato in reception")
dlg.var_tipo.set("Laptop"); dlg._build_fields()
assert dlg.text_note.get("1.0", "end").strip() == "consegnato in reception"
dlg.var_tipo.set(TIPO); dlg._build_fields()
assert dlg.var_imei.get() == "356938035643809"
dlg.var_restituito.set("Marco Bianchi"); dlg._ok()
r = dlg.result
assert r["imei"] == r["asset_tag"] == "356938035643809"
assert r["restituito_da"] == "Marco Bianchi" and r["seriale"] == ""
assert r["stanza"] == BAU and r["stato"] == DA_RISPEDIRE
app._run(lambda: app.store.add(r), "ok")

# ---- l'iPhone non si sposta, nemmeno forzandolo
app._run(lambda: app.store.add(new_item(tipo=TIPO, modello="Apple iPhone 13",
                                        imei="351234567890123", restituito_da="E. R.",
                                        stanza=KIOSK)), "ok")
assert app._item_by_tag("351234567890123")["stanza"] == BAU
assert app.store.move_to_room(["356938035643809", "IT-0101"], DR) == (1, 1, 0)
assert app._item_by_tag("356938035643809")["stanza"] == BAU
assert app._item_by_tag("IT-0101")["stanza"] == DR
avvisi.clear()
app.show_home(); app.tree.selection_set(["356938035643809"]); app.on_move()
assert "non possono essere spostati" in avvisi[-1][1]

# ---- lo stato di un iPhone non si tocca
try:
    app.store.set_stato("356938035643809", "Controllare"); raise SystemExit("iPhone modificato")
except InventoryError:
    pass
old = app._item_by_tag("356938035643809")
mod = dict(old); mod["modello"] = "Apple iPhone 14 Pro"
app._run(lambda: app.store.update("356938035643809", mod), "ok")
assert app._item_by_tag("356938035643809")["stato"] == DA_RISPEDIRE

# ---- contenitore iPhone: scheda in coda e vista dedicata
app.show_home()
schede = [w for f in app.body.winfo_children() for w in f.winfo_children()
          if isinstance(w, RoomCard)]
assert [c.labels[1].cget("text") for c in schede] == [BAU, KIOSK, DR, TIPO]
assert schede[3].labels[2].cget("text") == "2"
assert schede[3].note.cget("text") == "anche in %s" % BAU
schede[3].command()
assert app.view == "type" and app.var_room.get() == "Tutte"
assert sorted(i["asset_tag"] for i in app.visible) == ["351234567890123", "356938035643809"]
assert app._columns()[1] == ACTION_COLUMN
assert app.tree.heading(ACTION_COLUMN)["text"] == "Spedizione"
app.var_search.set("marco"); app.refresh_table()
assert [i["asset_tag"] for i in app.visible] == ["356938035643809"]
app.reset_filters()
assert app.view == "type" and len(app.visible) == 2
app.show_room(BAU)
assert "356938035643809" in app.tree.get_children(), "gli iPhone restano visibili in BAU"
app.destroy()
print("DIALOG OK")
