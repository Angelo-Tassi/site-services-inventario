"""Un'installazione appena estratta deve essere completa e funzionante.

Il caso segnalato: estratto lo zip, mancavano il contenitore iPhone e i
prestiti, perche' i valori predefiniti non li prevedevano.
"""
import json, os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario import config
from inventario.store import InventoryStore, is_iphone, new_item, rows_from_workbook
from inventario.ui import App

RADICE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def installazione_vergine():
    """Come dopo aver estratto lo zip: cartelle vuote, nessuna impostazione."""
    finta = tempfile.mkdtemp()
    os.makedirs(os.path.join(finta, "Produzione"))
    os.makedirs(os.path.join(finta, "Backup"))
    config.app_dir = lambda: finta
    os.environ["APPDATA"] = tempfile.mkdtemp()
    os.environ.pop("INVENTARIO_FILE", None)
    return finta

finta = installazione_vergine()
percorso = config.default_data_path()
assert percorso == os.path.join(finta, "Produzione", "Inventario.xlsx"), percorso
assert not os.path.exists(percorso), "prima dell'avvio non c'e' nessun inventario"

s = InventoryStore(percorso)
assert s.create_if_missing() is True

# ---- le impostazioni vengono scritte, non lasciate implicite
impostazioni = config.shared_config_path(percorso)
assert os.path.exists(impostazioni), "il file delle impostazioni deve essere creato"
dati = json.load(open(impostazioni, encoding="utf-8"))
assert set(dati) == {"rooms", "types", "loan_rooms", "iphone_room", "states"}, dati

cfg = config.load_shared_config(percorso)
assert cfg["rooms"] == [fixture.BAU, fixture.KIOSK, fixture.DR], cfg["rooms"]
assert cfg["loan_rooms"] == [fixture.KIOSK], cfg["loan_rooms"]
assert cfg["iphone_room"] == fixture.BAU
assert len(cfg["states"]) == 5

# ---- il tipo iPhone c'e' da subito
tipo = next((t for t in cfg["types"] if is_iphone(t)), None)
assert tipo is not None, cfg["types"]

# ---- e quindi il contenitore, il modulo e le regole funzionano
app = App(percorso); app._initial_load()
assert app.iphone_type() == tipo
assert app.iphone_room() == fixture.BAU
assert app.cfg["loan_rooms"] == [fixture.KIOSK]
schede = [w for f in app.body.winfo_children() for w in f.winfo_children()
          if w.__class__.__name__ == "RoomCard"]
nomi = [c.labels[1].cget("text") for c in schede]
assert nomi == [fixture.BAU, fixture.KIOSK, fixture.DR, tipo], nomi

# un iPhone si inserisce e si comporta come deve
app._run(lambda: app.store.add(new_item(tipo=tipo, modello="Apple iPhone 14",
                                        imei="356938035643809", restituito_da="M. B.")), "ok")
telefono = app._item_by_tag("356938035643809")
assert telefono["stanza"] == fixture.BAU and telefono["stato"] == "Da Rispedire"
app.show_iphones()
assert len(app.visible) == 1
assert app.action_label(telefono) == "Conferma spedizione"

# i prestiti sono attivi nella stanza giusta
app.show_room(fixture.KIOSK)
assert app.loan_column_visible(), "la colonna Prestito deve esserci nel kiosk"

# ---- l'importazione dei file di collaudo funziona su questa installazione
items, esito = rows_from_workbook(
    os.path.join(RADICE, "Collaudo", "Inventario_di_prova.xlsx"), cfg["rooms"])
assert len(items) == 30 and esito["da_tag"] == 30, esito
app._run(lambda: app.store.import_items(items, "merge"), "ok")
app.store.load()
conteggi = {}
for i in app.store.items:
    conteggi[i["stanza"]] = conteggi.get(i["stanza"], 0) + 1
assert conteggi == {fixture.BAU: 11, fixture.KIOSK: 10, fixture.DR: 10}, conteggi

# ---- una voce svuotata di proposito resta svuotata, le altre no
config.save_shared_config(percorso, {"rooms": cfg["rooms"], "types": cfg["types"],
                                     "loan_rooms": [], "iphone_room": fixture.BAU,
                                     "states": cfg["states"]})
assert config.load_shared_config(percorso)["loan_rooms"] == [], "nessuna stanza con prestiti"
parziale = {"rooms": ["Solo una"]}
json.dump(parziale, open(impostazioni, "w", encoding="utf-8"))
cfg = config.load_shared_config(percorso)
assert cfg["rooms"] == ["Solo una"], cfg["rooms"]
assert any(is_iphone(t) for t in cfg["types"]), "i tipi tornano ai predefiniti"
assert cfg["states"], "gli stati tornano ai predefiniti"
app.destroy()
print("INSTALLAZIONE NUOVA OK")
