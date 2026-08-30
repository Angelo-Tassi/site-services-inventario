"""Come si determina il percorso del file dati."""
import json, os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from inventario import config
from inventario.store import InventoryStore

finta_app = tempfile.mkdtemp()
profilo = tempfile.mkdtemp()
config.app_dir = lambda: finta_app
os.environ.pop("INVENTARIO_FILE", None)
os.environ["APPDATA"] = profilo

assert config.load_data_path() is None, "cartella vuota: si chiede all'utente"

dati = os.path.join(finta_app, "Inventario.xlsx")
InventoryStore(dati).create_if_missing()
assert config.load_data_path() == dati, "basta il file accanto al programma"
assert not os.path.exists(config.local_config_path()), "nessun file di configurazione necessario"

altro = os.path.join(tempfile.mkdtemp(), "Altro.xlsx")
InventoryStore(altro).create_if_missing()
assert config.save_data_path(altro) == config.local_config_path()
assert config.load_data_path() == altro, "il percorso salvato ha la precedenza"

os.remove(altro)
assert config.load_data_path() == dati, "percorso sparito: si torna al file accanto"

os.remove(config.local_config_path())
os.chmod(finta_app, 0o555)
try:
    InventoryStore(altro).create_if_missing()
    scritto = config.save_data_path(altro)
    assert scritto == config.user_config_path(), scritto
    assert json.load(open(scritto))["data_path"] == altro
    assert config.load_data_path() == altro
finally:
    os.chmod(finta_app, 0o755)

os.environ["INVENTARIO_FILE"] = "/percorso/di/rete/Inventario.xlsx"
assert config.load_data_path() == "/percorso/di/rete/Inventario.xlsx"
os.environ.pop("INVENTARIO_FILE")

# impostazioni condivise
cfg = config.load_shared_config(dati)
assert cfg["iphone_room"] in cfg["rooms"]
assert cfg["states"][0] == "Disponibile"
config.save_shared_config(dati, {"rooms": ["A", "B"], "types": ["Laptop"],
                                 "loan_rooms": ["B"], "iphone_room": "sparita",
                                 "states": ["Disponibile"]})
cfg = config.load_shared_config(dati)
assert cfg["iphone_room"] == "A", "stanza iPhone inesistente: ripiega sulla prima"
print("CONFIG OK")
