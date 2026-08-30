"""Le copie di sicurezza prima delle operazioni distruttive."""
import os, sys, tempfile, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from datetime import datetime
from inventario import config
from inventario.store import InventoryError, InventoryStore, new_item

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
RADICE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ---- la cartella e' dentro il programma e ha la sua .gitignore
protezione = os.path.join(RADICE, "Backup", ".gitignore")
assert os.path.exists(protezione), "manca Backup/.gitignore"
regole = open(protezione, encoding="utf-8").read()
assert "*" in regole and "!.gitignore" in regole, regole

finta_app = tempfile.mkdtemp()
config.app_dir = lambda: finta_app
os.environ["APPDATA"] = tempfile.mkdtemp()
cartella = config.backup_dir()
assert cartella == os.path.join(finta_app, "Backup"), cartella
assert os.path.isdir(cartella)
assert not os.path.exists(os.path.join(cartella, ".scrivibile")), "la prova va rimossa"

# ---- il nome porta la data del file salvato, non quella della copia
p = os.path.join(tempfile.mkdtemp(), "Inventario.xlsx")
s = InventoryStore(p, iphone_room=BAU)
s.create_if_missing()
for tag, stanza in (("IT-0101", BAU), ("IT-0106", KIOSK), ("DR-0201", DR)):
    s.add(new_item(tag, "Laptop", "T14", "SN" + tag, stanza))
s.add(new_item(tipo="Iphone", modello="iPhone 14", imei="356938035643809",
               restituito_da="M. B."))
quando = datetime.fromtimestamp(os.path.getmtime(p))
atteso = "Inventario_%s.xlsx" % quando.strftime("%Y-%m-%d_%H-%M-%S")

copia = s.copia_di_sicurezza()
assert os.path.dirname(copia) == cartella, copia
assert os.path.basename(copia) == atteso, (os.path.basename(copia), atteso)

# ---- due copie dello stesso file non si sovrascrivono
seconda = s.copia_di_sicurezza()
assert seconda != copia and os.path.exists(seconda)
assert "(2)" in os.path.basename(seconda), os.path.basename(seconda)

# ---- la copia contiene davvero i dati di prima
prima = InventoryStore(copia); prima.load()
assert len(prima.items) == 4, len(prima.items)

# ---- il reset ci passa
eliminati, tenuti, dal_reset = s.reset()
assert os.path.dirname(dal_reset) == cartella
s.load()
assert eliminati == 3 and tenuti == 1
recuperabile = InventoryStore(dal_reset); recuperabile.load()
assert len(recuperabile.items) == 4, "dalla copia si recupera tutto"

# ---- e anche l'importazione che sostituisce
s.import_items([new_item("IT-0900", "Laptop", "T14", "PF900", KIOSK)], "merge")
time.sleep(1.1)
s.import_items([new_item("IT-0901", "Laptop", "T14", "PF901", KIOSK)], "replace")
copie = sorted(f for f in os.listdir(cartella) if f.endswith(".xlsx"))
assert len(copie) == 4, copie
assert all(f.startswith("Inventario_20") for f in copie), copie

# ---- senza posto dove scrivere, l'operazione si annulla
config.backup_dir = lambda: None
prima_del_tentativo = len(s.items)
try:
    s.reset(); raise SystemExit("reset eseguito senza copia di sicurezza")
except InventoryError as e:
    assert "annullata" in str(e), str(e)
s.load()
assert len(s.items) == prima_del_tentativo, "i dati non devono essere toccati"
print("BACKUP OK")
