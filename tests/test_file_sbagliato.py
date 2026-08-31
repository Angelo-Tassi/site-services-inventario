"""Un foglio da importare non deve poter diventare l'inventario.

Il caso vero, trovato sul campo: al primo avvio il programma chiedeva dove
fosse l'inventario, l'utente ha indicato il file di prova, e da li' in poi le
stanze restavano vuote e i nomi delle stanze comparivano in elenco come
dispositivi. Nessuno dei due sintomi rimandava alla causa.
"""
import os, shutil, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario import config
from inventario.store import (InventoryStore, righe_separatore,
                              sembra_un_foglio_da_importare)

STANZE = [fixture.BAU, fixture.KIOSK, fixture.DR]
RADICE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROVA = os.path.join(RADICE, "Collaudo", "Inventario_di_prova.xlsx")

# ---- il foglio di prova viene riconosciuto per quello che e'
copia = os.path.join(tempfile.mkdtemp(), "prova.xlsx")
shutil.copy(PROVA, copia)
da_importare, motivo = sembra_un_foglio_da_importare(copia, STANZE)
assert da_importare, "il foglio di prova deve essere riconosciuto"
assert "separatore" in motivo, motivo

# ---- e le sue righe separatore si elencano
items = InventoryStore(copia).load()
assert len(items) == 33, len(items)          # 30 dispositivi + 3 separatori
separatori = righe_separatore(items, STANZE)
assert separatori == ["SITE SERVICES BAU", "DIGITAL KIOSK",
                      "MAGAZZINO DISASTER RECOVERY"], separatori
assert not any(i.get("stanza") for i in items), "nessuno ha una stanza: e' il sintomo"

# ---- un inventario vero non viene mai scambiato per un foglio da importare
vero = fixture.build()
da_importare, motivo = sembra_un_foglio_da_importare(vero, STANZE)
assert not da_importare, motivo
assert righe_separatore(InventoryStore(vero).load(), STANZE) == []

# ---- un inventario vuoto appena creato non da' falsi allarmi
nuovo = os.path.join(tempfile.mkdtemp(), "Inventario.xlsx")
InventoryStore(nuovo).create_if_missing()
assert sembra_un_foglio_da_importare(nuovo, STANZE)[0] is False

# ---- al primo avvio l'inventario si crea da solo, senza chiedere niente
finta_app = tempfile.mkdtemp()
vero_app_dir = config.app_dir
config.app_dir = lambda: finta_app
try:
    assert config.load_data_path() is None, "non c'e' ancora nessun inventario"
    atteso = config.default_data_path()
    assert atteso == os.path.join(finta_app, "Produzione", "Inventario.xlsx"), atteso
    os.makedirs(os.path.dirname(atteso))
    InventoryStore(atteso).create_if_missing()
    assert config.load_data_path() == atteso, "ora si trova da solo"
finally:
    config.app_dir = vero_app_dir

print("FILE SBAGLIATO OK")
