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
config.load_data_path = lambda: os.path.join(finta_app, "Produzione", "Inventario.xlsx")
cartella = config.backup_dir()
assert cartella == os.path.join(finta_app, "Backup"), cartella

# ---- con il programma installato in locale, le copie restano con i dati
altrove = tempfile.mkdtemp()
os.makedirs(os.path.join(altrove, "Produzione"))
config.load_data_path = lambda: os.path.join(altrove, "Produzione", "Inventario.xlsx")
sulla_share = config.backup_dir()
assert sulla_share == os.path.join(altrove, "Produzione", "Backup"), sulla_share
assert not sulla_share.startswith(finta_app), "le copie non devono finire sulla postazione"
config.load_data_path = lambda: os.path.join(finta_app, "Produzione", "Inventario.xlsx")
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

# ---------------------------------------------------------------- ripristino
from inventario.ui import App, RestoreDialog, RoomsDialog
from tkinter import messagebox

avvisi = []
messagebox.showinfo = lambda t, m, **k: avvisi.append((t, m))
messagebox.showwarning = lambda t, m, **k: avvisi.append((t, m))
# il promemoria della copia locale non deve fermare una suite automatica
messagebox.askyesno = lambda t, m, **k: avvisi.append((t, m)) or False

config.backup_dir = lambda: cartella          # ripristina la vera funzione
app = App(fixture.build()); app._initial_load()
partenza = len(app.store.items)
assert partenza == 13

# una copia dello stato buono, poi un pasticcio
buona = app.store.copia_di_sicurezza()
app._run(lambda: app.store.import_items(
    [new_item("IT-9%02d" % n, "Laptop", "Doppione", "SN%d" % n, KIOSK) for n in range(20)],
    "merge"), "ok")
app.store.load()
assert len(app.store.items) == partenza + 20, len(app.store.items)

# l'elenco delle copie parte dalla piu' recente e sa quanti dispositivi contiene
copie = app.store.copie_disponibili()
assert copie, "nessuna copia trovata"
assert copie[0][0] == buona, copie[0]
assert copie[0][2] == partenza, copie[0]
assert all(copie[i][1] >= copie[i + 1][1] for i in range(len(copie) - 1)), "ordine sbagliato"

# la finestra di scelta elenca le copie e restituisce quella selezionata
dlg = RestoreDialog(app, copie, len(app.store.items))
righe = dlg.elenco.get_children()
assert len(righe) == len(copie)
assert dlg.elenco.selection() == ("0",), "la piu' recente e' gia' selezionata"
assert str(copie[0][2]) in dlg.elenco.item("0", "values")
dlg._ok()
assert dlg.result == buona

# il ripristino riporta indietro e salva lo stato sbagliato
quanti, precedente = app._run(lambda: app.store.restore(buona), "ok")
app.store.load()
assert quanti == partenza and len(app.store.items) == partenza
assert not any(i["modello"] == "Doppione" for i in app.store.items)
assert os.path.exists(precedente)
sbagliato = InventoryStore(precedente); sbagliato.load()
assert len(sbagliato.items) == partenza + 20, "anche l'errore resta recuperabile"

# una copia illeggibile viene rifiutata senza toccare niente
finto = os.path.join(cartella, "Inventario_2020-01-01_00-00-00.xlsx")
open(finto, "w").write("non e' un foglio Excel")
try:
    app.store.restore(finto); raise SystemExit("copia illeggibile accettata")
except InventoryError as e:
    assert "non e' un inventario leggibile" in str(e), str(e)
app.store.load()
assert len(app.store.items) == partenza
assert all(p != finto for p, _q, _n in app.store.copie_disponibili()), \
    "una copia illeggibile non va nemmeno elencata"

# e una copia sparita
try:
    app.store.restore(os.path.join(cartella, "mai_esistita.xlsx"))
    raise SystemExit("copia inesistente accettata")
except InventoryError as e:
    assert "non esiste" in str(e)

# le impostazioni sanno chiedere il ripristino
dlg = RoomsDialog(app, app.cfg["rooms"], app.cfg["types"], [KIOSK], BAU, "it")
dlg._ripristina()
assert dlg.result == {"ripristina": True}
app.destroy()
print("RIPRISTINO OK")
