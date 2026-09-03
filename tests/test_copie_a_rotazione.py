"""Delle copie automatiche se ne tengono dieci, non una in piu'.

Servono a tornare indietro di qualche passo dopo un errore, non a fare da
archivio storico: quello e' il compito della copia che il tecnico si salva in
locale. Senza un limite la cartella Backup sulla share cresceva a ogni
eliminazione, e la copia utile finiva in fondo a un elenco di centinaia.
"""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario import config
from inventario.store import (COPIE_DA_TENERE, InventoryStore, new_item,
                              e_una_copia_automatica)

BAU = fixture.BAU
assert COPIE_DA_TENERE == 10, COPIE_DA_TENERE

# ---- si riconosce una copia automatica dal nome, non dall'estensione
assert e_una_copia_automatica("Inventario_2026-09-02_18-30-00.xlsx")
assert e_una_copia_automatica("Inventario_2026-09-02_18-30-00 (2).xlsx")
assert e_una_copia_automatica("Inventario di prova_2026-01-01_00-00-00.xlsx")
for estraneo in ("Inventario.xlsx",                   # l'inventario stesso
                 "Note del collega.xlsx",             # messo li' a mano
                 "Inventario_2026-13-45_99-99-99.xlsx",  # data impossibile
                 "~$Inventario_2026-09-02_18-30-00.xlsx",  # lock di Excel
                 "Inventario_2026-09-02_18-30-00.txt"):
    assert not e_una_copia_automatica(estraneo), estraneo

# ---- la rotazione vera
finta_app = tempfile.mkdtemp()
config.app_dir = lambda: finta_app
os.environ["APPDATA"] = tempfile.mkdtemp()
dati = os.path.join(finta_app, "Produzione", "Inventario.xlsx")
os.makedirs(os.path.dirname(dati))
config.load_data_path = lambda: dati
cartella = config.backup_dir()

s = InventoryStore(dati, iphone_room=BAU)
s.create_if_missing()
s.add(new_item("IT-0101", "Laptop", "T14", "PF1", BAU))

# due file che nessuno deve toccare: non li ha scritti il programma
a_mano = os.path.join(cartella, "Da non buttare.xlsx")
open(a_mano, "w").close()
vecchio_inventario = os.path.join(cartella, "Inventario.xlsx")
open(vecchio_inventario, "w").close()

# tredici copie, ognuna con la sua data: il nome la porta dentro
fatte = []
for n in range(13):
    quando = 1750000000 + n * 3600
    os.utime(dati, (quando, quando))
    fatte.append(s.copia_di_sicurezza())

rimaste = sorted(f for f in os.listdir(cartella) if e_una_copia_automatica(f))
assert len(rimaste) == COPIE_DA_TENERE, rimaste
# restano le dieci piu' recenti, sparite le tre piu' vecchie
assert [os.path.basename(p) for p in fatte[3:]] == rimaste, (fatte, rimaste)
for perduta in fatte[:3]:
    assert not os.path.exists(perduta), perduta
assert s.copie_scartate == [fatte[2]], s.copie_scartate

# ---- quello che non ha scritto il programma resta dov'e'
assert os.path.exists(a_mano), "un file messo a mano non si cancella"
assert os.path.exists(vecchio_inventario), "solo le copie con la data si ruotano"

# ---- la copia nuova c'e' sempre, e si ruota dopo averla scritta
ultima = fatte[-1]
assert os.path.exists(ultima), "la copia appena fatta non puo' sparire"

# ---- l'elenco di Ripristina vede esattamente quelle rimaste
elencate = [p for p, _quando, _quanti in s.copie_disponibili()
            if e_una_copia_automatica(os.path.basename(p))]
assert len(elencate) == COPIE_DA_TENERE, elencate
assert elencate[0] == fatte[-1], "la piu' recente per prima"

# ---- una copia che non si riesce a cancellare non ferma niente
vero_remove = os.remove

def remove_che_si_rifiuta(percorso):
    # solo le copie: backup_dir() scrive e cancella un file di prova, e
    # bloccare anche quello direbbe un'altra cosa da quella che si sta provando
    if e_una_copia_automatica(os.path.basename(percorso)):
        raise OSError("aperta in Excel da un collega")
    return vero_remove(percorso)

os.remove = remove_che_si_rifiuta
try:
    quando = 1750000000 + 99 * 3600
    os.utime(dati, (quando, quando))
    copia = s.copia_di_sicurezza()
    assert os.path.exists(copia), "la copia va scritta comunque"
    assert s.copie_scartate == [], s.copie_scartate
finally:
    os.remove = vero_remove
assert len([f for f in os.listdir(cartella) if e_una_copia_automatica(f)]) == 11

print("COPIE A ROTAZIONE OK")
