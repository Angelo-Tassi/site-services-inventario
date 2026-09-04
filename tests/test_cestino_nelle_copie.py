"""Il cestino viaggia con l'inventario: nelle copie locali e in quelle automatiche.

Gli eliminati di recente stanno in un file a parte accanto ai dati, e per un po'
sono rimasti fuori da tutto: la copia locale portava via inventario e
impostazioni ma non il cestino, e le copie automatiche solo il .xlsx. Chi si
fosse ritrovato senza la cartella di rete avrebbe recuperato i dispositivi ma
perso l'unico posto da cui si ripesca quello tolto per sbaglio.

Inventario ed eliminati sono lo stesso stato condiviso: si salvano insieme e
tornano indietro insieme, o tornando indietro sui soli dispositivi il cestino
resterebbe pieno di roba nel frattempo rientrata in inventario.
"""
import json, os, sys, tempfile, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario import config
from inventario.store import (NOME_DATI_NELLO_ZIP, NOME_ELIMINATI_NELLO_ZIP,
                              NOME_IMPOSTAZIONI_NELLO_ZIP, InventoryStore, new_item)

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR

# ============================ la copia locale ============================
percorso = fixture.build()
store = InventoryStore(percorso, iphone_room=BAU)
store.load()
store.delete(["IT-0101", "IT-0102"])
# eliminati nello stesso istante: l'ordine fra i due non e' garantito
assert sorted(v["asset_tag"] for v in store.eliminati()) == ["IT-0101", "IT-0102"]

fuori = tempfile.mkdtemp()
archivio = os.path.join(fuori, "Copia.zip")
store.copia_in(archivio)
dentro = zipfile.ZipFile(archivio).namelist()
assert dentro == [NOME_DATI_NELLO_ZIP, NOME_IMPOSTAZIONI_NELLO_ZIP,
                  NOME_ELIMINATI_NELLO_ZIP], dentro
assert store.eliminati_nella_copia is True

# ---- l'anteprima lo conta, prima di toccare qualsiasi cosa
rapporto = store.anteprima_copia_locale(archivio)
assert rapporto["eliminati"] == 2, rapporto["eliminati"]
assert rapporto["dispositivi"] == 11, rapporto["dispositivi"]

# ---- e il ripristino lo rimette com'era
store.delete(["IT-0103"])
store.ripristina_eliminati(["IT-0101"])
assert sorted(v["asset_tag"] for v in store.eliminati()) == ["IT-0102", "IT-0103"]
store.ripristina_da_copia_locale(archivio)
assert sorted(v["asset_tag"] for v in store.eliminati()) == ["IT-0101", "IT-0102"], \
    "il cestino deve tornare quello della copia"
assert store.eliminati_ripristinati == 2, store.eliminati_ripristinati

# ---- anche la copia in solo .xlsx se lo porta accanto
solo_dati = os.path.join(fuori, "SoloDati.xlsx")
store.copia_in(solo_dati)
accanto = os.path.join(fuori, "SoloDati_eliminati.json")
assert os.path.exists(accanto), sorted(os.listdir(fuori))
assert len(json.load(open(accanto))) == 2

# ---- una copia senza cestino non azzera quello di adesso
senza = os.path.join(fuori, "Senza.xlsx")
store.copia_in(senza)
os.remove(os.path.join(fuori, "Senza_eliminati.json"))
prima = [v["asset_tag"] for v in store.eliminati()]
store.ripristina_da_copia_locale(senza)
assert [v["asset_tag"] for v in store.eliminati()] == prima, \
    "senza cestino nella copia, quello di adesso resta dov'e'"
assert store.eliminati_ripristinati == 0

# ---- e da un ripristino nessuno esce in due posti insieme
# Il caso vero: si elimina un dispositivo, poi si torna a una copia presa
# quando quel dispositivo c'era ancora. Il ripristino lo rimette in
# inventario, e se la sua voce restasse nel cestino da li' se ne potrebbe
# "ripristinare" una seconda copia.
due_posti = InventoryStore(fixture.build(), iphone_room=BAU)
due_posti.load()
vecchia = os.path.join(fuori, "PrimaDiTutto.zip")
due_posti.copia_in(vecchia)                    # cestino vuoto: ci va lo stesso
assert zipfile.ZipFile(vecchia).namelist() == [
    NOME_DATI_NELLO_ZIP, NOME_IMPOSTAZIONI_NELLO_ZIP, NOME_ELIMINATI_NELLO_ZIP], \
    "il cestino viaggia anche quando e' vuoto"
due_posti.delete(["IT-0101"])
due_posti.load()
assert [v["asset_tag"] for v in due_posti.eliminati()] == ["IT-0101"]
due_posti.ripristina_da_copia_locale(vecchia)
due_posti.load()
assert [i for i in due_posti.items if i["asset_tag"] == "IT-0101"], "e' tornato"
assert not due_posti.eliminati(), "e non e' rimasto anche nel cestino"

# ---- e se la copia e' vecchia e il cestino non ce l'ha, ci pensa la pulizia
antica = os.path.join(fuori, "Antica.xlsx")
due_posti.copia_in(antica)
os.remove(os.path.join(fuori, "Antica_eliminati.json"))
due_posti.delete(["IT-0102"])
due_posti.load()
due_posti.ripristina_da_copia_locale(antica)
due_posti.load()
assert [i for i in due_posti.items if i["asset_tag"] == "IT-0102"], "e' tornato"
assert not due_posti.eliminati(), "tolto dal cestino: e' in inventario"
assert [v["asset_tag"] for v in due_posti.tolti_dal_ripristino] == ["IT-0102"], \
    due_posti.tolti_dal_ripristino

# ============================ le copie automatiche ============================
finta_app = tempfile.mkdtemp()
config.app_dir = lambda: finta_app
os.environ["APPDATA"] = tempfile.mkdtemp()
dati = os.path.join(finta_app, "Produzione", "Inventario.xlsx")
os.makedirs(os.path.dirname(dati))
config.load_data_path = lambda: dati
cartella = config.backup_dir()

s = InventoryStore(dati, iphone_room=BAU)
s.create_if_missing()
for tag in ("IT-0101", "IT-0102", "IT-0103"):
    s.add(new_item(tag, "Laptop", "T14", "SN" + tag, BAU))
s.load()

s.delete(["IT-0101"])
copia = s.copia_di_sicurezza()
gemello = os.path.splitext(copia)[0] + "_eliminati.json"
assert os.path.exists(gemello), sorted(os.listdir(cartella))
assert [v["asset_tag"] for v in json.load(open(gemello))] == ["IT-0101"]

# ---- tornando indietro, torna indietro tutto lo stato insieme
s.delete(["IT-0102"])
assert sorted(v["asset_tag"] for v in s.eliminati()) == ["IT-0101", "IT-0102"]
quanti, _precedente = s.restore(copia)
assert quanti == 2, quanti
assert sorted(i["asset_tag"] for i in s.load()) == ["IT-0102", "IT-0103"]
assert [v["asset_tag"] for v in s.eliminati()] == ["IT-0101"], \
    "IT-0102 e' rientrato in inventario: non puo' restare anche nel cestino"
assert s.eliminati_ripristinati == 1

# ---- una copia vecchia, senza gemello, lascia il cestino com'e'
vecchia = s.copia_di_sicurezza()
os.remove(os.path.splitext(vecchia)[0] + "_eliminati.json")
prima = [v["asset_tag"] for v in s.eliminati()]
s.restore(vecchia)
assert [v["asset_tag"] for v in s.eliminati()] == prima
assert s.eliminati_ripristinati == 0

# ---- la rotazione porta via il gemello insieme alla sua copia
for n in range(14):
    quando = 1750000000 + n * 3600
    os.utime(dati, (quando, quando))
    s.copia_di_sicurezza()
xlsx = sorted(f for f in os.listdir(cartella) if f.endswith(".xlsx"))
gemelli = sorted(f for f in os.listdir(cartella) if f.endswith("_eliminati.json"))
assert len(xlsx) == 10, xlsx
assert len(gemelli) <= len(xlsx), gemelli
for g in gemelli:
    atteso = g[:-len("_eliminati.json")] + ".xlsx"
    assert atteso in xlsx, "gemello orfano rimasto indietro: %s" % g

print("CESTINO NELLE COPIE OK")
