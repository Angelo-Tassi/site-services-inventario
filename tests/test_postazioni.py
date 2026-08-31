"""Programma sulle postazioni, inventario sulla share: il collegamento.

Quello che deve reggere: la configurazione viaggia con la cartella del
programma, un inventario gia' esistente non viene mai toccato, e se la share
non risponde il programma non si inventa una copia locale - sarebbe il modo
piu' silenzioso di far lavorare un tecnico su dati che nessun altro vede.
"""
import json, os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario import config, configura
from inventario.store import InventoryError, InventoryStore, new_item

vero_app_dir = config.app_dir
postazione = tempfile.mkdtemp()          # la cartella del programma, in locale
share = tempfile.mkdtemp()               # la cartella di rete condivisa
config.app_dir = lambda: postazione
os.environ.pop("INVENTARIO_FILE", None)

try:
    # ---- il percorso si scrive in tutti i modi in cui verrebbe naturale
    atteso = os.path.join(share, config.NOME_PRODUZIONE, config.DATA_FILE_NAME)
    for scritto in (share, share + os.sep,
                    os.path.join(share, config.NOME_PRODUZIONE), atteso):
        assert configura.percorso_inventario(scritto) == atteso, scritto

    # ---- primo collegamento: l'inventario non c'e' e viene creato vuoto
    assert config.configured_data_path() == (None, None)
    percorso, esisteva = configura.collega(share)
    assert percorso == atteso and not esisteva
    assert os.path.exists(percorso)
    assert InventoryStore(percorso).load() == []

    # ---- la configurazione sta accanto al programma, per viaggiare con lui
    scritto = config.local_config_path()
    assert os.path.dirname(scritto) == postazione
    assert json.load(open(scritto))["data_path"] == atteso
    assert config.load_data_path() == atteso
    assert config.configured_data_path()[0] == atteso

    # ---- le copie di sicurezza restano con i dati, sulla share
    assert config.backup_dir() == os.path.join(share, config.NOME_PRODUZIONE,
                                               config.NOME_BACKUP)

    # ---- un inventario gia' popolato non viene mai toccato
    store = InventoryStore(percorso, iphone_room=fixture.BAU)
    store.add(new_item("IT-9001", "Laptop", "T14 Gen 5", "PF9", fixture.BAU, ""))
    percorso2, esisteva2 = configura.collega(share)
    assert percorso2 == atteso and esisteva2, "il secondo giro non deve ricreare"
    assert [i["asset_tag"] for i in InventoryStore(percorso).load()] == ["IT-9001"]

    # ---- la lingua gia' scelta non viene persa dal collegamento
    config.save_language("en")
    configura.collega(share)
    assert json.load(open(scritto)).get("lingua") == "en"
    assert json.load(open(scritto))["data_path"] == atteso
    config.save_language("it")

    # ---- share irraggiungibile: si dice, non si ripiega
    try:
        configura.collega(os.path.join(share, "non", "esiste"))
        raise AssertionError("doveva rifiutare")
    except InventoryError as exc:
        assert "non si raggiunge" in str(exc), exc

    # ---- share che sparisce: il programma sa che l'inventario e' altrove e
    # non deve creare una copia locale
    os.rename(os.path.join(share, config.NOME_PRODUZIONE),
              os.path.join(share, "Produzione_spostata"))
    assert config.load_data_path() is None, "il file non c'e' piu'"
    quale, sorgente = config.configured_data_path()
    assert quale == atteso, "ma resta scritto quale sarebbe"
    assert sorgente == scritto
    # e' proprio su questa differenza che main() si ferma invece di creare
    assert not os.path.exists(os.path.join(postazione, config.NOME_PRODUZIONE))
finally:
    config.app_dir = vero_app_dir

print("POSTAZIONI OK")
