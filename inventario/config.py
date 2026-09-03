"""Configurazione dell'applicazione.

Il percorso del file dati si determina in quest'ordine:

  1. la variabile d'ambiente INVENTARIO_FILE (utile per prove e collegamenti);
  2. il file inventario_percorso.json accanto al programma, o nel profilo utente;
  3. Inventario.xlsx nella stessa cartella del programma - e' il caso normale
     quando l'eseguibile e i dati stanno insieme nella cartella di rete;
  4. altrimenti si chiede all'utente, una volta sola.

Le impostazioni condivise (stanze, tipi, prestiti) vivono accanto al file dati,
cosi' sono uguali per tutti gli utenti.
"""

import json
import os

from . import store

APP_NAME = "Inventario"

DEFAULT_ROOMS = ["Site Services BAU", "Digital Kiosk", "Magazzino Disaster Recovery"]
# "Iphone" non e' un tipo come gli altri: la sua presenza accende il
# contenitore in home, il modulo con l'IMEI, la spedizione e la
# conservazione. Deve esserci da subito in un'installazione nuova.
DEFAULT_TYPES = ["Laptop", "Tablet", "Iphone"]
# Stanze in cui e' attiva la gestione dei prestiti.
DEFAULT_LOAN_ROOMS = ["Digital Kiosk"]
# Gli iPhone stanno sempre qui e non possono essere spostati altrove.
DEFAULT_IPHONE_ROOM = "Site Services BAU"


def app_dir():
    """Cartella del programma: quella che contiene Inventario.py."""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


DATA_FILE_NAME = "Inventario.xlsx"


def local_config_path():
    return os.path.join(app_dir(), "inventario_percorso.json")


def user_config_path():
    """Ripiego per quando la cartella del programma e' in sola lettura."""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "inventario_percorso.json")


NOME_BACKUP = "Backup"
NOME_PRODUZIONE = "Produzione"


def backup_dir():
    """Cartella delle copie di sicurezza.

    Le copie stanno con i dati, non con il programma: se l'eseguibile e'
    installato sulle singole postazioni e l'inventario vive sulla share, i
    backup devono restare sulla share, uno solo per tutti. Quando programma e
    dati stanno insieme - il caso di partenza - la cartella e' quella di sempre
    dentro l'applicazione.

    Se non si riesce a scrivere si ripiega sul profilo utente: una copia deve
    poter essere scritta, altrimenti l'operazione distruttiva viene annullata.
    """
    candidate = []
    percorso = load_data_path()
    dati = os.path.dirname(os.path.abspath(percorso)) if percorso else None
    dentro_il_programma = bool(
        dati and os.path.abspath(dati).startswith(os.path.abspath(app_dir())))
    if dentro_il_programma or not dati:
        candidate.append(os.path.join(app_dir(), NOME_BACKUP))
    if dati:
        candidate.append(os.path.join(dati, NOME_BACKUP))
        candidate.append(os.path.join(app_dir(), NOME_BACKUP))
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    candidate.append(os.path.join(base, "Inventario", NOME_BACKUP))
    for cartella in candidate:
        try:
            if not os.path.isdir(cartella):
                os.makedirs(cartella)
            prova = os.path.join(cartella, ".scrivibile")
            with open(prova, "w") as fh:
                fh.write("")
            os.remove(prova)
            return cartella
        except OSError:
            continue
    return None


def production_dir():
    """Cartella dell'inventario di produzione, dentro quella del programma."""
    return os.path.join(app_dir(), NOME_PRODUZIONE)


def default_data_path():
    """Il file dati che il programma apre quando nessuno gli dice altro.

    Sta in `Produzione/` accanto al programma: cosi' l'inventario vero e' uno
    solo, dentro la cartella condivisa, e non una copia sulla postazione di
    ciascun tecnico. Le installazioni piu' vecchie, che lo tenevano accanto
    all'eseguibile, continuano a funzionare.
    """
    nuovo = os.path.join(production_dir(), DATA_FILE_NAME)
    if os.path.exists(nuovo):
        return nuovo
    vecchio = os.path.join(app_dir(), DATA_FILE_NAME)
    if os.path.exists(vecchio):
        return vecchio
    return nuovo


def _read_config_path(path):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh).get("data_path") or None
    except (OSError, ValueError):
        return None


def load_language():
    """La lingua e' una preferenza personale: sta accanto al programma o nel profilo."""
    for sorgente in (local_config_path(), user_config_path()):
        try:
            with open(sorgente, "r", encoding="utf-8") as fh:
                lingua = json.load(fh).get("lingua")
        except (OSError, ValueError):
            continue
        if lingua:
            return lingua
    return "it"


def save_language(lingua):
    for target in (local_config_path(), user_config_path()):
        try:
            try:
                with open(target, "r", encoding="utf-8") as fh:
                    dati = json.load(fh)
            except (OSError, ValueError):
                dati = {}
            dati["lingua"] = lingua
            with open(target, "w", encoding="utf-8") as fh:
                json.dump(dati, fh, indent=2, ensure_ascii=False)
            return target
        except OSError:
            continue
    return None


def configured_data_path():
    """L'inventario che questa installazione ha il compito di aprire.

    Ritorna (percorso, sorgente) e non guarda se il file esista: dice solo che
    cosa e' stato deciso. Serve a distinguere due situazioni che non vanno mai
    confuse - "nessuno ha ancora detto dove sta l'inventario", e "l'inventario
    e' sulla share ma adesso la share non risponde". Nella seconda, creare un
    inventario nuovo in locale significherebbe far lavorare un tecnico su una
    copia che nessun altro vede.
    """
    path = os.environ.get("INVENTARIO_FILE")
    if path:
        return path, "variabile d'ambiente INVENTARIO_FILE"
    for source in (local_config_path(), user_config_path()):
        saved = _read_config_path(source)
        if saved:
            return saved, source
    return None, None


def save_configured_data_path(path):
    """Scrive nella configurazione quale inventario deve aprire il programma.

    Ritorna il file scritto, o None se non si e' potuto scrivere da nessuna
    parte. Si scrive accanto al programma, cosi' la cartella si copia gia'
    configurata su tutte le postazioni.
    """
    for target in (local_config_path(), user_config_path()):
        try:
            try:
                with open(target, "r", encoding="utf-8") as fh:
                    dati = json.load(fh)
            except (OSError, ValueError):
                dati = {}
            dati["data_path"] = path
            with open(target, "w", encoding="utf-8") as fh:
                json.dump(dati, fh, indent=2, ensure_ascii=False)
            return target
        except OSError:
            continue
    return None


def load_data_path():
    """Percorso del file dati da usare, oppure None se va ancora scelto."""
    forzato = os.environ.get("INVENTARIO_FILE")
    if forzato:
        return forzato          # scelta esplicita: vale anche se non esiste ancora
    scelto, _ = configured_data_path()
    if scelto and os.path.exists(scelto):
        return scelto
    accanto = default_data_path()
    if os.path.exists(accanto):
        return accanto
    return None


def save_data_path(path):
    """Ricorda la scelta; se la cartella del programma non e' scrivibile,
    ripiega sul profilo dell'utente. La lingua gia' scelta non si perde."""
    return save_configured_data_path(path)


def shared_config_path(data_path):
    folder = os.path.dirname(os.path.abspath(data_path))
    return os.path.join(folder, "inventario_impostazioni.json")


def deleted_path(data_path):
    """Il file degli eliminati di recente, accanto ai dati e quindi condiviso.

    Sta sulla rete come i dati: un dispositivo eliminato da un tecnico deve
    poterlo ripescare un altro. Non e' dentro l'.xlsx perche' non e' inventario:
    quei record non devono comparire in nessuna esportazione ne' in nessuna
    ricerca.
    """
    folder = os.path.dirname(os.path.abspath(data_path))
    return os.path.join(folder, "inventario_eliminati.json")


def load_shared_config(data_path):
    """Impostazioni condivise; ritorna sempre un dizionario valido."""
    cfg = {"rooms": list(DEFAULT_ROOMS), "types": list(DEFAULT_TYPES),
           "loan_rooms": list(DEFAULT_LOAN_ROOMS),
           "iphone_room": DEFAULT_IPHONE_ROOM,
           "states": list(store.STATI)}
    try:
        with open(shared_config_path(data_path), "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, ValueError):
        data = {}
    for key in ("rooms", "types", "loan_rooms", "states"):
        if key not in data:
            continue          # voce assente: restano i valori predefiniti
        values = [str(v).strip() for v in data.get(key) or [] if str(v).strip()]
        if values or key == "loan_rooms":
            # solo loan_rooms puo' essere svuotata di proposito: significa
            # "nessuna stanza gestisce prestiti"
            cfg[key] = values
    stanza = str(data.get("iphone_room") or cfg["iphone_room"]).strip()
    # Se la stanza degli iPhone non esiste (o e' stata rinominata) si ripiega
    # sulla prima, cosi' i telefoni restano sempre in una stanza valida.
    cfg["iphone_room"] = stanza if stanza in cfg["rooms"] else cfg["rooms"][0]
    return cfg


def save_shared_config(data_path, cfg):
    tmp = shared_config_path(data_path) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, indent=2, ensure_ascii=False)
    os.replace(tmp, shared_config_path(data_path))
