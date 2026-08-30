"""Archivio dati dell'inventario.

I dati risiedono in un unico file .xlsx sulla cartella di rete: e' allo stesso
tempo il database e l'inventario apribile in Excel.

Sicurezza sugli accessi concorrenti:
  * ogni modifica avviene dentro un lock esclusivo (file .lock creato con
    O_CREAT|O_EXCL, che su SMB e' atomico);
  * dentro il lock i dati vengono sempre riletti da disco e l'operazione
    (aggiunta / modifica / eliminazione) viene riapplicata sui dati freschi,
    quindi due utenti che lavorano su schede diverse non si sovrascrivono;
  * la scrittura avviene su un file temporaneo nella stessa cartella e poi
    os.replace(), quindi il file non resta mai a meta'.
"""

import calendar
import getpass
import json
import os
import platform
import socket
import time
import uuid
from datetime import datetime

from openpyxl import Workbook, load_workbook

SHEET_NAME = "Inventario"

FIELDS = ["asset_tag", "tipo", "modello", "seriale", "imei", "restituito_da",
          "stanza", "stato", "prestato_a", "prestato_il", "spedito_il", "note"]

# Per un iPhone l'identita' e' l'IMEI, non l'asset tag aziendale.
TIPO_IPHONE = "iphone"

DISPONIBILE = "Disponibile"
NON_DISPONIBILE = "Non disponibile"
# Gli iPhone in nostro possesso sono sempre in attesa di essere rispediti.
DA_RISPEDIRE = "Da Rispedire"
# ...finche' non partono davvero per il servizio telefonia.
SPEDITO = "Spedito al servizio telefonia"

# Un dispositivo spedito resta consultabile in inventario per tre mesi.
MESI_CONSERVAZIONE = 3

TESTO_SPEDIZIONE = (
    "Il dispositivo e' stato rispedito al servizio telefonia il %s. "
    "Resta in inventario per consultazione fino al %s, data dalla quale potra' "
    "essere eliminato."
)

# Stati scegliibili a mano. Gli altri due sono automatici: NON_DISPONIBILE
# mentre c'e' un prestito in corso, DA_RISPEDIRE per gli iPhone.
STATI = [
    DISPONIBILE,
    "In attesa ritiro",
    "Guasto in attesa tecnico",
    "Da rebuildare",
    "Controllare",
]
AUDIT_FIELDS = ["modificato_il", "modificato_da"]
ALL_FIELDS = FIELDS + AUDIT_FIELDS

HEADERS = {
    "asset_tag": "Asset Tag",
    "tipo": "Tipo",
    "modello": "Modello",
    "seriale": "Numero di serie",
    "imei": "IMEI",
    "restituito_da": "Restituito da",
    "stanza": "Stanza",
    "stato": "Stato",
    "prestato_a": "In prestito a",
    "prestato_il": "Prestato il",
    "spedito_il": "Spedito il",
    "note": "Note",
    "modificato_il": "Ultima modifica",
    "modificato_da": "Modificato da",
}

# Intestazioni accettate in importazione (minuscole, senza spazi doppi).
HEADER_ALIASES = {
    "asset_tag": ["asset tag", "assettag", "asset", "tag", "etichetta", "inventario"],
    "tipo": ["tipo", "tipologia", "categoria", "device type", "type"],
    "modello": ["modello", "model", "descrizione", "dispositivo", "device"],
    "seriale": ["numero di serie", "seriale", "serial", "serial number", "s/n",
                "sn", "matricola", "service tag"],
    "imei": ["imei", "imei/meid", "meid", "codice imei"],
    "restituito_da": ["restituito da", "proprietario", "consegnato da",
                      "riconsegnato da", "owner"],
    "stanza": ["stanza", "room", "locale", "ubicazione", "posizione", "location"],
    "stato": ["stato", "status", "disponibilita", "disponibilita'"],
    "prestato_a": ["in prestito a", "prestato a", "prestito", "assegnato a",
                   "utilizzatore", "borrower", "assigned to"],
    "prestato_il": ["prestato il", "data prestito", "in prestito dal",
                    "loan date", "borrowed on"],
    "spedito_il": ["spedito il", "data spedizione", "rispedito il",
                   "shipped on"],
    "note": ["note", "nota", "commenti", "notes"],
    "modificato_il": ["ultima modifica", "modificato il", "data"],
    "modificato_da": ["modificato da", "utente"],
}

# Parole ignorate quando si ricavano i tag dai nomi delle stanze.
_PAROLE_VUOTE = {"DEL", "DELLA", "DEI", "DELLE", "DI", "DA", "IN", "PER", "LA", "IL"}


def tag_stanze(rooms):
    """Tag riconosciuti in importazione, ricavati dai nomi delle stanze.

    Per ogni stanza vale il nome completo, piu' ogni singola parola che sia
    inequivocabile: "Digital Kiosk" si scrive anche solo KIOSK, "Site Services
    BAU" anche solo BAU, "Magazzino Disaster Recovery" anche solo DISASTER.
    """
    tags = {}
    parole = {}
    for stanza in rooms:
        nome = clean(stanza).upper()
        if not nome:
            continue
        tags[nome] = stanza
        for parola in nome.split():
            if len(parola) < 3 or parola in _PAROLE_VUOTE:
                continue
            parole.setdefault(parola, set()).add(stanza)
    for parola, stanze in parole.items():
        if len(stanze) == 1 and parola not in tags:
            tags[parola] = list(stanze)[0]
    return tags


def riga_tag(row, tags):
    """Se la riga e' un separatore di stanza ritorna la stanza, altrimenti None.

    Un separatore e' una riga con una sola cella scritta, che contiene il tag.
    """
    valori = [clean(c) for c in (row or ()) if clean(c)]
    if len(valori) != 1:
        return None
    testo = valori[0].upper().strip(" :-\u2013\u2014.\t")
    return tags.get(testo)


LOCK_TIMEOUT = 20.0     # secondi di attesa prima di rinunciare
LOCK_STALE_AFTER = 120  # secondi dopo i quali un lock e' considerato abbandonato


class InventoryError(Exception):
    """Errore mostrabile all'utente."""


class LockBusy(InventoryError):
    pass


class BloccoIphoneNonSpedito(InventoryError):
    """Un iPhone si elimina solo dopo essere stato rispedito."""

    def __init__(self, item):
        self.item = item
        InventoryError.__init__(
            self,
            "%s non e' ancora stato rispedito al servizio telefonia\n"
            "e non puo' essere eliminato dall'inventario.\n\n"
            "Registra prima la spedizione con il pulsante Conferma spedizione,\n"
            "nel contenitore Iphone. Da quel momento restera' consultabile per\n"
            "%d mesi, e poi potra' essere eliminato."
            % (item["asset_tag"], MESI_CONSERVAZIONE))


class BloccoConservazione(InventoryError):
    """Un dispositivo spedito non si elimina prima dei tre mesi di conservazione."""

    def __init__(self, item, sblocco):
        self.item = item
        self.sblocco = sblocco
        InventoryError.__init__(
            self,
            "%s e' stato rispedito al servizio telefonia il %s.\n\n"
            "Va conservato in inventario per consultazione: potrai eliminarlo\n"
            "a partire dal %s."
            % (item["asset_tag"], item["spedito_il"], sblocco.strftime("%d/%m/%Y")))


def current_user():
    try:
        user = getpass.getuser()
    except Exception:
        user = "sconosciuto"
    try:
        host = socket.gethostname()
    except Exception:
        host = platform.node() or "?"
    return "%s@%s" % (user, host)


def norm_tag(value):
    return " ".join(str(value or "").split()).upper()


def clean(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y %H:%M")
    return " ".join(str(value).split())


def is_iphone(tipo):
    return clean(tipo).lower() == TIPO_IPHONE


def new_item(asset_tag="", tipo="", modello="", seriale="", stanza="", note="",
             prestato_a="", prestato_il="", imei="", restituito_da="", stato="",
             spedito_il=""):
    item = {
        "asset_tag": norm_tag(asset_tag),
        "tipo": clean(tipo),
        "modello": clean(modello),
        "seriale": clean(seriale),
        "imei": clean(imei),
        "restituito_da": clean(restituito_da),
        "stanza": clean(stanza),
        "stato": clean(stato),
        "prestato_a": clean(prestato_a),
        "prestato_il": clean(prestato_il),
        "spedito_il": clean(spedito_il),
        "note": clean(note),
        "modificato_il": "",
        "modificato_da": "",
    }
    return normalize_state(normalize_identity(item))


def normalize_identity(item):
    """Per un iPhone l'IMEI e' l'identificativo: se manca l'asset tag, lo sostituisce."""
    if not item.get("asset_tag") and item.get("imei"):
        item["asset_tag"] = norm_tag(item["imei"])
    return item


def normalize_iphone(item):
    """Un iPhone ha l'IMEI e basta: niente asset tag, seriale o prestiti.

    L'asset tag resta valorizzato in memoria come chiave interna - serve a
    identificare la riga in ogni operazione - ma vale sempre l'IMEI e non viene
    mai mostrato ne' scritto sul file.
    """
    if is_iphone(item.get("tipo")):
        item["seriale"] = ""
        item["prestato_a"] = ""
        item["prestato_il"] = ""
        if item.get("imei"):
            item["asset_tag"] = norm_tag(item["imei"])
    return item


def valore_visibile(item, field):
    """Il valore da mostrare o scrivere: per un iPhone l'asset tag resta vuoto."""
    if field == "asset_tag" and is_iphone(item.get("tipo")):
        return ""
    return item.get(field, "")


def normalize_state(item, stati=None):
    """Mette a posto lo stato.

    Due casi sono automatici e vincono sempre: un iPhone e' "Da Rispedire", un
    dispositivo in prestito e' "Non disponibile". Negli altri casi si tiene lo
    stato scelto dall'utente, purche' sia fra quelli previsti.
    """
    ammessi = list(stati or STATI)
    normalize_iphone(item)
    if is_iphone(item.get("tipo")):
        item["stato"] = SPEDITO if item.get("spedito_il") else DA_RISPEDIRE
    elif item.get("prestato_a"):
        item["stato"] = NON_DISPONIBILE
    elif clean(item.get("stato")) not in ammessi:
        item["stato"] = DISPONIBILE
    return item


def is_on_loan(item):
    return bool(item.get("prestato_a"))


def is_shipped(item):
    return bool(item.get("spedito_il"))


def _somma_mesi(quando, mesi):
    anno = quando.year + (quando.month - 1 + mesi) // 12
    mese = (quando.month - 1 + mesi) % 12 + 1
    giorno = min(quando.day, calendar.monthrange(anno, mese)[1])
    return quando.replace(year=anno, month=mese, day=giorno)


def eliminabile_dal(item):
    """Data dalla quale un dispositivo spedito puo' essere eliminato, o None."""
    quando = clean(item.get("spedito_il"))
    if not quando:
        return None
    for formato in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return _somma_mesi(datetime.strptime(quando, formato), MESI_CONSERVAZIONE)
        except ValueError:
            continue
    return None


def puo_essere_eliminato(item, adesso=None):
    """(True/False, data di sblocco).

    Un iPhone non ancora rispedito non si elimina mai: va prima registrata la
    spedizione, e da quel momento restano tre mesi di conservazione. In quel
    caso la data di sblocco e' None, perche' non e' ancora determinabile.
    """
    if is_iphone(item.get("tipo")) and not is_shipped(item):
        return False, None
    sblocco = eliminabile_dal(item)
    if sblocco is None:
        return True, None
    return (adesso or datetime.now()) >= sblocco, sblocco


def testo_spedizione(item):
    """La frase da mostrare accanto a un dispositivo spedito."""
    sblocco = eliminabile_dal(item)
    if sblocco is None:
        return ""
    return TESTO_SPEDIZIONE % (item["spedito_il"], sblocco.strftime("%d/%m/%Y"))


class _Lock(object):
    """Lock esclusivo basato su un file accanto ai dati."""

    def __init__(self, data_path):
        self.path = os.path.join(
            os.path.dirname(os.path.abspath(data_path)),
            "." + os.path.basename(data_path) + ".lock",
        )
        self.fd = None

    def _holder(self):
        try:
            with open(self.path, "r", encoding="utf-8") as fh:
                return json.load(fh)
        except (OSError, ValueError):
            return {}

    def __enter__(self):
        deadline = time.time() + LOCK_TIMEOUT
        while True:
            try:
                self.fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                info = json.dumps(
                    {"utente": current_user(), "pid": os.getpid(), "ts": time.time()}
                )
                os.write(self.fd, info.encode("utf-8"))
                os.close(self.fd)
                self.fd = None
                return self
            except FileExistsError:
                holder = self._holder()
                # L'eta' si misura sul file, non sul contenuto: subito dopo la
                # creazione il file e' ancora vuoto e non va scambiato per
                # abbandonato.
                try:
                    age = time.time() - os.stat(self.path).st_mtime
                except OSError:
                    continue
                if age > LOCK_STALE_AFTER:
                    # Lock abbandonato (PC spento, crash): lo rimuoviamo.
                    try:
                        os.remove(self.path)
                    except OSError:
                        pass
                    continue
                if time.time() >= deadline:
                    raise LockBusy(
                        "L'inventario e' in uso da %s. Riprova tra qualche istante."
                        % (holder.get("utente") or "un altro utente")
                    )
                time.sleep(0.4)
            except OSError as exc:
                raise InventoryError(
                    "Impossibile accedere alla cartella di rete:\n%s" % exc
                )

    def __exit__(self, *exc):
        try:
            os.remove(self.path)
        except OSError:
            pass
        return False


class InventoryStore(object):
    def __init__(self, data_path, iphone_room=None, stati=None):
        self.path = os.path.abspath(data_path)
        # Stanza in cui gli iPhone stanno per forza; None disattiva il vincolo.
        self.iphone_room = iphone_room
        self.stati = list(stati or STATI)
        self.items = []
        self._stamp = None

    def _enforce_iphone_room(self, item):
        """Un iPhone appartiene sempre alla sua stanza, comunque lo si registri."""
        if self.iphone_room and is_iphone(item.get("tipo")):
            item["stanza"] = self.iphone_room
        return item

    # ------------------------------------------------------------ lettura

    def exists(self):
        return os.path.exists(self.path)

    def create_if_missing(self):
        if self.exists():
            return False
        folder = os.path.dirname(self.path)
        if folder and not os.path.isdir(folder):
            raise InventoryError("La cartella %s non esiste." % folder)
        self._write([])
        return True

    def _disk_stamp(self):
        try:
            st = os.stat(self.path)
        except OSError:
            return None
        return (st.st_mtime, st.st_size)

    def changed_on_disk(self):
        """True se qualcun altro ha scritto il file dopo la nostra lettura."""
        return self._disk_stamp() != self._stamp

    def load(self):
        self.items = self._read()
        self._stamp = self._disk_stamp()
        return self.items

    def _read(self):
        if not self.exists():
            return []
        try:
            wb = load_workbook(self.path, read_only=True, data_only=True)
        except Exception as exc:
            raise InventoryError("Impossibile leggere %s:\n%s" % (self.path, exc))
        try:
            ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.worksheets[0]
            rows = ws.iter_rows(values_only=True)
            try:
                header = next(rows)
            except StopIteration:
                return []
            mapping = map_headers(header)
            items = []
            for row in rows:
                item = _row_to_item(row, mapping)
                if item:
                    items.append(item)
            return items
        finally:
            wb.close()

    # ---------------------------------------------------------- scrittura

    def _write(self, items):
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append([HEADERS[f] for f in ALL_FIELDS])
        for item in items:
            ws.append([valore_visibile(item, f) for f in ALL_FIELDS])
        _style_sheet(ws, len(items))
        tmp = "%s.tmp-%d-%s" % (self.path, os.getpid(), uuid.uuid4().hex[:8])
        try:
            wb.save(tmp)
            os.replace(tmp, self.path)
        except Exception as exc:
            try:
                os.remove(tmp)
            except OSError:
                pass
            raise InventoryError("Impossibile salvare %s:\n%s" % (self.path, exc))
        finally:
            wb.close()

    def _apply(self, operation):
        """Esegue `operation(items)` su dati freschi, dentro il lock, e salva.

        `operation` puo' sollevare InventoryError per annullare tutto.
        Ritorna il valore restituito da `operation`.
        """
        with _Lock(self.path):
            items = self._read()
            result = operation(items)
            items.sort(key=lambda it: (it.get("stanza", ""), it.get("asset_tag", "")))
            self._write(items)
            self.items = items
            self._stamp = self._disk_stamp()
        return result

    # ---------------------------------------------------------- operazioni

    def add(self, item):
        item = dict(item)
        item["asset_tag"] = norm_tag(item.get("asset_tag"))
        if not item["asset_tag"]:
            raise InventoryError("L'asset tag e' obbligatorio.")

        def op(items):
            if any(it["asset_tag"] == item["asset_tag"] for it in items):
                raise InventoryError(
                    "L'asset tag %s e' gia' presente nell'inventario." % item["asset_tag"]
                )
            self._enforce_iphone_room(item)
            normalize_state(item, self.stati)
            _stamp_item(item)
            items.append(item)

        return self._apply(op)

    def update(self, old_tag, item):
        old_tag = norm_tag(old_tag)
        item = dict(item)
        item["asset_tag"] = norm_tag(item.get("asset_tag"))
        if not item["asset_tag"]:
            raise InventoryError("L'asset tag e' obbligatorio.")

        def op(items):
            index = _index_of(items, old_tag)
            if index is None:
                raise InventoryError(
                    "L'articolo %s non esiste piu': e' stato eliminato da un altro utente."
                    % old_tag
                )
            if item["asset_tag"] != old_tag and _index_of(items, item["asset_tag"]) is not None:
                raise InventoryError(
                    "L'asset tag %s e' gia' presente nell'inventario." % item["asset_tag"]
                )
            self._enforce_iphone_room(item)
            normalize_state(item, self.stati)
            _stamp_item(item)
            items[index] = item

        return self._apply(op)

    def delete(self, tags):
        """Elimina i dispositivi, salvo quelli spediti da meno di tre mesi."""
        wanted = set(norm_tag(t) for t in tags)

        def op(items):
            for it in items:
                if it["asset_tag"] not in wanted:
                    continue
                libero, sblocco = puo_essere_eliminato(it)
                if libero:
                    continue
                if sblocco is None:
                    raise BloccoIphoneNonSpedito(it)
                raise BloccoConservazione(it, sblocco)
            before = len(items)
            items[:] = [it for it in items if it["asset_tag"] not in wanted]
            return before - len(items)

        return self._apply(op)

    def ship(self, tag):
        """Registra la spedizione al servizio telefonia."""
        tag = norm_tag(tag)

        def op(items):
            index = _index_of(items, tag)
            if index is None:
                raise InventoryError("Il dispositivo %s non esiste piu' nell'inventario." % tag)
            item = items[index]
            if not is_iphone(item.get("tipo")):
                raise InventoryError(
                    "La spedizione al servizio telefonia riguarda solo gli iPhone.")
            if is_shipped(item):
                raise InventoryError(
                    "%s risulta gia' spedito il %s." % (tag, item["spedito_il"]))
            item["spedito_il"] = datetime.now().strftime("%d/%m/%Y %H:%M")
            normalize_state(item, self.stati)
            _stamp_item(item)
            return testo_spedizione(item)

        return self._apply(op)

    def move_to_room(self, tags, room):
        """Sposta i dispositivi selezionati.

        Ritorna (spostati, iphone_lasciati_fermi): gli iPhone non si spostano.
        """
        wanted = set(norm_tag(t) for t in tags)
        room = clean(room)

        def op(items):
            moved = bloccati = 0
            for it in items:
                if it["asset_tag"] not in wanted:
                    continue
                if self.iphone_room and is_iphone(it.get("tipo")):
                    if it.get("stanza") != self.iphone_room:
                        it["stanza"] = self.iphone_room     # rimette a posto
                        _stamp_item(it)
                    bloccati += 1
                    continue
                if it.get("stanza") != room:
                    it["stanza"] = room
                    _stamp_item(it)
                    moved += 1
            return moved, bloccati

        return self._apply(op)

    def lend(self, tag, person):
        """Registra il prestito di un dispositivo a una persona."""
        tag = norm_tag(tag)
        person = clean(person)
        if not person:
            raise InventoryError("Indica il nome della persona a cui presti il dispositivo.")

        def op(items):
            index = _index_of(items, tag)
            if index is None:
                raise InventoryError("Il dispositivo %s non esiste piu' nell'inventario." % tag)
            item = items[index]
            if is_iphone(item.get("tipo")):
                raise InventoryError("Gli iPhone non vengono dati in prestito.")
            if is_on_loan(item):
                raise InventoryError(
                    "%s risulta gia' in prestito a %s dal %s."
                    % (tag, item["prestato_a"], item["prestato_il"]))
            item["prestato_a"] = person
            item["prestato_il"] = datetime.now().strftime("%d/%m/%Y %H:%M")
            normalize_state(item, self.stati)
            _stamp_item(item)
            return item["prestato_il"]

        return self._apply(op)

    def give_back(self, tag):
        """Chiude il prestito e rimette il dispositivo fra i disponibili."""
        tag = norm_tag(tag)

        def op(items):
            index = _index_of(items, tag)
            if index is None:
                raise InventoryError("Il dispositivo %s non esiste piu' nell'inventario." % tag)
            item = items[index]
            if not is_on_loan(item):
                raise InventoryError("%s non risulta in prestito." % tag)
            person = item["prestato_a"]
            item["prestato_a"] = ""
            item["prestato_il"] = ""
            item["stato"] = DISPONIBILE
            normalize_state(item, self.stati)
            _stamp_item(item)
            return person

        return self._apply(op)

    def set_stato(self, tag, stato):
        """Cambia lo stato di un dispositivo (tendina nell'elenco)."""
        tag = norm_tag(tag)
        stato = clean(stato)
        if stato not in self.stati:
            raise InventoryError("Stato non previsto: %s." % stato)

        def op(items):
            index = _index_of(items, tag)
            if index is None:
                raise InventoryError("Il dispositivo %s non esiste piu' nell'inventario." % tag)
            item = items[index]
            if is_iphone(item.get("tipo")):
                atteso = SPEDITO if is_shipped(item) else DA_RISPEDIRE
                raise InventoryError(
                    "Lo stato degli iPhone e' sempre \"%s\" e non si cambia." % atteso)
            if is_on_loan(item):
                raise InventoryError(
                    "%s e' in prestito a %s: registra prima il rientro."
                    % (tag, item["prestato_a"]))
            if item.get("stato") == stato:
                return False
            item["stato"] = stato
            _stamp_item(item)
            return True

        return self._apply(op)

    def set_note(self, tag, note):
        """Aggiorna soltanto le note (modifica al volo dall'elenco)."""
        tag = norm_tag(tag)
        note = clean(note)

        def op(items):
            index = _index_of(items, tag)
            if index is None:
                raise InventoryError("Il dispositivo %s non esiste piu' nell'inventario." % tag)
            if items[index].get("note", "") == note:
                return False
            items[index]["note"] = note
            _stamp_item(items[index])
            return True

        return self._apply(op)

    def import_items(self, incoming, mode="merge"):
        """mode: 'merge' aggiorna/aggiunge, 'replace' sostituisce tutto."""

        def op(items):
            added = updated = 0
            if mode == "replace":
                # Gli iPhone non arrivano mai da un'importazione: vanno tenuti,
                # altrimenti una sostituzione li cancellerebbe.
                items[:] = [it for it in items if is_iphone(it.get("tipo"))]
            index = {it["asset_tag"]: i for i, it in enumerate(items)}
            for raw in incoming:
                item = dict(raw)
                item["asset_tag"] = norm_tag(item.get("asset_tag"))
                if not item["asset_tag"]:
                    continue
                normalize_identity(item)
                self._enforce_iphone_room(item)
                normalize_state(item, self.stati)
                _stamp_item(item)
                if item["asset_tag"] in index:
                    items[index[item["asset_tag"]]] = item
                    updated += 1
                else:
                    index[item["asset_tag"]] = len(items)
                    items.append(item)
                    added += 1
            return added, updated

        return self._apply(op)


# ------------------------------------------------------------- utilita'


def _index_of(items, tag):
    for i, it in enumerate(items):
        if it["asset_tag"] == tag:
            return i
    return None


def _stamp_item(item):
    # con i secondi due inserimenti nello stesso minuto restano in ordine
    item["modificato_il"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    item["modificato_da"] = current_user()


def map_headers(header_row):
    """Mappa indice colonna -> campo, riconoscendo le intestazioni note."""
    mapping = {}
    for idx, cell in enumerate(header_row or ()):
        name = clean(cell).lower()
        if not name:
            continue
        for field, aliases in HEADER_ALIASES.items():
            if field in mapping.values():
                continue
            if name in aliases:
                mapping[idx] = field
                break
    return mapping


def _row_to_item(row, mapping):
    item = new_item()
    for idx, field in mapping.items():
        if idx < len(row):
            item[field] = norm_tag(row[idx]) if field == "asset_tag" else clean(row[idx])
    normalize_identity(item)
    normalize_state(item)
    return item if item["asset_tag"] else None


def rows_from_workbook(path, rooms=None):
    """Legge un file .xlsx/.xlsm esterno.

    Ritorna (items, esito), dove esito e' un dizionario con le righe scartate,
    quelle che hanno preso la stanza da un separatore, gli iPhone ignorati, le
    colonne del file che non sono state riconosciute e le righe senza modello.

    Se nel foglio compaiono righe-separatore con il nome (o l'abbreviazione) di
    una stanza, tutte le righe successive fino al separatore seguente vengono
    assegnate a quella stanza: e' il modo per dividere per stanza un unico
    inventario. Gli iPhone non si importano: sono gestiti solo a mano.
    """
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise InventoryError("Impossibile leggere il file:\n%s" % exc)
    try:
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.worksheets[0]
        rows = ws.iter_rows(values_only=True)
        try:
            header = next(rows)
        except StopIteration:
            return [], 0
        mapping = map_headers(header)
        if not ({"asset_tag", "imei"} & set(mapping.values())):
            raise InventoryError(
                "Nel file non e' stata trovata la colonna \"Asset Tag\" (o \"IMEI\").\n"
                "La prima riga deve contenere le intestazioni delle colonne."
            )
        tags = tag_stanze(rooms or [])
        items = []
        esito = {"scartate": 0, "da_tag": 0, "iphone": 0, "senza_modello": 0,
                 "colonne_ignorate": [clean(c) for i, c in enumerate(header)
                                      if clean(c) and i not in mapping]}
        stanza_corrente = None
        for row in rows:
            if row is None or all(c is None or clean(c) == "" for c in row):
                continue
            stanza = riga_tag(row, tags)
            if stanza is not None:
                stanza_corrente = stanza
                continue
            item = _row_to_item(row, mapping)
            if not item:
                esito["scartate"] += 1
                continue
            if is_iphone(item.get("tipo")):
                esito["iphone"] += 1
                continue
            if stanza_corrente:
                item["stanza"] = stanza_corrente
                esito["da_tag"] += 1
            if not item.get("modello"):
                esito["senza_modello"] += 1
            items.append(item)
        return items, esito
    finally:
        wb.close()


def _style_sheet(ws, row_count):
    """Formattazione minima del file dati (l'export di stampa e' piu' curato)."""
    from openpyxl.styles import Alignment, Font, PatternFill

    fill = PatternFill("solid", fgColor="1F4E79")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"
    widths = {"asset_tag": 18, "tipo": 12, "modello": 32, "seriale": 20,
              "imei": 20, "restituito_da": 22, "stanza": 24, "stato": 26,
              "prestato_a": 24, "prestato_il": 18, "spedito_il": 18, "note": 38,
              "modificato_il": 20, "modificato_da": 24}
    for i, field in enumerate(ALL_FIELDS, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = widths[field]
    if row_count:
        ws.auto_filter.ref = "A1:%s%d" % (
            ws.cell(row=1, column=len(ALL_FIELDS)).column_letter, row_count + 1)
