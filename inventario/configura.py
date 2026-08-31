"""Collega questa installazione all'inventario condiviso sulla share.

Si esegue una volta sola, e quello che scrive - `inventario_percorso.json`
accanto al programma - viaggia con la cartella: si configura una copia e la si
distribuisce gia' pronta a tutte le postazioni.

Non copia dati e non ne cancella: se sulla share l'inventario non c'e' ancora,
lo crea vuoto; se c'e', lo lascia esattamente com'e'.
"""

import os
import sys

from . import config
from .store import InventoryStore, InventoryError


def _chiedi(domanda):
    try:
        return input(domanda).strip().strip('"').strip()
    except (EOFError, KeyboardInterrupt):
        return ""


def percorso_inventario(cartella):
    """Da una cartella condivisa al file dell'inventario dentro `Produzione`.

    Accetta indifferentemente la cartella condivisa, la cartella `Produzione` o
    direttamente il file: chi la scrive a mano intende sempre la stessa cosa.
    """
    cartella = cartella.rstrip("\\/")
    if cartella.lower().endswith(".xlsx"):
        return cartella
    if os.path.basename(cartella).lower() == config.NOME_PRODUZIONE.lower():
        return os.path.join(cartella, config.DATA_FILE_NAME)
    return os.path.join(cartella, config.NOME_PRODUZIONE, config.DATA_FILE_NAME)


def collega(cartella, creare=True):
    """Punta questa installazione all'inventario indicato.

    Ritorna (percorso, gia_esisteva). Solleva InventoryError con un messaggio
    leggibile se la cartella non si raggiunge o non si puo' scrivere.
    """
    percorso = percorso_inventario(cartella)
    dentro = os.path.dirname(percorso)
    padre = os.path.dirname(dentro)
    if not os.path.isdir(padre):
        raise InventoryError(
            "La cartella condivisa non si raggiunge:\n  %s\n\n"
            "Aprila prima da Esplora risorse: se non si apre da li', non si\n"
            "apre nemmeno da qui. Controlla il percorso e la connessione." % padre)
    gia_esisteva = os.path.exists(percorso)
    if not gia_esisteva:
        if not creare:
            raise InventoryError("Sulla share non c'e' nessun inventario:\n  %s"
                                 % percorso)
        try:
            if not os.path.isdir(dentro):
                os.makedirs(dentro)
            InventoryStore(percorso).create_if_missing()
        except OSError as exc:
            raise InventoryError(
                "Sulla cartella condivisa non si puo' scrivere:\n  %s\n\n%s\n\n"
                "Serve il permesso di Modifica su quella cartella: chiedilo a\n"
                "chi amministra la share." % (dentro, exc))
    scritto = config.save_configured_data_path(percorso)
    if not scritto:
        raise InventoryError(
            "Non si e' potuto salvare la configurazione ne' accanto al\n"
            "programma ne' nel profilo utente.")
    return percorso, gia_esisteva


def main():
    print()
    print("=" * 68)
    print("  Collega questa postazione all'inventario condiviso")
    print("=" * 68)
    print()
    print("Serve il percorso della cartella condivisa in cui sta - o in cui")
    print("va creato - l'inventario di tutti. Per esempio:")
    print()
    print("   \\\\server\\Condivisa\\Inventario")
    print("   F:\\Inventario")
    print()
    attuale, sorgente = config.configured_data_path()
    if attuale:
        print("Adesso questa installazione apre:")
        print("   %s" % attuale)
        print("   (scritto in %s)" % sorgente)
        print()

    cartella = _chiedi("Cartella condivisa: ")
    if not cartella:
        print("\nAnnullato: non e' stato cambiato niente.")
        return 1
    try:
        percorso, gia_esisteva = collega(cartella)
    except InventoryError as exc:
        print()
        print("NON RIUSCITO")
        print()
        print(str(exc))
        print()
        return 1

    print()
    print("-" * 68)
    print("Fatto. Questa postazione ora apre l'inventario condiviso:")
    print()
    print("   %s" % percorso)
    print()
    if gia_esisteva:
        print("L'inventario era gia' li' e non e' stato toccato.")
    else:
        print("Sulla share non c'era ancora: e' stato creato vuoto.")
        print("Se hai gia' un inventario, chiudi il programma e copialo li'")
        print("sopra, con lo stesso nome %s." % config.DATA_FILE_NAME)
    print()
    print("Le copie di sicurezza andranno in Backup accanto ai dati, sulla")
    print("share: una sola serie per tutti.")
    print()
    print("Questa cartella del programma si puo' ora copiare cosi' com'e' su")
    print("tutte le altre postazioni: la configurazione viaggia con lei.")
    print("-" * 68)
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
