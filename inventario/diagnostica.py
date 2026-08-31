"""Raccolta di informazioni sull'installazione, da mandare a chi assiste.

Non tocca niente: legge, prova a importare il file di prova in memoria e
scrive un rapporto. Serve a rispondere alle domande che, davanti a un difetto,
non si possono indovinare da lontano: quale copia del programma sta girando,
dove tiene i dati, quali stanze conosce, cosa legge davvero dal foglio Excel e
dove sta la scrivania dell'utente.
"""

import os
import sys
from datetime import datetime


def _sezione(righe, titolo):
    righe.append("")
    righe.append("--- %s " % titolo + "-" * max(0, 60 - len(titolo)))


def _desktop():
    """Il desktop come lo intende Windows, non come lo si costruisce a mano."""
    tentativi = []
    try:
        import ctypes
        from ctypes import wintypes
        buf = ctypes.create_unicode_buffer(1024)
        # CSIDL_DESKTOPDIRECTORY = 0x10: la cartella vera, anche con OneDrive
        ctypes.windll.shell32.SHGetFolderPathW(None, 0x10, None, 0, buf)
        tentativi.append(("SHGetFolderPath", buf.value))
    except Exception as exc:
        tentativi.append(("SHGetFolderPath", "non disponibile (%s)" % exc))
    profilo = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    tentativi.append(("USERPROFILE\\Desktop", os.path.join(profilo, "Desktop")))
    uno = os.environ.get("OneDrive")
    if uno:
        tentativi.append(("OneDrive\\Desktop", os.path.join(uno, "Desktop")))
    return tentativi


def _misura_interfaccia(percorso):
    """Apre l'interfaccia per un istante e misura la tabella.

    Serve a rispondere alla domanda che da lontano non si puo' indovinare: la
    tabella e' vuota perche' non ci sono righe, o perche' e' alta zero pixel?
    Con i font e l'ingrandimento di Windows le proporzioni non sono quelle del
    computer di chi sviluppa.
    """
    righe = []
    try:
        import tkinter as tk
        from .ui import App
    except Exception as exc:
        return ["interfaccia non disponibile: %r" % exc]

    esito = []

    def misura(app):
        try:
            app.update()
            esito.append("scala di Windows   : %.2f" % app.tk.call("tk", "scaling"))
            esito.append("schermo            : %dx%d"
                         % (app.winfo_screenwidth(), app.winfo_screenheight()))
            esito.append("finestra           : %dx%d"
                         % (app.winfo_width(), app.winfo_height()))
            from tkinter import ttk
            altezza_riga = ttk.Style(app).lookup("Inv.Treeview", "rowheight")
            esito.append("altezza di una riga: %s" % altezza_riga)
            tree = app.tree
            wrap = tree.master
            esito.append("cornice tabella    : %dx%d"
                         % (wrap.winfo_width(), wrap.winfo_height()))
            esito.append("tabella            : %dx%d"
                         % (tree.winfo_width(), tree.winfo_height()))
            esito.append("righe nella tabella: %d" % len(tree.get_children()))
            esito.append("tabella visibile   : %s" % bool(tree.winfo_ismapped()))
            larghezza_colonne = sum(int(tree.column(c, "width"))
                                    for c in tree.cget("columns"))
            esito.append("larghezza colonne  : %d (la tabella ne mostra %d)"
                         % (larghezza_colonne, tree.winfo_width()))
            if tree.winfo_height() < 40:
                esito.append("")
                esito.append("  ESITO: la tabella e' alta %d pixel, troppo poco per"
                             % tree.winfo_height())
                esito.append("  mostrare anche una sola riga. I dispositivi ci sono,")
                esito.append("  ma non c'e' spazio per disegnarli: e' questo il")
                esito.append("  motivo dello schermo bianco.")
            elif not tree.get_children():
                esito.append("")
                esito.append("  ESITO: la tabella ha spazio ma non contiene righe.")
            else:
                esito.append("")
                esito.append("  ESITO: la tabella ha spazio e contiene righe.")
        except Exception as exc:
            esito.append("misura non riuscita: %r" % exc)
        finally:
            try:
                app.destroy()
            except Exception:
                pass

    try:
        app = App(percorso)
        app._initial_load()
        app.after(400, lambda: misura(app))
        # se qualcosa si inceppa la finestra si chiude comunque
        app.after(8000, lambda: app.destroy())
        app.mainloop()
    except Exception as exc:
        return ["apertura dell'interfaccia non riuscita: %r" % exc]
    return esito or ["nessuna misura raccolta"]


def raccogli():
    from . import __version__, config
    from .store import (InventoryStore, righe_separatore,
                        rows_from_workbook)

    righe = ["RAPPORTO DI DIAGNOSTICA - Inventario Site Services",
             "generato il %s" % datetime.now().strftime("%d/%m/%Y %H:%M:%S")]

    _sezione(righe, "programma")
    righe.append("versione            : %s" % __version__)
    righe.append("cartella programma  : %s" % config.app_dir())
    righe.append("file di avvio       : %s" % os.path.abspath(sys.argv[0] or "?"))
    righe.append("python              : %s" % sys.executable)
    righe.append("versione python     : %s" % sys.version.split()[0])
    righe.append("questo modulo       : %s" % os.path.abspath(__file__))

    _sezione(righe, "cosa c'e' nella cartella del programma")
    try:
        for nome in sorted(os.listdir(config.app_dir())):
            pieno = os.path.join(config.app_dir(), nome)
            righe.append("  %-38s %s" % (nome, "<cartella>" if os.path.isdir(pieno)
                                         else "%d byte" % os.path.getsize(pieno)))
    except OSError as exc:
        righe.append("  non elencabile: %s" % exc)
    _sezione(righe, "dati")
    percorso = config.load_data_path()
    if percorso:
        righe.append("percorso inventario : %s" % percorso)
    else:
        percorso = config.default_data_path()
        righe.append("percorso inventario : nessuno ancora scelto")
        righe.append("verra' creato in    : %s" % percorso)
    righe.append("esiste              : %s" % os.path.exists(percorso or ""))
    if percorso and os.path.exists(percorso):
        righe.append("dimensione          : %d byte" % os.path.getsize(percorso))
        righe.append("modificato          : %s" % datetime.fromtimestamp(
            os.path.getmtime(percorso)).strftime("%d/%m/%Y %H:%M:%S"))
    righe.append("cartella copie      : %s" % config.backup_dir())
    cartella = os.path.dirname(percorso or "") or "."
    righe.append("si puo' scrivere    : %s" % os.access(cartella, os.W_OK))

    _sezione(righe, "impostazioni condivise")
    cfg = config.load_shared_config(percorso)
    righe.append("file impostazioni   : %s" % config.shared_config_path(percorso))
    righe.append("esiste              : %s"
                 % os.path.exists(config.shared_config_path(percorso)))
    for chiave in ("rooms", "types", "loan_rooms", "iphone_room"):
        righe.append("%-20s: %r" % (chiave, cfg.get(chiave)))

    _sezione(righe, "inventario attuale")
    try:
        store = InventoryStore(percorso, iphone_room=cfg.get("iphone_room"),
                               stati=cfg.get("states"))
        items = store.load()
        righe.append("dispositivi         : %d" % len(items))
        conteggio = {}
        for it in items:
            conteggio[it.get("stanza") or "(senza stanza)"] = \
                conteggio.get(it.get("stanza") or "(senza stanza)", 0) + 1
        for stanza, quanti in sorted(conteggio.items()):
            righe.append("  %-34s %d" % (stanza, quanti))
        sospetti = righe_separatore(items, cfg.get("rooms"))
        if sospetti:
            righe.append("")
            righe.append("  ATTENZIONE: questo file NON e' un inventario, e' un foglio")
            righe.append("  da importare. Contiene le righe separatore di stanza")
            righe.append("  (%s)," % ", ".join(sospetti[:3]))
            righe.append("  che qui compaiono come se fossero dispositivi: per questo")
            righe.append("  le stanze restano vuote.")
    except Exception as exc:
        righe.append("lettura non riuscita: %r" % exc)

    _sezione(righe, "lettura del file di prova")
    prova = os.path.join(config.app_dir(), "Collaudo", "Inventario_di_prova.xlsx")
    righe.append("file                : %s" % prova)
    righe.append("esiste              : %s" % os.path.exists(prova))
    if os.path.exists(prova):
        righe.append("dimensione          : %d byte" % os.path.getsize(prova))
        try:
            letti, esito = rows_from_workbook(prova, cfg.get("rooms"))
            righe.append("righe lette         : %d" % len(letti))
            righe.append("stanze riconosciute : %r" % esito.get("stanze_trovate"))
            righe.append("righe con stanza    : %d" % esito.get("da_tag", 0))
            righe.append("colonne ignorate    : %r" % esito.get("colonne_ignorate"))
            righe.append("prime righe lette   :")
            for it in letti[:4]:
                righe.append("   %-14s %-8s %-30s -> %s"
                             % (it["asset_tag"], it.get("tipo", ""),
                                (it.get("modello") or "")[:30], it.get("stanza") or "?"))
            if not esito.get("stanze_trovate"):
                righe.append("")
                righe.append("  ESITO: nessun separatore riconosciuto. E' questo il")
                righe.append("  motivo per cui le stanze restano vuote.")
            else:
                righe.append("")
                righe.append("  ESITO: la lettura del file funziona correttamente.")
        except Exception as exc:
            righe.append("lettura non riuscita: %r" % exc)

    _sezione(righe, "misura dell'interfaccia")
    for riga in _misura_interfaccia(percorso):
        righe.append(riga)

    _sezione(righe, "dove si trova il desktop")
    for come, dove in _desktop():
        righe.append("%-22s: %s" % (come, dove))
        righe.append("%-22s  esiste: %s" % ("", os.path.isdir(dove)))
        if os.path.isdir(dove):
            collegamenti = [n for n in os.listdir(dove) if n.lower().endswith(".lnk")
                            and "inventario" in n.lower()]
            righe.append("%-22s  collegamenti Inventario: %r" % ("", collegamenti))

    righe.append("")
    righe.append("-" * 66)
    righe.append("Manda questo file a chi ti assiste.")
    return "\n".join(righe)


def main():
    from . import config
    testo = raccogli()
    destinazione = os.path.join(config.app_dir(), "Diagnostica.txt")
    try:
        with open(destinazione, "w", encoding="utf-8") as fh:
            fh.write(testo)
    except OSError:
        # cartella non scrivibile: si ripiega sul desktop, poi sulla home
        for base in [d for _, d in _desktop() if os.path.isdir(d)] \
                + [os.path.expanduser("~")]:
            destinazione = os.path.join(base, "Diagnostica.txt")
            try:
                with open(destinazione, "w", encoding="utf-8") as fh:
                    fh.write(testo)
                break
            except OSError:
                continue
    print(testo)
    print("\nRapporto salvato in: %s" % destinazione)
    return destinazione


if __name__ == "__main__":
    main()
