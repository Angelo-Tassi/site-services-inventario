"""Il registro della tastiera: gli ultimi eventi, per quando si blocca.

Su Windows, "la tastiera non risponde piu'" ha una decina di cause possibili -
un widget con il fuoco che e' stato distrutto, una presa (grab) rimasta a una
finestra che non c'e' piu', una finestra ricostruita sotto le dita, un menu che
non ha restituito il fuoco - e da lontano, senza un PC Windows, non si
distinguono. Questo modulo tiene in memoria gli ultimi trecento eventi che
contano: quale widget riceve i tasti, dove va e da dove viene il fuoco, quando
si aprono e chiudono le finestre, quando l'elenco viene ricostruito e perche'.

Non scrive mai su disco da solo, salvo quando vede il sintomo: un tasto premuto
mentre nessun widget ha il fuoco, o mentre la presa punta a una finestra morta.
Chi ha il problema preme F12 e trova il file accanto al programma, o sul
desktop; Diagnostica.bat lo raccoglie insieme al resto.

Non registra mai i caratteri scritti: di un tasto si tiene il widget che lo ha
ricevuto e la classe del tasto - lettera, cifra, comando - non quale.
"""

import collections
import os
import time
from datetime import datetime

QUANTI = 300
NOME_FILE = "Tastiera.log"
# fra una scrittura automatica e l'altra: il sintomo, se c'e', si ripete a
# ogni tasto, e un file per tasto non serve a nessuno
PAUSA_AUTOMATICA = 60

_TASTI_DI_COMANDO = {
    "Return", "KP_Enter", "Escape", "Tab", "BackSpace", "Delete", "Up", "Down",
    "Left", "Right", "Home", "End", "Prior", "Next", "Insert", "F1", "F2", "F3",
    "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", "Shift_L",
    "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock",
    "Win_L", "Win_R", "Meta_L", "Meta_R", "space",
}


def _classe_del_tasto(keysym):
    """Che genere di tasto e', senza dire quale."""
    if not keysym:
        return "?"
    if keysym in _TASTI_DI_COMANDO:
        return keysym
    if len(keysym) == 1:
        if keysym.isdigit():
            return "cifra"
        if keysym.isalpha():
            return "lettera"
        return "segno"
    return "altro"


class Registro(object):
    def __init__(self, quanti=QUANTI):
        self.eventi = collections.deque(maxlen=quanti)
        self._ultima_automatica = 0.0
        self.root = None
        self.scritti = []          # i file scritti, per dirlo a chi chiede

    # ------------------------------------------------------------ annotare

    def nota(self, cosa, **dettagli):
        self.eventi.append((time.time(), cosa, dettagli))

    def righe(self):
        fuori = []
        for quando, cosa, dettagli in self.eventi:
            orario = datetime.fromtimestamp(quando).strftime("%H:%M:%S.%f")[:-3]
            resto = "  ".join("%s=%s" % (k, v) for k, v in dettagli.items())
            fuori.append("%s  %-14s %s" % (orario, cosa, resto))
        return fuori

    # ------------------------------------------------------------ osservare

    def installa(self, root):
        """Si mette in ascolto su tutta l'applicazione, senza mai intervenire.

        `add="+"` e nessun "break": il registro guarda passare gli eventi, non
        li ferma. F12 e Ctrl+Maiusc+D scrivono il file.
        """
        self.root = root
        root.bind_all("<KeyPress>", self._tasto, add="+")
        root.bind_all("<FocusIn>", self._fuoco_dentro, add="+")
        root.bind_all("<FocusOut>", self._fuoco_fuori, add="+")
        root.bind_all("<F12>", lambda e: self.scrivi("F12"), add="+")
        root.bind_all("<Control-Shift-D>", lambda e: self.scrivi("Ctrl+Maiusc+D"),
                      add="+")
        root.bind_all("<Control-Shift-d>", lambda e: self.scrivi("Ctrl+Maiusc+D"),
                      add="+")
        self.nota("avvio")

    def _percorso(self, widget):
        try:
            return str(widget)
        except Exception:
            return "?"

    def _fuoco_dentro(self, evento):
        self.nota("fuoco ->", a=self._percorso(evento.widget))

    def _fuoco_fuori(self, evento):
        self.nota("fuoco <-", da=self._percorso(evento.widget))

    def _tasto(self, evento):
        widget = self._percorso(evento.widget)
        self.nota("tasto", tipo=_classe_del_tasto(evento.keysym), su=widget)
        self._controlla_il_sintomo(widget)

    def _controlla_il_sintomo(self, widget):
        """Un tasto e' arrivato: c'e' qualcuno che puo' riceverlo davvero?"""
        if self.root is None:
            return
        sintomo = None
        try:
            fuoco = self.root.focus_get()
        except KeyError:
            fuoco = "(widget interno di Tk)"
        except Exception:
            fuoco = None
        if fuoco is None:
            sintomo = "tasto premuto ma nessun widget ha il fuoco"
        else:
            try:
                presa = self.root.grab_current()
                if presa is not None and not presa.winfo_exists():
                    sintomo = "la presa punta a una finestra che non esiste piu'"
            except Exception:
                pass
        if sintomo:
            self.nota("SINTOMO", quale=sintomo, tasto_su=widget)
            adesso = time.time()
            if adesso - self._ultima_automatica >= PAUSA_AUTOMATICA:
                self._ultima_automatica = adesso
                self.scrivi("automatico: " + sintomo)

    # ------------------------------------------------------------ scrivere

    def _dove(self):
        """Accanto al programma, altrimenti sul desktop, altrimenti a casa."""
        candidati = []
        try:
            from . import config
            candidati.append(config.app_dir())
        except Exception:
            pass
        try:
            from .diagnostica import _desktop
            candidati.extend(d for _n, d in _desktop() if os.path.isdir(d))
        except Exception:
            pass
        candidati.append(os.path.expanduser("~"))
        return candidati

    def testo(self, motivo=""):
        from . import __version__
        intestazione = [
            "REGISTRO DELLA TASTIERA - Inventario Site Services v%s" % __version__,
            "scritto il %s  (%s)" % (datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                                     motivo or "a richiesta"),
            "Ogni riga: orario, evento, dettagli. I caratteri scritti non ci sono:",
            "di un tasto si tiene solo il widget che lo ha ricevuto e il genere.",
            ""]
        stato = []
        if self.root is not None:
            try:
                stato.append("fuoco adesso   : %s" % self.root.focus_get())
            except KeyError:
                stato.append("fuoco adesso   : (widget interno di Tk)")
            except Exception as exc:
                stato.append("fuoco adesso   : ? (%s)" % exc)
            try:
                stato.append("presa adesso   : %s" % self.root.grab_current())
            except Exception as exc:
                stato.append("presa adesso   : ? (%s)" % exc)
            try:
                stato.append("finestre aperte: %s"
                             % ", ".join(str(w) for w in self.root.winfo_children()
                                         if w.winfo_class() == "Toplevel"))
            except Exception:
                pass
        return "\n".join(intestazione + stato + [""] + self.righe()) + "\n"

    def scrivi(self, motivo=""):
        """Scrive il registro su file e ritorna il percorso, o None."""
        testo = self.testo(motivo)
        for cartella in self._dove():
            destinazione = os.path.join(cartella, NOME_FILE)
            try:
                with open(destinazione, "a", encoding="utf-8") as fh:
                    fh.write(testo)
                    fh.write("=" * 72 + "\n")
            except OSError:
                continue
            self.scritti.append(destinazione)
            self.nota("scritto", dove=destinazione, motivo=motivo)
            return destinazione
        return None


# L'unico registro dell'applicazione: chi vuole annotare qualcosa lo importa
# da qui, senza doversi passare l'oggetto.
registro = Registro()
