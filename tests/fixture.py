"""Costruisce un inventario di prova indipendente dai dati dell'utente."""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from inventario import config
from inventario import lingua as lang
from inventario.store import InventoryStore, new_item

# Le suite partono sempre in italiano e non leggono mai la configurazione vera:
# la lingua e' una preferenza salvata sul computer e non deve cambiare l'esito
# dei test, ne' i test devono modificarla.
lang.imposta(lang.ITALIANO)
load_language_reale = config.load_language          # per chi vuole provarla davvero
config.load_language = lambda: lang.ITALIANO

BAU, KIOSK, DR = "Site Services BAU", "Digital Kiosk", "Magazzino Disaster Recovery"
TIPO_IPHONE = "Iphone"

DEMO = [
 ("IT-0101","Laptop","Lenovo ThinkPad T14 Gen 4","PF4A1B2C",BAU,"Postazione reception"),
 ("IT-0102","Laptop","Lenovo ThinkPad T14 Gen 4","PF4A1B7D",BAU,""),
 ("IT-0103","Laptop","Lenovo ThinkPad T14 Gen 5","PF5K9M3E",BAU,"Batteria da sostituire"),
 ("IT-0104","Tablet","Dell Latitude 7320 Detachable","8H2KLM3",BAU,"Con tastiera e pennino"),
 ("IT-0105","Tablet","Dell Latitude 7230 Rugged Extreme","4RT9WQ2",BAU,""),
 ("IT-0106","Laptop","Lenovo ThinkPad T14 Gen 5","PF5K9M8F",KIOSK,"Postazione kiosk 1"),
 ("IT-0107","Laptop","Lenovo ThinkPad T14 Gen 4","PF4A2C1G",KIOSK,""),
 ("IT-0108","Laptop","Lenovo ThinkPad T14 Gen 5","PF5L4N2H",KIOSK,"Postazione kiosk 2"),
 ("IT-0109","Tablet","Dell Latitude 7320 Detachable","8H2KLP9",KIOSK,""),
 ("IT-0110","Tablet","Dell Latitude 7230 Rugged Extreme","4RT9WX7",KIOSK,"Con custodia"),
 ("DR-0201","Laptop","Lenovo ThinkPad T14 Gen 4","PF4B7T1J",DR,"Scorta sigillata"),
 ("DR-0202","Laptop","Lenovo ThinkPad T14 Gen 5","PF5M2P4K",DR,"Scorta sigillata"),
 ("DR-0203","Tablet","Dell Latitude 7320 Detachable","8H3NQR5",DR,"Kit continuita"),
]

def build():
    d = tempfile.mkdtemp()
    p = os.path.join(d, "Inventario.xlsx")
    s = InventoryStore(p, iphone_room=BAU)
    s.create_if_missing()
    for row in DEMO:
        s.add(new_item(*row))
    s.lend("IT-0107", "Marco Bianchi")
    s.lend("IT-0110", "Elena Rossi")
    config.save_shared_config(p, {"rooms": [BAU, KIOSK, DR],
                                  "types": ["Laptop", "Tablet", TIPO_IPHONE],
                                  "loan_rooms": [KIOSK],
                                  "iphone_room": BAU})
    return p


# ------------------------------------------------------- finestre che bloccano
# Una suite che apre un messagebox si ferma li' e aspetta un clic che nessuno
# dara' mai: il programma resta fermo finche' run_all.py non lo uccide dopo
# due minuti, e l'esito dice "BLOCCATA" senza dire su che cosa. Peggio, chi
# lancia i test si trova le finestre in faccia e deve chiuderle a mano.
#
# Qui si risponde da soli, con la risposta che non fa succedere niente:
#   - askyesno    -> No     (il promemoria della copia locale non parte)
#   - askokcancel -> Ok     (le conferme vanno avanti)
#   - le finestre "scegli un file" -> annullate.
# Ogni risposta viene registrata in `popup`, cosi' una suite puo' verificare
# che cosa e' stato chiesto invece di limitarsi a non bloccarsi.
#
# Una suite che vuole altre risposte le imposta dopo l'import, come ha sempre
# fatto: l'ultima assegnazione vince.

popup = []


def _risposta(tipo, valore):
    def finta(titolo="", messaggio="", **kwargs):
        popup.append((tipo, str(titolo), str(messaggio)))
        return valore
    return finta


def silenzia_i_popup():
    """Risponde da sola a ogni finestra che aspetterebbe un clic."""
    from tkinter import filedialog, messagebox

    for nome, valore in (("showinfo", "ok"), ("showwarning", "ok"),
                         ("showerror", "ok"), ("askyesno", False),
                         ("askokcancel", True), ("askyesnocancel", False),
                         ("askretrycancel", False), ("askquestion", "no")):
        setattr(messagebox, nome, _risposta(nome, valore))
    for nome in ("asksaveasfilename", "askopenfilename", "askdirectory",
                 "askopenfilenames"):
        if hasattr(filedialog, nome):
            setattr(filedialog, nome, _risposta(nome, ""))

    # Le finestre del programma - la scheda, il riepilogo, il reset - aspettano
    # con wait_window: se una suite ne apre una senza prevederlo, si chiude da
    # sola dopo tre secondi come se si fosse premuto Annulla. Meglio un test che
    # fallisce dicendo che cosa mancava, che una suite ferma per due minuti.
    from inventario import ui

    if not getattr(ui._Modal, "_show_sorvegliato", False):
        vero_show = ui._Modal.show

        def show_con_scadenza(self):
            self.after(3000, lambda: self.winfo_exists() and self._cancel())
            popup.append(("modale", self.title(), ""))
            return vero_show(self)

        ui._Modal.show = show_con_scadenza
        ui._Modal._show_sorvegliato = True


silenzia_i_popup()
