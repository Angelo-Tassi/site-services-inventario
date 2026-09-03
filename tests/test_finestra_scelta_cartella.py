"""La finestra "scegli la cartella" deve venire in primo piano, su Windows.

Un tecnico ha segnalato dal campo: al primo avvio di "Collega inventario
condiviso.bat" la finestra di scelta non si vedeva - bisognava andarla a
cercare come icona nella barra delle applicazioni, e la stessa cosa succedeva
alla domanda di conferma dopo, nel terminale. E' un comportamento noto di
Tkinter su Windows quando la finestra "padre" (qui tenuta nascosta con
withdraw) non e' mai stata mostrata prima di aprire un dialogo: il -topmost
prima di mostrarla forza il primo piano.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from tkinter import filedialog
from inventario import configura

# ---- la finestra nascosta che ospita il dialogo passa in primo piano prima
# di mostrarlo, non dopo: altrimenti il dialogo puo' aprirsi gia' dietro
visti = []
def finta_askdirectory(parent=None, title=None, **kwargs):
    visti.append((parent.attributes("-topmost"), parent.winfo_exists()))
    return "/tmp/una-cartella-qualsiasi"

vera_askdirectory = filedialog.askdirectory
filedialog.askdirectory = finta_askdirectory
try:
    scelta = configura.scegli_cartella()
finally:
    filedialog.askdirectory = vera_askdirectory

assert scelta == "/tmp/una-cartella-qualsiasi", scelta
assert visti, "askdirectory non e' stata chiamata"
topmost, esisteva = visti[0]
assert topmost == 1, "la finestra padre deve essere -topmost prima di mostrare il dialogo"
assert esisteva, "la finestra padre deve esistere ancora quando si mostra il dialogo"

# ---- annullando il dialogo (askdirectory ritorna vuoto), si torna a mano
filedialog.askdirectory = lambda **k: ""
try:
    assert configura.scegli_cartella() == ""
finally:
    filedialog.askdirectory = vera_askdirectory

# ---- fuori da Windows, riportare in primo piano il terminale non fa niente
# e non solleva errori: e' li' solo per chi esegue davvero da cmd.exe
vero_platform = sys.platform
sys.platform = "darwin"
try:
    configura._porta_avanti_la_console()      # non deve sollevare
finally:
    sys.platform = vero_platform

print("FINESTRA SCELTA CARTELLA OK")
