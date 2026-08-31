"""Il rapporto di diagnostica dice le cose che servono a capire un difetto."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario import config, diagnostica

percorso = fixture.build()
vero = config.load_data_path
config.load_data_path = lambda: percorso
try:
    testo = diagnostica.raccogli()
finally:
    config.load_data_path = vero

# ---- le domande a cui il rapporto deve rispondere
from inventario import __version__
for atteso in ("versione            : " + __version__,
               "cartella programma", "percorso inventario", "cartella copie",
               "si puo' scrivere", "rooms", "iphone_room",
               "dispositivi         : 13",
               "lettura del file di prova", "dove si trova il desktop"):
    assert atteso in testo, atteso

# ---- il percorso dei dati e' quello vero, non un valore di ripiego
assert percorso in testo

# ---- le stanze con il loro conteggio
for stanza, quanti in ((fixture.BAU, 5), (fixture.KIOSK, 5), (fixture.DR, 3)):
    assert "%-34s %d" % (stanza, quanti) in testo, stanza

# ---- il file di prova viene letto e l'esito e' esplicito
assert "stanze riconosciute : ['Site Services BAU', 'Digital Kiosk'," in testo
assert "righe lette         : 30" in testo
assert "ESITO: la lettura del file funziona correttamente." in testo

# ---- non contiene niente di segreto oltre ai percorsi
assert "password" not in testo.lower()

print("DIAGNOSTICA OK")
