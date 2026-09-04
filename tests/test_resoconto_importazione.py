"""Il messaggio dopo l'importazione racconta tutto quello che e' successo.

A operazione avvenuta i numeri non bastano: chi ha appena riscritto l'inventario
di tutti deve poter controllare che sia successo quello che si aspettava, e
ritrovare la copia di sicurezza se non lo e'.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario.ui import App

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
app = App(fixture.build())
app._initial_load()

risultato = {"aggiunti": 12, "eliminati": 0, "copia": None,
             "gia_presenti": [{"asset_tag": "IT-0101", "stanza": BAU},
                              {"asset_tag": "IT-0102", "stanza": KIOSK}]}
esito = {"scartate": 2, "iphone": 1, "doppioni": ["IT-1", "IT-2"],
         "senza_modello": 4, "colonne_ignorate": ["Costo", "Fornitore"]}
righe = app._resoconto_importazione(risultato, esito,
                                    {"stanza": None, "mode": "merge"}, scartati=5)
testo = "\n".join(righe)

# ---- che cosa e' stato fatto
assert "Unione in tutto l'inventario." in testo, testo
assert "Aggiunti: 12" in testo, testo
# le righe gia' in inventario non entrano, e si dice dov'e' quello che c'e' gia'
assert "NON IMPORTATI, gia' in inventario: 2" in testo, testo
assert ("IT-0101  ->  %s" % BAU) in testo, testo

# ---- e che cosa non e' stato caricato, con il motivo di ognuno
assert "2 senza asset tag" in testo, testo
assert "1 iPhone" in testo, testo
assert "5 di altre stanze" in testo, testo
assert "2 doppioni nel foglio" in testo, testo
assert "4 senza modello" in testo, testo
assert "Costo, Fornitore" in testo, testo

# ---- e come si trova l'inventario adesso, stanza per stanza
assert "In inventario adesso: 13 dispositivi." in testo, testo
for stanza in (BAU, KIOSK, DR):
    assert stanza in testo, stanza

# ---- una sostituzione dice quanti ne ha eliminati e dove sta la copia
risultato = {"aggiunti": 30, "aggiornati": 0, "eliminati": 13,
             "copia": "/percorso/Backup/Inventario_2026-09-02.xlsx"}
righe = app._resoconto_importazione(risultato, {},
                                    {"stanza": KIOSK, "mode": "replace"})
testo = "\n".join(righe)
assert "Sostituzione in %s." % KIOSK in testo, testo
assert "Eliminati prima del caricamento: 13" in testo, testo
assert "Copia di sicurezza del file precedente:" in testo
assert "/percorso/Backup/Inventario_2026-09-02.xlsx" in testo

# ---- quando non c'e' niente da segnalare, non si inventano righe
righe = app._resoconto_importazione({"aggiunti": 1, "aggiornati": 0,
                                     "eliminati": 0, "copia": None},
                                    {}, {"stanza": None, "mode": "merge"})
testo = "\n".join(righe)
assert "Righe non caricate" not in testo, testo
assert "Colonne del foglio non riconosciute" not in testo, testo

app.destroy()
print("RESOCONTO IMPORTAZIONE OK")
