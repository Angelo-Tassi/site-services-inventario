"""Ogni vista mostra solo le colonne che in quella vista possono avere un valore.

Una colonna vuota per costruzione - il prestito dove i prestiti non esistono,
l'IMEI dove non ci sono telefoni, la stanza dentro una stanza sola - toglie
spazio a quello che si deve leggere. La scelta dipende dalla configurazione, non
dai dati del momento: una stanza vuota deve mostrare le stesse colonne di quando
sara' piena.
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario.store import ALL_FIELDS, new_item
from inventario.ui import ACTION_COLUMN, CHECK_COLUMN, App

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
IPHONE = ("imei", "restituito_da", "spedito_il")
PRESTITO = ("prestato_a", "prestato_il")

app = App(fixture.build())
app._initial_load()

def campi():
    return app._campi_visibili()

# ---- home: una panoramica, non una scheda. Dice che cos'e' un dispositivo,
# dov'e' e come sta; il resto si guarda dentro la stanza che lo riguarda
app.show_home()
assert campi() == ["asset_tag", "tipo", "modello", "seriale", "stanza",
                   "stato", "note"], campi()
# l'IMEI e' l'identificativo dei soli telefoni: si guarda dove stanno loro
assert "imei" not in campi()
assert app._columns() == [CHECK_COLUMN] + campi()
# lo stato riassume da solo il prestito e la spedizione
prestati = [i for i in app.store.items if i.get("prestato_a")]
assert prestati and all(i["stato"] == "In prestito" for i in prestati), prestati
# e l'ordinamento continua a funzionare anche su una colonna che non si vede
assert app.sort_field == "modificato_il" and "modificato_il" not in campi()
app.sort_by("modello")
assert app.visible[0]["modello"] <= app.visible[-1]["modello"]
app.sort_by("modificato_il")

# ---- dentro una stanza la colonna Stanza sparisce sempre: e' nel titolo
for stanza in (BAU, KIOSK, DR):
    app.show_room(stanza)
    assert "stanza" not in campi(), stanza

# ---- Site Services BAU: ci stanno gli iPhone, non ci sono prestiti
app.show_room(BAU)
assert all(c in campi() for c in IPHONE), campi()
assert not any(c in campi() for c in PRESTITO), campi()
assert "asset_tag" in campi() and "seriale" in campi()   # i laptop ce l'hanno
assert ACTION_COLUMN not in app._columns()

# ---- Digital Kiosk: si prestano dispositivi, non ci sono telefoni
app.show_room(KIOSK)
assert all(c in campi() for c in PRESTITO), campi()
assert not any(c in campi() for c in IPHONE), campi()
assert app._columns()[:2] == [CHECK_COLUMN, ACTION_COLUMN]

# ---- Magazzino Disaster Recovery: ne' prestiti ne' telefoni, la piu' pulita
app.show_room(DR)
assert not any(c in campi() for c in IPHONE + PRESTITO), campi()
assert campi() == ["asset_tag", "tipo", "modello", "seriale", "stato", "note",
                   "modificato_il", "modificato_da"], campi()

# ---- contenitore iPhone: niente asset tag, niente seriale, niente prestiti,
# e nemmeno tipo e stanza, che sarebbero uguali su ogni riga
app.show_iphones()
assert campi() == ["modello", "imei", "restituito_da", "stato", "spedito_il",
                   "note", "modificato_il", "modificato_da"], campi()
assert ACTION_COLUMN in app._columns()

# ---- le colonne non dipendono dai dati: una stanza svuotata non le cambia
app.show_room(DR)
prima = campi()
app.store.delete([i["asset_tag"] for i in app.store.items if i.get("stanza") == DR])
app.on_refresh(); app.show_room(DR)
assert not app.visible, "la stanza deve essere vuota"
assert campi() == prima, "le colonne sono cambiate con i dati"

# ---- e seguono invece la configurazione: dando i prestiti al magazzino,
# le due colonne compaiono
app.cfg["loan_rooms"] = [KIOSK, DR]
app.show_room(DR)
assert all(c in campi() for c in PRESTITO), campi()
app.cfg["loan_rooms"] = [KIOSK]

# ---- senza il tipo Iphone configurato, le colonne dei telefoni spariscono
# anche dalla home
app.cfg["types"] = ["Laptop", "Tablet"]
app.store.items = [i for i in app.store.items if i.get("tipo") != fixture.TIPO_IPHONE]
app.show_home()
assert not any(c in campi() for c in IPHONE), campi()

app.destroy()
print("COLONNE PER VISTA OK")
