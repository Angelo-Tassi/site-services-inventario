"""La copia locale dell'inventario, contro la share che sparisce.

Deve essere un inventario vero e aggiornato all'istante in cui si chiede: non
un estratto, non una versione tenuta in memoria dal programma.
"""
import json, os, sys, tempfile, threading, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario import config
from inventario.store import InventoryError, InventoryStore, new_item

percorso = fixture.build()
store = InventoryStore(percorso, iphone_room=fixture.BAU)
store.load()
fuori = tempfile.mkdtemp()

# ---- la copia e' un inventario completo, riapribile
destinazione = os.path.join(fuori, "Copia.xlsx")
salvato, impostazioni, quanti = store.copia_in(destinazione)
assert salvato == destinazione and os.path.exists(salvato)
assert quanti == 13, quanti
riletti = InventoryStore(salvato).load()
assert len(riletti) == 13
assert [i["asset_tag"] for i in riletti] == [i["asset_tag"] for i in store.load()]
# i prestiti in corso ci sono ancora: e' l'inventario, non un estratto
assert [i["asset_tag"] for i in riletti if i.get("prestato_a")] == ["IT-0107", "IT-0110"]

# ---- accanto ai dati viaggiano le impostazioni
assert impostazioni == os.path.join(fuori, "Copia_impostazioni.json"), impostazioni
salvate = json.load(open(impostazioni))
assert salvate["rooms"] == [fixture.BAU, fixture.KIOSK, fixture.DR], salvate
assert salvate["loan_rooms"] == [fixture.KIOSK]

# ---- e' aggiornata al momento della richiesta, non a quando il programma ha letto
store.add(new_item("IT-9999", "Laptop", "T14 Gen 5", "PF99", fixture.BAU, "aggiunto dopo"))
seconda = os.path.join(fuori, "Copia2.xlsx")
_, _, quanti2 = store.copia_in(seconda)
assert quanti2 == 14, quanti2
assert "IT-9999" in [i["asset_tag"] for i in InventoryStore(seconda).load()]
# la prima copia non e' cambiata: e' una fotografia
assert len(InventoryStore(destinazione).load()) == 13

# ---- una cartella che non esiste si dice, non si crea di nascosto
try:
    store.copia_in(os.path.join(fuori, "non", "esiste", "Copia.xlsx"))
    raise AssertionError("doveva rifiutare")
except InventoryError as exc:
    assert "non esiste" in str(exc), exc

# ---- la copia aspetta chi sta salvando, invece di prendere un file a meta'
from inventario.store import _Lock
tenuto = threading.Event()
rilasciare = threading.Event()

def chi_salva():
    with _Lock(store.path):
        tenuto.set()
        rilasciare.wait(5)

t = threading.Thread(target=chi_salva); t.start()
tenuto.wait(5)
inizio = time.time()
timer = threading.Timer(0.4, rilasciare.set); timer.start()
terza = os.path.join(fuori, "Copia3.xlsx")
_, _, quanti3 = store.copia_in(terza)          # deve bloccarsi finche' l'altro ha finito
durata = time.time() - inizio
t.join(); timer.cancel()
assert durata >= 0.3, "non ha aspettato il lock: %.2f s" % durata
assert quanti3 == 14, quanti3

# ---- la copia sta fuori dal programma e dalla share, dove l'ha chiesta l'utente
assert not os.path.abspath(destinazione).startswith(os.path.dirname(percorso))
assert not os.path.abspath(destinazione).startswith(os.path.abspath(config.app_dir()))

print("COPIA LOCALE OK")
