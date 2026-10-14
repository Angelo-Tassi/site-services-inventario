"""Una scheda aperta non cancella la funzione importata da un'altra postazione."""
import os
import sys
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario.store import InventoryError, InventoryStore
from inventario.ui import App, ItemDialog

percorso = fixture.build()
app = App(percorso)
try:
    app._initial_load()
    app.update()
    app.tree.selection_set("IT-0101")

    def salva_scheda(dialogo):
        # La scheda parte con funzione vuota. Mentre resta aperta, un collega
        # importa le funzioni; chi salva cambia soltanto la nota.
        collega = InventoryStore(percorso)
        collega.aggiorna_funzioni({"IT-0101": "PC Refresh"})
        return dict(app.selected_items()[0], note="Nota dalla scheda")

    with patch.object(ItemDialog, "show", salva_scheda):
        app.on_edit()
    riletti = InventoryStore(percorso).load()
    dispositivo = next(i for i in riletti if i["asset_tag"] == "IT-0101")
    assert dispositivo["funzione"] == "PC Refresh", dispositivo
    assert dispositivo["note"] == "Nota dalla scheda"
    # Chi cambia davvero la funzione la salva. Due modifiche diverse sullo
    # stesso campo chiedono invece di riaprire, senza annullare quella altrui.
    archivio = InventoryStore(percorso)
    prima = dict(next(i for i in archivio.load() if i["asset_tag"] == "IT-0101"))
    archivio.update("IT-0101", dict(prima, funzione="Standard"), originale=prima)
    prima = dict(next(i for i in archivio.load() if i["asset_tag"] == "IT-0101"))
    InventoryStore(percorso).aggiorna_funzioni({"IT-0101": "PC Refresh"})
    try:
        archivio.update("IT-0101", dict(prima, funzione=""), originale=prima)
    except InventoryError as exc:
        assert "Riapri la scheda" in str(exc), exc
    else:
        raise AssertionError("Una modifica concorrente alla funzione e' stata sovrascritta")
    assert next(i for i in InventoryStore(percorso).load()
                if i["asset_tag"] == "IT-0101")["funzione"] == "PC Refresh"
finally:
    app.destroy()

print("PERSISTENZA FUNZIONI OK")
