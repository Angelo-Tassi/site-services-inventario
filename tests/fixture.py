"""Costruisce un inventario di prova indipendente dai dati dell'utente."""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from inventario import config
from inventario import lingua as lang
from inventario.store import InventoryStore, new_item

# Le suite partono sempre in italiano: la lingua e' una preferenza salvata sul
# computer, e non deve cambiare l'esito dei test.
lang.imposta(lang.ITALIANO)

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
