"""Tutto quello che il programma dice ha una traduzione inglese.

Una stringa dimenticata non sparisce - resta in italiano - e proprio per questo
non se ne accorge nessuno finche' non la vede un utente inglese. Qui si legge il
codice e si controlla che ogni testo passi dalla traduzione e che la traduzione
esista davvero.
"""
import ast, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario.lingua import EN, INTESTAZIONI_EN, STATI_EN
from inventario.store import HEADERS, STATI

RADICE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODULI = ("inventario/ui.py", "inventario/store.py", "inventario/configura.py")

def stringhe_tradotte(percorso):
    albero = ast.parse(open(os.path.join(RADICE, percorso), encoding="utf-8").read())
    fuori = []
    for nodo in ast.walk(albero):
        if isinstance(nodo, ast.Call) and isinstance(nodo.func, ast.Name) \
           and nodo.func.id == "T" and nodo.args:
            primo = nodo.args[0]
            if isinstance(primo, ast.Constant) and isinstance(primo.value, str):
                fuori.append(primo.value)
    return fuori

# ---- ogni testo passato alla traduzione ha la sua voce inglese
mancanti = []
for modulo in MODULI:
    for testo in stringhe_tradotte(modulo):
        if testo.strip() and testo not in EN:
            mancanti.append((modulo, testo))
assert not mancanti, "\n".join("%s: %r" % (m, t[:60]) for m, t in mancanti)

# ---- e ce n'e' abbastanza da essere credibile
assert len(stringhe_tradotte("inventario/ui.py")) > 300
assert len(stringhe_tradotte("inventario/store.py")) > 20, \
    "anche gli errori dell'archivio finiscono in una finestra"

# ---- i messaggi di errore dell'archivio passano dalla traduzione
albero = ast.parse(open(os.path.join(RADICE, "inventario/store.py"),
                        encoding="utf-8").read())
crudi = []
for nodo in ast.walk(albero):
    if isinstance(nodo, ast.Call) and isinstance(nodo.func, ast.Name) \
       and nodo.func.id == "InventoryError" and nodo.args:
        primo = nodo.args[0]
        if isinstance(primo, ast.BinOp):
            primo = primo.left
        if isinstance(primo, ast.Constant) and isinstance(primo.value, str):
            crudi.append(primo.value[:50])
assert not crudi, "messaggi di errore non tradotti: %s" % crudi

# ---- colonne e stati hanno la loro traduzione
for campo, titolo in HEADERS.items():
    assert titolo in INTESTAZIONI_EN, campo
for stato in STATI + ["In prestito", "Da Rispedire", "Spedito al servizio telefonia"]:
    assert stato in STATI_EN, stato

# ---- e nessuna traduzione e' rimasta uguale all'italiano per distrazione
uguali = [k for k, v in EN.items()
          if k == v and len(k) > 12 and not k.startswith("%s")]
assert not uguali, uguali

print("TRADUZIONE COMPLETA OK (%d voci)" % len(EN))
