"""Schermate, schede stanza, colonne, colori delle righe e selezione."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fixture
from inventario import theme
from inventario.store import NON_DISPONIBILE, new_item
from inventario.ui import ACTION_COLUMN, CHECK_COLUMN, CHECK_OFF, CHECK_ON, App, RoomCard

BAU, KIOSK, DR = fixture.BAU, fixture.KIOSK, fixture.DR
app = App(fixture.build()); app._initial_load()
TIPO = app.iphone_type()

def schede():
    return [w for f in app.body.winfo_children() for w in f.winfo_children()
            if isinstance(w, RoomCard)]

# ---- home: tre stanze piu' il contenitore iPhone
assert app.view == "home" and len(app.visible) == 13
assert [c.labels[1].cget("text") for c in schede()] == [BAU, KIOSK, DR, TIPO]
assert [c.labels[2].cget("text") for c in schede()] == ["5", "5", "3", "0"]
assert [c.note.cget("text") for c in schede()[:3]] == ["", "2 in prestito", ""]

# ---- colonne: la casella c'e' sempre, l'azione solo dove serve
assert app._columns()[0] == CHECK_COLUMN
assert ACTION_COLUMN not in app._columns(), "in home niente colonna azione"
assert app.tree.heading(CHECK_COLUMN)["text"] == ""
assert str(app.tree.cget("selectmode")) == "browse"
for stanza in (BAU, DR):
    app.show_room(stanza)
    assert ACTION_COLUMN not in app._columns(), stanza
app.show_room(KIOSK)
assert app._columns()[:2] == [CHECK_COLUMN, ACTION_COLUMN]
assert app.tree.heading(ACTION_COLUMN)["text"] == "Prestito"

# ---- prestiti: righe rosse, pulsanti giusti, cella vuota
assert len(app.visible) == 5
def colore_riga(t):
    tags = app.tree.item(t, "tags")
    nome = tags[0].replace("_alt", "") if tags else ""
    return "" if nome == "odd" else nome
rossi = [t for t in app.tree.get_children() if colore_riga(t) == "loan"]
assert sorted(rossi) == ["IT-0107", "IT-0110"]
for tag in rossi:
    assert app._item_by_tag(tag)["stato"] == NON_DISPONIBILE
    assert app.tree.set(tag, ACTION_COLUMN) == "", "il testo sta sul pulsante"
assert app.action_label(app._item_by_tag("IT-0106")) == "Presta"
assert app.action_label(app._item_by_tag("IT-0107")) == "Registra rientro"
assert app.action_label(app._item_by_tag("IT-0101")) == "", "fuori dal kiosk niente pulsante"
app._sync_row_buttons(); assert isinstance(app._row_buttons, dict)

app._run(lambda: app.store.lend("IT-0106", "Giulia Verdi"), "ok")
assert app.action_label(app._item_by_tag("IT-0106")) == "Registra rientro"
app._run(lambda: app.store.give_back("IT-0106"), "ok")
assert app.action_label(app._item_by_tag("IT-0106")) == "Presta"

# ---- selezione a casella, una alla volta
app.show_home()
assert set(app.tree.set(t, CHECK_COLUMN) for t in app.tree.get_children()) == {CHECK_OFF}
app.tree.selection_set(["IT-0103"]); app._on_select()
assert app.tree.set("IT-0103", CHECK_COLUMN) == CHECK_ON
assert app.tree.set("IT-0104", CHECK_COLUMN) == CHECK_OFF
app.tree.selection_set(["IT-0104"]); app._on_select()
assert app.selected_tags() == ["IT-0104"]
assert app.tree.set("IT-0103", CHECK_COLUMN) == CHECK_OFF
app.refresh_table()
assert app.selected_tags() == ["IT-0104"] and app.tree.set("IT-0104", CHECK_COLUMN) == CHECK_ON

# ---- colori per tipo
app._run(lambda: app.store.add(new_item(tipo=TIPO, modello="Apple iPhone 14",
                                        imei="356938035643809", restituito_da="M. B.")), "ok")
app._run(lambda: app.store.add(new_item("IT-0999", "Tablet", "Samsung Galaxy Tab A9",
                                        "R52ZZ", BAU)), "ok")
def colore(t):
    """Il colore della riga, ignorando la banda alternata."""
    tags = app.tree.item(t, "tags")
    nome = tags[0].replace("_alt", "") if tags else ""
    return "" if nome == "odd" else nome
app.show_home()
assert colore("356938035643809") == "iphone"
assert colore("IT-0104") == "tablet" and colore("IT-0105") == "tablet"
assert colore("IT-0999") == "", "tablet non Dell: nessun colore"
assert colore("IT-0101") == "", "laptop: nessun colore"
assert colore("IT-0110") == "loan", "il prestito ha la precedenza sul tipo"
app.show_room(BAU)
assert colore("356938035643809") == "iphone" and colore("IT-0104") == "tablet"
assert str(app.tree.tag_configure("iphone", "background")) == theme.IPHONE_ROW
assert str(app.tree.tag_configure("tablet", "background")) == theme.TABLET_ROW

# ---- ordinamento e filtri
app.show_home()
app.sort_by("modello")
assert app.visible[0]["modello"].startswith("Apple")
app.sort_by("modello")
assert app.visible[0]["modello"].startswith("Samsung")
app.var_search.set("marco bianchi"); app.refresh_table()
assert [i["asset_tag"] for i in app.visible] == ["IT-0107"]
app.var_search.set("8h2klm3"); app.refresh_table()
assert [i["asset_tag"] for i in app.visible] == ["IT-0104"]
app.reset_filters()
assert len(app.visible) == 15
app.destroy()
print("UI OK")
