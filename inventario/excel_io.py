"""Esportazione e stampa in formato Excel."""

import os
import subprocess
import sys
import tempfile
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .store import (ALL_FIELDS, HEADERS, InventoryError, NON_DISPONIBILE,
                    SPEDITO, STATI, is_iphone, valore_visibile)

PRINT_FIELDS = ["asset_tag", "tipo", "modello", "seriale", "imei", "restituito_da",
                "stanza", "stato", "prestato_a", "prestato_il", "spedito_il", "note"]
PRINT_WIDTHS = {"asset_tag": 16, "tipo": 10, "modello": 26, "seriale": 16,
                "imei": 18, "restituito_da": 20, "stanza": 22, "stato": 14,
                "prestato_a": 20, "prestato_il": 15, "spedito_il": 15, "note": 24,
                "modificato_il": 18, "modificato_da": 24}

_HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
_BAND_FILL = PatternFill("solid", fgColor="F2F6FA")
_LOAN_FILL = PatternFill("solid", fgColor="FDEEEC")
_LOAN_TEXT = "A93226"
_SHIP_FILL = PatternFill("solid", fgColor="F6EFFB")
_SHIP_TEXT = "6C3483"
_THIN = Side(style="thin", color="B7C4D2")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _safe_sheet_title(name):
    for ch in "[]:*?/\\":
        name = name.replace(ch, "-")
    return (name or "Foglio")[:31]


def _write_table(ws, items, fields, title=None, subtitle=None):
    row = 1
    if title:
        ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=14)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(fields))
        row = 2
        if subtitle:
            cell = ws.cell(row=2, column=1, value=subtitle)
            cell.font = Font(size=9, color="666666")
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(fields))
            row = 3
        row += 1  # riga vuota di respiro

    header_row = row
    for col, field in enumerate(fields, start=1):
        cell = ws.cell(row=header_row, column=col, value=HEADERS[field])
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(vertical="center", horizontal="left")
        cell.border = _BORDER
    ws.row_dimensions[header_row].height = 20

    for i, item in enumerate(items):
        r = header_row + 1 + i
        on_loan = item.get("stato") == NON_DISPONIBILE
        spedito = item.get("stato") == SPEDITO
        for col, field in enumerate(fields, start=1):
            cell = ws.cell(row=r, column=col, value=valore_visibile(item, field))
            cell.border = _BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(field == "note"))
            if on_loan:
                cell.fill = _LOAN_FILL
                cell.font = Font(color=_LOAN_TEXT)
            elif spedito:
                cell.fill = _SHIP_FILL
                cell.font = Font(color=_SHIP_TEXT)
            elif i % 2:
                cell.fill = _BAND_FILL

    for col, field in enumerate(fields, start=1):
        ws.column_dimensions[get_column_letter(col)].width = PRINT_WIDTHS.get(field, 20)

    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    if items:
        last = get_column_letter(len(fields))
        ws.auto_filter.ref = "A%d:%s%d" % (header_row, last, header_row + len(items))
    return header_row


def _setup_print(ws, header_row, fields, footer_left):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "%d:%d" % (header_row, header_row)
    ws.print_area = "A1:%s%d" % (get_column_letter(len(fields)), ws.max_row)
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.6
    ws.oddFooter.left.text = footer_left
    ws.oddFooter.left.size = 8
    ws.oddFooter.right.text = "Pagina &P di &N"
    ws.oddFooter.right.size = 8


def _sort_key(item):
    return (item.get("stanza", ""), item.get("asset_tag", ""))


def export(items, path, group_by_room=False, rooms=None, full=True, for_print=False,
           con_iphone=False):
    """Scrive l'inventario in un file .xlsx.

    group_by_room: un foglio per stanza (piu' un foglio con il totale).
    full: include anche le colonne di tracciamento delle modifiche.
    for_print: aggiunge titolo, intestazioni ripetute e impaginazione A4.
    con_iphone: gli iPhone sono gestiti solo a mano e restano fuori dalle
        esportazioni; la stampa interna invece li include.
    """
    fields = list(ALL_FIELDS) if (full and not for_print) else list(PRINT_FIELDS)
    if not con_iphone:
        items = [i for i in items if not is_iphone(i.get("tipo"))]
    items = sorted(items, key=_sort_key)
    stamp = datetime.now().strftime("%d/%m/%Y %H:%M")

    wb = Workbook()
    wb.remove(wb.active)

    def add_sheet(name, subset, room=None):
        sheet_fields = [f for f in fields if not (room and f == "stanza")]
        ws = wb.create_sheet(_safe_sheet_title(name))
        title = subtitle = None
        if for_print:
            title = 'Site Services : Inventario Iphone, Laptop e Tablet' + (" - %s" % room if room else "")
            subtitle = "Stampato il %s - %d dispositivi" % (stamp, len(subset))
        header_row = _write_table(ws, subset, sheet_fields, title, subtitle)
        if for_print:
            _setup_print(ws, header_row, sheet_fields,
                         'Site Services : Inventario Iphone, Laptop e Tablet' + " - %s" % stamp)

    if group_by_room:
        room_names = list(rooms or [])
        for room in sorted(set(room_names) | set(i.get("stanza", "") for i in items)):
            subset = [i for i in items if i.get("stanza", "") == room]
            add_sheet(room or "Senza stanza", subset, room=room or "Senza stanza")
        if not wb.sheetnames:
            add_sheet("Inventario", items)
    else:
        add_sheet("Inventario", items)

    try:
        wb.save(path)
    except Exception as exc:
        raise InventoryError("Impossibile scrivere %s:\n%s" % (path, exc))
    finally:
        wb.close()
    return path


def build_print_file(items, group_by_room=False, rooms=None):
    """Genera in una cartella temporanea il file pronto da stampare."""
    name = "Inventario_stampa_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(tempfile.gettempdir(), name)
    return export(items, path, group_by_room=group_by_room, rooms=rooms,
                  full=False, for_print=True, con_iphone=True)


TEMPLATE_FIELDS = ["asset_tag", "tipo", "modello", "seriale", "stato", "note"]
TEMPLATE_TIPI = ["Laptop", "Tablet"]
RIGHE_PER_STANZA = 8


def build_template(path, rooms, stati=None):
    """Genera il modello vuoto da compilare e reimportare.

    Contiene solo le colonne che servono a laptop e tablet, gia' divise per
    stanza con le righe-separatore, e le tendine sui campi a scelta fissa.
    Gli iPhone non compaiono: si inseriscono a mano dal programma.
    """
    from openpyxl.worksheet.datavalidation import DataValidation

    stati = list(stati or STATI)
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventario"

    intestazioni = [HEADERS[f] for f in TEMPLATE_FIELDS]
    ws.append(intestazioni)
    for cella in ws[1]:
        cella.font = Font(bold=True, color="FFFFFF")
        cella.fill = _HEADER_FILL
        cella.alignment = Alignment(vertical="center")
        cella.border = _BORDER
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    fill_tag = PatternFill("solid", fgColor="DCE6F1")
    riga = 2
    for stanza in rooms:
        cella = ws.cell(row=riga, column=1, value=str(stanza).upper())
        cella.font = Font(bold=True, color="1F4E79")
        cella.fill = fill_tag
        ws.merge_cells(start_row=riga, start_column=1,
                       end_row=riga, end_column=len(TEMPLATE_FIELDS))
        ws.row_dimensions[riga].height = 18
        riga += 1
        for _ in range(RIGHE_PER_STANZA):
            for colonna in range(1, len(TEMPLATE_FIELDS) + 1):
                ws.cell(row=riga, column=colonna).border = _BORDER
            riga += 1

    ultima = riga - 1
    tipi = DataValidation(type="list", formula1='"%s"' % ",".join(TEMPLATE_TIPI),
                          allow_blank=True, showDropDown=False)
    tipi.error = "Scegli Laptop o Tablet."
    tipi.errorTitle = "Tipo non valido"
    ws.add_data_validation(tipi)
    tipi.add("B2:B%d" % ultima)

    scelte = DataValidation(type="list", formula1='"%s"' % ",".join(stati),
                            allow_blank=True, showDropDown=False)
    scelte.error = "Scegli uno degli stati previsti."
    scelte.errorTitle = "Stato non valido"
    ws.add_data_validation(scelte)
    scelte.add("E2:E%d" % ultima)

    larghezze = {"asset_tag": 18, "tipo": 14, "modello": 38, "seriale": 22,
                 "stato": 24, "note": 42}
    for i, campo in enumerate(TEMPLATE_FIELDS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = larghezze[campo]

    guida = wb.create_sheet("Istruzioni")
    testo = [
        ("Modello di inventario - laptop e tablet", True),
        ("", False),
        ("Compila il foglio \"Inventario\" e importalo dal programma con", False),
        ("Importa xls...  Le righe azzurre con il nome della stanza sono", False),
        ("separatori: tutto cio' che scrivi sotto una di esse finisce in", False),
        ("quella stanza, fino al separatore successivo.", False),
        ("", False),
        ("Regole", True),
        ("- Asset Tag e Modello sono indispensabili; il numero di serie e'", False),
        ("  vivamente consigliato.", False),
        ("- L'asset tag identifica il dispositivo: importando due volte lo", False),
        ("  stesso asset tag, la scheda viene aggiornata invece che duplicata.", False),
        ("- Tipo e Stato hanno la tendina: usa i valori proposti.", False),
        ("- Puoi aggiungere righe sotto un separatore, o spostare i separatori.", False),
        ("- Non cambiare i nomi delle colonne nella prima riga.", False),
        ("- Le righe lasciate vuote vengono semplicemente ignorate.", False),
        ("", False),
        ("Gli iPhone non si importano", True),
        ("Sono gestiti solo a mano dal programma e non compaiono ne' nelle", False),
        ("importazioni ne' nelle esportazioni. Se ne inserisci qui, vengono", False),
        ("ignorati.", False),
    ]
    for numero, (frase, grassetto) in enumerate(testo, start=1):
        cella = guida.cell(row=numero, column=1, value=frase)
        if grassetto:
            cella.font = Font(bold=True, size=12, color="1F4E79")
    guida.column_dimensions["A"].width = 78

    try:
        wb.save(path)
    except Exception as exc:
        raise InventoryError("Impossibile scrivere %s:\n%s" % (path, exc))
    finally:
        wb.close()
    return path


def send_to_printer(path):
    """Invia il file alla stampante predefinita.

    Ritorna True se la stampa e' stata avviata, False se il file e' stato
    soltanto aperto (l'utente stampera' da Excel).
    """
    if sys.platform.startswith("win"):
        try:
            os.startfile(path, "print")  # noqa: attributo solo Windows
            return True
        except Exception:
            os.startfile(path)  # noqa
            return False
    opener = "open" if sys.platform == "darwin" else "xdg-open"
    try:
        subprocess.Popen([opener, path])
    except Exception as exc:
        raise InventoryError("Impossibile aprire il file di stampa:\n%s" % exc)
    return False


def open_file(path):
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa
        return
    opener = "open" if sys.platform == "darwin" else "xdg-open"
    subprocess.Popen([opener, path])
