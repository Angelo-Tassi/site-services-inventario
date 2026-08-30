"""Esportazione e stampa in formato Excel."""

import os
import subprocess
import sys
import tempfile
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .store import ALL_FIELDS, HEADERS, InventoryError, NON_DISPONIBILE

PRINT_FIELDS = ["asset_tag", "tipo", "modello", "seriale", "imei", "restituito_da",
                "stanza", "stato", "prestato_a", "prestato_il", "note"]
PRINT_WIDTHS = {"asset_tag": 16, "tipo": 10, "modello": 26, "seriale": 16,
                "imei": 18, "restituito_da": 20, "stanza": 22, "stato": 14,
                "prestato_a": 20, "prestato_il": 15, "note": 26,
                "modificato_il": 18, "modificato_da": 24}

_HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
_BAND_FILL = PatternFill("solid", fgColor="F2F6FA")
_LOAN_FILL = PatternFill("solid", fgColor="FBE3E1")
_LOAN_TEXT = "A93226"
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
        for col, field in enumerate(fields, start=1):
            cell = ws.cell(row=r, column=col, value=item.get(field, ""))
            cell.border = _BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(field == "note"))
            if on_loan:
                cell.fill = _LOAN_FILL
                cell.font = Font(color=_LOAN_TEXT)
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


def export(items, path, group_by_room=False, rooms=None, full=True, for_print=False):
    """Scrive l'inventario in un file .xlsx.

    group_by_room: un foglio per stanza (piu' un foglio con il totale).
    full: include anche le colonne di tracciamento delle modifiche.
    for_print: aggiunge titolo, intestazioni ripetute e impaginazione A4.
    """
    fields = list(ALL_FIELDS) if (full and not for_print) else list(PRINT_FIELDS)
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
                  full=False, for_print=True)


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
