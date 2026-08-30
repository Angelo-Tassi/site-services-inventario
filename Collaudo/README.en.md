# Test files for the trial run

*[Versione italiana](README.md)*

Two Excel sheets ready to import, to try the program without touching the real
inventory.

| File | Contents |
| --- | --- |
| `Inventario_di_prova.xlsx` | 30 devices, **10 per room**, split by the separator rows. All regular. |
| `Inventario_di_prova_con_difetti.xlsx` | The same layout, but deliberately carrying the cases the program has to report. |

They are regenerated with:

```bash
.venv/bin/python tests/genera_file_di_prova.py
```

If you need to prepare a file of your own instead of using these, the full
format guide is here:
**[How to prepare the Excel file](https://angelo-tassi.github.io/site-services-inventario/formato-xls.html?lang=en)**.

## Before you start

Run the trials on a **fake** inventory, not on the one in the network folder.
The quickest way is to create an empty folder, copy `Inventario.exe` into it and
open it: on first run it offers to create a new `Inventario.xlsx` there.

---

# How to test the import

## 1. First load

1. *Import xls...*
2. **The whole inventory** + **Merge**, then *Choose the file*
3. select `Inventario_di_prova.xlsx`
4. the summary must say **30 valid rows** and **30 rows took their room from the
   separators**
5. confirm with *Import*

**What you should get.** Three room cards on the home page with **10** each, and
the full inventory below. Dell tablets are blue, laptops are not.

## 2. Merging does not duplicate

Repeat step 1 exactly. At the end the message must say **Added: 0, Updated: 30**,
and the card totals must stay 10, 10 and 10. Anything duplicated would push the
counts up.

## 3. Replacing a single room

1. *Import xls...* > **A single room** > `Digital Kiosk` > **Replace**
2. choose `Inventario_di_prova.xlsx` again
3. the summary says **10 valid rows**, **20 rows from other rooms discarded**,
   and warns that **10** devices will be deleted first

**What you should get.** Digital Kiosk still holds **10** devices, the ones
listed under the `DIGITAL KIOSK` row: the sheet declares the rooms, so the
separators rule and the other rooms' rows are thrown away. The other two rooms
stay as they were, with 10 each.

A file named `Inventario_prima_del_reset_...xlsx` must have appeared in the data
folder: that is the backup copy.

### When the sheet declares no rooms

Open `Inventario_di_prova.xlsx`, **delete all three separator rows** and save
under another name. Import it again with **A single room** > `Digital Kiosk`.

This time the summary says **30 valid rows** and warns that *the sheet declares
no rooms: every row will go into Digital Kiosk*. That is the case where your
choice is the only information available.

## 4. Replacing everything

1. *Import xls...* > **The whole inventory** > **Replace**
2. choose `Inventario_di_prova.xlsx`
3. the summary asks you to type `DELETE EVERYTHING`

Try confirming **without** typing it first, or typing "yes": it must refuse and
stay open. Then type the phrase and confirm: you are back to 30 devices spread
10, 10 and 10.

## 5. The warnings on imperfect files

Import `Inventario_di_prova_con_difetti.xlsx` with **The whole inventory** +
**Merge**. The summary must show, all together:

- **3 valid rows found**
- **1 row ignored**: the identifier is missing
- **3 rows took their room from the separators**
- **1 iPhone ignored**: they are only entered by hand
- the red box with the **unrecognised columns**: `Costo`, `Fornitore`,
  `Centro di costo`
- the warning that **1 row has no model**

On confirmation 3 devices go in: two in Site Services BAU (one of them without a
model, to be filled in by hand) and one in Digital Kiosk. The iPhone does **not**
go in.

## 6. Importing into a room

This is the trial of the **Import this room's data** button, next to the room
name once you open it. It must give the **same result** as the *A single room*
option of step 3: it is the same rule, reached two different ways.

1. enter **Digital Kiosk**
2. press *Import this room's data*, choose **Merge**
3. select `Inventario_di_prova.xlsx`

**What you should get.** The summary says **10 valid rows** and **20 rows from
other rooms discarded**. Only the ten devices listed under the `DIGITAL KIOSK`
row go in: those of BAU and the warehouse are thrown away, even though they are
in the same sheet.

### The room row is required

Take `Inventario_di_prova.xlsx`, **delete the `DIGITAL KIOSK` row** and save
under another name. Then try importing it into Digital Kiosk.

A warning must appear saying that no row in the sheet names that room, that
**nothing was imported**, and explaining to add an otherwise empty row with
`DIGITAL KIOSK` in the first cell. The warning also lists the rooms it found
instead. Check that the inventory is unchanged.

Try the short form too: instead of `DIGITAL KIOSK` write just `KIOSK`. It must
work the same.

## 7. Reset and reload

1. *Reset inventory*, type `DELETE EVERYTHING`
2. check the final message: it says how many were deleted and where it saved the
   backup
3. import `Inventario_di_prova.xlsx` again

If you had added an iPhone by hand before the reset, it must **survive**:
iPhones are never deleted and never reimported.

## 8. Exporting

With the inventory loaded from step 1, try the three forms from *Export xls...*.

**A single list.** One sheet with all 30 devices.

**One sheet per room.** One file with three sheets, one per room, 10 devices
each. Open each sheet and check that **A1** carries the room name, the export
date below it, and that the *Room* column is there and always reports the same
room.

**A separate file for each room.** Pick a folder: three files
`Inventario_<Room_name>_<date>.xlsx` must come out, each with its room in A1.

Then the counter-check: **import back** the file with the three sheets, with *The
whole inventory* + *Merge*. It must read all three and recognise 30 devices,
because the sheet name counts as a separator row.

Finally, enter a room and use the **Export this room to xls** shortcut: same
result as the *A single room* option. No iPhone must appear in any form.

## 9. Language

Switch language from the **Language** dropdown at the top right of the header,
and check that **everything** turns Italian or English: toolbar buttons, column
headers, room card text, the status bar at the bottom, and the messages that
appear when pressing *Delete*, *Move to room* or *Reset inventory* without having
ticked anything.

Then switch back from the same dropdown, or from *Settings*: they are the same
thing. The choice survives reopening the program.

**Exporting in English while working in Italian.** With the interface in Italian,
*Export xls...* and tick **Export the files in English**: in the file the headers
and statuses must be in English, while the room names stay exactly as you wrote
them. Import that file back: it must go in without errors and the statuses must
return to Italian.

---

## If something does not add up

Report the problem with: which step you were on, what you expected, what
happened, and if you can a screenshot of the summary. On macOS, the startup log
is in `avvio.log` in the program folder.
