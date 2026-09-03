# Test files for the trial run

*[Versione italiana](README.md)*

This folder is part of the program and travels with it: it is there to try the
import with fake data, before loading the real inventory.

## Where the files are

They are in the **`Collaudo`** folder, inside the program folder on the
workstation:

```
C:\Inventario\                      the program, on the workstation
    Inventario.py
    inventario\
    python\
    inventario_percorso.json     says which inventory to open
    Collaudo\
        Inventario_di_prova.xlsx
        Inventario_di_prova_con_difetti.xlsx
        README.md                the Italian version
        README.en.md             this document

\\server\Shared\Inventory\          the network folder: data only
    Produzione\
        Inventario.xlsx          the real inventory, one for everybody
        Backup\                  the backup copies
```

If you downloaded `Inventario-windows.zip` from the Releases page, the
`Collaudo` folder is already inside. Working from source, it is `Collaudo/` in
the project root.

When you are asked to choose a file, point here: **nothing needs to be copied
anywhere else**.

## What they contain

### `Inventario_di_prova.xlsx` - 30 devices, all regular

A single sheet called *Inventario*, with the column headers in the first row and
the devices split into three blocks by the blue separator rows:
`SITE SERVICES BAU`, `DIGITAL KIOSK`, `MAGAZZINO DISASTER RECOVERY`. Ten devices
per block.

| Column | Contents |
| --- | --- |
| Asset Tag | `IT-BAU-101` ... `IT-KSK-210` ... `IT-DRC-310`, one series per room |
| Type | seven Laptops and three Tablets per room |
| Notes | some filled in, some empty |
| Status | a mix of the five statuses, to see how they behave |
| Model/Description | Lenovo ThinkPad T14 Gen 4 and Gen 5, Dell Latitude 7320 Detachable and 7230 Rugged Extreme |
| Serial number | plausible serials, all different |

The columns are in the same order as the
[template to fill in](../docs/Import_template.xlsx) and as the list in the
program. The order does not matter anyway: the import recognises columns by
name.

No iPhones: phones are entered by hand only.

### `Inventario_di_prova_con_difetti.xlsx` - the cases to be reported

Same layout, but carrying **on purpose** everything the program has to
recognise and tell you about. Besides the two separators there are six rows:

| What is in it | What the program must do |
| --- | --- |
| Three extra columns: `Costo`, `Fornitore`, `Centro di costo` | ignore them, and list them before importing |
| A row with no model (`IT-BAU-902`) | import it anyway, saying how many are missing one |
| A row with neither asset tag nor IMEI | discard it and count it |
| A row of type `Iphone`, with an IMEI | ignore it: phones are not imported |
| A completely empty row | skip it without counting it |
| The separators written in short form, `BAU` and `KIOSK` | recognise them all the same |

These are the same files the automated tests use: if the program's behaviour
changes, the tests fail and these instructions get rewritten. They are
regenerated with `.venv/bin/python tests/genera_file_di_prova.py`.

## The test files stay here

Do not move them and do not copy them elsewhere: they are part of the program,
they are updated along with it and they serve anyone who has to run the trial
again.

**These files are loaded only with  Import xls...  from inside the program.**
They are not inventories: they are sheets to import, and they contain the
separator rows that divide devices by room. Open one *as* an inventory and those
rows become devices, and no device has a room. The program now notices and tells
you, but the rule stands: they are imported, not opened.

## Real inventories live outside

The Excel sheet with the final inventory - the one you will actually load, with
the real devices - **must not be put in the program folder**. Keep it in a
folder of your own, on your own computer: `Documents\Inventory`, the Desktop,
wherever you prefer.

Two reasons:

- the program folder is **replaced at every update**: a file left there is lost
  without warning;
- it is a local folder on your own workstation: whatever you leave there nobody
  else sees, and it is no place for anything that matters.

The same goes for the files you **export** from the program and for any backup
copies you decide to keep: save them in a folder of your own.

## Not to be confused: the file the program uses

A different thing is the file the program works on, the one it reads and writes
continuously. You do not import it: it simply opens it. It lives on the **shared
network folder**, at `Produzione\Inventario.xlsx`, and it is the same one for
every technician.

Which one it is is written in `inventario_percorso.json`, next to the program on
the workstation; to change it, run `Collega inventario condiviso.bat` again.

---

# How to test the import

The checks below work on whatever inventory the program is currently using. Do
them before loading the final data: step 7 empties everything, so this is the
right moment to try while the inventory is still fake.

## 1. First load

1. *Import xls...*
2. **The whole inventory** + **Merge**, then *Choose the file*
3. open the `Collaudo` folder next to the program and pick
   `Inventario_di_prova.xlsx`
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
stay as they were.

A file named `Inventario_<date>.xlsx` must have appeared in the **`Backup`**
folder, inside the program folder: that is the backup copy, taken before
anything was touched.

### When the sheet declares no rooms

Open `Inventario_di_prova.xlsx`, **delete the three separator rows** and save
under another name **in a folder of your own**, so the test file stays
untouched. Import it again with **A single room** > `Digital Kiosk`.

This time the summary says **30 valid rows** and warns that *the sheet declares
no rooms: every row will go into Digital Kiosk*.

## 4. The warnings on imperfect files

Import `Inventario_di_prova_con_difetti.xlsx` with **The whole inventory** +
**Merge**. The summary must show, all together:

- **3 valid rows found**
- **1 row ignored**: the identifier is missing
- **3 rows took their room from the separators**
- **1 iPhone ignored**: they are only entered by hand
- the red box with the **unrecognised columns**: `Costo`, `Fornitore`,
  `Centro di costo`
- the warning that **1 row has no model**

On confirmation 3 devices go in, one of them without a model to be filled in by
hand. The iPhone does **not** go in.

## 5. Importing into a room

The **Import this room's data** button sits next to the room name once you open
it. It must give the **same result** as step 3: the same rule, reached two
different ways.

1. enter **Digital Kiosk**
2. press *Import this room's data*, choose **Merge**
3. select `Inventario_di_prova.xlsx`

**What you should get.** The summary says **10 valid rows** and **20 rows from
other rooms discarded**.

### The room row is required

Take `Inventario_di_prova.xlsx`, **delete the `DIGITAL KIOSK` row** and save
under another name in a folder of your own. Try importing it into Digital Kiosk again: a warning must
appear saying no row in the sheet names that room, that **nothing was imported**,
and explaining to add an otherwise empty row with `DIGITAL KIOSK` in the first
cell. Check that the inventory is unchanged.

Try the short form too: instead of `DIGITAL KIOSK` write just `KIOSK`. It must
work the same.

## 6. Exporting

With the inventory loaded from step 1, try the three forms from *Export xls...*.

**A single list.** One sheet with all 30 devices.

**One sheet per room.** One file with three sheets, 10 devices each. In every
sheet, **A1** must carry the room name, the export date below it, and the *Room*
column must be there reporting always the same room.

**A separate file for each room.** Pick a folder: three files
`Inventario_<Room_name>_<date>.xlsx` must come out.

Then the counter-check: **import back** the file with the three sheets, with
*The whole inventory* + *Merge*. It must read them all and recognise 30 devices.

### Sending by e-mail

Every export ends with the same question. Choose **Send by e-mail with
Outlook**: a new message must open with the file already attached, for you to
fill in and send. Close the message without sending it, this is only a check.

Repeat with the export **into one file per room**: this time the attachment must
be a single `.zip` archive holding the three files.

If Outlook is not on that computer, a warning must appear saying so and
reminding you that the file was created anyway: check that it really is there.

## 7. Going back from a mistake

This is the check that really matters on site: getting an import wrong and
putting it right.

1. import `Inventario_di_prova.xlsx` twice with **The whole inventory** +
   **Replace**, so at least two copies exist
2. now make the mess: import `Inventario_di_prova.xlsx` choosing **A single
   room** > `Site Services BAU` > **Merge**. All 30 devices end up in BAU, which
   had 10
3. open **Settings** > **Restore from a copy...** and pick the first in the
   list, which is the latest saved copy

**What you should get.** The list shows every copy newest first, with date, time
and device count compared to now. On confirmation the room cards go back to 10,
10 and 10, and the final message says where it saved the wrong state - because
that stays recoverable too.

**Where the three buttons went.** *Save a local copy...*, *Restore from a
copy...* and *Reset inventory* live inside **Settings**, in the *Copies and
restore* box: they are no longer in the toolbar.

## 8. Language

Switch language from the **Language** dropdown at the top right of the header,
and check that **everything** changes: buttons, column headers, room cards, the
status bar, and the messages that appear when pressing *Delete* or *Move to
room* without having ticked anything. Then switch back.

With the interface in Italian, try *Export xls...* ticking **Export the files in
English**: in the file, headers and statuses must be in English, while the room
names stay exactly as you wrote them. Import that file back: it must go in
without errors and the statuses must return to Italian.

---

# Loading the final inventory

Once the trial is over, the inventory is full of fake data. It has to be emptied
before loading the real thing, or the two get mixed: the test devices would stay
in alongside yours.

There are two ways. They do the same thing; only the moment of the emptying
differs.

## 9. Multiple selection

This checks that you can work on several devices at once, and that before doing
so you read exactly what will happen.

1. in the list, click a row: the circle in the first column lights up;
2. **Ctrl+click** two other distant rows: all three must stay lit, with the amber
   background, and *3 selected* must appear next to the list title;
3. click any row **without holding anything down**: the selection starts again
   from that one alone;
4. click a row, then **Shift+click** a row further down: the whole range lights
   up;
5. with three rows picked press **Move to room...**, choose a room and read the
   summary: which rooms they come from, how many iPhones stay put, and how the
   rooms stand before and after. Press **Cancel**;
6. select two laptops **and an iPhone never shipped back**, then press
   **Delete**: the summary must list the two laptops among those disappearing and
   the iPhone among the skipped, with the reason. Confirm;
7. check that **only the two laptops** are gone and that the iPhone is still in
   its container;
8. the last message must say how many devices are left and where the backup copy
   is: open it and check they are all still there;
9. go into the **Digital Kiosk**, press **Lend** on a laptop and give a name.
   Then, with that laptop selected, try:
   - **Move to room...** - if it is the only one chosen the program says it is on
     loan and does not even ask for the room; together with others, the summary
     lists it among those that *stay where they are because they are on loan*,
     and after confirming the others move and it does not;
   - **Delete** - it says you cannot, and that the return must be registered
     first;
   - **Edit**, changing the **room** - refused with the same reason, while
     correcting the **note** works;
10. press **Register return** on that laptop and repeat the move: now it travels
    like any other.

**What must happen:** the selection is not lost on its own - not even when the
list reloads - both operations always show the summary before acting, a
protected iPhone does not stop the others being deleted, and **a device on loan
is neither moved nor deleted by any route** until the return is registered.

## 10. Copy and paste

This checks that the fields accept what comes from Excel, and that an identifier
can be taken out of the list.

1. open `Collaudo\Inventario_di_prova.xlsx` in Excel and **copy a column** of
   asset tags, for example from `IT-BAU-101` to `IT-BAU-105`;
2. in the program press **`Delete +`** and paste with **Ctrl+V** into the box:
   the five rows must appear. Press *Check* and then **Cancel**: nothing is being
   deleted, this is only testing the paste;
3. in the list, **double-click the Notes cell** of a device and paste any text
   with Ctrl+V: it must go in. `Esc` to cancel;
4. **double-click the Type cell**: the dropdown offers the types, but you can
   also type or paste a different value;
5. **right-click** in any text field: the menu must open with *Copy*, *Cut*,
   *Paste*, *Select all*;
6. select a row in the list and press **Ctrl+C**: the row goes to the clipboard,
   ready to paste into Excel. **Right-clicking the row** offers *Copy the
   identifier*, which copies the asset tag alone;
7. run any export: in the final window the file paths can be **selected and
   copied**.

**What must happen:** every field accepts Ctrl+C and Ctrl+V, the right button
always opens the menu, and what you copy from the list pastes into Excel with no
touch-ups.

## 11. Automatic copies never go past ten

This checks that the copies folder does not grow forever.

1. open `Produzione\Backup\` on the share and count the files named
   `Inventario_<date>_<time>.xlsx`;
2. do eleven operations that save a copy - eleven *Delete* of any device will
   do, or eleven imports with **Replace**;
3. go back to the folder and count again.

**What you should get.** The files with the date in the name are **ten**, never
more, and they are the ten most recent: the oldest goes when a new one arrives.
A file you had put there by hand, with any name, **must still be in place**: the
rotation only touches the copies the program wrote.

If a copy is open in Excel on another PC the program cannot delete it: the
operation goes ahead anyway, and that copy will go on the next round. It is not
an error.

---

## 12. Starting again from a local copy

This checks the way out of the worst case: the network folder gone, and with it
the automatic backups that lived inside it.

1. open **Settings** > **Save a local copy...** and save the zip **on your own
   PC**, not on the share. The message must say how many devices it holds;
2. check the file: opening it with a double click, inside there must be
   `Inventario.xlsx` and `inventario_impostazioni.json`;
3. now make the mess: in **Settings** rename a room and save, then delete two or
   three devices;
4. open **Settings** > **Restore from a local file...** and pick the zip from
   step 1;
5. read the summary **before confirming**: it must say how many devices come
   back compared to now, what date the copy is from, how every room ends up
   (before and after) and which settings change, with a `*` on the ones that
   differ. Confirm;
6. check that both the devices and **the room name** you changed have come back;
7. try the bad case too: pick any file that is not a copy - a document, an
   image. It must say it is not a readable copy and **touch nothing**.

**What must happen:** everything comes back from the single file you took away,
rooms included. A backup copy of the previous state is kept anyway, and its path
is written in the final message.

---

## 13. Renaming a room or a type

This checks that the devices follow the room and the type when their name
changes.

1. look at the home page and note how many devices are in the **Digital
   Kiosk**;
2. open **Settings**, change the `Digital Kiosk` line to `Digital Kiosk 2` and
   press **Save**;
3. read the message: it must say `Digital Kiosk  ->  Digital Kiosk 2` with how
   many devices it moved;
4. on the home page there must be a **Digital Kiosk 2** card with the same
   number as before, and **no card with the old name**;
5. go into the room: the **Loan** column with the *Lend* / *Register return*
   buttons must still be there - the room with loans followed the new name;
6. reopen **Settings** and try typing the same room name twice: it must refuse,
   naming the repeated one;
7. now the trickiest case: rename **Site Services BAU**, which is the iPhone
   room. After saving, reopen **Settings**: the *iPhone room* field must carry
   the new name. Go into the **Iphone** container: the phones must still be
   there, and adding a new one must put it in the room with the new name;
8. move a line to another position without changing any name and save: it
   **must move nothing**, reordering is not renaming;
9. now the types. In **Settings** change `Laptop` into `Notebook` and save: the
   message must say how many devices it updated, and in the list the *Type*
   column must read `Notebook` on every former laptop. Open the record of one of
   them: asset tag, serial number, room, status, notes and loan must be
   **exactly what they were**;
10. finally try renaming the `Iphone` type into something else: it **must
    refuse**, explaining that it is the word the program recognises phones by and
    telling you which name to put back.

**What must happen:** no device is left behind with the old name, no card appears
for a room that no longer exists, and the phones stay phones.

---

## 14. Recently deleted

This checks the safety net for a deletion made by mistake.

1. on the home screen delete any three devices, confirming the summary;
2. search for one of those three in the **Search** box of the main list: it
   **must not appear**. Do an export too: it must not be there;
3. press **RECENTLY DELETED**, the black button next to *General duplicate
   check*. In brackets there must be **3**, the ones you just deleted: they must
   be at the top, with asset tag and type;
4. select a row and read the line at the bottom: it must say when and by whom it
   was deleted, and which room it would go back to;
5. press **Restore** on the row: the device goes back **into its room**. Open its
   record and check that serial number, notes and status are what they were;
6. select the other two with **Ctrl+click** and press **Restore the selected
   ones** at the top. A **single popup** must open: above, the list of the ones
   coming back **grouped by the room** they land in, and below the *Cancel* and
   *Restore* buttons. On confirmation both come back, and under the list you read
   how many returned and where. Try **cancelling** too: nothing must go back
   in;
7. now delete more than ten devices and reopen the bin: the list must be **split
   into pages of ten**, with *Previous* and *Next*;
8. type in the **search** box at the top the asset tag of one that is on the
   **second** page: it must find it all the same. The search looks at the whole
   bin, not the page you are on;
9. finally open **Settings** and remove from the list a room that still holds
   devices. It must ask for confirmation saying how many there are; on
   confirmation they end up in the bin with a **red background**. Try restoring
   one: the program must **ask you which room** to put it in.

**What must happen:** what is in the bin is not visible anywhere else in the
program, and from there you come back without losing anything on the record.

---

## 15. Removing a room that has a role

This checks the relocation: the iPhone room and the room with loans cannot
disappear leaving the devices stranded.

1. open **Settings** and remove **Site Services BAU** from the list, which is the
   iPhone room. Save;
2. the program must **ask which room to move them to**, saying how many devices
   are inside. Choose `Magazzino Disaster Recovery`;
3. check that the **Iphone** container still holds every phone, **shipped or
   not**, and that they are in the room you chose. The other devices that were in
   BAU must be there too;
4. in **Settings** the *iPhone room* field must carry the new name;
5. repeat with the **Digital Kiosk**, which is the room with loans, having lent a
   device first. The program asks where to move them:
   - the device **on loan** moves too, and the loan **stays open** - the name of
     whoever has it is still there;
   - the room you chose takes the Kiosk's place among the **rooms with loans**:
     going into it, the *Loan* column with its buttons must be there;
6. do the round again and this time press **Cancel** at the question: it **must
   save nothing**, not even the other changes made in that window.

**What must happen:** no device is left in a room that no longer exists, and none
ends up in recently deleted - they were moved, not deleted.

---

## Way A - reset, then import

Use this if you want to see the inventory empty before loading.

1. open **Settings** and press **Reset inventory**
2. the warning says how many devices will be deleted, **for every user**
3. type `DELETE EVERYTHING` in full in the box, and confirm
4. the program saves a copy of the data file by itself, in the `Backup` folder
   inside the program folder, and then empties
5. check that the room cards are all at zero
6. now *Import xls...* > **The whole inventory** > **Merge**, and choose your
   file

## Way B - an import that replaces everything

One step, identical result.

1. *Import xls...*
2. **The whole inventory** + **Replace**
3. choose your file
4. the summary says how many rows it read and how many devices will be deleted
   before loading
5. type `DELETE EVERYTHING` in full and confirm

Here too the copy goes into `Backup` before anything is touched. If the copy
fails, the operation is cancelled and nothing is changed.

**If you get it wrong**, the copy is right there: open it with the program or
with Excel, it is a complete inventory. The `Backup` folder is never emptied on
its own, so it is worth clearing it out now and then.

## What survives either way

**iPhones are never deleted**, neither by the reset nor by a replacement, in any
state. They do not come from an import - they are entered by hand only - so
deleting them here would mean losing them for good. The program says how many it
kept.

If you added fake iPhones during the trial, remove them by hand before starting:
record the shipment with *Confirm shipment*, and you will be able to delete them
three months later. Alternatively, run the trial without adding any iPhone.

---

## If something does not add up

Report the problem with: which step you were on, what you expected, what
happened, and if you can a screenshot of the summary. On macOS the startup log
is in `avvio.log` in the program folder.
