# Site Services : iPhone, Laptop and Tablet Inventory

### [Open the project page](https://angelo-tassi.github.io/site-services-inventario/?lang=en) &nbsp;·&nbsp; [Download the program](https://github.com/Angelo-Tassi/site-services-inventario/releases/latest/download/Inventario-windows.zip) &nbsp;·&nbsp; [Manual](https://angelo-tassi.github.io/site-services-inventario/manuale.html?lang=en) &nbsp;·&nbsp; [Italiano](README.md)

> **Project page:** <https://angelo-tassi.github.io/site-services-inventario/?lang=en>
> From there you download the program and the Excel template, and read the
> guides, in Italian and English.

Desktop application for Windows that manages the inventory of the devices we
physically hold: iPhones, laptops and tablets, split by room, with loan
tracking, import, export and printing in Excel format.

> **Alpha version.** It works and it is tested, but it is on its first round in
> the field: expect adjustments. Report anything that looks wrong by opening an
> issue.

## First of all: how it is meant to be used

**[Keep it current, don't rebuild it](https://angelo-tassi.github.io/site-services-inventario/tenerlo-aggiornato.html?lang=en)**
- the working method, before any feature.

This program is not for putting the inventory back together before the review:
it is for never having to do that again. Every physical movement - a device
coming in, going out, going on loan, breaking down, being shipped back - must be
recorded **as it happens, by whoever performs it**, in about ten seconds. At
review time there is nothing to prepare: you export and send.

If instead movements pile up for three months and everything is rebuilt with an
import, the effort of keeping an inventory is wasted: the file goes off already
out of date, and nobody knows where a device ended up any more.

## Download

**[Download `Inventario-windows.zip`](../../releases/latest/download/Inventario-windows.zip)**
from the Releases page. Inside are the official python.org Python, signed by the
Python Software Foundation, and the program in plain sight as `.py` files: no
executable built by us, so nothing unsigned to get past corporate security.

You may also want the **[Excel template](docs/Modello_inventario.xlsx)** if you
have laptops and tablets already recorded elsewhere to load in bulk.

## In short

- **No server, no database.** The data lives in a single `.xlsx` on the network
  folder: it is both the store and the inventory you open in Excel.
- **Folder permissions are the permissions.** No separate user list to keep.
- **Several people at once.** Every save goes through an exclusive lock and
  rewrites the file atomically; the list refreshes by itself.
- **Loans**, **statuses**, **types**, **rooms**, **notes** and **descriptions**
  editable on the spot, with a double-click on the cell.
- **Italian and English**: the *Language* setting translates the interface,
  the columns and the statuses.

---

---

## How to install it

**The program goes on the workstations. The inventory goes on the share. One
inventory, for everybody.**

That split is the whole point, and it is worth saying why.

The inventory has to be **one**: if every technician kept a copy, after half a
day they would all differ and none would be true. So the `.xlsx` file lives on
the network folder, and every change - an addition, a loan, a return - is
written straight there, where everybody sees it.

The program, on the other hand, has nothing to share: it is the same code on
every PC. Keeping it on the share brings no benefit and two concrete problems:
an executable started from the network is the pattern corporate security watches
most closely, and its files stay **locked** by Windows while anybody has it open
- even from another computer - so they cannot be updated.

```
On each workstation                     On the network folder
------------------------------          ----------------------------------
C:\Inventario\                          \\server\Shared\Inventory\
    Inventario.py                            Produzione\
    inventario\                                  Inventario.xlsx      <- the inventory
    python\                                      inventario_impostazioni.json
    inventario_percorso.json                 Backup\                  <- the copies
    Collaudo\
```

### Step 1 - prepare one copy, once

On any PC:

1. download the zip and **extract it into a local folder**, for example
   `C:\Inventario`. Not onto the share;
2. double-click **`Collega inventario condiviso.bat`** ("link shared
   inventory");
3. **a window opens**: browse to the network folder and pick it. No paths to
   copy.

If the window does not open, the path can always be pasted by hand -
`\\server\Shared\Inventory`, or the mapped drive letter, `F:\Inventory`.
Quotes, reversed slashes and one trailing slash too many do no harm.

The program takes it from there: if the inventory is not on the share yet, it
creates an empty one at `Produzione\Inventario.xlsx`; if it is already there, it
leaves it alone. Then it writes `inventario_percorso.json` beside itself, and
that is where it remembers which inventory to open.

You need **Modify** permission on that network folder. If you do not have it,
the program says so clearly instead of failing halfway.

### Step 2 - bring the inventory you already have

If an inventory already exists - the file you have been working on - copy it
onto the share **with the right name in the right place**:

```
\\server\Shared\Inventory\Produzione\Inventario.xlsx
```

The name must be exactly `Inventario.xlsx`, and it must sit inside `Produzione`.
Close the program before replacing it, and keep a copy of the old file until you
have checked that the inventory opens and every device is there.

If instead your inventory is a sheet with other columns, or split by room with
separator rows, **do not copy it there**: that one is loaded from inside the
program with *Import xls...*, which is a different thing. See
[The import template](#the-import-template).

### Step 3 - hand the folder out to the workstations

The `C:\Inventario` folder you prepared is already configured: the configuration
travels with it. Copy it as it is onto every workstation - by hand, by script,
by software package or by GPO - always to the same local path.

Then, on each workstation, double-click **`Crea collegamento sul desktop.bat`**:
it puts the icon on the user's desktop and leaves a copy in the folder.

The shortcut points at **`Avvia Inventario.bat`**, not straight at `pythonw.exe`:
on some workstations a shortcut to an executable with arguments is refused by
security policy, while one to a `.bat` file always works.

**If nothing appears on the desktop** - it happens where automatic creation is
blocked - do it by hand and it will not fail: right-click
`Avvia Inventario.bat` > *Show more options* > *Send to* >
*Desktop (create shortcut)*. Alternatively drag onto the desktop the
`Inventario dispositivi.lnk` copy left in the folder, which can also be handed
to other users without running anything.

From then on the technician double-clicks the icon and works on everybody's
inventory, knowing nothing about network paths.

### Changing inventory later

It can also be done from the program, touching no files: *Settings* > **Link
shared inventory...**, browse to the folder and confirm. The program says first
what will happen - if the inventory is already there it is left alone, if not an
empty one is created - and then it has to be reopened.

It is the convenient route when each technician installs the program
themselves: they open it, pick the production inventory folder, and they are
done.

### What the network folder needs

| Who | On what | Permission |
| --- | --- | --- |
| the technicians | `\\server\Shared\Inventory\` and everything in it | **Modify** |

That is all. No execute permission, because nothing is executed from the share;
no user list to keep inside the program: whoever can reach the folder can use
the inventory, whoever cannot does not open it.

### If the share does not answer

The program **stops and says so**, naming the inventory it expected to find. It
does not create a local one: working on a copy nobody else sees would be the
quietest way to lose a day's work.

## Running it

Double-click the **Inventario dispositivi** icon on the desktop.

The window opens on the shared inventory. The version number is in the title
bar: it tells you, when in doubt, which copy you are using.

## Updating the program

Now that the program lives on the workstations, an update is a copy into a
folder of your own, where you do have permissions:

1. download the new package;
2. on each workstation, **replace the local program folder**;
3. **do not touch `inventario_percorso.json`**: that is the line saying where
   the shared inventory is. If you overwrite it, run
   `Collega inventario condiviso.bat` again.

**The data is never touched**: it lives on the share, and no program update goes
near it. The desktop shortcuts stay valid too, because they point at the same
local path.

## Windows security warnings

The package contains no executable built by us, which removes the
most common cause of a warning. One precaution remains, at download time:

**Unblock the zip before extracting it.** Right-click the downloaded file >
*Properties* > at the bottom of the General tab, tick **Unblock** > *Apply*.

Windows marks everything extracted from a marked archive as "downloaded from the
Internet", and that mark follows the files. Unblocking the archive before opening
it stops the mark reaching anything.

**If the "Unblock" box is not there, there is nothing to unblock**: it means
Windows never applied the mark - which happens when the download goes through a
corporate proxy - and the extracted files are already clean.

If running unsigned programs is governed by security policy in your company
(AppLocker, Windows Defender Application Control), you need an allowance from the
administrators: a binary signed by the Python Software Foundation, installed
locally, is however the easiest case to get approved.

## A sheet to import is not an inventory

They are two different things and must never be swapped:

- **the inventory** is `Produzione\Inventario.xlsx` on the share: the program
  opens it and writes into it;
- **a sheet to import** is an Excel file holding devices to load, often split by
  rows carrying a room name. It is loaded with *Import xls...*, from inside the
  program.

If a sheet to import is opened *as* an inventory, the separator rows become
devices and no device has a room. The program notices and warns, but the rule
stands: sheets are imported, not opened.

## When something is off: the diagnostic

The program folder holds **`Diagnostica.bat`**. Double-click it and it writes
`Diagnostica.txt` beside it, opening it in Notepad. It changes nothing.

Inside is what it takes to understand a problem without guessing: the version in
use and where it starts from, which inventory it opens and whether it can write
there, the rooms it knows, what it actually reads from an Excel file, the size of
the window and of the table, and where Windows puts the user's desktop.

The file holds paths and room names, nothing confidential: send it to whoever
helps you and it answers in one go questions that otherwise cost days.

## How the package is built

There is no need to build it: [GitHub Actions](.github/workflows/build-windows.yml)
does it on a Windows machine at every published version. The package is the
official python.org embeddable Python plus the program; the build checks the
signature of `pythonw.exe`, tries a real start, and refuses to publish if even
one file is left read-only.

## Permissions on the network folder

One permission, on one object.

| Object | NTFS permission | Why |
| --- | --- | --- |
| the **shared folder** and everything in it | **Modify** | the program creates, replaces and deletes files there, not just writes inside `Inventario.xlsx` |

No execute permission is needed: nothing runs from the share, the program lives
on the workstations.

*Modify* and not *Write*, because every save is three operations:

1. it creates the lock file `.Inventario.xlsx.lock`, then **deletes it**;
2. it writes a temporary file `Inventario.xlsx.tmp-...`;
3. it **replaces** `Inventario.xlsx` with the temporary one.

Granting only *Write* on the file is not enough: creating and deleting files in
the folder would be missing, and saves would fail. The SMB share, not just NTFS,
must grant write access too: the more restrictive of the two wins.

**Read-only users.** Anyone with only *Read* opens the program and consults the
inventory fine; it fails as soon as they try to change something. It is a
legitimate way to give consultation access.

If somebody keeps `Inventario.xlsx` open in Excel, saves can fail because Windows
locks the file: close Excel and try again. To look at the data in Excel safely,
use *Export xls...*.


## What it does

- **Home** with one card per room: name, number of devices and the breakdown by
  type. Clicking a card opens that room's inventory; below the cards the home
  page still shows the whole inventory.
- After the rooms there is the **Iphone** card, same shape but not a room: a
  shortcut that gathers every phone. iPhones stay recorded in their room and
  appear there as usual.
- Device list with **Asset Tag**, **Type**, **Model/Description**, **Serial number**,
  **IMEI**, **Returned by**, **Room**, **Status**, **On loan to**, **Lent on**,
  **Shipped on** and **Notes**, plus the date and author of the last change.
- **Coloured rows**, always: light green for iPhones, blue for Dell tablets,
  purple for devices already shipped back, red for those on loan. When they
  overlap the most urgent wins: loan, then shipment, then type.
- **Selection**: the first column is a checkbox. A click picks one row;
  **Ctrl+click** adds others, however far apart, and **Shift+click** takes the
  whole range up to that row. The keys must be held down: a click without them
  starts again from a single row, as on Windows. How many are picked is written
  next to the list title.
- **Delete** and **Move to room...** act on all the selected rows, and show a
  summary of what they are about to do before proceeding. *Edit* stays one
  device at a time: a record opens on its own.
- **Loans**: the *Loan* column, with a real button on every row, appears **only
  inside a room that handles loans**.
- **Add** asks first *what* you are adding, from a dropdown of the configured
  types, and only then *how*: by hand or with the barcode reader.
- **Editing in place**, with no window to open: double-click the *Notes* or
  *Model/Description* cell to type (`Enter` saves, `Esc` cancels), *Status* for
  the dropdown. On *Type* the dropdown offers the configured types but you can
  also type or paste a different value, which then appears among the filters
  too. Double-clicking any other column opens the full
  record.
- **Delete +** removes many devices at once: you paste the codes from Excel,
  read what disappears and from which room, and confirm by typing `ACCETTO`.
- Search across asset tag, model, serial number, IMEI, notes, status, who
  returned the device and who has it on loan; filters by room and type.
- **The list starts from the most recent**: the last device added or changed is
  always on top.
- **Import** from Excel, **export** and **print**, all described below.

### Required fields

| Type | Required field | The rest |
| --- | --- | --- |
| Laptop, Tablet, ... | **Asset Tag** | model, serial number, notes: filled in later |
| iPhone | **IMEI** | model, returned by, notes: filled in later |

Only the **identifier** is required: it is the one thing without which the device
does not exist in the inventory. Model and serial number are often not to hand at
the moment an arrival is recorded, and demanding them postpones the entry - which
means losing the row. The **room** does not block saving: it is a dropdown, and
if it is not chosen it starts from the first, so no device is left without one.

An iPhone has **no asset tag, no serial number, and is never lent**. The IMEI is
its only identifier. The rule lives in the data store: even when importing or
editing a record those fields are cleared, and trying to lend an iPhone is
refused. Inside the *Iphone* container those columns do not appear at all.

---

## Settings

The **Settings** button, top right, is the only place where the program is
configured. What you decide there applies to **every user**, because it is saved
in `inventario_impostazioni.json` next to the data file - except the language,
which is a personal preference of the single computer.

| Field | What it does |
| --- | --- |
| **Rooms** | the list of rooms, one per line. The order is the order of the cards on the home page |
| **Device types** | the entries of the *Type* dropdown, one per line. Out of the box: *Laptop*, *Tablet* and *Iphone* |
| **Rooms with loans** | where the *Loan* column appears. They must be names present among the rooms |
| **iPhone room** | where phones always end up |
| **Language** | Italiano or English. The same dropdown is also in the window header |

### Creating or renaming a room

Add a line in the **Rooms** box and save: the card shows up on the home page
right away, empty. To rename one, change the text of its line.

Careful: renaming a room **does not move the devices**, which keep the old name
and appear in a separate card. To bring them across, open them and use *Move to
room...*, or - faster with many devices - export the old room, delete the
devices and import the file into the new room.

If you rename the room named as the *iPhone room* without updating that field,
the program notices and falls back to the first room in the list, so the phones
are never left in a room that does not exist.

### Adding a device type

One line in **Device types** is enough. A type called **iPhone** - written
however you like, case does not matter - switches on all the phone rules by
itself: IMEI instead of asset tag and serial number, no loans, locked room, the
container on the home page, shipment and retention.

### Turning loans on in a room

Write the room name in the **Rooms with loans** box. If the name matches no
room, saving is refused with a warning: that is how you notice a typo instead of
discovering later that the buttons never showed up.

### Adding devices

That does not go through the settings: use **Add** in the toolbar, which asks
for the type first and then whether to type it in or scan it. To load many at
once there is the Excel import, and to empty everything before a full reload
there is *Reset inventory*.

### Statuses

The list of statuses is not in the window: it is the `states` entry of
`inventario_impostazioni.json`, which opens in any text editor. The automatic
statuses - *On loan*, *To be shipped back*, *Shipped to the phone service*
- are not editable.

### The settings file

It is written by itself when the program creates a new inventory, next to the
data file in `Produzione`, with the starting values: three rooms, the three
types, loans on the Digital Kiosk, iPhones in Site Services BAU and the five
statuses. That way the configuration is visible and the same for every
technician, instead of depending on what the program carries inside.

If an entry is removed from the file, the starting values apply again for that
entry. The exception is `loan_rooms`, which can deliberately stay empty: it means
no room handles loans.

## Language

The **Language** dropdown sits in two places: in the window **header**, beside
the title at the top right, and in **Settings**. It switches between Italian and
English and changes everything: buttons, messages, warnings, column names in the
list and the statuses. The program redraws itself immediately, no restart.

The choice belongs **to the computer**: someone working on the same inventory
from another PC can keep it in Italian. The data in the file always stays in
Italian - statuses included - so two users with different languages read the
same inventory without conflicts.

### Exporting in English while working in Italian

The export window carries the **Export the files in English** checkbox: it is
ticked by default when the interface is in English, and can be switched on by
hand when you work in Italian but the file goes to someone who reads English.

It translates the column headers and the statuses. It does **not** translate
room and type names, which are your own text. A file exported in English imports
back without trouble: the English headers are recognised and the statuses return
to their Italian form.

---

## Adding with the barcode reader

Pressing **Add** you first pick the device type, then choose between typing it in
and scanning. What gets read depends on the type.

**Laptops and tablets** - three steps, numbered at the top of the window:

1. **Scan the asset tag**
2. **Scan the serial number**
3. **Type the model**, which is not on the label

**iPhones** - one step only: **scan the IMEI**, their only identifier.

In both cases you end up on the record already filled in: what remains are the
fields the barcode does not carry - model and who returned it for an iPhone,
type and room for the others - then *Save*. The room offered is the one you are
looking at, and for iPhones it is always theirs.

Barcode readers behave like a keyboard: they type into the field and confirm by
themselves, so you move from one code to the next without touching the mouse.

### If a code will not read

Every scanning step carries the **"I can't scan it - type it instead"** button:
it turns that window into manual typing for that field alone, without losing the
steps already done and without leaving the procedure. The field cannot be left
empty: confirming an empty one is refused and the window stays put.

## Statuses

| Status | When |
| --- | --- |
| Available | the starting value |
| Awaiting collection | ready, waiting to be picked up |
| Faulty, awaiting technician | out of order, waiting for the repair |
| To be rebuilt | to be reinstalled before going out again |
| To be checked | to be verified |

You pick it when adding the device, and change it at any time **without opening
any window**: double-click the *Status* cell in the list and choose from the
dropdown that appears in its place. It works from every screen.

Three statuses are automatic and win over everything: **On loan** while a
loan is in progress, **To be shipped back** for iPhones still with us, and
**Shipped to the phone service** for those already gone. In those cases the
dropdown is filled in and locked, with the reason beside it, and trying to
change it from the list is refused in the status bar at the bottom, without
opening a window.

The status is never written by hand into the file: if an import brings an
unexpected status, it is reset to *Available*.

## Two things the program does by itself

**An identifier is never repeated.** Adding a device with an asset tag - or an
IMEI - already present, the program **inserts nothing** and says where the one
holding it is, with its model and room: you see at once whether it is a typo or
a device already recorded.

**Every twenty devices touched, it reminds you of the local copy.** Additions,
edits and deletions are counted together - and it counts *records*, so a bulk
deletion of thirty counts all thirty. On the twentieth the program asks whether
you want to save a copy on your own computer, and you can say no: it comes back
after another twenty. Saving a copy resets the count.

It is there because the automatic copies live on the network folder next to the
data: they cover human error, not the folder disappearing.

## Deleting several devices at once

**`Delete +`**, next to *Delete*, is for removing thirty devices without hunting
them one by one - and without giving up knowing what disappears.

1. **paste the codes** taken from an Excel sheet, one per line. A column of asset
   tags is fine, or whole rows: the first code matching a device in the inventory
   is the one used;
2. press **Check**. The program touches nothing and shows:
   - the devices that **will be deleted**, grouped by room, with the model;
   - those **skipped because they cannot be deleted** - a device on loan, an
     iPhone not yet shipped back, one still in retention - with the reason and
     the date;
   - those **skipped because they are not in the inventory**, listed;
3. to go ahead you type **`ACCETTO`**. A backup copy is saved before deleting,
   and at the end the program says how many devices are left.

Duplicates count once, empty lines are ignored, and iPhones keep every one of
their protections.

## Selecting several devices

A click picks one row. **Ctrl+click** adds others, however far apart;
**Shift+click** takes the whole range up to that row. The keys must be held
down: as soon as you click without them the selection starts again from a single
row - the Windows behaviour.

The chosen rows are recognised by the filled circle in the first column and by
the amber background, the one tint no category uses: a selected row is never
mistaken for a tablet or an iPhone. How many there are is written next to the
list title.

The selection holds **only the visible rows**: filtering or searching thins it
out, and clearing the filter does not bring it back. That is deliberate - it
stops you deleting devices you are not looking at.

**Delete** and **Move to room...** act on several rows. *Edit* stays one device
at a time, and *Delete +* is a different thing: that one is for pasting codes
taken from Excel.

### The summary before acting

Neither operation starts without showing what it will do.

**Delete** lists the devices that will disappear, grouped by room, with the
model; then those **skipped because they cannot be deleted** - a device on loan,
an iPhone not yet shipped back, one still in retention - with the reason; and how
many devices will be left. A **backup copy** is saved before proceeding, and at
the end the program says where it is.

If the selection contains a protected device, **only the others disappear**: that
one stays, and the summary says why.

**Move to room...** lists the devices split by room of origin, how many stay put
and for what reason, and how the rooms stand before and after. If they are
already all in the chosen room it says so, instead of letting you find out at
the end.

### Devices on loan are untouchable

**A device out on loan is neither moved nor deleted.** It holds everywhere:
Delete, Delete +, Move to room... and the *Room* field inside the record too,
which would otherwise be the back door for moving it anyway.

The reason is that while it is in someone's hands the inventory is the only
trace of who has it and where it is booked: moving or deleting it would wipe
that trace out. Loans are opened and closed in the **Digital Kiosk**, the room
where loans are active; to free the device press **Register return** on its row,
and from that moment it moves and deletes like any other.

Every other change stays possible: notes, model and type can be corrected while
the device is out.

## Duplicates

The program does not create them. **Adding a device refuses** an identifier
already present, saying where the existing one is and what it is, and without
inserting anything; an **import updates** the record instead of duplicating it.
If a sheet contains the same identifier twice the last row wins, and the summary
says so before importing.

They get in another way: the data file is an `.xlsx` anyone can open and correct
by hand. That is what **`General duplicate check`** is for, next to the template
button on the home screen:

- it compares the **asset tag** of every device - that is the identifier, and
  for iPhones it is the IMEI - and looks for any that appears more than once;
- it shows **every group**: which record it keeps and which it deletes, with
  room, model and the date of the last change. It keeps the most recent one;
- it asks for confirmation and **saves a backup copy** before deleting;
- at the end it says what it deleted, what it could not touch - protected
  iPhones - how many devices there were and how many there are now;
- it separately flags **serial numbers repeated** on different asset tags: the
  serial does not identify the device, so those are **not duplicates** and are
  not touched - but they are almost always a typo.

## The columns of each room

Each view shows **only** the columns that can hold a value there. A column empty
by construction carries no information: it takes space away from what you have
to read.

The **complete inventory on the home screen** is an overview: it says what a
device is, where it is and how it is - asset tag, type, room, notes, status,
model, serial - in a row you can read without scrolling sideways. The finer
questions are asked inside the room they belong to, and in the overview the
status sums them up: *On loan*, *Shipped to the phone service*.

| Where you are | What goes away |
| --- | --- |
| **Home** | IMEI, loans, returned by, shipped on, last change |
| **Site Services BAU** | Room, On loan to, Lent on |
| **Digital Kiosk** | Room, IMEI, Returned by, Shipped on |
| **Magazzino Disaster Recovery** | Room, loans and the iPhone fields |
| **Iphone container** | Asset tag, serial, loans, room and type |

Every column is **as wide as it needs to be** to show what it holds in full: no
truncated text, no columns to widen by hand. An empty column stays at least as wide
as its own name. If the list outgrows the window you scroll with the bar at the
bottom or with Shift + wheel.

A **coloured vertical line** divides one column from the next, and the same tint
appears as a small bar in the heading: it is there so you do not lose the column
while scrolling a wide list. The colour groups by meaning - blue for identifiers,
green for what the device is, purple for where it is, red for how it is, amber
for the loan. The text stays black on white: the colour lives in the dividers and
the headings, not inside the cells.

The **Room** column always disappears inside a room: it would be the same on
every row, and the name is already written above the list. The same goes for
**type** inside the Iphone container.

The columns depend on **how the inventory is configured**, not on the devices
present at that moment: an empty room shows the same columns it will show when
it is full, and nothing dances while you work. Turn loans on for another room in
the settings and the two columns appear there.

### The same rule in the files that come out

An exported file says **what we have, where it is and what there is to know**.
Four columns: asset tag, type, room, notes - the notes travel with the device,
because they are what a row has that is particular to it.

Status, model, serial number, loans, IMEI, shipments and last change stay out:
they serve whoever is working in front of the list, inside the room they belong
to, not whoever receives the file. **Printing** does carry them, because it is
made for whoever is working.

**An inventory cannot be rebuilt from an export**, because those fields are not
in the file. To put the whole inventory back you use *Save a local copy...*,
which copies the real file and carries everything with it.

## Copy and paste

It works **everywhere**, in every field, with the keyboard or the right mouse
button: the codes come from an Excel sheet and go back to one, and that is the
round trip you make every day.

| Where | What you can do |
| --- | --- |
| in **any text field** | `Ctrl+C` copies, `Ctrl+V` pastes, `Ctrl+X` cuts, `Ctrl+A` selects all; the **right button** opens the same menu |
| in the **cells edited in place** | double-click and paste as in any other program; `Enter` saves, `Esc` cancels |
| in **read-only fields** | no writing, but **always copying**: the quickest way to grab the path of a file you have just exported |
| from the **device list** | `Ctrl+C` copies the selected row, in columns and ready for Excel; the right button also offers *Copy the identifier* |
| in **`Delete +`** | it is made for pasting: a column of asset tags taken from Excel, one per line |

The shortcuts are bound explicitly by the program rather than left to Tk's
defaults, which differ between systems and keyboard layouts: `Ctrl` and `Cmd`
both work, everywhere.

## Loans

Opening a room listed among the *rooms with loans* (the Digital Kiosk by
default) the list gains the *Loan* column, with a button on every row that
changes with the state of the device. Outside that room the column does not
exist.

*Lend* asks for the person's name and records date and time. While the device is
out, the row is red and the status is *On loan*. *Register return* closes
the loan and puts it back among the available ones.

While it is on loan the device **is neither moved nor deleted**: it is already
booked where it needs to be, and whoever looks for it must find it there. The
program skips it and says who has it; to free it press *Register return*.

## Shipping iPhones

In the **Iphone** container every phone not yet shipped back carries the
**Confirm shipment** button on its row, to be pressed when it actually leaves
for the phone service. On confirmation the program records **day and time** in
the *Shipped on* column, sets the status to *Shipped to the phone service*,
colours the row **purple** and shows the sentence that stays valid for the
device.

### When an iPhone can be deleted

An iPhone can be deleted only after it has been shipped back, and in any case
not before three months from that date.

| Situation | Trying to delete it |
| --- | --- |
| Not shipped back yet | Warning pointing at the *Confirm shipment* button. There is no unlock date yet |
| Shipped less than three months ago | Warning with the exact date from which it will be possible |
| Shipped more than three months ago | Deleted normally |

The rule applies **to iPhones only**: laptops and tablets are always deletable.
The constraint lives in the data store, not in the interface, so there is no
screen from which to work around it.

---

## How importing works

Pressing *Import xls...* the first thing that appears is not the file picker,
but the choice of **what** to load and **how**.

**What** - the whole inventory, or a single room chosen from the dropdown.

**How** - *Merge* adds the new ones and updates those already there;
*Replace* empties first, then loads only what the file contains.

Once the file is chosen, a summary shows how many rows were read, what was
ignored and - for a replacement - how many devices will be deleted. Up to that
point nothing has been written: cancelling leaves the inventory as it was.

### Protections on replacement

- Before every replacement a **copy of the data file** is saved in the `Backup`
  folder, inside the program folder, carrying the date of the file being saved.
  If the copy fails, the operation is cancelled.
- Replacing **the whole inventory** requires typing `DELETE EVERYTHING` in full.
  For a single room the confirmation with the numbers in plain sight is enough.
- **iPhones are never deleted**, in either mode: they do not come from an
  import, so a replacement would lose them forever.

### Importing into a single room

Two ways, behaving **identically**: from inside the room, with **Import this
room's data** beside its name, or from the home page with *Import xls...* and
the **A single room** option.

| The sheet | What happens |
| --- | --- |
| **declares the rooms** with separator rows | only the chosen room's section is loaded; every other row is **discarded**, even in the same file |
| **declares no room** | the choice applies: every row goes into the room you picked |
| declares rooms, but **not the chosen one** | **nothing** is imported, and a warning explains which line to add, listing the rooms it found instead |

### Splitting a single inventory by room

If you have one sheet with every device and no *Room* column, you do not need to
add one: split the list with **separator rows**. A row with **a single cell
filled in**, containing a room name, assigns that room to every row that
follows, until the next separator.

| Asset Tag | Type | Model/Description | Serial number |
| --- | --- | --- | --- |
| **BAU** | | | |
| IT-0101 | Laptop | Lenovo ThinkPad T14 Gen 4 | PF4A1B2C |
| **KIOSK** | | | |
| IT-0106 | Laptop | Lenovo ThinkPad T14 Gen 5 | PF5K9M8F |

The recognised tags come **from the configured room names**: the full name works,
and so does any single word that is unambiguous. Case and a trailing colon do not
matter. A row with more than one filled cell is never a separator.

### How to prepare the Excel file

The complete guide, with examples and tables, lives on its own page:
**[How to prepare the Excel file](https://angelo-tassi.github.io/site-services-inventario/formato-xls.html)**
(source: [`docs/formato-xls.html`](docs/formato-xls.html)). It covers the names
recognised for each column, the separator rows that split the rooms, the most
common mistakes and what happens to every row.

In short: one header row with the column names, one row per device, and at least
the **Asset Tag** (or the **IMEI** for phones). Everything else is optional.
Starting from the [template](docs/Modello_inventario.xlsx) you do not even need
to read it.

### When the file has different columns

| In the file | What the program does |
| --- | --- |
| Extra columns (cost, supplier, cost centre...) | ignores them and **lists them for you** before importing |
| Names with different case or spacing | recognises them anyway |
| Two columns for the same field | uses the first and reports the second among the ignored ones |
| Missing **model** | imports anyway and tells you how many rows are left without |
| Missing **asset tag** (or IMEI) | stops with an error and imports nothing |
| A title before the table | skips it and looks for the headers in the first 12 rows |
| Several sheets | reads them all; a sheet named after a room counts as a separator |
| A sheet with no table (instructions, notes) | ignores it |
| Empty rows | skips them |

The delicate point is the **unrecognised columns**: if your file calls the model
*Item description*, that data would be silently lost. That is why the import
window, before asking for confirmation, shows a box with the names of the columns
it did not understand and invites you to rename them.

### iPhones stay out of import and export

iPhones are entered **by hand only**. They are not in the template, they are
ignored if present in an imported file (the import window says how many), and
they never end up in any export, not even the one split by room. A **replacement**
of the inventory does not delete them.

They do appear in the **printout**, which is internal consultation, and of course
in the on-screen inventory and in the data file.

### Building the inventory from an existing Excel file

The whole inventory can be created by importing: choose *Import xls...*, **The
whole inventory** + **Merge**, and pick the file. A header row is required; at
least one column between **Asset Tag** and **IMEI** must be there. The others are
recognised by name, in Italian or English, for example:

| Field | Accepted names |
| --- | --- |
| Asset Tag | Asset Tag, Asset, Tag, Etichetta, Inventario |
| Type | Type, Device type, Tipo, Tipologia, Categoria |
| Model/Description | Model, Model/Description, Device, Modello, Descrizione, Dispositivo |
| Serial number | Serial number, Serial, S/N, SN, Service tag, Numero di serie, Seriale, Matricola |
| IMEI | IMEI, IMEI/MEID, MEID, Codice IMEI |
| Room | Room, Location, Stanza, Locale, Ubicazione, Posizione |
| Status | Status, Stato, Disponibilita' |
| Notes | Notes, Note, Nota, Commenti |

Rows with neither an asset tag nor an IMEI are counted and discarded; the summary
says how many. The full guide, with examples, is
[How to prepare the Excel file](https://angelo-tassi.github.io/site-services-inventario/formato-xls.html?lang=en).

### The import template

To load laptops and tablets in bulk, start from the ready-made template:

- from the program, the **Download the import template** button on the home page,
  to the right of the *Full inventory* title: it comes out in the interface
  language;
- from the project page, or directly at
  [`docs/Import_template.xlsx`](docs/Import_template.xlsx) (English) and
  [`docs/Modello_inventario.xlsx`](docs/Modello_inventario.xlsx) (Italian).

It has **the same columns as an exported file** - *Asset Tag*, *Type*, *Room*,
*Notes* - with dropdowns on *Type* and *Room*, the separator rows ready and an
*Istruzioni* sheet. That way you export, fix things in Excel and import back
without changing format.

The columns the template does not have - *Model/Description*, *Serial number*,
*Status* - **stay importable**: if your sheet contains them they are recognised
by name and loaded all the same. The template generated by the program mirrors the rooms and statuses
configured at that moment.

---

## If the file has different columns from ours

The import is tolerant and does not stop over a "dirty" file. Here is exactly
what happens.

| In the file | What the program does |
| --- | --- |
| Extra columns (cost, supplier, cost centre...) | ignores them and **lists them for you** before importing |
| Names with different case or spacing | recognises them anyway: `  ASSET TAG `, `tipo`, `MoDeLLo`, `s/n` all work |
| Two columns for the same field | uses the first and reports the second among the ignored ones |
| The **model** is missing | imports anyway and tells you how many rows are left without |
| The **asset tag** (or IMEI) is missing | stops with an error and imports nothing |
| A title before the table | skips it and looks for the headings in the first 12 rows |
| Several sheets | reads them all; a sheet titled like a room counts as a separator |
| A sheet with no table (instructions, notes) | ignores it |
| Empty rows | skips them without counting them |
| Rows with no identifier | counts them as discarded and carries on |

The delicate point is the **unrecognised columns**: if your file calls a field
something the program does not know, that data is not imported - and the summary
before importing says so, so you can rename the column and try again.

## Creating the inventory from an existing Excel file

There is no need to enter the devices by hand: *Import xls...* accepts any Excel
file whose **first row holds the headings**. The only required column is the
asset tag - or, for iPhones, the IMEI, which takes its place; the others are
recognised by name, in Italian or in English, for example:

| Column | Headings recognised |
| --- | --- |
| Asset Tag | Asset Tag, Asset, Tag, Etichetta, Inventario |
| Type | Tipo, Tipologia, Categoria, Type |
| Model/Description | Model, Model/Description, Modello, Descrizione, Dispositivo |
| Serial number | Numero di serie, Seriale, Serial Number, S/N, Matricola, Service Tag |
| IMEI | IMEI, IMEI/MEID, MEID, Codice IMEI |
| Returned by | Restituito da, Proprietario, Consegnato da, Owner |
| Room | Stanza, Room, Locale, Ubicazione, Posizione |
| On loan to | In prestito a, Prestato a, Assegnato a, Borrower |
| Status | Stato, Status, Disponibilita' |
| Notes | Note, Nota, Commenti, Notes |

Everything else is ignored, and the summary before importing lists the columns
it did not understand.

## How exporting works

As with importing, *Export xls...* opens a window with two questions first.

**What** - the whole inventory, or a single room chosen from the dropdown.

**In what form**, when exporting everything:

| | |
| --- | --- |
| A single list | every device in one sheet |
| One sheet per room | a single file, with one sheet per room inside |
| A separate file for each room | you pick a folder, and out comes one file per room, named `Inventario_<Room>_<date>.xlsx` |

Rooms with no devices produce neither empty sheets nor empty files.

**Every sheet says whose it is.** The room name appears in three places: it is
the sheet name, it is written at the top of the first row, and below it there is
the export date with the number of devices. The *Room* column always stays in
the table.

**Everything can be imported back.** Whatever the export produces can be
reloaded: a file with several sheets is read in full, and each sheet name counts
as that room's separator row.

Inside a room there is also the **Export this room to xls** shortcut, which skips
the window and produces that room's file directly.

### Sending it by e-mail

Every export, whichever form you chose, ends with the same question: **Send by
e-mail with Outlook**, *Open the file*, or nothing.

Choosing to send opens a **new Outlook message with the file already attached**.
Recipient, subject and text are yours to write, and sending stays with you: the
program prepares the message and stops there. If the export produced several
files - one per room - they are collected into a **zip archive**, because Outlook
accepts a single attachment from the command line.

If Outlook is not installed on that computer, the program says so and reminds you
that **the file was created anyway**: just attach it by hand.

There is no menu entry for sending: it is an option at the end of the export,
because what you send is what you have just produced.

## Going back after a mistake

If an import goes wrong - duplicated devices, the wrong room, a file that was
not the right one - there is no need to fix it by hand: you go back to the good
version.

**From the toolbar**, the **Restore** button offers the **latest saved copy**,
telling you when it is from and how many devices it held compared to the ones
there now. That is the frequent case: the last destructive operation is undone
in two clicks.

**From *Settings* > *Restore from a copy...*** you choose among all the copies
available instead, listed newest first with date, time and number of devices:
useful when the mistake goes back a few steps.

Restoring **saves the current state first** into a new copy, so even a wrong
restore can be undone. An unreadable or missing copy is refused without touching
the inventory.

Restoring acts on the shared inventory: what you roll back, every technician
sees.

## The local copy

The automatic copies live on the network folder, next to the data. They cover
human error - one reset too many, a botched import - but they do **not** cover
the network folder disappearing, or somebody deleting inside it: in that case
they go too.

That is what **`Save a local copy...`** in the toolbar is for. It saves, wherever
you decide - your PC, a USB stick, off the network - a copy of the inventory
**as it is at that second**:

- it is taken from the file on the network at the moment you ask, not from what
  the program had read earlier;
- if another technician is saving right then, the program waits for them to
  finish: you never get a file caught mid-write;
- the settings file is saved next to the data, with the same name and the
  `_impostazioni.json` suffix: the data alone would not be enough to put the
  inventory back as it was;
- **it is a complete inventory, not an extract**: it opens in Excel, and it is
  reloaded with *Restore* or with *Import xls...* in Replace mode.

The suggested name carries the date and time,
`Inventario_2026-08-31_18-30.xlsx`. It is worth doing before every big
operation, and now and then out of habit.

## Automatic backup copies

Before every operation that deletes data - the **reset** and every **replacing
import**, on the whole inventory or on a single room - the program duplicates
the data file into the **`Backup`** folder, next to the data on the network.

The name carries **the date of the file being saved**, not the date of the copy:
`Inventario_2026-08-31_09-12-45.xlsx`. That way two resets in a row on the same
inventory do not produce two identical files, and looking for a version you go
by when the contents date from rather than when somebody pressed a button. If a
copy with that name already exists, a numbered one is added.

**If the copy fails, the operation is cancelled** and nothing is touched. If the
program folder is read-only, the program falls back to a `Backup` folder next to
the data file, then to the user profile: a copy has to be writable.

To recover, open the backup file with the program or with Excel: it is a
complete inventory, not a special format. The copies **never end up in the
repository**: the folder carries its own rule excluding them.

They have to be cleared out by hand now and then: nobody deletes them for you.

## Emptying the inventory to reload it

The **Reset inventory** button, top right, is there to start over before a full
reimport. Before anything happens the program shows a warning with how many
devices will be deleted **for every user** and asks you to type `DELETE
EVERYTHING` in full, then **saves a copy** of the data file in the same network
folder with date and time in the name. If the copy fails, the reset is cancelled
and nothing is touched.

**iPhones always stay.** The reset deletes none of them, in any state: they do
not come from an import, so deleting them here would mean losing them for good.
That holds for those shipped more than three months ago too, which could be
deleted by hand.

## Printing

A4 landscape, headers repeated on every page, page numbers and date. It goes to
the default Windows printer; elsewhere the document is opened so it can be
printed from the spreadsheet application.

---

## Opening the files on a Mac with Numbers

Everything the program produces is **standard `.xlsx`** (Office Open XML).
Numbers opens it with a double-click; the same goes for LibreOffice and Google
Sheets. Two things to know: Numbers **does not import the dropdowns** of *Type*
and *Status* in the template - the allowed values stay written in the
*Istruzioni* sheet - and Numbers saves in its own `.numbers` format, which the
program cannot read: to reimport a file edited there, use *File > Export to >
Excel*.

## What is where

The **network folder** holds only the data:

| Path | Contents |
| --- | --- |
| `Produzione\Inventario.xlsx` | the inventory; already the file to consult, openable in Excel |
| `Produzione\inventario_impostazioni.json` | rooms, types, loan rooms, iPhone room, statuses |
| `Produzione\Backup\` | copies saved before every reset and every replacing import |
| `.Inventario.xlsx.lock` | present only for a fraction of a second during a save |

The **workstation** holds only the program:

| Path | Contents |
| --- | --- |
| `Inventario.py`, `inventario\` | the program |
| `python\` | the official python.org Python |
| `inventario_percorso.json` | which inventory to open, and the language preference |
| `Collaudo\` | the test files and the test instructions |

The program on the workstation can be replaced at any moment with no
consequences: everything that matters is on the share.

## Simultaneous access

Several people can keep the application open at the same time.

- Every save happens inside an **exclusive lock** and rewrites the file
  atomically, so the inventory can never be left half-written.
- Inside the lock the data is **re-read from disk** and the change is reapplied
  on top of it: two users editing different records do not overwrite each other.
- The on-screen list refreshes by itself every 15 seconds when somebody else
  saves; `F5` forces a reload.
- A lock left behind by a machine that went down is ignored and removed after
  two minutes.

If somebody keeps `Inventario.xlsx` open in Excel, saves may fail because
Windows locks the file: close Excel and try again. To read the data in Excel
safely, use *Export xls...*.

## Shortcuts

| Key | Action |
| --- | --- |
| `Ctrl+N` | add device |
| `Ctrl+F` | jump to the search box |
| `Ctrl+P` | print the current view |
| `Esc` | back to the home page |
| `Del` | delete the ticked devices |
| click on the checkbox | tick or untick the row |
| `Ctrl`+click | add a row to the selection, however far apart |
| `Shift`+click | select the whole range up to that row |
| double-click on notes | edit the note in the list |
| double-click on model/description | edit the description in the list |
| double-click on status | dropdown to change the status in the list |
| double-click on type | dropdown to change the type, or type/paste one |
| double-click anywhere else | open the device record |
| `Ctrl+C` on the list | copy the selected row, ready for Excel |
| right-click on the list | copy the identifier or the whole row |
| `Ctrl+C` / `Ctrl+V` in a field | copy and paste; the right button opens the menu |

## Where to keep the real inventory

**On the network folder, at `Produzione\Inventario.xlsx`. Never inside the
program folder.**

The program on the workstation carries demo data and test files: the real
inventory lives elsewhere, one of it, where everybody sees it. That way it does
not end up in a repository, does not get overwritten by a test, and no program
update goes near it.

Which inventory to open is written in `inventario_percorso.json`, beside the
program. To change it, run `Collega inventario condiviso.bat` again.

## Trying the import

The **`Collaudo/`** folder travels with the program: in the Windows package it
sits next to `Inventario.py`. It contains two Excel sheets ready to import - a
regular one with **30 devices, 10 per room**, and one deliberately carrying the
cases the program has to report - and step-by-step instructions in
[`Collaudo/README.en.md`](Collaudo/README.en.md).

The checks are run on the included files, which stay where they are. The
instructions also explain where to keep the **real inventories** - in a personal
folder on your own computer, never in the program folder, which is replaced at
every update - and **how to load the final inventory** once the trial is over:
reset first, or an import that replaces everything.

## Automated tests

```bash
.venv/bin/python tests/run_all.py
```

Eighteen suites covering the data store and simultaneous access, path
resolution, screens and colours, the device record, loans, shipments, import,
export and translation. They run without needing a window on screen and never
touch real data: each one builds its own inventory in a temporary folder.

## Code layout

```
Inventario.py            application entry point
inventario/config.py     data file path, shared settings, language
inventario/store.py      reading and writing the .xlsx, lock, operations
inventario/excel_io.py   export, template, print layout, printing
inventario/lingua.py     Italian and English translations
inventario/theme.py      palette, fonts and interface styles
inventario/ui.py         graphical interface
tests/                   automated test suites
Collaudo/                sample files and manual testing instructions
```
