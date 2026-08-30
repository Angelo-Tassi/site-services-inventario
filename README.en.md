# Site Services : iPhone, Laptop and Tablet Inventory

*[Versione italiana](README.md)*

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

**[Download the latest version](../../releases/latest)** from the Releases page.
You may also want the **[Excel template](docs/Modello_inventario.xlsx)** for
loading laptops and tablets already recorded elsewhere.

The package contains `Inventario.exe`: a single file carrying Python and every
library. **Nothing has to be installed on the PCs.** Copy it into the shared
network folder, double-click it, and whoever can reach that folder can use the
inventory.

```
\\server\Shared\Inventory\
    Inventario.exe                  the program
    Inventario.xlsx                 the data, openable in Excel too
    inventario_impostazioni.json    rooms, types, loans, statuses
```

## In short

- **No server, no database.** The data lives in a single `.xlsx` on the network
  folder: it is both the store and the inventory you open in Excel.
- **Folder permissions are the permissions.** No separate user list to keep.
- **Several people at once.** Every save goes through an exclusive lock and
  rewrites the file atomically; the list refreshes by itself.
- **Loans**, **statuses**, **rooms** and **notes** editable on the spot.
- **Italian and English**: the *Language* setting translates the interface,
  the columns and the statuses.

---

## Permissions on the network folder

There is no installation: you copy the executable and that is it. Two different
permissions are needed, on two different objects.

| Object | NTFS permission | Why |
| --- | --- | --- |
| `Inventario.exe` | **Read and execute** | without the execute right Windows will not start a program, readable or not |
| the **folder** | **Modify** | the program creates, replaces and deletes files there, it does not merely write inside `Inventario.xlsx` |

*Modify* on the folder is required because every save is three operations, not
one: it creates the lock file `.Inventario.xlsx.lock` and then **deletes** it,
writes a temporary `Inventario.xlsx.tmp-...`, and **replaces** the inventory
with the temporary one. Granting only *Write* on the file is not enough. The SMB
share permission counts too: the stricter of the two wins.

**Read-only users** can open the program and browse the inventory; they only
fail when they try to change something. It is a legitimate way to grant
consultation access.

Two things that happen on Windows: an executable opened from a network path may
raise the *"the publisher could not be verified"* warning, which goes away by
adding the server to the *Local intranet* sites; and in some environments
running from a share is forbidden by security policy (AppLocker or software
restriction policies), in which case the NTFS permission is there but the
program still will not start, and an exception from the administrators is needed.

## Running it

Double-click **`Inventario.exe`** in the network folder.

On first run, if it does not find the inventory yet, it offers to create
`Inventario.xlsx` right beside itself; from then on everybody opens it without
being asked anything. The quickest desktop shortcut: right-click the executable
> *Show more options* > *Send to* > *Desktop (create shortcut)*. Alternatively
`Crea collegamento sul desktop.bat`, included in the package, creates one and
leaves a copy in the network folder for the others to drag onto their desktop.

### Building the executable yourself

Not needed if you download the release: it is built automatically by
[GitHub Actions](.github/workflows/build-windows.yml) on a Windows machine at
every published version. To build it yourself, on any Windows PC with Python,
double-click `Compila EXE per Windows.bat`.

### Running from source

Needs Python 3.8+ and `openpyxl`. On macOS the system Python uses **Tk 8.5**,
Apple's deprecated build, which makes windows quit right after opening: use the
Python from [python.org](https://www.python.org/downloads/macos/), which carries
Tk 9.0, and rebuild the virtual environment with it.

```bash
git clone https://github.com/Angelo-Tassi/site-services-inventario.git
cd site-services-inventario
pip install -r requirements.txt
python Inventario.py
```

---

## What it does

- **Home** with one card per room: name, number of devices and the breakdown by
  type. Clicking a card opens that room's inventory; below the cards the home
  page still shows the whole inventory.
- After the rooms there is the **Iphone** card, same shape but not a room: a
  shortcut that gathers every phone. iPhones stay recorded in their room and
  appear there as usual.
- Device list with **Asset Tag**, **Type**, **Model**, **Serial number**,
  **IMEI**, **Returned by**, **Room**, **Status**, **On loan to**, **Lent on**,
  **Shipped on** and **Notes**, plus the date and author of the last change.
- **Coloured rows**, always: light green for iPhones, blue for Dell tablets,
  purple for devices already shipped back, red for those on loan. When they
  overlap the most urgent wins: loan, then shipment, then type.
- **Tick to select**: the first column is a checkbox, and you work on **one
  device at a time**.
- **Loans**: the *Loan* column, with a real button on every row, appears **only
  inside a room that handles loans**.
- **Add** asks first *what* you are adding, from a dropdown of the configured
  types, and only then *how*: by hand or with the barcode reader.
- **Notes and status editable in place**: double-click the cell.
- Search across asset tag, model, serial number, IMEI, notes, status, who
  returned the device and who has it on loan; filters by room and type.
- **The list starts from the most recent**: the last device added or changed is
  always on top.
- **Import** from Excel, **export** and **print**, all described below.

### Required fields

| Type | Required |
| --- | --- |
| Laptop, Tablet, ... | Asset Tag, Model, Serial number, Room |
| iPhone | IMEI, Model, Returned by, Room |

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
| **Device types** | the entries of the *Type* dropdown, one per line |
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
statuses - *Not available*, *To be shipped back*, *Shipped to the phone service*
- are not editable.

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

Three statuses are automatic and win over everything: **Not available** while a
loan is in progress, **To be shipped back** for iPhones still with us, and
**Shipped to the phone service** for those already gone. In those cases the
dropdown is filled in and locked, with the reason beside it, and trying to
change it from the list is refused in the status bar at the bottom, without
opening a window.

The status is never written by hand into the file: if an import brings an
unexpected status, it is reset to *Available*.

## Loans

Opening a room listed among the *rooms with loans* (the Digital Kiosk by
default) the list gains the *Loan* column, with a button on every row that
changes with the state of the device. Outside that room the column does not
exist.

*Lend* asks for the person's name and records date and time. While the device is
out, the row is red and the status is *Not available*. *Register return* closes
the loan and puts it back among the available ones.

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

- Before every replacement a **copy of the data file** is saved in the network
  folder, with date and time in the name. If the copy fails, the operation is
  cancelled.
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

| Asset Tag | Type | Model | Serial number |
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

### The import template

To load laptops and tablets in bulk, start from the ready-made template:

- from the program, the **Download the import template** button on the home page,
  to the right of the *Full inventory* title: it comes out in the interface
  language;
- from the project page, or directly at
  [`docs/Import_template.xlsx`](docs/Import_template.xlsx) (English) and
  [`docs/Modello_inventario.xlsx`](docs/Modello_inventario.xlsx) (Italian).

It carries only the columns laptops and tablets need, already split by room by
the separator rows, with dropdowns on *Type* and *Status* and an *Istruzioni*
sheet. The template generated by the program mirrors the rooms and statuses
configured at that moment.

---

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
| `Del` | delete the ticked device |
| click on the checkbox | tick or untick the row |
| double-click on notes | edit the note in the list |
| double-click on status | dropdown to change the status in the list |
| double-click anywhere else | open the device record |

## Where to keep the real inventory

**Not inside this folder.** The repository holds the program and some
demonstration data; the real inventory belongs elsewhere - on the shared network
folder, or in a local folder while you are still trying things out. That way the
data never ends up in a public repository and is never overwritten by a test.

The path is chosen on first run and remembered in `inventario_percorso.json`. To
change it, delete that file and restart, or edit its `data_path` entry.

The `Esempio/` folder contains an `Inventario.xlsx` with thirteen **fake**
devices across the three rooms, two of them on loan in the Digital Kiosk. It is
there to show the program to whoever opens it for the first time, and can be
regenerated with `.venv/bin/python tests/genera_esempio.py`.

## Trying the import

The **`Collaudo/`** folder travels with the program: in the Windows package it
sits next to `Inventario.exe`. It contains two Excel sheets ready to import - a
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
