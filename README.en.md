# Site Services : iPhone, Laptop and Tablet Inventory

### [Open the project page](https://angelo-tassi.github.io/site-services-inventario/?lang=en) &nbsp;·&nbsp; [Download the program](https://github.com/Angelo-Tassi/site-services-inventario/releases/latest/download/Inventario-windows-senza-exe.zip) &nbsp;·&nbsp; [Italiano](README.md)

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

The Releases page offers **two different packages** that do exactly the same
thing. They differ only in what gets executed.

| Package | What is inside | When to pick it |
| --- | --- | --- |
| **[`Inventario-windows-senza-exe.zip`](../../releases/latest/download/Inventario-windows-senza-exe.zip)** | the official python.org Python, signed by the Python Software Foundation, and the program in plain sight as `.py` files | **recommended.** No executable built by us: there is nothing unsigned to get past security |
| [`Inventario-windows.zip`](../../releases/latest/download/Inventario-windows.zip) | `Inventario.exe`, the program packed into a single executable | if you prefer one file to start and warnings are not a concern |

You may also want the **[Excel template](docs/Modello_inventario.xlsx)** for
loading laptops and tablets already recorded elsewhere.

Something native does run either way - on Windows it has to - but in the first
package it is `python\pythonw.exe`, that is **the official python.org binary**,
with its signature and its reputation. Antivirus software knows it.
`Inventario.exe`, by contrast, is a file born anew with every version, signed by
nobody and never seen before: that is what gets it looked at with suspicion, not
the program itself.

**Nothing has to be installed on the PCs** in either case: Python and the
libraries travel inside the package. Extract it into the shared network folder,
and whoever can reach that folder can use the inventory.

With the recommended package, the share holds this:

```
\\server\Shared\Inventory\
    Inventario.py                   the program, readable: opens in Notepad
    inventario\                     the rest of the program, also in plain sight
    python\                         the official python.org Python
    Crea collegamento sul desktop.bat   creates the desktop shortcut
    LEGGIMI-PRIMA.txt               installation instructions
    Produzione\
        Inventario.xlsx             the real inventory, one for everybody
        inventario_impostazioni.json  rooms, types, loans, statuses
    Backup\                        copies saved before every destructive operation
    Collaudo\                      the test files
```

With the `Inventario.exe` package instead:

```
\\server\Shared\Inventory\
    Inventario.exe                  the program
    _internal\                      Python and the libraries: do not move, do not rename
    Produzione\
        Inventario.xlsx             the real inventory, one for everybody
        inventario_impostazioni.json  rooms, types, loans, statuses
    Backup\                        copies saved before every destructive operation
    Collaudo\                      the test files
```

**There is one inventory, and it lives on the share.** Every technician opens the
same file: there are no local copies on individual workstations, and every change
- an addition, a loan, a restore - is written straight there where everybody sees
it. That is why saves go through a lock and the list refreshes by itself every
fifteen seconds.

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

## Windows security warnings

The program is not signed with a certificate: without a signature, Windows warns
before running anything that came from the Internet. **It is not a fault of the
program, and it can be avoided.**

### The one step that removes them all

**Unblock the zip before extracting it.** Right-click
`Inventario-windows.zip` > *Properties* > at the bottom of the General tab, tick
**Unblock** > *Apply*.

Windows marks everything extracted from a blocked archive as "downloaded from
the Internet", and that mark follows the files onto the share: it is what makes
*"Windows protected your PC"* appear on every workstation. Unblocking the archive
before opening it stops the mark reaching the extracted files, and nobody sees a
warning again.

If you already extracted without unblocking, delete the extracted folder, unblock
the zip and extract again: unblocking the files one by one is not enough.

### The other two cases

**"The publisher could not be verified"**, opening a program from a network path:
it goes away by adding the server to the *Local intranet* sites - Internet
Options > Security > Local intranet > Sites > Advanced > `\\server`.

**Running from a share forbidden by security policy** (AppLocker or software
restriction policies): here the NTFS permission is there but the program still
will not start, and an exception from the administrators is needed.

### Can the executable be avoided altogether?

No, and it is worth understanding why. Something has to run on the PC: the
alternatives are a Python script, which would need Python installed on every
workstation - exactly what we set out to avoid - or a `.bat` file, which Windows
treats with the same suspicion and some antivirus products like even less.

The package does use the **"onedir"** form though: the executable sits next to its
libraries in the `_internal` folder, instead of unpacking itself into a temporary
folder at every start. It starts faster from a share and gives antivirus software
far less to complain about, since self-extracting files are what they look at
hardest.

**The definitive fix is a code-signing certificate.** With one, the executable is
signed by the organisation and SmartScreen says nothing, on any workstation and
without unblocking anything. It costs a few hundred euros a year and has to be
requested from whoever runs corporate IT; if you get one, it is added to the build
in a single step.

## Where to put the program

There are two ways to deploy it, and they differ only in where the executable
lives. In both, **there is one inventory, on the share**.

### A. Everything on the network folder

This is the starting arrangement: extract the package onto the share and put the
shortcut on the desktops. One place to update, nothing to touch on the
workstations.

The drawback is that the executable is **started from the network**, and
corporate security systems watch that behaviour closely: an unsigned binary run
from a share is one of the patterns they flag most often.

### B. Program on the workstations, data on the share

If security is a concern - or if somebody has already flagged the executable -
this is the quieter arrangement, and the program supports it with no changes.

1. copy the program folder (everything in the zip except `Produzione`) onto
   each workstation, for example into `C:\Program Files\Inventario`
   or `%LOCALAPPDATA%\Inventario`;
2. **leave only the data on the share**: the `Produzione` folder with
   `Inventario.xlsx` and the settings inside;
3. on each workstation, next to the executable, create
   `inventario_percorso.json` with the network path:

```json
{ "data_path": "\\\\server\\Shared\\Inventory\\Produzione\\Inventario.xlsx" }
```

   Alternatively start the program once and point it at the file when asked: the
   choice is remembered by itself.

No executable is left on the share, and nobody starts one from the network. The
**backup copies follow the data**: they go into `Produzione\Backup` on the
share, one set for everybody, not onto individual workstations.

The price is that an update has to be distributed to every workstation - by copy,
by software package or by GPO - instead of replacing a single folder.

### If the executable gets flagged

The program is not signed, so it can be intercepted. There are three routes, in
order of solidity:

1. **a code-signing certificate**: solves it at the root, on any workstation and
   with either arrangement;
2. **an allow rule** agreed with whoever runs security - AppLocker or Windows
   Defender Application Control - by path or by file hash;
3. **using the package without an executable of ours**, which removes the
   unsigned binary, and optionally **moving the program locally**, that is
   arrangement B, which also removes the "executable started from a share"
   pattern. Together the two leave little to flag.

It is worth talking to whoever runs security **before** deploying to many
workstations: an unsigned internal tool is a normal situation, usually settled
with one allow rule.

## Running it

Double-click **`Inventario.exe`** in the network folder.

On first run, if it does not find the inventory yet, it offers to create
`Inventario.xlsx` right beside itself; from then on everybody opens it without
being asked anything. The quickest desktop shortcut: right-click the executable
> *Show more options* > *Send to* > *Desktop (create shortcut)*. Alternatively
`Crea collegamento sul desktop.bat`, included in the package, creates one and
leaves a copy in the network folder for the others to drag onto their desktop.

### Building the package yourself

Not needed if you download the release: it is built automatically by
[GitHub Actions](.github/workflows/build-windows.yml) on a Windows machine at
every published version. To build it yourself, on any Windows PC with Python,
double-click `Compila EXE per Windows.bat`: it produces the `Distribuzione`
folder, ready to be copied onto the share.

### Updating the program

Rebuild with `Compila EXE per Windows.bat`, or download the new package from the
Releases page, and replace the contents in the network folder. **The data is
untouched**: it lives in `Produzione`, and the `Backup` and `Collaudo` folders
stay where they are. The desktop shortcuts keep working, because they point at
the same path.

#### If `Inventario.exe` will not delete

It is not a missing permission: **Windows locks the executable while a process
started from it is running**, even when that process runs on somebody else's
computer. One technician who left the program open is enough. With this package
the files inside `_internal` are locked too.

It also happens that the program was closed but the server still holds the
connection open: a stale SMB handle.

**How to find who is holding it**, from the server hosting the share:

*Computer Management* > *Shared Folders* > **Open Files**. Sort by name and look
for `Inventario`: you see the user and the computer using it. From the same
window, right-click > *Close Open File* releases the lock.

With PowerShell, on the server:

```powershell
Get-SmbOpenFile | Where-Object Path -like "*Inventario*" |
    Select-Object ClientUserName, ClientComputerName, Path
```

and to release it:

```powershell
Get-SmbOpenFile | Where-Object Path -like "*Inventario*" | Close-SmbOpenFile -Force
```

**If nobody appears to have it open** and the file stays locked, the handle is
stale on the server side: closing it from the *Open Files* window usually does
it. As a last resort, restart the Server service (`Restart-Service LanmanServer`),
which disconnects every SMB session on that machine: do it out of hours.

**The way to avoid the problem** is to update when nobody is using it: it takes
a few seconds, and a glance at *Open Files* before starting tells you straight
away whether you can go ahead.

#### The file stays locked even after a reboot

If you rebooted your computer and the file still will not delete, the lock is not
on your PC. Before looking elsewhere, one test tells you straight away which
problem you have.

**Check the read-only attribute first**: right-click the file > *Properties* >
at the bottom of the General tab, the **Read-only** box. If it is ticked, clear
it and try again: it is the most banal cause and also the most common, and has
nothing to do with locks. From the command line:

```
attrib -R "\\server\Shared\Inventory\Inventario.exe"
```

If that was not it, **try renaming** `Inventario.exe` to `Inventario_old.exe`.

| What happens | Which problem it is | What to do |
| --- | --- | --- |
| The rename **succeeds** | It is a lock: Windows refuses to delete an executable in use, but allows renaming it | Put the new `Inventario.exe` in its place and delete the old one once whoever held it has closed |
| The rename **fails** with *access denied* | Not a lock, a **permission**: you do not have *Modify* on the folder | Ask whoever runs the share for NTFS modify rights, and check the share permissions too: the stricter of the two wins |
| The rename fails with *file open in another program* | Lock confirmed, from another workstation or from the server | Find who is holding it with *Open Files* or `Get-SmbOpenFile`, as above |

If the share is on a NAS rather than a Windows server, *Open Files* is not there:
the equivalent lives in the NAS admin panel, usually under *SMB service* or
*Connections*. Failing that, rebooting the NAS out of hours releases every handle.

Check as well whether **Offline Files** is enabled on your PC: in that case you
are working on a local copy of the share, and deletions behave oddly until
synchronisation has caught up.

#### The way that unblocks you anyway

You are not obliged to delete that file to move on. Install the new package **in
a new folder**, next to the old one:

1. create `\\server\Shared\Inventory2\` and extract the updated package there,
   remembering to unblock the zip first;
2. **move** - do not copy - the `Produzione` folder from the old installation to
   the new one: the real inventory is in there. Move `Backup` too, if you want to
   keep the copies;
3. redo the desktop shortcuts pointing at the new `Inventario.exe`;
4. delete the old folder once it frees up, at your leisure.

The inventory is not lost and nobody is left without the program. This is also
the procedure to use when an update cannot wait for everyone to close.

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
statuses - *Not available*, *To be shipped back*, *Shipped to the phone service*
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

### Building the inventory from an existing Excel file

The whole inventory can be created by importing: choose *Import xls...*, **The
whole inventory** + **Merge**, and pick the file. A header row is required; at
least one column between **Asset Tag** and **IMEI** must be there. The others are
recognised by name, in Italian or English, for example:

| Field | Accepted names |
| --- | --- |
| Asset Tag | Asset Tag, Asset, Tag, Etichetta, Inventario |
| Type | Type, Device type, Tipo, Tipologia, Categoria |
| Model | Model, Device, Modello, Descrizione, Dispositivo |
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

## Backup copies

Before every operation that deletes data - the **reset** and every **replacing
import**, on the whole inventory or on a single room - the program duplicates
the data file into the **`Backup`** folder, inside the program folder.

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

## What is in the network folder

| File | Contents |
| --- | --- |
| `Inventario.exe` | the program; on its own, with no prerequisites |
| `Produzione\Inventario.xlsx` | the data; it is the inventory itself, openable in Excel |
| `Produzione\inventario_impostazioni.json` | rooms, types, rooms with loans, iPhone room, statuses |
| `Backup\` | the copies saved before every reset and every replacing import |
| `Collaudo\` | the test files and the trial instructions |
| `.Inventario.xlsx.lock` | present for a fraction of a second during a save |
| `inventario_percorso.json` | which file the program opens, and the language preference |

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
