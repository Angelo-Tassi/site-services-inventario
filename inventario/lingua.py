"""Traduzione dell'interfaccia e dei file prodotti.

Il testo sorgente e' l'italiano: `T("Salva")` ritorna "Salva" in italiano e
"Save" in inglese. Cosi' una stringa non ancora tradotta resta leggibile invece
di sparire, e aggiungere una lingua non richiede di toccare le chiamate.

I dati restano sempre in italiano nel file: stati, nomi delle stanze e dei tipi
sono contenuto, non interfaccia. Vengono tradotti soltanto al momento di
scrivere un'esportazione in inglese, e le importazioni riconoscono entrambe le
lingue.
"""

ITALIANO = "it"
INGLESE = "en"

LINGUE = [("Italiano", ITALIANO), ("English", INGLESE)]

_corrente = ITALIANO

# --------------------------------------------------------------- interfaccia

EN = {
    # --- barra e schermate
    "Aggiungi": "Add",
    "Modifica": "Edit",
    "Elimina": "Delete",
    "Sposta in stanza...": "Move to room...",
    "Importa xls...": "Import xls...",
    "Esporta xls...": "Export xls...",
    "Stampa": "Print",
    "Aggiorna": "Refresh",
    "Impostazioni": "Settings",
    "Reset inventario": "Reset inventory",
    "‹  Home": "‹  Home",
    "Cerca": "Search",
    "Stanza": "Room",
    "Tipo": "Type",
    "Tutte": "All",
    "Tutti": "All",
    "Azzera filtri": "Clear filters",
    "(senza stanza)": "(no room)",
    "Inventario completo": "Full inventory",
    "Scarica il modello di importazione": "Download the import template",
    "Esporta questa stanza in xls": "Export this room to xls",
    "Importa i dati di questa stanza": "Import this room's data",
    "Laptop e tablet in nostro possesso": "Laptops and tablets we hold",
    "Inventario di %s": "%s inventory",
    "Telefoni in nostro possesso - registrati in %s":
        "Phones we hold - recorded in %s",
    "%d dispositivi": "%d devices",
    "%d dispositivi di %d": "%d devices out of %d",
    "tutti i telefoni, ovunque siano registrati":
        "all phones, wherever they are recorded",
    "anche in %s": "also in %s",
    "nessun dispositivo": "no devices",
    "%d in prestito": "%d on loan",
    "nessun prestito in corso": "no loans in progress",
    "Apri l'inventario  ›": "Open the inventory  ›",
    "Prestito": "Loan",
    "Spedizione": "Shipment",
    "Presta": "Lend",
    "Registra rientro": "Register return",
    "Conferma spedizione": "Confirm shipment",

    # --- scheda dispositivo
    "Nuovo dispositivo": "New device",
    "Modifica dispositivo": "Edit device",
    "Tipo *": "Type *",
    "Asset Tag *": "Asset Tag *",
    "IMEI *": "IMEI *",
    "Modello *": "Model *",
    "Numero di serie *": "Serial number *",
    "Restituito da *": "Returned by *",
    "Stanza *": "Room *",
    "Stato": "Status",
    "Note": "Notes",
    "Salva": "Save",
    "Annulla": "Cancel",
    "I campi contrassegnati con * sono obbligatori.":
        "Fields marked with * are required.",
    "Gli iPhone restano sempre in %s.": "iPhones always stay in %s.",
    "Gli iPhone non si prestano: lo stato lo decide la spedizione.":
        "iPhones are never lent: the shipment decides the status.",
    "In prestito: lo stato torna modificabile dopo il rientro.":
        "On loan: the status becomes editable again after the return.",
    "Dati mancanti": "Missing data",

    # --- aggiunta e lettore di codici
    "Aggiungi dispositivo": "Add device",
    "Aggiungi iPhone": "Add iPhone",
    "Che cosa vuoi aggiungere?": "What do you want to add?",
    "Come vuoi aggiungerlo?": "How do you want to add it?",
    "Avanti": "Next",
    "Inserimento manuale": "Manual entry",
    "Scansiona con il lettore di codici": "Scan with the barcode reader",
    "Scansiona l'IMEI con il lettore di codici": "Scan the IMEI with the barcode reader",
    "Passo %d di %d": "Step %d of %d",
    "Scansiona %s": "Scan %s",
    "Scrivi %s": "Type %s",
    "l'asset tag": "the asset tag",
    "il numero di serie": "the serial number",
    "il modello del dispositivo": "the device model",
    "l'IMEI del telefono": "the phone's IMEI",
    "Asset tag": "Asset tag",
    "Numero di serie": "Serial number",
    "Modello/Descrizione": "Model/Description",
    "IMEI": "IMEI",
    "Non riesco a scansionare - inserisci a mano":
        "I can't scan it - type it instead",
    "Digita il valore e premi Invio.": "Type the value and press Enter.",
    "Campo vuoto": "Empty field",

    # --- importazione ed esportazione
    "Importa da Excel": "Import from Excel",
    "Importa in %s": "Import into %s",
    "Esporta in Excel": "Export to Excel",
    "Che cosa vuoi caricare": "What do you want to load",
    "Che cosa vuoi esportare": "What do you want to export",
    "Tutto l'inventario": "The whole inventory",
    "Una sola stanza": "A single room",
    "Come": "How",
    "In che forma": "In what form",
    "Unisci: aggiunge i nuovi e aggiorna quelli gia' presenti":
        "Merge: adds new ones and updates those already there",
    "Sostituisci: svuota prima, poi carica solo il file":
        "Replace: empties first, then loads only the file",
    "Un unico elenco, in un solo foglio": "A single list, in one sheet",
    "Un foglio per ogni stanza, nello stesso file":
        "One sheet per room, in the same file",
    "Un file separato per ogni stanza": "A separate file for each room",
    "Scegli il file": "Choose the file",
    "Importa": "Import",
    "Importa inventario": "Import inventory",
    "Esporta": "Export",
    "Esportazione completata": "Export complete",
    "Importazione completata": "Import complete",
    "Importazione non riuscita": "Import failed",
    "Esportazione non riuscita": "Export failed",
    "Esporta i file in inglese": "Export the files in English",
    "Intestazioni e stati in inglese. Nomi delle stanze e dei\n"
    "tipi restano come li hai scritti tu.":
        "Headers and statuses in English. Room and type names stay\n"
        "exactly as you wrote them.",

    # --- reset
    "Reset dell'inventario": "Inventory reset",
    "Stai per svuotare l'inventario condiviso":
        "You are about to empty the shared inventory",
    "Svuota l'inventario": "Empty the inventory",
    "Conferma non valida": "Invalid confirmation",
    "Inventario svuotato": "Inventory emptied",

    # --- impostazioni
    "Impostazioni inventario": "Inventory settings",
    "Stanze (una per riga)": "Rooms (one per line)",
    "Tipi di dispositivo": "Device types",
    "Stanze con prestito": "Rooms with loans",
    "Stanza degli iPhone": "iPhone room",
    "Lingua": "Language",
    "Vale solo per questo computer.": "Applies to this computer only.",
    "Stanza sconosciuta": "Unknown room",
    "Dato mancante": "Missing value",

    # --- varie
    "Errore": "Error",
    "Operazione non riuscita": "Operation failed",
    "Conferma eliminazione": "Confirm deletion",
    "Eliminazione non consentita": "Deletion not allowed",
    "Scegli stanza": "Choose room",
    "Sposta": "Move",
    "Rientro": "Return",
    "Presta dispositivo": "Lend device",
    "Nome della persona": "Person's name",
    "Registra prestito": "Register loan",
    "Modello creato": "Template created",
    "Modello non creato": "Template not created",
    "Reset": "Reset",
    "Nessun dispositivo": "No devices",
    "Manca la riga della stanza": "The room line is missing",
    "Spedizione registrata": "Shipment recorded",
    "Conferma spedizione?": "Confirm shipment?",
    "Campo mancante": "Missing field",
    "Cartella in cui salvare un file per stanza":
        "Folder in which to save one file per room",
    "Crea il file inventario": "Create the inventory file",
    "Data e ora del prestito vengono registrate in automatico.":
        "Date and time of the loan are recorded automatically.",
    "Esporta %s": "Export %s",
    "Esporta inventario": "Export inventory",
    "File da importare in %s": "File to import into %s",
    "Importazione": "Import",
    "Per confermare, scrivi   %s": "To confirm, type   %s",
    "Salva il modello di inventario": "Save the inventory template",
    "Seleziona il file da importare": "Select the file to import",
    "Seleziona il file inventario": "Select the inventory file",
    "Stampa non riuscita": "Printing failed",
    "Stanza mancante": "Room missing",
    "Una stanza sola sta in un file solo, con il suo nome in testa.":
        "A single room fits in a single file, with its name at the top.",
    "ELIMINA TUTTO": "DELETE EVERYTHING",
    # --- invio per e-mail
    'Invia per e-mail con Outlook':
        'Send by e-mail with Outlook',
    "Apre un messaggio nuovo con il file gia' allegato:\ndestinatario e testo li scrivi tu, l'invio resta a te.":
        'It opens a new message with the file already attached:\nrecipient and text are yours to write, sending stays with you.',
    'Apri il file':
        'Open the file',
    'Ho finito':
        "I'm done",
    '\n... e altri %d':
        '\n... and %d more',
    'Invio per e-mail':
        'Sending by e-mail',
    'Messaggio aperto in Outlook con %s allegato.     ':
        'Message opened in Outlook with %s attached.     ',
    '%d file scritti in:\n%s':
        '%d files written to:\n%s',
    '%s esportati in:\n%s':
        '%s exported to:\n%s',
    '%d dispositivi di %s esportati in:\n%s':
        '%d devices from %s exported to:\n%s',
    "Non c'e' nessun file da allegare.":
        'There is no file to attach.',
    "Outlook non e' stato trovato su questo computer.\n\nIl file e' stato creato lo stesso: allegalo a mano al messaggio.":
        'Outlook was not found on this computer.\n\nThe file was created anyway: attach it to the message by hand.',
    "Non riesco ad aprire Outlook:\n%s\n\nIl file e' stato creato lo stesso: allegalo a mano.":
        'I cannot open Outlook:\n%s\n\nThe file was created anyway: attach it by hand.',

    # --- ripristino
    'Da quale copia vuoi ripartire?':
        'Which copy do you want to go back to?',
    'Dispositivi':
        'Devices',
    'File':
        'File',
    "In inventario ci sono adesso %d dispositivi. La copia scelta\nprendera' il loro posto; lo stato attuale viene salvato prima,\ncosi' puoi tornare indietro anche da qui.":
        'There are %d devices in the inventory right now. The chosen copy\nwill take their place; the current state is saved first, so you\ncan step back from here too.',
    'Inventario ripristinato':
        'Inventory restored',
    'Nessuna copia':
        'No copy',
    "Non c'e' ancora nessuna copia di sicurezza.\n\nNe viene salvata una a ogni reset e a ogni importazione\nche sostituisce i dati.":
        'There is no backup copy yet.\n\nOne is saved at every reset and at every import that\nreplaces the data.',
    'Ripristina':
        'Restore',
    'Ripristina da una copia':
        'Restore from a copy',
    'Ripristina da una copia...':
        'Restore from a copy...',
    "Ripristina l'ultima copia":
        'Restore the latest copy',
    "Ripristinati %d dispositivi dalla copia del %s.\n\nLo stato precedente e' stato salvato in:\n%s":
        '%d devices restored from the copy of %s.\n\nThe previous state was saved in:\n%s',
    'Salvata il':
        'Saved on',
    "Scegli una copia dall'elenco.":
        'Choose a copy from the list.',
    "Ultima copia salvata: %s\nContiene %d dispositivi; adesso in inventario ce ne sono %d.\n\nL'inventario di tutti tornera' com'era in quel momento.\nLo stato attuale viene salvato prima, cosi' puoi tornare indietro.\n\nProcedere?":
        "Latest copy saved: %s\nIt holds %d devices; the inventory currently holds %d.\n\nEveryone's inventory will go back to how it was at that moment.\nThe current state is saved first, so you can step back.\n\nGo ahead?",
    "File Excel": "Excel files",
    "Colonne non riconosciute, il cui contenuto non verra' importato:\n%s\n"
    "Se una di queste e' un dato che ti serve, rinominala come la colonna\n"
    "corrispondente dell'inventario e riprova.":
        "Unrecognised columns, whose contents will not be imported:\n%s\n"
        "If one of these is data you need, rename it like the matching\n"
        "column of the inventory and try again.",
    "Modello_inventario.xlsx": "Import_template.xlsx",
    "Tutti i file": "All files",

    # --- etichette e messaggi con formattazione
    '  %d iPhone lasciati in %s.':
        '  %d iPhones left in %s.',
    '  |  visualizzati: %d':
        '  |  shown: %d',
    ' di %d':
        ' of %d',
    ' e altre %d':
        ' and %d more',
    ' in %s':
        ' in %s',
    '%d dispositivi     %s':
        '%d devices     %s',
    '%d dispositivi  (%s)%s     File: %s':
        '%d devices  (%s)%s     File: %s',
    '%d dispositivi di %s':
        '%d devices from %s',
    "%s e' in prestito a %s: registra prima il rientro.":
        '%s is on loan to %s: register the return first.',
    '%s prestato a %s il %s.     %s':
        '%s lent to %s on %s.     %s',
    '%s rientrato da %s.     %s':
        '%s returned by %s.     %s',
    '%s: %s.':
        '%s: %s.',
    ', un foglio per stanza':
        ', one sheet per room',
    'Aggiunto %s.':
        '%s added.',
    'Elenco ricaricato.':
        'List reloaded.',
    "Eliminare %s - %s dall'inventario?":
        'Delete %s - %s from the inventory?',
    "Eliminare %s dall'inventario?":
        'Delete %s from the inventory?',
    'Eliminato %s.':
        '%s deleted.',
    'Gli iPhone non vengono toccati.':
        'iPhones are not touched.',
    'Inventario aggiornato da un altro utente.':
        'Inventory updated by another user.',
    "L'inventario":
        'The inventory',
    'Lo stato degli iPhone e\' sempre "%s".':
        'iPhones are always "%s".',
    'Nessuno spostamento.':
        'Nothing moved.',
    'Nota aggiornata su %s.':
        'Note updated on %s.',
    'Salvato %s.':
        '%s saved.',
    'Sostituzione':
        'Replacement',
    'Sposta %s in:':
        'Move %s to:',
    'Spostati %d dispositivi in %s.':
        '%d devices moved to %s.',
    'Stampa inviata alla stampante predefinita.     ':
        'Sent to the default printer.     ',
    'Una copia del file dati viene salvata prima di procedere.':
        'A copy of the data file is saved before going ahead.',
    'Unione':
        'Merge',
    "Verranno prima eliminati %d dispositivi gia' in inventario%s.":
        '%d devices already in the inventory%s will be deleted first.',
    'altre/nessuna: %d':
        'other/none: %d',
    "tutto l'inventario":
        'the whole inventory',
    'nessun dispositivo':
        'no devices',
    '%d in prestito':
        '%d on loan',
    "Verranno eliminati %d dispositivi, per tutti gli utenti.":
        "%d devices will be deleted, for every user.",
    "L'operazione non si annulla dal programma.":
        "The operation cannot be undone from the program.",
    "Restano dentro %d iPhone: il reset non li elimina mai,\n"
    "perche' non potrebbero essere ricaricati da un file.":
        "%d iPhones stay in: the reset never deletes them,\n"
        "because they could not be reloaded from a file.",
    "Eliminati %d dispositivi.": "%d devices deleted.",
    "\nMantenuti %d iPhone.": "\n%d iPhones kept.",
    "Aggiunti: %d": "Added: %d",
    "Aggiornati: %d": "Updated: %d",
    "Eliminati prima del caricamento: %d": "Deleted before loading: %d",
    "Copia di sicurezza del file precedente:": "Backup of the previous file:",
    "Scartate %d righe di altre stanze.": "%d rows from other rooms discarded.",
    "In %s - aggiunti: %d, aggiornati: %d": "In %s - added: %d, updated: %d",
    "%d righe valide trovate.": "%d valid rows found.",
    "%d righe ignorate: manca l'identificativo.":
        "%d rows ignored: the identifier is missing.",
    "%d righe hanno preso la stanza dai separatori nel foglio.":
        "%d rows took their room from the separators in the sheet.",
    "%d iPhone ignorati: si inseriscono solo a mano.":
        "%d iPhones ignored: they are only entered by hand.",
    "%d righe di altre stanze scartate.": "%d rows from other rooms discarded.",

    # --- messaggi
    "\n\nCopia di sicurezza del file precedente:\n%s\n\nOra puoi ricaricare l'inventario con Importa xls...":
        '\n\nBackup of the previous file:\n%s\n\nYou can now reload the inventory with Import xls...',
    '%d dispositivi di %s esportati in:\n%s\n\nAprirlo ora?':
        '%d devices from %s exported to:\n%s\n\nOpen it now?',
    '%d file scritti in:\n%s\n\n%s\n\nAprire la cartella?':
        '%d files written to:\n%s\n\n%s\n\nOpen the folder?',
    '%s\n\nCompila il foglio "Inventario" e reimportalo con\nImporta xls...  Le righe con il nome della stanza dividono\nl\'elenco: quello che scrivi sotto finisce in quella stanza.\n\nAprirlo ora?':
        '%s\n\nFill in the "Inventario" sheet and import it back with\nImport xls...  The rows carrying a room name split the\nlist: whatever you write below one goes into that room.\n\nOpen it now?',
    "%s - %s\n\nIl dispositivo e' stato rispedito al servizio telefonia il %s e\nva conservato in inventario per consultazione.\n\nPotrai eliminarlo a partire dal %s.":
        '%s - %s\n\nThe device was shipped back to the phone service on %s and\nmust be kept in the inventory for reference.\n\nYou will be able to delete it from %s.',
    '%s - %s\n\nIn prestito a %s dal %s.\nRegistrare il rientro?':
        '%s - %s\n\nOn loan to %s since %s.\nRegister the return?',
    "%s - %s\n\nQuesto iPhone non e' ancora stato rispedito al servizio\ntelefonia, quindi non puo' essere eliminato dall'inventario.\n\nRegistra prima la spedizione con il pulsante Conferma\nspedizione, nel contenitore Iphone. Da quel momento restera'\nconsultabile per %d mesi, e poi potra' essere eliminato.":
        '%s - %s\n\nThis iPhone has not been shipped back to the phone service\nyet, so it cannot be deleted from the inventory.\n\nRecord the shipment first, with the Confirm shipment button\nin the Iphone container. From then on it stays available for\nreference for %d months, and can then be deleted.',
    "%s - %s\n\nRegistrare la spedizione al servizio telefonia?\n\nData e ora vengono registrate adesso. Il dispositivo resta in\ninventario per consultazione per %d mesi, poi potra' essere eliminato.":
        '%s - %s\n\nRecord the shipment to the phone service?\n\nDate and time are recorded now. The device stays in the\ninventory for reference for %d months, then it can be deleted.',
    '%s esportati in:\n%s\n\nAprirlo ora?':
        '%s exported to:\n%s\n\nOpen it now?',
    '%s non contiene dispositivi da esportare.':
        '%s contains no devices to export.',
    '%s non contiene dispositivi.':
        '%s contains no devices.',
    "%s non puo' restare vuoto.\n\nRiprova la scansione, oppure usa il pulsante per inserirlo a mano.":
        '%s cannot be left empty.\n\nTry scanning again, or use the button to type it in.',
    "%s risulta gia' spedito il %s.":
        '%s was already shipped on %s.',
    'Gli iPhone restano sempre in %s e non possono essere spostati.':
        'iPhones always stay in %s and cannot be moved.',
    "Il dispositivo non e' stato inserito.\nMancano questi dati obbligatori:\n\n%s\n\nCompilali e premi di nuovo Salva.":
        'The device was not added.\nThese required fields are missing:\n\n%s\n\nFill them in and press Save again.',
    'Il dispositivo non risulta in prestito.':
        'The device is not on loan.',
    "Il documento e' stato aperto in Excel.\nUsa File > Stampa per inviarlo alla stampante.":
        'The document was opened in Excel.\nUse File > Print to send it to the printer.',
    "Il foglio dichiara le stanze con le righe-separatore, ma nessuna\nindica %s: non so quali dispositivi siano suoi.\nNon e' stato importato niente.\n\nAggiungi al file una riga vuota con scritto soltanto\n\n        %s\n\nnella prima cella, e sotto elenca i dispositivi della stanza.\n\nNel file ho trovato invece: %s.":
        "The sheet declares rooms with separator rows, but none of them\nnames %s, so I cannot tell which devices belong to it.\nNothing was imported.\n\nAdd to the file an otherwise empty row containing only\n\n        %s\n\nin the first cell, and list the room's devices below it.\n\nWhat I found instead: %s.",
    'Il foglio non dichiara stanze: tutte le righe finiranno in %s.':
        'The sheet declares no rooms: every row will go into %s.',
    "Il tipo decide i campi da compilare e cosa si puo'\nleggere con il lettore di codici.":
        'The type decides which fields are asked for and what\nthe barcode reader can read.',
    'Impossibile salvare le impostazioni:\n%s':
        'Could not save the settings:\n%s',
    'Indica almeno una stanza.':
        'Enter at least one room.',
    'Indica il nome della persona.':
        "Enter the person's name.",
    'Inquadra il codice: il lettore compila il campo e\nconferma da solo. Puoi anche digitarlo.':
        'Point at the barcode: the reader fills the field and\nconfirms by itself. You can also type it.',
    "L'inventario e' gia' vuoto.":
        'The inventory is already empty.',
    "La gestione dei prestiti e' attiva solo per: %s.":
        'Loans are enabled only for: %s.',
    "La riga %s c'e', ma sotto non ci sono dispositivi validi.\nNon e' stato importato niente.":
        'The %s row is there, but there are no valid devices below it.\nNothing was imported.',
    'La scansione compila asset tag e numero di serie con il\nlettore di codici a barre.':
        'Scanning fills in asset tag and serial number with the\nbarcode reader.',
    "La scansione legge l'IMEI dal codice a barre: un iPhone non\nha asset tag ne' numero di serie.":
        'Scanning reads the IMEI from the barcode: an iPhone has\nneither an asset tag nor a serial number.',
    'La spedizione al servizio telefonia riguarda solo gli iPhone.':
        'Shipping to the phone service applies to iPhones only.',
    "La stanza degli iPhone (%s) non e' nell'elenco delle stanze.":
        'The iPhone room (%s) is not in the list of rooms.',
    "Le impostazioni sono salvate accanto al file dati e valgono per tutti gli utenti.\nNelle stanze con prestito ogni riga dell'elenco ha il pulsante Presta / Registra rientro.\nGli iPhone vengono registrati sempre nella stanza indicata qui sopra e non si spostano.":
        'Settings are saved next to the data file and apply to every user.\nIn rooms with loans, each row of the list carries the Lend / Register return button.\niPhones are always recorded in the room named above and never move.',
    'Nessuna riga valida trovata nel file.':
        'No valid rows found in the file.',
    'Nessuna stanza contiene dispositivi.':
        'No room contains any devices.',
    "Non c'e' niente da eliminare: tutti i %d dispositivi in inventario\nsono iPhone protetti dalla conservazione.":
        'There is nothing to delete: all %d devices in the inventory\nare iPhones protected by the retention rule.',
    "Non c'e' nulla da stampare nella vista corrente.":
        'There is nothing to print in the current view.',
    'Non ci sono tipi di dispositivo configurati.':
        'No device types are configured.',
    "Non ho ancora trovato il file dell'inventario.\n\nDi norma si chiama Inventario.xlsx e sta nella stessa cartella del\nprogramma:\n%s\n\nSi'  = apri un file inventario gia' esistente\nNo   = crea qui un nuovo inventario vuoto\nAnnulla = esci":
        "I haven't found the inventory file yet.\n\nIt is normally called Inventario.xlsx and sits in the same folder\nas the program:\n%s\n\nYes    = open an existing inventory file\nNo     = create a new empty inventory here\nCancel = quit",
    'Ogni foglio porta in testa il nome della stanza,\nla data e il numero di dispositivi.':
        'Each sheet carries the room name at the top,\nwith the date and the number of devices.',
    "Per svuotare l'inventario devi scrivere esattamente:\n\n%s":
        'To empty the inventory you must type exactly:\n\n%s',
    "Prima di procedere il programma salva una copia del file dati\nnella cartella Backup, dentro quella del programma: se qualcosa\nva storto, l'inventario si recupera da li'.":
        "Before going ahead the program saves a copy of the data file\nin the Backup folder, inside the program folder: if anything\ngoes wrong, the inventory can be recovered from there.",
    'Prima di una sostituzione una copia del file dati va in Backup.\nGli iPhone non vengono mai toccati: si inseriscono solo a mano.':
        'Before any replacement a copy of the data file goes into Backup.\niPhones are never touched: they are only entered by hand.',
    "Queste stanze con prestito non sono nell'elenco delle stanze:\n%s":
        'These rooms with loans are not in the list of rooms:\n%s',
    'Scegli la stanza.':
        'Choose the room.',
    'Se il foglio dichiara le stanze con le righe-separatore, viene\ncaricata solo la sezione della stanza scelta e il resto si scarta.\nSe non le dichiara, tutte le righe finiscono nella stanza scelta.':
        'If the sheet declares rooms with separator rows, only the section\nof the chosen room is loaded and the rest is discarded.\nIf it declares none, every row goes into the chosen room.',
    'Se il foglio dichiara le stanze, viene caricata solo la\nsezione di questa stanza e il resto si scarta. Se non le\ndichiara, tutte le righe finiscono qui.':
        "If the sheet declares rooms, only this room's section is\nloaded and the rest is discarded. If it declares none,\nevery row ends up here.",
    'Site Services : Inventario Iphone, Laptop e Tablet':
        'Site Services : iPhone, Laptop and Tablet Inventory',
    'Spunta il dispositivo da modificare, oppure fai doppio clic sulla riga.':
        'Tick the device you want to edit, or double-click its row.',
    'Spunta il dispositivo da spostare.':
        'Tick the device you want to move.',
    'Spunta il dispositivo su cui vuoi agire.':
        'Tick the device you want to act on.',
    "Stai per svuotare l'inventario di tutti.\n\nPer procedere scrivi esattamente:\n%s":
        "You are about to empty everyone's inventory.\n\nTo go ahead, type exactly:\n%s",
    'Stampare una pagina separata per ogni stanza?\n\nNo = un unico elenco.':
        'Print a separate page for each room?\n\nNo = a single list.',
    'Collega inventario condiviso...':
        'Link shared inventory...',
    "Scegli la cartella condivisa dell'inventario":
        'Choose the shared inventory folder',
    "Gia' collegato":
        'Already linked',
    "E' gia' questo l'inventario aperto:\n\n%s":
        'This is already the inventory in use:\n\n%s',
    'Collega inventario condiviso':
        'Link shared inventory',
    "L'inventario e' gia' li' e non verra' toccato.":
        'The inventory is already there and will not be touched.',
    "Li' non c'e' ancora nessun inventario: ne verra' creato uno vuoto.":
        'There is no inventory there yet: an empty one will be created.',
    "%s\n\n%s\n\nDa adesso questa postazione lavorera' su quel file, e\nl'inventario aperto ora non verra' piu' usato ne' modificato.\n\nIl programma va chiuso e riaperto. Procedo?":
        '%s\n\n%s\n\nFrom now on this workstation will work on that file, and the\ninventory open now will no longer be used or changed.\n\nThe program has to be closed and reopened. Go ahead?',
    'Collegamento non riuscito':
        'Linking failed',
    'Collegato':
        'Linked',
    'Questa postazione ora apre:\n\n%s\n\nRiapri il programma per lavorarci.':
        'This workstation now opens:\n\n%s\n\nReopen the program to work on it.',
    'Inventario non raggiungibile':
        'Inventory not reachable',
    "Questo programma deve aprire l'inventario condiviso:\n\n%s\n\nIn questo momento non si raggiunge. Di solito e' la cartella\ndi rete che non risponde, o la connessione.\n\nControlla di vedere quella cartella da Esplora risorse, poi\nriapri il programma. Non viene creato nessun inventario\nlocale: si lavora tutti sullo stesso file.\n\nIl percorso e' scritto in:\n%s":
        'This program has to open the shared inventory:\n\n%s\n\nIt cannot be reached right now. Usually it is the network\nfolder not answering, or the connection.\n\nCheck that you can see that folder in File Explorer, then\nreopen the program. No local inventory is created: everybody\nworks on the same file.\n\nThe path is written in:\n%s',
    'Salva copia in locale...':
        'Save a local copy...',
    "Salva una copia dell'inventario":
        'Save a copy of the inventory',
    'Copia non riuscita':
        'Copy failed',
    'Copia salvata':
        'Copy saved',
    '%d dispositivi, come sono in questo momento.':
        '%d devices, exactly as they are right now.',
    "Accanto ai dati e' stato salvato anche il file delle\nimpostazioni: stanze, tipi e stati per rimetterlo\ncom'era.":
        'The settings file was saved next to the data as well:\nrooms, types and statuses, to put it back the way it\nwas.',
    "E' un inventario completo: si apre in Excel, e in caso di\nguaio si ricarica con Ripristina o con Importa xls...\nin modalita' Sostituisci.":
        'It is a complete inventory: it opens in Excel, and if\nsomething goes wrong it is reloaded with Restore, or with\nImport xls... in Replace mode.',
    'La cartella non esiste:\n%s':
        'The folder does not exist:\n%s',
    "L'inventario non c'e' piu':\n%s":
        'The inventory is gone:\n%s',
    'Non riesco a salvare la copia:\n%s':
        'I cannot save the copy:\n%s',
    'Elimina +':
        'Delete +',
    "Elimina piu' dispositivi":
        'Delete several devices',
    'Incolla qui i dispositivi da eliminare':
        'Paste the devices to delete here',
    'Uno per riga. Va bene incollare una colonna di asset tag da\nExcel, o righe intere: viene letto il primo codice che\ncorrisponde a un dispositivo in inventario.':
        'One per line. Pasting a column of asset tags from Excel is fine,\nor whole rows: the first code matching a device in the\ninventory is the one used.',
    'Controlla':
        'Check',
    'Non hai incollato niente.':
        'You have not pasted anything.',
    'Righe incollate: %d':
        'Lines pasted: %d',
    'VERRANNO ELIMINATI: %d':
        'WILL BE DELETED: %d',
    '  %s - %d dispositivi':
        '  %s - %d devices',
    '   [in prestito a %s]':
        '   [on loan to %s]',
    "Non c'e' niente da eliminare.":
        'There is nothing to delete.',
    "SALTATI perche' non si possono eliminare: %d":
        'SKIPPED because they cannot be deleted: %d',
    "SALTATI perche' non sono in inventario: %d":
        'SKIPPED because they are not in the inventory: %d',
    '    e altri %d':
        '    and %d more',
    'Per eliminare questi %d dispositivi, scrivi   %s':
        'To delete these %d devices, type   %s',
    "Stai per eliminare %d dispositivi dall'inventario di tutti.\n\nPer procedere scrivi esattamente:\n%s":
        "You are about to delete %d devices from everybody's inventory.\n\nTo go ahead type exactly:\n%s",
    'Eliminazione annullata':
        'Deletion cancelled',
    'Eliminati %d dispositivi.':
        '%d devices deleted.',
    'Eliminazione completata':
        'Deletion complete',
    '%d dispositivi eliminati.\n\nIn inventario ne restano %d.\n\nCopia di sicurezza del file precedente:\n%s':
        '%d devices deleted.\n\n%d are left in the inventory.\n\nBackup copy of the previous file:\n%s',
    'Dove finiscono:':
        'Where they go:',
    '%d nuovi':
        '%d new',
    '%d aggiornati':
        '%d updated',
    'niente':
        'nothing',
    'Saltate:':
        'Skipped:',
    '  %d senza identificativo':
        '  %d with no identifier',
    'Prima eliminati: %d':
        'Deleted first: %d',
    'In inventario adesso: %d':
        'In the inventory now: %d',
    "Dopo l'importazione: %d":
        'After the import: %d',
    'Conviene una copia':
        'Time for a copy',
    "Hai modificato %d dispositivi dall'ultima copia locale.\n\nLe copie automatiche stanno sulla cartella di rete, accanto\nai dati: se sparisce quella, spariscono anche loro.\n\nVuoi salvare adesso una copia dell'inventario sul tuo\ncomputer? Ci vogliono cinque secondi.":
        'You have changed %d devices since the last local copy.\n\nThe automatic copies live on the network folder, next to\nthe data: if that goes, they go with it.\n\nDo you want to save a copy of the inventory on your own\ncomputer now? It takes five seconds.',
}

# ------------------------------------------------- colonne dei file prodotti

INTESTAZIONI_EN = {
    "Asset Tag": "Asset Tag",
    "Tipo": "Type",
    "Modello/Descrizione": "Model/Description",
    "Numero di serie": "Serial number",
    "IMEI": "IMEI",
    "Restituito da": "Returned by",
    "Stanza": "Room",
    "Stato": "Status",
    "In prestito a": "On loan to",
    "Prestato il": "Lent on",
    "Spedito il": "Shipped on",
    "Note": "Notes",
    "Ultima modifica": "Last change",
    "Modificato da": "Changed by",
}

STATI_EN = {
    "Disponibile": "Available",
    "In prestito": "On loan",
    "Da Rispedire": "To be shipped back",
    "Spedito al servizio telefonia": "Shipped to the phone service",
    "In attesa ritiro": "Awaiting collection",
    "Guasto in attesa tecnico": "Faulty - awaiting technician",
    "Da rebuildare": "To be rebuilt",
    "Controllare": "To be checked",
}


# ------------------------------------------------------------------ funzioni


def imposta(lingua):
    global _corrente
    _corrente = lingua if lingua in (ITALIANO, INGLESE) else ITALIANO
    return _corrente


def corrente():
    return _corrente


def T(testo, lingua=None):
    """La stringa nella lingua richiesta; l'italiano e' il testo sorgente."""
    scelta = lingua or _corrente
    if scelta == INGLESE:
        return EN.get(testo, testo)
    return testo


def intestazione(nome, lingua=None):
    """Il nome di una colonna nei file prodotti."""
    if (lingua or _corrente) == INGLESE:
        return INTESTAZIONI_EN.get(nome, nome)
    return nome


def stato(valore, lingua=None):
    """Uno stato, tradotto solo per la visualizzazione o l'export in inglese."""
    if (lingua or _corrente) == INGLESE:
        return STATI_EN.get(valore, valore)
    return valore


def nome_lingua(codice):
    for nome, valore in LINGUE:
        if valore == codice:
            return nome
    return codice
