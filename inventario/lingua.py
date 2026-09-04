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
    'Copia':
        'Copy',
    'Taglia':
        'Cut',
    'Incolla':
        'Paste',
    'Seleziona tutto':
        'Select all',
    "Copia l'identificativo":
        'Copy the identifier',
    'Copia la riga':
        'Copy the row',
    'Modifica...':
        'Edit...',
    "Non c'e' nessuna riga selezionata.":
        'No row is selected.',
    'Copiato negli appunti: %s.':
        'Copied to the clipboard: %s.',
    "Il tipo non puo' restare vuoto.":
        'The type cannot be left empty.',
    'Controllo generale duplicati':
        'General duplicate check',
    'Nessun duplicato':
        'No duplicates',
    "Ho controllato %d dispositivi: ognuno compare una volta sola.\n\nNessun numero di serie e' ripetuto su due dispositivi diversi.":
        'I checked %d devices: each appears only once.\n\nNo serial number is repeated on two different devices.',
    'Controllati %d dispositivi.':
        '%d devices checked.',
    "DOPPIONI TROVATI: %d identificativi, %d righe in piu'.":
        'DUPLICATES FOUND: %d identifiers, %d extra rows.',
    'TIENE':
        'KEEPS',
    'elimina':
        'deletes',
    '  ... e altri %d identificativi':
        '  ... and %d more identifiers',
    "Si tiene la registrazione modificata piu' di recente.":
        'The most recently changed record is the one kept.',
    'NUMERI DI SERIE RIPETUTI su dispositivi diversi: %d':
        'SERIAL NUMBERS REPEATED on different devices: %d',
    'Questi non vengono toccati: vanno guardati a mano.':
        'These are not touched: they need a look by hand.',
    'Controllo duplicati':
        'Duplicate check',
    "\n\nProcedo a eliminare le righe in piu'?\nUna copia di sicurezza viene salvata prima.":
        '\n\nShall I delete the extra rows?\nA backup copy is saved first.',
    'Duplicati rimossi.':
        'Duplicates removed.',
    'Duplicati rimossi':
        'Duplicates removed',
    "Eliminate %d righe in piu'.":
        '%d extra rows deleted.',
    '    ... e altre %d':
        '    ... and %d more',
    'NON eliminate, protette: %d':
        'NOT deleted, protected: %d',
    'Dispositivi prima: %d':
        'Devices before: %d',
    'Dispositivi adesso: %d':
        'Devices now: %d',
    "%d identificativi compaiono piu' volte nel foglio (%s): vale l'ultima riga.":
        '%d identifiers appear more than once in the sheet (%s): the last row wins.',
    ' e altri':
        ' and others',
    '%s in %s.':
        '%s into %s.',
    'Righe non caricate, o caricate con riserva:':
        'Rows not loaded, or loaded with a reservation:',
    '  %d iPhone: si inseriscono solo a mano':
        '  %d iPhones: entered by hand only',
    '  %d di altre stanze':
        '  %d from other rooms',
    "  %d doppioni nel foglio: tenuta l'ultima riga":
        '  %d duplicates in the sheet: the last row was kept',
    '  %d senza modello: caricate lo stesso':
        '  %d with no model: loaded anyway',
    'Colonne del foglio non riconosciute: %s':
        'Sheet columns not recognised: %s',
    'In inventario adesso: %d dispositivi.':
        'In the inventory now: %d devices.',
    '%d righe valide trovate.':
        '%d valid rows found.',
    '%d righe hanno preso la stanza dai separatori nel foglio.':
        '%d rows took their room from the separators in the sheet.',
    '%d iPhone ignorati: si inseriscono solo a mano.':
        '%d iPhones ignored: they are entered by hand only.',
    '%d righe di altre stanze scartate.':
        '%d rows from other rooms discarded.',
    '%d righe non hanno il modello del dispositivo: verranno importate\ncon quel campo vuoto, da completare a mano.':
        '%d rows have no device model: they will be imported with that\nfield empty, to be completed by hand.',
    'Un unico elenco, in un solo foglio':
        'A single list, in one sheet',
    'Un foglio per ogni stanza, nello stesso file':
        'One sheet per room, in the same file',
    'Un file separato per ogni stanza':
        'A separate file for each room',
    '%s  →  %s':
        '%s  →  %s',
    '%s - %s':
        '%s - %s',
    '%s aggiornato su %s.':
        '%s updated on %s.',
    'Asset Tag':
        'Asset Tag',
    'Modello':
        'Model',
    'Restituito da':
        'Returned by',
    "Non c'e' un altro tipo in cui trasformarlo.":
        'There is no other type to turn it into.',
    "Questo non e' un inventario":
        'This is not an inventory',
    "Questo file non e' un inventario":
        'This file is not an inventory',
    "%s\n\n%s\n\nUn foglio del genere si CARICA in un inventario con\nImporta xls..., non si apre come inventario: aprendolo, le\nrighe separatore diventerebbero dispositivi e nessun\ndispositivo avrebbe una stanza.\n\nScegli Annulla, lascia creare l'inventario vuoto, e importa\nquesto file da dentro il programma.":
        '%s\n\n%s\n\nA sheet like this is LOADED into an inventory with\nImport xls..., it is not opened as one: opening it would turn\nthe separator rows into devices and leave every device\nwithout a room.\n\nChoose Cancel, let the empty inventory be created, and import\nthis file from inside the program.',
    "Il file aperto e' un foglio da IMPORTARE, non un inventario:\n%s\n\nContiene le righe che dividono i dispositivi per stanza (%s),\nche qui compaiono in elenco come se fossero dispositivi. Per\nquesto le stanze restano vuote.\n\nCome sistemare:\n1. chiudi il programma;\n2. cancella il file inventario_percorso.json accanto al programma,\n   se c'e': e' li' che resta memorizzata la scelta sbagliata;\n3. riapri: l'inventario vuoto viene creato da solo in Produzione;\n4. carica questo foglio con  Importa xls...  da dentro il programma.":
        'The file you opened is a sheet to IMPORT, not an inventory:\n%s\n\nIt contains the rows that split devices by room (%s), which\nappear here in the list as if they were devices. That is why\nthe rooms stay empty.\n\nHow to fix it:\n1. close the program;\n2. delete inventario_percorso.json next to the program, if it is\n   there: that is where the wrong choice is remembered;\n3. reopen: the empty inventory is created by itself in Produzione;\n4. load this sheet with  Import xls...  from inside the program.',
    "Nel foglio non c'e' nessuna riga che dichiari una stanza.\nI %d dispositivi verranno importati SENZA STANZA: le schede\ndelle stanze resteranno vuote.\n\nUna riga separatore e' una riga con scritto solo il nome della\nstanza, per esempio  Site Services BAU  (vanno bene anche BAU,\nKIOSK, DISASTER). Vale per tutte le righe che la seguono.":
        'The sheet contains no row declaring a room.\nThe %d devices will be imported WITH NO ROOM: the room cards\nwill stay empty.\n\nA separator row is a row carrying only the room name, for\nexample  Site Services BAU  (BAU, KIOSK, DISASTER work too).\nIt applies to every row that follows it.',
    "%s e' gia' in inventario%s.\n\nNon e' stato inserito niente: due dispositivi non possono avere lo stesso identificativo.\n\nSe e' un dispositivo diverso, controlla il codice; se e' lo stesso, modificalo invece di reinserirlo.":
        '%s is already in the inventory%s.\n\nNothing was added: two devices cannot share the same identifier.\n\nIf it is a different device, check the code; if it is the same one, edit it instead of adding it again.',
    "%s non e' un inventario leggibile:\n%s\n\nNon e' stato ripristinato niente.":
        '%s is not a readable inventory:\n%s\n\nNothing was restored.',
    '%s non risulta in prestito.':
        '%s is not recorded as being on loan.',
    "%s risulta gia' in prestito a %s dal %s.":
        '%s is already on loan to %s since %s.',
    'Gli iPhone non vengono dati in prestito.':
        'iPhones are not lent out.',
    "Il campo %s non si modifica dall'elenco.":
        'The %s field is not edited from the list.',
    "Il dispositivo %s non esiste piu' nell'inventario.":
        'Device %s is no longer in the inventory.',
    'Impossibile accedere alla cartella di rete:\n%s':
        'Cannot reach the network folder:\n%s',
    'Impossibile leggere %s:\n%s':
        'Cannot read %s:\n%s',
    'Impossibile leggere il file:\n%s':
        'Cannot read the file:\n%s',
    'Impossibile salvare %s:\n%s':
        'Cannot save %s:\n%s',
    'Indica il nome della persona a cui presti il dispositivo.':
        'Give the name of the person you are lending the device to.',
    "L'articolo %s non esiste piu': e' stato eliminato da un altro utente.":
        'Item %s no longer exists: another user deleted it.',
    "L'asset tag %s e' gia' presente nell'inventario.":
        'Asset tag %s is already in the inventory.',
    "L'asset tag e' obbligatorio.":
        'The asset tag is required.',
    'La cartella %s non esiste.':
        'The folder %s does not exist.',
    "La copia %s non esiste piu'.":
        'Copy %s no longer exists.',
    'Lo stato degli iPhone e\' sempre "%s" e non si cambia.':
        'The status of an iPhone is always "%s" and cannot be changed.',
    'Nel file non e\' stata trovata la colonna "Asset Tag" (o "IMEI").\nCi deve essere una riga con le intestazioni delle colonne.':
        'The "Asset Tag" (or "IMEI") column was not found in the file.\nThere must be a row with the column headings.',
    "Non riesco a creare la cartella delle copie di sicurezza.\n\nL'operazione e' stata annullata: nessun dato e' stato toccato.":
        'I cannot create the backup folder.\n\nThe operation was cancelled: no data was touched.',
    "Non riesco a creare la copia di sicurezza:\n%s\n\nL'operazione e' stata annullata: nessun dato e' stato toccato.":
        'I cannot create the backup copy:\n%s\n\nThe operation was cancelled: no data was touched.',
    "Non riesco a ripristinare la copia:\n%s\n\nL'inventario e' rimasto com'era.":
        'I cannot restore the copy:\n%s\n\nThe inventory was left as it was.',
    "Non si passa da iPhone a un altro tipo, ne' viceversa.\n\nUn iPhone e' identificato dall'IMEI, gli altri dispositivi dall'asset tag\ne dal numero di serie: il passaggio cancellerebbe l'identificativo.\n\nElimina il dispositivo e reinseriscilo con il tipo giusto.":
        'You cannot switch from iPhone to another type, or the other way round.\n\nAn iPhone is identified by its IMEI, the other devices by asset tag\nand serial number: the switch would erase the identifier.\n\nDelete the device and enter it again with the right type.',
    'Stato non previsto: %s.':
        'Unexpected status: %s.',
    "La cartella condivisa non si raggiunge:\n  %s\n\nAprila prima da Esplora risorse: se non si apre da li', non si\napre nemmeno da qui. Controlla il percorso e la connessione.":
        'The shared folder cannot be reached:\n  %s\n\nOpen it in File Explorer first: if it does not open there, it will\nnot open from here either. Check the path and the connection.',
    "Non si e' potuto salvare la configurazione ne' accanto al\nprogramma ne' nel profilo utente.":
        'The configuration could not be saved, neither next to the program\nnor in the user profile.',
    "Sulla cartella condivisa non si puo' scrivere:\n  %s\n\n%s\n\nServe il permesso di Modifica su quella cartella: chiedilo a\nchi amministra la share.":
        'The shared folder cannot be written to:\n  %s\n\n%s\n\nModify permission on that folder is needed: ask whoever\nadministers the share.',
    "Sulla share non c'e' nessun inventario:\n  %s":
        'There is no inventory on the share:\n  %s',
    "Ho confrontato l'ASSET TAG di %d dispositivi: ognuno compare una\nvolta sola. E' quello l'identificativo, e su quello si controlla.\n\nNemmeno un numero di serie risulta ripetuto, ma quella e' solo una\nverifica in piu': il seriale non identifica il dispositivo.":
        'I compared the ASSET TAG of %d devices: each appears only once.\nThat is the identifier, and that is what the check is on.\n\nNo serial number is repeated either, but that is only an extra\ncheck: the serial does not identify the device.',
    "Confrontato l'asset tag di %d dispositivi.":
        'Compared the asset tag of %d devices.',
    "DOPPIONI TROVATI: %d asset tag registrati piu' volte, %d righe in piu'.":
        'DUPLICATES FOUND: %d asset tags recorded more than once, %d extra rows.',
    '  ... e altri %d asset tag':
        '  ... and %d more asset tags',
    'Nota a margine - numeri di serie ripetuti su asset tag\ndiversi: %d':
        'Side note - serial numbers repeated on different asset\ntags: %d',
    "Il seriale non identifica il dispositivo, quindi questi\nNON sono duplicati e non vengono toccati. Di solito pero'\nsono un errore di battitura, e conviene guardarli.":
        'The serial does not identify the device, so these are NOT\nduplicates and are not touched. They usually are a typo\nthough, and are worth a look.',
    "%d righe ignorate: manca l'asset tag.":
        '%d rows ignored: the asset tag is missing.',
    "%d asset tag compaiono piu' volte nel foglio (%s): vale l'ultima riga.":
        '%d asset tags appear more than once in the sheet (%s): the last row wins.',
    '  %d senza asset tag':
        '  %d with no asset tag',
    "Copia l'asset tag":
        'Copy the asset tag',
    "Copia l'IMEI":
        'Copy the IMEI',
    '  -  %d selezionati':
        '  -  %d selected',
    'Spunta i dispositivi da eliminare.':
        'Tick the devices to delete.',
    'Spunta i dispositivi da spostare.':
        'Tick the devices to move.',
    'In inventario resteranno %d dispositivi.':
        '%d devices will be left in the inventory.',
    'Eliminare %d dispositivi?':
        'Delete %d devices?',
    'Eliminare questo dispositivo?':
        'Delete this device?',
    "Stai per eliminare %d dispositivi dall'inventario di tutti.\nUna copia del file dati viene salvata prima di procedere.":
        "You are about to delete %d devices from everybody's inventory.\nA copy of the data file is saved before proceeding.",
    "Il dispositivo sparisce dall'inventario di tutti.\nUna copia del file dati viene salvata prima di procedere.":
        "The device disappears from everybody's inventory.\nA copy of the data file is saved before proceeding.",
    'Eliminati %d dispositivi.\n\nIn inventario ne restano %d.\n\nCopia di sicurezza del file precedente:\n%s':
        '%d devices deleted.\n\n%d are left in the inventory.\n\nBackup copy of the previous file:\n%s',
    'SPOSTATI IN %s: %d':
        'MOVED TO %s: %d',
    '  da %s - %d dispositivi':
        '  from %s - %d devices',
    "Nessun cambiamento: sono gia' tutti in %s.":
        'Nothing changes: they are already all in %s.',
    "Gia' in %s, restano dove sono: %d":
        'Already in %s, they stay put: %d',
    "RESTANO FERMI perche' sono iPhone: %d":
        'THEY STAY PUT because they are iPhones: %d',
    'Gli iPhone sono sempre registrati in %s.':
        'iPhones are always recorded in %s.',
    'Come restano le stanze:':
        'How the rooms end up:',
    'Conferma spostamento':
        'Confirm move',
    'Sposta %d dispositivi in:':
        'Move %d devices to:',
    'Spostare %d dispositivi in %s?':
        'Move %d devices to %s?',
    'Spostare questo dispositivo in %s?':
        'Move this device to %s?',
    'iPhone non ancora rispedito al servizio telefonia':
        'iPhone not yet sent back to the phone service',
    'Copie e ripristino':
        'Copies and restore',
    "Delle copie automatiche se ne tengono %d: quando ne arriva\n"
    "una nuova, la piu' vecchia viene cancellata.":
        'Of the automatic copies %d are kept: when a new one arrives,\n'
        'the oldest is deleted.',
    'COME RESTANO LE STANZE:':
        'HOW THE ROOMS END UP:',
    'TORNANO ANCHE LE IMPOSTAZIONI:':
        'THE SETTINGS COME BACK TOO:',
    "L'inventario di tutti torna com'era nella copia.\n"
    'Lo stato attuale viene salvato prima, cosi\' puoi tornare indietro.':
        "Everyone's inventory goes back to how it was in the copy.\n"
        'The current state is saved first, so you can go back.',
    'La copia contiene %d dispositivi; adesso in inventario ce ne sono %d.':
        'The copy holds %d devices; the inventory now has %d.',
    'La copia non porta le impostazioni: stanze, tipi e stati\n'
    "restano quelli di adesso.":
        'The copy does not carry the settings: rooms, types and statuses\n'
        'stay as they are now.',
    'Stanze':
        'Rooms',
    "Le righe con * cambiano rispetto a com'e' adesso.":
        'The lines marked * change from how it is now.',
    'Sono tornate anche le impostazioni: stanze, tipi e stati.':
        'The settings came back too: rooms, types and statuses.',
    'La copia non portava le impostazioni: stanze e tipi sono rimasti\n'
    'quelli di prima.':
        'The copy did not carry the settings: rooms and types stayed as\n'
        'they were.',
    'Risale al %s.':
        'It dates from %s.',
    'Scegli la copia da cui ripristinare':
        'Choose the copy to restore from',
    'Ripristinati %d dispositivi da %s.\n\n%s\n\n'
    "Lo stato precedente e' stato salvato in:\n%s":
        'Restored %d devices from %s.\n\n%s\n\n'
        'The previous state was saved in:\n%s',
    "Dentro ci sono l'inventario e le impostazioni: stanze,\n"
    "tipi e stati per rimetterlo com'era.":
        'Inside are the inventory and the settings: rooms, types and\n'
        'statuses to put it back as it was.',
    'Da qui si riparte anche se la cartella di rete sparisce:\n'
    'Impostazioni > Ripristina da un file locale...\n\n'
    'Per guardare i dati in Excel, apri lo zip con un doppio\n'
    "clic: l'inventario dentro e' un .xlsx normale.":
        'From here you can start again even if the network folder is gone:\n'
        'Settings > Restore from a local file...\n\n'
        'To look at the data in Excel, open the zip with a double click:\n'
        'the inventory inside is an ordinary .xlsx.',
    "E' un inventario completo: si apre in Excel, e in caso di\n"
    'guaio si ricarica con Ripristina da un file locale...\n'
    "o con Importa xls... in modalita' Sostituisci.":
        'It is a complete inventory: it opens in Excel, and if something\n'
        'goes wrong it is reloaded with Restore from a local file...\n'
        'or with Import xls... in Replace mode.',
    'Ripristino non riuscito':
        'Restore failed',
    'Ripristina da un file locale...':
        'Restore from a local file...',
    'Ripristina da un file locale':
        'Restore from a local file',
    '(niente)':
        '(none)',
    "Copia dell'inventario":
        'Inventory copy',
    'Copia completa (zip)':
        'Complete copy (zip)',
    'Solo i dispositivi (xlsx)':
        'Devices only (xlsx)',
    'Delle copie automatiche se ne tengono %d: quando ne arriva\n'
    "una nuova, la piu' vecchia viene cancellata.\n"
    'Stanno sulla rete accanto ai dati: se sparisce quella\n'
    'cartella, si riparte da un file locale.':
        'Of the automatic copies %d are kept: when a new one arrives,\n'
        'the oldest is deleted.\n'
        'They live on the network next to the data: if that folder goes,\n'
        'you start again from a local file.',
    'Ripristinare tutto da %s?':
        'Restore everything from %s?',
    "Il file %s non esiste piu'.":
        'The file %s does not exist any more.',
    "%s non e' una copia leggibile: il file e' rovinato o non e'\n"
    'uno di quelli salvati da questo programma.\n\n'
    "Non e' stato ripristinato niente.":
        '%s is not a readable copy: the file is damaged, or it is not one\n'
        'of those saved by this program.\n\n'
        'Nothing was restored.',
    'Non riesco a leggere la copia:\n%s':
        'I cannot read the copy:\n%s',
    "Nella copia %s non c'e' nessun inventario.\n\n"
    "Non e' stato ripristinato niente.":
        'There is no inventory inside the copy %s.\n\n'
        'Nothing was restored.',
    'Stanza ripetuta':
        'Room repeated',
    'Stanza rinominata':
        'Room renamed',
    "Questa stanza compare due volte nell'elenco:\n%s\n\n"
    'Due stanze non possono chiamarsi allo stesso modo.':
        'This room appears twice in the list:\n%s\n\n'
        'Two rooms cannot have the same name.',
    '%s  ->  %s   (%d dispositivi)':
        '%s  ->  %s   (%d devices)',
    '%s\n\nI dispositivi che ci stavano dentro sono stati spostati\n'
    'nella stanza con il nome nuovo.':
        '%s\n\nThe devices that were inside have been moved to the room\n'
        'with the new name.',
    'Tipo ripetuto':
        'Type repeated',
    'Il tipo iPhone non si rinomina':
        'The iPhone type cannot be renamed',
    'Rinomina completata':
        'Renaming done',
    "Questo tipo compare due volte nell'elenco:\n%s\n\n"
    'Due tipi non possono chiamarsi allo stesso modo.':
        'This type appears twice in the list:\n%s\n\n'
        'Two types cannot have the same name.',
    'Stai rinominando "%s" in "%s".\n\n'
    "E' la parola con cui il programma riconosce i telefoni: da li'\n"
    "vengono l'IMEI al posto dell'asset tag, la stanza obbligata, la\n"
    'spedizione al servizio telefonia e il fatto che non si eliminino.\n\n'
    'Rimetti "%s" e salva.':
        'You are renaming "%s" into "%s".\n\n'
        'It is the word the program recognises phones by: from it come the\n'
        'IMEI instead of the asset tag, the fixed room, the shipment to the\n'
        'phone service and the fact that they cannot be deleted.\n\n'
        'Put "%s" back and save.',
    'STANZE:':
        'ROOMS:',
    'TIPI DI DISPOSITIVO:':
        'DEVICE TYPES:',
    '%s\n\nI dispositivi sono stati aggiornati con il nome nuovo.\n'
    "Non hanno perso niente: cambia solo l'etichetta.":
        '%s\n\nThe devices have been updated with the new name.\n'
        'They lost nothing: only the label changes.',
    'Il tipo "%s" non si rinomina.\n\n'
    "E' la parola con cui il programma riconosce i telefoni: da li'\n"
    "vengono l'IMEI al posto dell'asset tag, la stanza obbligata, la\n"
    'spedizione al servizio telefonia e il fatto che non si eliminino.\n'
    'Cambiandola, i telefoni gia\' registrati smetterebbero di essere\n'
    'telefoni.':
        'The type "%s" cannot be renamed.\n\n'
        'It is the word the program recognises phones by: from it come the\n'
        'IMEI instead of the asset tag, the fixed room, the shipment to the\n'
        'phone service and the fact that they cannot be deleted.\n'
        'Change it and the phones already registered would stop being\n'
        'phones.',
    'Eliminati di recente':
        'Recently deleted',
    '‹  Precedenti':
        '‹  Previous',
    'Successivi  ›':
        'Next  ›',
    'Chiudi':
        'Close',
    'Nessun dispositivo eliminato di recente.':
        'No device deleted recently.',
    '%d-%d di %d':
        '%d-%d of %d',
    'torna in %s':
        'goes back to %s',
    'Eliminato il %s da %s  -  %s':
        'Deleted on %s by %s  -  %s',
    'Scegli i dispositivi da ripristinare.':
        'Choose the devices to restore.',
    "Questi dispositivi erano in una stanza che non esiste piu':\n%s\n\n"
    'In che stanza rimetterli?':
        'These devices were in a room that no longer exists:\n%s\n\n'
        'Which room should they go back to?',
    'Ripristina (%d)':
        'Restore (%d)',
    'Ripristinati %d dispositivi.':
        'Restored %d devices.',
    'SALTATI: %d':
        'SKIPPED: %d',
    'Stanza tolta con dispositivi dentro':
        'Room removed with devices inside',
    "Stai togliendo dall'elenco stanze che non sono vuote:\n\n%s\n\n"
    'Questi dispositivi finiscono in Eliminati di recente, da dove\n'
    'si ripristinano scegliendo una stanza. Restano li\' %d giorni.\n\n'
    'Procedere?':
        'You are removing rooms from the list that are not empty:\n\n%s\n\n'
        'These devices go into Recently deleted, from where you restore them\n'
        'by choosing a room. They stay there %d days.\n\n'
        'Go ahead?',
    "Restano qui %d giorni dall'eliminazione, e al massimo %d.\n"
    'Non compaiono in nessuna ricerca, esportazione o stampa.':
        'They stay here %d days from deletion, and at most %d.\n'
        'They appear in no search, no export and no printout.',
    'Ripristinati %d dispositivi dagli eliminati di recente.':
        'Restored %d devices from the recently deleted.',
    "non e' piu' fra gli eliminati di recente":
        'is no longer among the recently deleted',
    'non aveva una stanza: indica dove rimetterlo':
        'had no room: say where to put it back',
    "esiste gia' in inventario":
        'already exists in the inventory',
    "nessuna stanza: la sua e' stata tolta, te la chiedera'":
        'no room: its own was removed, it will ask you',
    'Ripristino':
        'Restore',
    'Ripristina i selezionati':
        'Restore the selected ones',
    'Copia le righe (%d)':
        'Copy the rows (%d)',
    '%d selezionati':
        '%d selected',
    'ELIMINATI DI RECENTE (%d)':
        'RECENTLY DELETED (%d)',
    'Dove vanno gli iPhone?':
        'Where do the iPhones go?',
    "Stai togliendo %s, che e' la stanza degli iPhone.\n\n"
    'I telefoni devono stare in una stanza: scegli quale, e ci\n'
    'andranno tutti, rispediti e non - insieme agli altri %d\n'
    'dispositivi che ci sono dentro.':
        'You are removing %s, which is the iPhone room.\n\n'
        'The phones have to live in a room: choose which one, and they will\n'
        'all go there, shipped or not - together with the other %d devices\n'
        'inside it.',
    'Dove vanno i prestiti?':
        'Where do the loans go?',
    'Dove vanno i dispositivi in prestito?':
        'Where do the devices on loan go?',
    "Stai togliendo %s, e li' dentro c'e' ancora almeno un\n"
    'dispositivo in prestito.\n\n'
    "Un dispositivo in prestito non si elimina, quindi non puo'\n"
    'finire negli eliminati di recente: scegli dove spostare i\n'
    'suoi %d dispositivi. Il prestito resta aperto.':
        'You are removing %s, and there is still at least one device on loan\n'
        'inside it.\n\n'
        'A device on loan is not deleted, so it cannot end up in recently\n'
        'deleted: choose where to move its %d devices. The loan stays open.',
    "Stai togliendo %s, che e' una stanza con prestito.\n\n"
    "Scegli dove spostare i suoi %d dispositivi: la stanza\n"
    "scelta prendera' il suo posto fra quelle con prestito, e\n"
    'anche i dispositivi in prestito si sposteranno.':
        'You are removing %s, which is a room with loans.\n\n'
        'Choose where to move its %d devices: the room you pick takes its\n'
        'place among the rooms with loans, and the devices on loan move\n'
        'too.',
    'Stanza traslocata':
        'Room moved',
    '%d iPhone, rispediti e non.':
        '%d iPhones, shipped or not.',
    '%d erano in prestito e si sono spostati lo stesso:\n'
    'il prestito resta aperto, cambia solo dove risultano\n'
    'registrati.':
        '%d were on loan and moved all the same:\n'
        'the loan stays open, only where they are booked changes.',
    'Si sono spostati tutti.':
        'They all moved.',
    'TORNANO IN INVENTARIO: %d':
        'GOING BACK INTO THE INVENTORY: %d',
    "Non e' tornato dentro niente.":
        'Nothing went back in.',
    "   [la sua stanza non c'e' piu']":
        '   [its room is gone]',
    'Ripristinati %d dispositivi':
        'Restored %d devices',
    "Qualcuno non e' tornato dentro":
        'Some did not go back in',
    'Conferma ripristino':
        'Confirm restore',
    "Tornano nell'inventario di tutti, con la scheda che avevano.":
        "They go back into everyone's inventory, with the record they had.",
    'Ripristinare questo dispositivo?':
        'Restore this device?',
    'Ripristinare %d dispositivi?':
        'Restore %d devices?',
    '(nessuna)':
        '(none)',
    'Stanza dei prestiti':
        'Loan room',
    'Le impostazioni sono salvate accanto al file dati e valgono per tutti gli utenti.\n'
    "Nella stanza dei prestiti ogni riga dell'elenco ha il pulsante Presta / Registra rientro.\n"
    'Gli iPhone vengono registrati sempre nella stanza indicata qui sopra e non si spostano.':
        'The settings are saved next to the data file and apply to every user.\n'
        'In the loan room every row of the list carries the Lend / Register return button.\n'
        'iPhones are always registered in the room named above and do not move.',
    'Ripristina tutto':
        'Restore everything',
    'Ripristina tutto (%d)':
        'Restore everything (%d)',
    'Ripristinare tutti i %d dispositivi del cestino?':
        'Restore all %d devices in the bin?',
    'Ripristinare tutti i %d risultati della ricerca?':
        'Restore all %d results of the search?',
    'Torna anche il cestino della copia: %d eliminati di recente,\n'
    'al posto di quelli di adesso.':
        'The copy\'s bin comes back too: %d recently deleted, in place of the\n'
        'ones there now.',
    'La copia non porta il cestino: gli eliminati di recente\n'
    'restano quelli di adesso.':
        'The copy does not carry the bin: the recently deleted stay as they\n'
        'are now.',
    "\n\nE' tornato anche il cestino di quel momento: %d eliminati\n"
    'di recente.':
        '\n\nThe bin from that moment came back too: %d recently deleted.',
    "\nE' tornato anche il cestino: %d eliminati di recente.":
        '\nThe bin came back too: %d recently deleted.',
    "Dentro ci sono l'inventario, le impostazioni - stanze,\n"
    "tipi e stati per rimetterlo com'era - e gli eliminati\n"
    'di recente.':
        'Inside are the inventory, the settings - rooms, types and statuses to\n'
        'put it back as it was - and the recently deleted.',
    'TOLTI DAGLI ELIMINATI DI RECENTE: %d':
        'TAKEN OUT OF THE RECENTLY DELETED: %d',
    "Erano nel cestino e l'importazione li ha rimessi in":
        'They were in the bin and the import put them back into the',
    'inventario: non possono stare in tutti e due i posti.':
        'inventory: they cannot be in both places.',
    'Erano nel cestino ma sono in inventario: un dispositivo non':
        'They were in the bin but they are in the inventory: a device cannot',
    "puo' stare in tutti e due i posti.":
        'be in both places.',
    '    ...e altri %d':
        '    ...and %d more',
    'Tolto dagli eliminati di recente':
        'Taken out of the recently deleted',
    "%s era anche fra gli eliminati di recente.\n\n"
    "L'ho tolto da li': adesso e' in inventario, in %s.":
        '%s was among the recently deleted too.\n\n'
        'I took it out of there: it is in the inventory now, in %s.',
    "Erano gia' in inventario":
        'They were already in the inventory',
    "%s era gia' in inventario, in %s.\n\n"
    "Non c'era niente da ripristinare: l'ho tolto dagli\n"
    "eliminati di recente, perche' un dispositivo non puo'\n"
    'stare insieme in elenco e nel cestino.':
        '%s was already in the inventory, in %s.\n\n'
        'There was nothing to restore: I took it out of the recently\n'
        'deleted, because a device cannot be in the list and in the bin\n'
        'at the same time.',
    "Questi erano gia' in inventario, e li ho tolti dagli\n"
    "eliminati di recente: un dispositivo non puo' stare\n"
    'insieme in elenco e nel cestino.\n':
        'These were already in the inventory, and I took them out of the\n'
        'recently deleted: a device cannot be in the list and in the bin\n'
        'at the same time.\n',
    "NON IMPORTATI, gia' in inventario: %d":
        'NOT IMPORTED, already in the inventory: %d',
    "La scheda di un dispositivo gia' registrato non si":
        'The record of a device already registered is not',
    "riscrive da un foglio: e' rimasta com'era.":
        'rewritten from a sheet: it stayed as it was.',
    "  %d gia' in inventario: non vengono importate":
        '  %d already in the inventory: they are not imported',
    "%d gia' in inventario":
        '%d already in the inventory',
    '      ...e altre %d':
        '      ...and %d more',
    'in conservazione fino al %s':
        'kept until %s',
'in prestito a %s: registra prima il rientro':
        'on loan to %s: register the return first',
    "%s e' in prestito a %s dal %s.\n\n"
    "Un dispositivo in prestito non si sposta e non si elimina:\n"
    "registra prima il rientro con il pulsante Registra rientro,\n"
    "nella stanza dove e' stato prestato.":
        '%s is on loan to %s since %s.\n\n'
        'A device on loan is neither moved nor deleted:\n'
        'register the return first with the Register return button,\n'
        'in the room where it was lent.',
    "RESTANO FERMI perche' sono in prestito: %d":
        'STAY WHERE THEY ARE because they are on loan: %d',
    'Registra prima il rientro, poi si potranno spostare.':
        'Register the return first, then they can be moved.',
    '%d dispositivi sono in prestito e non si spostano.\n\n'
    'Registra prima il rientro, poi si potranno spostare.':
        '%d devices are on loan and do not move.\n\n'
        'Register the return first, then they can be moved.',
    "%s e' in prestito a %s: registra prima il rientro,\n"
    "poi si potra' spostare.":
        '%s is on loan to %s: register the return first,\n'
        'then it can be moved.',
    '  %d in prestito lasciati dove sono.':
        '  %d on loan left where they are.',
    '%s - %s\n\n'
    "Il dispositivo e' in prestito a %s dal %s.\n\n"
    "Un dispositivo in prestito non si elimina: l'inventario e'\n"
    "l'unica traccia di chi ce l'ha. Registra prima il rientro\n"
    'con il pulsante Registra rientro, poi potrai eliminarlo.':
        '%s - %s\n\n'
        'The device is on loan to %s since %s.\n\n'
        'A device on loan is not deleted: the inventory is the only\n'
        'trace of who has it. Register the return first with the\n'
        'Register return button, then you will be able to delete it.',
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
