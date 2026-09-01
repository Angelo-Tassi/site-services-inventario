# File di prova per il collaudo

*[English version](README.en.md)*

Questa cartella fa parte del programma e viaggia con lui: serve a provare
l'importazione con dati finti, prima di caricare l'inventario vero.

## Dove sono i file

Sono nella cartella **`Collaudo`**, dentro la cartella del programma sulla
postazione:

```
C:\Inventario\                      il programma, sulla postazione
    Inventario.py
    inventario\
    python\
    inventario_percorso.json     dice quale inventario aprire
    Collaudo\
        Inventario_di_prova.xlsx
        Inventario_di_prova_con_difetti.xlsx
        README.md                questo documento
        README.en.md             la versione inglese

\\server\Condivisa\Inventario\      la cartella di rete: solo dati
    Produzione\
        Inventario.xlsx          l'inventario vero, uno solo per tutti
        Backup\                  le copie di sicurezza
```

Se hai scaricato lo zip `Inventario-windows.zip` dalla pagina Releases, la
cartella `Collaudo` e' gia' dentro. Se lavori dai sorgenti, e' `Collaudo/` nella
radice del progetto.

Quando ti verra' chiesto di scegliere un file, punta qui: **non serve copiare
niente altrove**.

## Cosa contengono

### `Inventario_di_prova.xlsx` - 30 dispositivi, tutto regolare

Un foglio unico chiamato *Inventario*, con l'intestazione delle colonne nella
prima riga e i dispositivi divisi in tre blocchi dalle righe-separatore
azzurre: `SITE SERVICES BAU`, `DIGITAL KIOSK`, `MAGAZZINO DISASTER RECOVERY`.
Dieci dispositivi per blocco.

| Colonna | Contenuto |
| --- | --- |
| Asset Tag | `IT-BAU-101` ... `IT-KSK-210` ... `IT-DRC-310`, uno per ogni stanza |
| Tipo | sette Laptop e tre Tablet per stanza |
| Stato | un misto dei cinque stati, per vedere come si comportano |
| Note | alcune compilate, altre vuote |
| Modello/Descrizione | Lenovo ThinkPad T14 Gen 4 e Gen 5, Dell Latitude 7320 Detachable e 7230 Rugged Extreme |
| Numero di serie | seriali plausibili, tutti diversi |

Le colonne sono nello stesso ordine del [modello da compilare](../docs/Modello_inventario.xlsx)
e dell'elenco nel programma. L'ordine comunque non conta: l'importazione
riconosce le colonne dal nome.

Nessun iPhone: i telefoni si inseriscono solo a mano.

### `Inventario_di_prova_con_difetti.xlsx` - i casi da segnalare

Stessa impostazione, ma con dentro **apposta** tutto quello che il programma
deve saper riconoscere e dire. Oltre ai due separatori ci sono sei righe:

| Cosa c'e' | Cosa deve fare il programma |
| --- | --- |
| Tre colonne in piu': `Costo`, `Fornitore`, `Centro di costo` | ignorarle, ed elencarle prima di importare |
| Una riga senza modello (`IT-BAU-902`) | importarla lo stesso, dicendo quante ne mancano |
| Una riga senza asset tag ne' IMEI | scartarla e contarla |
| Una riga di tipo `Iphone`, con IMEI | ignorarla: i telefoni non si importano |
| Una riga completamente vuota | saltarla senza contarla |
| I separatori scritti in forma breve, `BAU` e `KIOSK` | riconoscerli lo stesso |

Sono i file usati anche dai test automatici: se il comportamento del programma
cambia, i test falliscono e queste istruzioni vengono riscritte. Si rigenerano
con `.venv/bin/python tests/genera_file_di_prova.py`.

## I file di prova restano qui

Non spostarli e non copiarli altrove: fanno parte del programma, vengono
aggiornati insieme a lui e servono a chiunque debba rifare il collaudo.

**Questi file si caricano solo con  Importa xls...  da dentro il programma.**
Non sono inventari: sono fogli da importare, e contengono le righe separatore
che dividono i dispositivi per stanza. Se ne apri uno *come* inventario, quelle
righe diventano dispositivi e nessun dispositivo ha una stanza. Il programma
ormai se ne accorge e te lo dice, ma la regola resta: si importano, non si
aprono.

## Gli inventari veri stanno fuori

Il foglio Excel con l'inventario definitivo - quello che caricherai davvero, con
i dispositivi reali - **non va messo nella cartella del programma**. Tienilo in
una cartella tua, sul tuo computer: `Documenti\Inventario`, il Desktop, dove
preferisci.

Due motivi:

- la cartella del programma viene **sostituita a ogni aggiornamento**: un file
  lasciato li' si perde senza preavviso;
- e' una cartella locale della tua postazione: quello che ci lasci non lo vede
  nessun altro, e non e' un posto dove tenere qualcosa che conta.

Vale anche per i file che **esporti** dal programma e per le copie di sicurezza
che decidi di conservare: salvali in una cartella tua.

## Da non confondere: il file che il programma usa

Diverso e' il file su cui il programma lavora, quello che legge e scrive di
continuo. Non lo importi: lo apre e basta. Sta sulla **cartella di rete
condivisa**, in `Produzione\Inventario.xlsx`, ed e' lo stesso per tutti i
tecnici.

Quale sia e' scritto in `inventario_percorso.json`, accanto al programma sulla
postazione; per cambiarlo si ripassa `Collega inventario condiviso.bat`.

---

# Come testare l'importazione

Le prove che seguono lavorano sull'inventario che il programma sta usando in
quel momento. Falle prima di caricare i dati definitivi: al punto 7 si svuota
tutto, quindi e' il momento giusto per provare mentre l'inventario e' ancora
finto.

## 1. Primo caricamento

1. *Importa xls...*
2. **Tutto l'inventario** + **Unisci**, poi *Scegli il file*
3. apri la cartella `Collaudo` accanto al programma e scegli
   `Inventario_di_prova.xlsx`
4. il riepilogo deve dire **30 righe valide** e **30 righe hanno preso la stanza
   dai separatori**
5. conferma con *Importa*

**Cosa deve risultare.** In home tre schede stanza con **10** ciascuna, e
l'inventario completo sotto. I tablet Dell sono azzurri, i laptop no.

## 2. L'unione non duplica

Ripeti esattamente il punto 1. Al termine il messaggio deve dire
**Aggiunti: 0, Aggiornati: 30**, e i totali delle schede devono restare 10, 10 e
10. Se qualcosa fosse duplicato, i conteggi salirebbero.

## 3. Sostituire una sola stanza

1. *Importa xls...* > **Una sola stanza** > `Digital Kiosk` > **Sostituisci**
2. scegli di nuovo `Inventario_di_prova.xlsx`
3. il riepilogo dice **10 righe valide**, **20 righe di altre stanze scartate**,
   e avverte che verranno prima eliminati **10** dispositivi

**Cosa deve risultare.** Digital Kiosk contiene ancora **10** dispositivi,
quelli elencati sotto la riga `DIGITAL KIOSK`: il foglio dichiara le stanze,
quindi comandano i separatori e le righe delle altre stanze vengono buttate via.
Le altre due stanze restano com'erano.

Nella cartella **`Backup`**, dentro quella del programma, deve essere comparso
un file `Inventario_<data>.xlsx`: e' la copia di sicurezza, presa prima di
toccare qualsiasi cosa.

### Quando il foglio non dichiara le stanze

Apri `Inventario_di_prova.xlsx`, **cancella le tre righe-separatore** e salva
con un altro nome **in una cartella tua**, per lasciare intatto il file di
prova. Reimportalo con **Una sola stanza** > `Digital Kiosk`.

Stavolta il riepilogo dice **30 righe valide** e avverte che *il foglio non
dichiara stanze: tutte le righe finiranno in Digital Kiosk*.

## 4. Le segnalazioni sui file imperfetti

Importa `Inventario_di_prova_con_difetti.xlsx` con **Tutto l'inventario** +
**Unisci**. Il riepilogo deve mostrare, tutte insieme:

- **3 righe valide trovate**
- **1 riga ignorata**: manca l'identificativo
- **3 righe hanno preso la stanza dai separatori**
- **1 iPhone ignorato**: si inseriscono solo a mano
- il riquadro rosso con le **colonne non riconosciute**: `Costo`, `Fornitore`,
  `Centro di costo`
- l'avviso che **1 riga non ha il modello**

Confermando entrano 3 dispositivi, di cui uno senza modello da completare a
mano. L'iPhone **non** entra.

## 5. Importare dentro una stanza

Il pulsante **Importa i dati di questa stanza** sta accanto al nome della stanza
quando la apri. Deve dare lo **stesso risultato** del punto 3: e' la stessa
regola, raggiunta da due strade diverse.

1. entra in **Digital Kiosk**
2. premi *Importa i dati di questa stanza*, scegli **Unisci**
3. seleziona `Inventario_di_prova.xlsx`

**Cosa deve risultare.** Il riepilogo dice **10 righe valide** e **20 righe di
altre stanze scartate**.

### La riga della stanza e' obbligatoria

Prendi `Inventario_di_prova.xlsx`, **cancella la riga `DIGITAL KIOSK`** e salva
con un altro nome in una cartella tua. Riprova a importarlo dentro Digital Kiosk: deve comparire un
avviso che dice che nel foglio non c'e' nessuna riga per quella stanza, che
**non e' stato importato niente**, e che spiega di aggiungere una riga vuota con
scritto `DIGITAL KIOSK` nella prima cella. Controlla che l'inventario sia
rimasto identico.

Prova anche la forma breve: al posto di `DIGITAL KIOSK` scrivi solo `KIOSK`.
Deve funzionare lo stesso.

## 6. Esportazione

Con l'inventario carico dal punto 1, prova le tre forme da *Esporta xls...*.

**Un unico elenco.** Un foglio solo con tutti e 30 i dispositivi.

**Un foglio per ogni stanza.** Un file con tre fogli, 10 dispositivi ciascuno.
In ogni foglio, **A1** deve contenere il nome della stanza, sotto la data di
esportazione, e la colonna *Stanza* deve esserci e riportare sempre la stessa
stanza.

**Un file separato per ogni stanza.** Scegli una cartella: devono uscire tre
file `Inventario_<Nome_stanza>_<data>.xlsx`.

Poi la controprova: **reimporta** il file con i tre fogli, con *Tutto
l'inventario* + *Unisci*. Deve leggerli tutti e riconoscere 30 dispositivi.

### L'invio per e-mail

Ogni esportazione finisce con la stessa domanda. Scegli **Invia per e-mail con
Outlook**: deve aprirsi un messaggio nuovo con il file gia' allegato, da
compilare e mandare tu. Chiudi il messaggio senza inviarlo, e' solo una prova.

Ripeti con l'esportazione **in un file per ogni stanza**: stavolta l'allegato
deve essere un unico archivio `.zip` con dentro i tre file.

Se sul computer non c'e' Outlook, deve comparire un avviso che lo dice e che
ricorda che il file e' stato creato comunque: verifica che ci sia davvero.

## 7. Tornare indietro da un errore

E' la prova che serve davvero in sede: sbagliare un'importazione e rimediare.

1. importa `Inventario_di_prova.xlsx` due volte con **Tutto l'inventario** +
   **Sostituisci**, cosi' esistono almeno due copie
2. adesso combina il pasticcio: importa `Inventario_di_prova.xlsx` scegliendo
   **Una sola stanza** > `Site Services BAU` > **Unisci**. Tutti e 30 i
   dispositivi finiscono in BAU, che ne aveva 10
3. premi **Ripristina** nella barra in alto

**Cosa deve risultare.** L'avviso dice di quando e' l'ultima copia e quanti
dispositivi conteneva rispetto a ora. Confermando, le schede stanza tornano
10, 10 e 10, e il messaggio finale dice dove ha salvato lo stato sbagliato -
perche' anche quello resta recuperabile.

Poi prova **Impostazioni** > **Ripristina da una copia...**: l'elenco mostra
tutte le copie dalla piu' recente, con data, ora e numero di dispositivi.
Sceglierne una e confermare deve dare lo stesso risultato.

## 8. La lingua

Cambia lingua dalla tendina **Lingua** in alto a destra nell'intestazione, e
controlla che diventi inglese **tutto**: pulsanti, intestazioni delle colonne,
schede stanza, barra di stato, e i messaggi che compaiono premendo *Elimina* o
*Sposta in stanza* senza aver spuntato niente. Poi torna in italiano.

Con l'interfaccia in italiano, prova *Esporta xls...* spuntando **Esporta i file
in inglese**: nel file intestazioni e stati devono essere in inglese, mentre i
nomi delle stanze restano come li hai scritti. Reimporta quel file: deve entrare
senza errori e gli stati devono tornare in italiano.

---

# Caricare l'inventario definitivo

Finito il collaudo, l'inventario e' pieno di dati finti. Prima di caricare
quelli veri va svuotato, altrimenti i due si mescolano: i dispositivi di prova
resterebbero dentro insieme ai tuoi.

Ci sono due modi. Fanno la stessa cosa; cambia solo quando avviene lo
svuotamento.

## Modo A - reset, poi importazione

Da usare se vuoi verificare che l'inventario sia vuoto prima di caricare.

1. premi **Reset inventario**, in alto a destra
2. l'avviso dice quanti dispositivi verranno eliminati, **per tutti gli utenti**
3. scrivi per esteso `ELIMINA TUTTO` nella casella, e conferma
4. il programma salva da solo una copia del file dati nella cartella `Backup`,
   dentro quella del programma, e poi svuota
5. controlla che le schede stanza siano tutte a zero
6. ora *Importa xls...* > **Tutto l'inventario** > **Unisci**, e scegli il tuo
   file

## Modo B - importazione che sostituisce tutto

Un passaggio solo, il risultato e' identico.

1. *Importa xls...*
2. **Tutto l'inventario** + **Sostituisci**
3. scegli il tuo file
4. il riepilogo dice quante righe ha letto e quanti dispositivi verranno
   eliminati prima di caricare
5. scrivi per esteso `ELIMINA TUTTO` e conferma

Anche qui la copia finisce in `Backup` prima che venga toccato qualsiasi cosa.
Se la copia non riesce, l'operazione si annulla e non viene cambiato niente.

**Se sbagli**, la copia e' li': aprila con il programma o con Excel, e' un
inventario completo. La cartella `Backup` non viene mai svuotata da sola, quindi
ogni tanto vale la pena ripulirla.

## Cosa sopravvive in ogni caso

**Gli iPhone non vengono mai eliminati**, ne' dal reset ne' da una sostituzione,
in nessuno stato. Non arrivano da un'importazione - si inseriscono solo a mano -
quindi cancellarli qui vorrebbe dire perderli per sempre. Il programma dice
quanti ne ha mantenuti.

Se durante il collaudo hai inserito iPhone finti, eliminali a mano prima di
partire: registra la spedizione con *Conferma spedizione*, e potrai eliminarli
tre mesi dopo. In alternativa, fai il collaudo senza inserire iPhone.

---

## Se qualcosa non torna

Segnala il problema con: a quale punto eri, cosa ti aspettavi, cosa e'
successo, e se puoi la schermata del riepilogo. Su macOS il diario di avvio e'
in `avvio.log` nella cartella del programma.
