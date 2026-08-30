# Site Services : Inventario Iphone, Laptop e Tablet

Applicazione desktop per Windows che gestisce l'inventario dei dispositivi
fisicamente in nostro possesso: iPhone, laptop e tablet, divisi per stanza, con
gestione dei prestiti, importazione, esportazione e stampa in formato Excel.

> **Versione alfa.** Funziona ed e' collaudata, ma e' al primo giro di prova sul
> campo: aspettati aggiustamenti. Segnala qualsiasi cosa non torni aprendo una
> issue.

## Scarica

**[Scarica l'ultima versione](../../releases/latest)** dalla pagina Releases.

Nel pacchetto trovi `Inventario.exe`: un file solo, che contiene Python e tutte
le librerie. **Sui PC non va installato niente.** Lo copi nella cartella di rete
condivisa, ci fai doppio clic, e chi ha accesso a quella cartella ha accesso
all'inventario.

```
\\server\Condivisa\Inventario\
    Inventario.exe                  il programma
    Inventario.xlsx                 i dati, apribili anche con Excel
    inventario_impostazioni.json    stanze, tipi, prestiti, stati
```

## In due parole

- **Nessun server, nessun database.** I dati stanno in un unico `.xlsx` sulla
  cartella di rete: e' insieme l'archivio e l'inventario che apri in Excel.
- **I permessi sono quelli della cartella.** Nessun elenco di utenti da gestire.
- **Piu' persone insieme.** Ogni salvataggio passa da un lock esclusivo e
  riscrive il file in modo atomico; l'elenco si aggiorna da solo.
- **Prestiti**, **stati**, **stanze** e **note** modificabili al volo.

---

## Compilare l'eseguibile da soli

Non serve, se scarichi la release: l'eseguibile viene costruito
automaticamente da [GitHub Actions](.github/workflows/build-windows.yml) su una
macchina Windows a ogni versione pubblicata.

Se preferisci compilarlo tu, su un qualsiasi PC Windows con Python installato
doppio clic su **`Compila EXE per Windows.bat`**. Scarica gli strumenti
necessari, compila e lascia tutto pronto nella cartella `Distribuzione`:

```
Distribuzione\
    Inventario.exe                      il programma, ~15 MB, si basta da solo
    Crea collegamento sul desktop.bat   utility per il collegamento
    Come funziona.txt                   questo documento
```

Copia il contenuto di `Distribuzione` nella cartella di rete condivisa. Fine.

## Come si lancia

Doppio clic su **`Inventario.exe`** nella cartella di rete.

Al primo avvio, se non trova ancora l'inventario, propone di crearlo li' accanto
con il nome `Inventario.xlsx`; da quel momento in poi tutti lo aprono senza che
venga chiesto piu' nulla. La cartella di rete diventa cosi':

```
\\server\Condivisa\Inventario\
    Inventario.exe                  il programma
    Inventario.xlsx                 i dati, apribili anche con Excel
    inventario_impostazioni.json    stanze, tipi, stanze con prestito
```

Chi ha accesso alla cartella ha accesso all'inventario: i permessi sono quelli
della cartella, non c'e' nessun altro elenco di utenti da gestire.

> Una precisazione: qualunque programma viene eseguito dal PC che lo apre, anche
> se il file risiede su una share. Quello che non serve, e che questa
> impostazione elimina, e' **installare** qualcosa sui PC.

## Collegamento sul desktop

Il modo piu' rapido, senza eseguire nulla: apri la cartella di rete, tasto
destro su `Inventario.exe` > *Mostra altre opzioni* > *Invia a* > *Desktop
(crea collegamento)*.

In alternativa, per prepararlo una volta e distribuirlo a tutti: doppio clic su
**`Crea collegamento sul desktop.bat`** nella cartella di rete. Crea il
collegamento sul desktop e ne lascia una copia nella cartella stessa: gli altri
utenti possono semplicemente trascinarla sul proprio desktop, oppure la si
distribuisce via GPO.

Il collegamento punta al percorso di rete, quindi resta valido anche quando il
programma viene aggiornato: basta sostituire `Inventario.exe` sulla share.

## Aggiornare il programma

Ricompila con `Compila EXE per Windows.bat` e sostituisci `Inventario.exe` nella
cartella di rete, con nessuno che lo sta usando in quel momento. I dati non si
toccano: stanno in `Inventario.xlsx`, che e' un file separato.

## Avvio dai sorgenti (per sviluppo)

Serve Python 3.8+ e `openpyxl` (`Installa_dipendenze.bat` su Windows).

- Windows: `Avvia Inventario.bat`
- macOS / Linux: `Avvia Inventario.command`, oppure `python3 Inventario.py`
- macOS, con l'icona sulla scrivania: `Inventario.app`

**Attenzione su macOS**: il Python di sistema (Command Line Tools) usa **Tk 8.5**,
la versione deprecata di Apple, che su macOS 14 e successivi fa terminare le
finestre subito dopo l'apertura, senza alcun messaggio. Serve un Python con
**Tk 8.6 o superiore**: quello di
[python.org](https://www.python.org/downloads/macos/) porta Tk 9.0.
Dopo averlo installato, ricostruisci l'ambiente con quell'interprete:

```bash
cd ~/Desktop/Inventario
rm -rf .venv
/Library/Frameworks/Python.framework/Versions/3.14/bin/python3 -m venv .venv
.venv/bin/pip install openpyxl
```

Un `.venv` creato in precedenza con il Python di sistema resta legato a Tk 8.5:
va rifatto, non basta installare il nuovo Python.

`Inventario.app` sceglie da solo il primo interprete con Tk utilizzabile che trova, e se
non ne trova nessuno lo dice con un avviso invece di chiudersi in silenzio;
il diario di avvio finisce in `avvio.log`.

Il problema riguarda **solo** l'esecuzione dai sorgenti su Mac: l'eseguibile
Windows costruito con PyInstaller si porta dietro Tcl/Tk 8.6 e non ne risente.

Il percorso del file dati si determina in quest'ordine: la variabile d'ambiente
`INVENTARIO_FILE`, poi un eventuale `inventario_percorso.json` accanto al
programma o nel profilo utente, poi `Inventario.xlsx` nella cartella del
programma, e solo in ultimo lo si chiede all'utente.

## Cosa fa

- **Home** con una scheda per stanza: nome, numero di dispositivi e ripartizione
  per tipo. Cliccando una scheda si apre l'inventario di quella stanza; sotto le
  schede la home mostra comunque l'inventario completo.
- In coda alle stanze c'e' la scheda **Iphone**, della stessa forma ma non e' una
  stanza: e' una scorciatoia che raccoglie automaticamente tutti i telefoni, per
  arrivarci con un clic. Gli iPhone restano registrati nella loro stanza e
  compaiono normalmente anche li'.
- Elenco dei dispositivi con **Asset Tag**, **Tipo**, **Modello**, **Numero di
  serie**, **IMEI**, **Restituito da**, **Stanza**, **Stato**, **In prestito a**,
  **Prestato il** e **Note**, piu' data e autore dell'ultima modifica.
  L'asset tag e' la chiave univoca dell'inventario.
- **Righe colorate**, sempre, sia nell'inventario completo sia dentro le stanze:
  verde chiaro gli iPhone, arancione chiaro i tablet Dell, **viola i dispositivi
  gia' rispediti**, rosso quelli in prestito. In caso di sovrapposizione vince
  l'informazione piu' urgente: prestito, poi spedizione, poi tipo. I tablet Dell si riconoscono dal modello, che
  deve contenere la parola *Dell* (come *Dell Latitude 7320 Detachable*).
- **Selezione con la casella**: la prima colonna dell'elenco e' una casella di
  spunta. Si lavora **su un dispositivo alla volta**: spuntandone un altro il
  primo si deseleziona, e cliccando di nuovo sulla casella si toglie la spunta.
- **Prestiti**: la colonna *Prestito*, con il pulsante vero su ogni riga,
  compare **solo quando si apre una stanza che gestisce i prestiti** (il Digital
  Kiosk). Nell'inventario completo, nelle altre stanze e nel contenitore iPhone
  la prima colonna e' soltanto la casella di selezione, senza intestazione.
  Vedi la sezione dedicata piu' sotto.
- **Nuovo** apre la scheda di inserimento, con il **tipo in cima**: e' lui a
  decidere quali campi vengono chiesti. Se manca un dato obbligatorio il
  programma elenca quali e non inserisce nulla finche' non sono compilati.

| Tipo | Campi obbligatori |
| --- | --- |
| Laptop, Tablet, ... | Asset Tag, Modello, Numero di serie, Stanza |
| iPhone | IMEI, Modello, Restituito da, Stanza |

  Scegliendo **iPhone** il modulo si ridisegna: al posto di *Asset Tag* e
  *Numero di serie* compaiono **IMEI** e **Restituito da** (il nome di chi lo ha
  consegnato). Cambiando tipo i valori gia' scritti e le note restano.

  La **stanza di un iPhone e' bloccata**: nella scheda il campo appare gia'
  compilato e non modificabile. La stanza si sceglie una volta per tutte da
  *Impostazioni* > *Stanza degli iPhone* (di serie *Site Services BAU*).

  Per un iPhone l'**IMEI fa da identificativo**: viene registrato nella sua
  colonna e usato anche come asset tag, cosi' il telefono resta univoco e
  ritrovabile come ogni altro dispositivo. Il tipo e' riconosciuto senza badare
  a maiuscole e minuscole: *iPhone*, *Iphone* e *IPHONE* sono la stessa cosa.
- **Note modificabili al volo**: doppio clic sulla cella *Note* per correggerla
  direttamente nell'elenco (`Invio` salva, `Esc` annulla). Doppio clic su una
  qualsiasi altra colonna apre la scheda completa.
- Aggiunta, modifica ed eliminazione; spostamento di un dispositivo da una
  stanza all'altra. Gli iPhone fanno eccezione: non si spostano, e se una
  selezione ne contiene qualcuno gli altri vengono spostati e il programma dice
  quanti telefoni ha lasciato dov'erano. Vale anche in importazione: un iPhone
  con una stanza diversa viene ricondotto alla sua.
- Ricerca libera (asset tag, modello, numero di serie, IMEI, note, stato, nome di
  chi ha restituito il dispositivo e di chi lo ha in prestito) e filtri per stanza
  e per tipo; ordinamento cliccando sulle intestazioni di colonna.
- **Importazione** da file Excel esistenti in modalita' *unisci* o *sostituisci*;
  le intestazioni piu' comuni sono riconosciute da sole (per il seriale: *Numero
  di serie*, *Seriale*, *Serial Number*, *S/N*, *Matricola*, *Service Tag*).
- **Esportazione** in `.xlsx` con **tutte** le colonne, in un unico elenco o con
  un foglio per stanza. Quello che esporti si puo' reimportare senza perdere nulla.
- **Stampa** su A4 orizzontale, con intestazioni ripetute a ogni pagina, numero di
  pagina e data; invio diretto alla stampante predefinita di Windows.
- Nomi delle stanze, tipi di dispositivo, stanze con prestito e stanza degli
  iPhone modificabili da *Impostazioni*, validi per tutti gli utenti.

Le stanze predefinite sono **Site Services BAU**, **Digital Kiosk** e
**Magazzino Disaster Recovery**; i prestiti sono attivi sul Digital Kiosk.

## Prestiti

Aprendo una stanza elencata fra le *stanze con prestito* (di serie il Digital
Kiosk) l'elenco guadagna la colonna *Prestito*, con su ogni riga un pulsante che
cambia in base allo stato del dispositivo. Fuori da quella stanza la colonna non
esiste:

| Stato | Pulsante | Cosa succede |
| --- | --- | --- |
| Disponibile | **Presta** | chiede il nome della persona e registra nome, data e ora accanto al dispositivo |
| Non disponibile | **Registra rientro** | chiude il prestito e riporta il dispositivo fra i disponibili |

Finche' il dispositivo e' fuori, la sua riga e' **evidenziata in rosso** e lo
stato risulta **Non disponibile**; l'evidenziazione si ritrova anche nei file
esportati e stampati. Il nome di chi ha in prestito un dispositivo e' incluso
nella ricerca, quindi basta digitare il cognome per trovare cosa ha in mano.

Tre stati sono automatici e vincono su tutto: **Non disponibile** mentre c'e' un
prestito in corso, **Da Rispedire** per gli iPhone ancora in casa e **Spedito**
per quelli gia' rispediti. Registrato il rientro,
il dispositivo torna *Disponibile*. Vedi *Stati* qui sotto per gli altri.

## Spedizione degli iPhone

Nel contenitore **Iphone** ogni telefono non ancora rispedito ha sul proprio rigo
il pulsante **SPEDITO**, da premere quando parte davvero verso il servizio
telefonia. La colonna compare solo li': nelle stanze non c'e'.

Alla conferma il programma registra **giorno e ora** nella colonna *Spedito il*,
porta lo stato a *Spedito*, colora la riga di **viola** e mostra la frase che
resta valida per il dispositivo:

> Il dispositivo e' stato rispedito al servizio telefonia il *gg/mm/aaaa hh:mm*.
> Resta in inventario per consultazione fino al *gg/mm/aaaa*, data dalla quale
> potra' essere eliminato.

### Quando un iPhone si puo' eliminare

Un iPhone si elimina soltanto dopo essere stato rispedito, e comunque non prima
di tre mesi da quella data. Sono due blocchi distinti, con due avvisi diversi:

| Situazione | Cosa succede provando a eliminarlo |
| --- | --- |
| Non ancora rispedito | Avviso: va prima registrata la spedizione con il pulsante *SPEDITO*. Non esiste ancora una data di sblocco. |
| Rispedito da meno di tre mesi | Avviso con la data esatta a partire dalla quale sara' possibile. |
| Rispedito da piu' di tre mesi | Si elimina normalmente. |

Nel frattempo il telefono resta visibile in inventario per eventuali
consultazioni. Passati i tre mesi il blocco cade da solo.

La regola vale **solo per gli iPhone**: laptop e tablet si eliminano sempre. Il
vincolo e' nell'archivio dati, non nell'interfaccia, quindi non c'e' schermata da
cui aggirarlo.

## Stati

Ogni dispositivo ha uno stato scegliibile fra:

| Stato | Quando |
| --- | --- |
| Disponibile | valore di partenza |
| In attesa ritiro | pronto, si aspetta che venga ritirato |
| Guasto in attesa tecnico | fuori uso, in attesa dell'intervento |
| Da rebuildare | da reinstallare prima di rimetterlo in giro |
| Controllare | da verificare |

Si sceglie all'inserimento, nella tendina della scheda, e si cambia in qualsiasi
momento **senza aprire nessuna finestra**: doppio clic sulla cella *Stato*
nell'elenco e si sceglie dalla tendina che compare al suo posto. Funziona da
ogni schermata: inventario completo, singola stanza, contenitore iPhone.

Restano fuori dalla scelta manuale gli stati automatici: un dispositivo **in
prestito** e' *Non disponibile* e lo stato torna modificabile solo dopo il
rientro; un **iPhone** e' *Da Rispedire* finche' non viene spedito, poi *Spedito*. In entrambi i casi la tendina
appare gia' compilata e bloccata, con la spiegazione accanto, e provando a
cambiarla dall'elenco il programma lo dice nella barra di stato in fondo, senza
aprire finestre.

Lo stato non si scrive mai a mano nel file: se un'importazione porta uno stato
non previsto, viene riportato a *Disponibile*. L'elenco degli stati e' in
`inventario_impostazioni.json`, alla voce `states`. Chi registra prestito e rientro
resta tracciato nelle colonne *Ultima modifica* e *Modificato da*.

Le righe delle stanze senza prestito non hanno alcun pulsante.

## Provare subito con dati di esempio

La cartella `Esempio/` contiene un `Inventario.xlsx` con tredici dispositivi di
prova divisi fra le tre stanze, di cui due gia' in prestito nel Digital Kiosk.
Per aprirlo senza toccare la configurazione:

```bash
INVENTARIO_FILE=Esempio/Inventario.xlsx python Inventario.py
```

Su Windows, dal Prompt dei comandi nella cartella dell'applicazione:

```bash
set INVENTARIO_FILE=Esempio\Inventario.xlsx && python Inventario.py
```

## Cosa c'e' nella cartella di rete

| File | Contenuto |
| --- | --- |
| `Inventario.exe` | il programma; da solo, senza prerequisiti |
| `Inventario.xlsx` | i dati; e' gia' l'inventario, apribile in Excel |
| `inventario_impostazioni.json` | stanze, tipi, stanze con prestito, stanza degli iPhone, stati |

## Test

```bash
.venv/bin/python tests/run_all.py
```

Cinque suite che coprono archivio dati e accessi concorrenti, risoluzione del
percorso, schermate e colori, scheda di inserimento e spedizioni. Girano senza
bisogno di una finestra a schermo e non toccano i dati reali: ognuna si
costruisce il proprio inventario in una cartella temporanea.
| `.Inventario.xlsx.lock` | presente solo per una frazione di secondo durante un salvataggio |

## Accessi contemporanei

Piu' persone possono tenere aperta l'applicazione insieme.

- Ogni salvataggio avviene dentro un **lock esclusivo** e riscrive il file in
  modo atomico (scrittura su file temporaneo + sostituzione), quindi
  l'inventario non puo' restare a meta'.
- Dentro il lock i dati vengono **riletti da disco** e la modifica viene
  riapplicata sui dati aggiornati: due utenti che modificano schede diverse non
  si sovrascrivono a vicenda.
- L'elenco a video si aggiorna da solo ogni 15 secondi quando qualcun altro
  salva; `F5` forza il ricaricamento.
- Un lock rimasto appeso (PC spento a meta' salvataggio) viene ignorato e
  rimosso dopo 2 minuti.

Perche' funzioni, gli utenti devono avere permessi di **lettura e scrittura**
sulla cartella (non solo sul file): l'applicazione vi crea file temporanei.

Se qualcuno tiene `Inventario.xlsx` aperto in Excel, i salvataggi possono
fallire perche' Windows blocca il file: chiudere Excel e riprovare. Per
consultare i dati in Excel senza rischi, usa *Esporta xls...*.

## Scorciatoie

| Tasto | Azione |
| --- | --- |
| `Ctrl+N` | nuovo dispositivo |
| `Ctrl+F` | vai alla casella di ricerca |
| `Ctrl+P` | stampa la vista corrente |
| `Esc` | torna alla home |
| `F5` | ricarica dalla rete |
| `Canc` | elimina il dispositivo spuntato |
| clic sulla casella | spunta o toglie la spunta alla riga |
| doppio clic sulle note | modifica la nota nell'elenco |
| doppio clic sullo stato | tendina per cambiare stato nell'elenco |
| doppio clic altrove | apre la scheda del dispositivo |

## Creare l'inventario da un file Excel esistente

Non serve inserire i dispositivi a mano: *Importa xls...* accetta un qualsiasi
file Excel la cui **prima riga contenga le intestazioni**. L'unica colonna
obbligatoria e' l'asset tag - o, per gli iPhone, l'IMEI, che ne fa le veci; le
altre sono riconosciute per nome, in italiano o in inglese, per esempio:

| Colonna | Intestazioni riconosciute |
| --- | --- |
| Asset Tag | Asset Tag, Asset, Tag, Etichetta, Inventario |
| Tipo | Tipo, Tipologia, Categoria, Type |
| Modello | Modello, Model, Descrizione, Dispositivo |
| Numero di serie | Numero di serie, Seriale, Serial Number, S/N, Matricola, Service Tag |
| IMEI | IMEI, IMEI/MEID, MEID, Codice IMEI |
| Restituito da | Restituito da, Proprietario, Consegnato da, Riconsegnato da, Owner |
| Stanza | Stanza, Room, Locale, Ubicazione, Posizione |
| In prestito a | In prestito a, Prestato a, Assegnato a, Borrower |
| Prestato il | Prestato il, Data prestito, Loan date |
| Note | Note, Nota, Commenti |

Le righe senza asset tag ne' IMEI vengono contate e scartate, le colonne che non servono
sono ignorate. Con la modalita' *Sostituisci* l'intero inventario viene creato
dal file in un colpo solo; con *Unisci* i dispositivi gia' presenti vengono
aggiornati e gli altri aggiunti.

## Struttura del codice

```
Compila EXE per Windows.bat       crea l'eseguibile autosufficiente
Crea collegamento sul desktop.bat collegamento alla cartella di rete
Avvia Inventario.bat              avvio dai sorgenti su Windows
Avvia Inventario.command          avvio dai sorgenti su macOS / Linux
Inventario.py                     avvio dell'applicazione
inventario/config.py     percorso del file dati e impostazioni condivise
inventario/store.py      lettura/scrittura del file .xlsx, lock, operazioni
inventario/excel_io.py   esportazione, impaginazione di stampa, invio a stampante
inventario/theme.py      palette, font e stili dell'interfaccia
inventario/ui.py         interfaccia grafica (home, vista stanza, dialoghi)
```

Per aggiungere un campo basta inserirlo in `FIELDS`, `HEADERS`, `HEADER_ALIASES`
e nelle larghezze colonna in `inventario/store.py`, in `PRINT_FIELDS` in
`inventario/excel_io.py`, e nella scheda `ItemDialog` in `inventario/ui.py`.

Un file `Inventario.xlsx` creato con una versione precedente resta leggibile: le
colonne mancanti risultano vuote e vengono aggiunte al primo salvataggio.
