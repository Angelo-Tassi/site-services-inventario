# Site Services : Inventario Iphone, Laptop e Tablet

### [Apri la pagina del progetto](https://angelo-tassi.github.io/site-services-inventario/) &nbsp;·&nbsp; [Scarica il programma](https://github.com/Angelo-Tassi/site-services-inventario/releases/latest/download/Inventario-windows.zip) &nbsp;·&nbsp; [Manuale](https://angelo-tassi.github.io/site-services-inventario/manuale.html) &nbsp;·&nbsp; [English](README.en.md)

> **Pagina del progetto:** <https://angelo-tassi.github.io/site-services-inventario/>
> Da li' si scarica il programma, il modello Excel e si leggono le guide, in
> italiano e in inglese.

Applicazione desktop per Windows che gestisce l'inventario dei dispositivi
fisicamente in nostro possesso: iPhone, laptop e tablet, divisi per stanza, con
gestione dei prestiti, importazione, esportazione e stampa in formato Excel.

> **Versione alfa.** Funziona ed e' collaudata, ma e' al primo giro di prova sul
> campo: aspettati aggiustamenti. Segnala qualsiasi cosa non torni aprendo una
> issue.

## Prima di tutto: come si usa

**[Tenerlo aggiornato, non ricostruirlo](https://angelo-tassi.github.io/site-services-inventario/tenerlo-aggiornato.html)**
- il metodo di lavoro, prima ancora delle funzioni.

Questo programma non serve a rifare l'inventario prima della review: serve a non
doverlo rifare mai piu'. Ogni movimento fisico - un dispositivo che entra, esce,
va in prestito, si guasta, viene rispedito - va registrato **quando avviene, da
chi lo compie**, in una decina di secondi. Alla review non si prepara niente: si
esporta e si manda.

Se invece si accumulano i movimenti per tre mesi e si ricostruisce tutto con
un'importazione, lo sforzo di tenere l'inventario e' sprecato: il file parte
gia' vecchio, e nessuno sa piu' dove sia finito un dispositivo.

## Scarica

**[Scarica `Inventario-windows.zip`](../../releases/latest/download/Inventario-windows.zip)**
dalla pagina Releases. Dentro ci sono il Python ufficiale di python.org,
firmato dalla Python Software Foundation, e il programma in chiaro come file
`.py`: nessun eseguibile costruito da noi, quindi niente di non firmato da far
passare alla sicurezza aziendale.

Serve anche il **[modello Excel da compilare](docs/Modello_inventario.xlsx)**,
se hai laptop e tablet gia' censiti altrove da caricare in blocco.

## In due parole

- **Nessun server, nessun database.** I dati stanno in un unico `.xlsx` sulla
  cartella di rete: e' insieme l'archivio e l'inventario che apri in Excel.
- **I permessi sono quelli della cartella.** Nessun elenco di utenti da gestire.
- **Piu' persone insieme.** Ogni salvataggio passa da un lock esclusivo e
  riscrive il file in modo atomico; l'elenco si aggiorna da solo.
- **Prestiti**, **stati**, **tipi**, **stanze**, **note** e **descrizioni**
  modificabili al volo, con un doppio clic sulla cella.

---

---

## Come si installa

**Il programma sta sulle postazioni. L'inventario sta sulla share. Uno solo,
per tutti.**

E' la divisione che conta, e vale la pena spiegare perche'.

L'inventario deve essere **uno**: se ogni tecnico ne tenesse una copia, dopo
mezza giornata sarebbero tutte diverse e nessuna vera. Per questo il file
`.xlsx` sta sulla cartella di rete, e ogni modifica - un'aggiunta, un prestito,
un rientro - viene scritta subito li', dove la vedono tutti.

Il programma, invece, non ha niente da condividere: e' lo stesso identico
codice su ogni PC. Tenerlo sulla share non porta nessun vantaggio e porta due
guai concreti: un eseguibile avviato dalla rete e' lo schema che i sistemi di
sicurezza aziendali guardano con piu' sospetto, e i suoi file restano
**bloccati** da Windows finche' qualcuno lo tiene aperto - anche da un altro
computer - quindi non si lasciano aggiornare.

```
Su ogni postazione                      Sulla cartella di rete
------------------------------          ----------------------------------
C:\Inventario\                          \\server\Condivisa\Inventario\
    Inventario.py                            Produzione\
    inventario\                                  Inventario.xlsx      <- l'inventario
    python\                                      inventario_impostazioni.json
    inventario_percorso.json                 Backup\                  <- le copie
    Collaudo\
```

### Passo 1 - prepara una copia, una volta sola

Su un PC qualsiasi:

1. scarica lo zip ed **estrailo in una cartella locale**, per esempio
   `C:\Inventario`. Non sulla share;
2. doppio clic su **`Collega inventario condiviso.bat`**;
3. **si apre una finestra**: sfoglia fino alla cartella di rete e scegli quella.
   Niente percorsi da copiare.

Se la finestra non si aprisse, il percorso si puo' sempre incollare a mano -
`\\server\Condivisa\Inventario`, o la lettera dell'unita' mappata,
`F:\Inventario`. Virgolette, barre al contrario e una barra finale di troppo
non danno fastidio.

Il programma ci pensa lui: se sulla share l'inventario non c'e' ancora, lo crea
vuoto in `Produzione\Inventario.xlsx`; se c'e' gia', non lo tocca. Poi scrive
`inventario_percorso.json` accanto a se stesso, ed e' li' che resta memorizzato
quale inventario aprire.

Serve il permesso di **Modifica** su quella cartella di rete. Se non ce l'hai,
il programma te lo dice chiaramente invece di fallire a meta'.

### Passo 2 - porta l'inventario che hai gia'

Se un inventario esiste gia' - il file su cui stavi lavorando - va copiato
sulla share **con il nome giusto e nel posto giusto**:

```
\\server\Condivisa\Inventario\Produzione\Inventario.xlsx
```

Il nome deve essere esattamente `Inventario.xlsx`, e deve stare dentro
`Produzione`. Chiudi il programma prima di sostituirlo, e conserva una copia
del file vecchio finche' non hai verificato che l'inventario si apra e i
dispositivi ci siano tutti.

Se invece il tuo inventario e' un foglio con altre colonne, o diviso per stanze
con le righe separatore, **non copiarlo li'**: quello si carica da dentro il
programma con *Importa xls...*, che e' un'altra cosa. Vedi
[Il modello di importazione](#il-modello-di-importazione).

### Passo 3 - distribuisci la cartella alle postazioni

La cartella `C:\Inventario` che hai preparato e' gia' configurata: la
configurazione viaggia con lei. Copiala tale e quale su ogni postazione - a
mano, con uno script, con un pacchetto software o via GPO - sempre nello stesso
percorso locale.

Su ogni postazione, poi, doppio clic su **`Crea collegamento sul desktop.bat`**:
mette l'icona sulla scrivania dell'utente e ne lascia una copia nella cartella.

Il collegamento punta ad **`Avvia Inventario.bat`**, non direttamente a
`pythonw.exe`: su alcune postazioni un collegamento a un eseguibile con
argomenti viene rifiutato dai criteri di sicurezza, mentre quello a un file
`.bat` funziona sempre.

**Se sul desktop non compare niente** - capita, dove la creazione automatica e'
bloccata - si fa a mano e non fallisce: tasto destro su
`Avvia Inventario.bat` > *Mostra altre opzioni* > *Invia a* >
*Desktop (crea collegamento)*. In alternativa si trascina sul desktop la copia
`Inventario dispositivi.lnk` rimasta nella cartella, che si puo' dare anche
agli altri utenti senza eseguire niente.

Da quel momento il tecnico fa doppio clic sull'icona e lavora sull'inventario di
tutti, senza sapere niente di percorsi di rete.

### Cambiare inventario dopo

Si puo' fare anche dal programma, senza toccare nessun file: *Impostazioni* >
**Collega inventario condiviso...**, si sfoglia fino alla cartella e si
conferma. Il programma dice prima cosa succedera' - se l'inventario e' gia' li'
non lo tocca, se non c'e' ne crea uno vuoto - e poi va riaperto.

E' la strada comoda quando ogni tecnico si installa il programma da solo: apre,
sceglie la cartella dell'inventario di produzione, e ha finito.

### Cosa serve sulla cartella di rete

| Chi | Su cosa | Permesso |
| --- | --- | --- |
| i tecnici | `\\server\Condivisa\Inventario\` e tutto quello che contiene | **Modifica** |

Basta questo. Nessun permesso di esecuzione, perche' dalla share non viene
eseguito niente; nessun elenco di utenti da gestire dentro il programma: chi ha
accesso alla cartella ha accesso all'inventario, chi non ce l'ha non lo apre.

### Se la share non risponde

Il programma **si ferma e lo dice**, indicando quale inventario si aspettava di
trovare. Non ne crea uno locale: lavorare su una copia che nessun altro vede
sarebbe il modo piu' silenzioso di perdere una giornata di lavoro.

## Come si lancia

Doppio clic sull'icona **Inventario dispositivi** sul desktop.

La finestra si apre sull'inventario condiviso. Il numero di versione e' scritto
nella barra del titolo: serve a sapere, davanti a un dubbio, quale copia si sta
usando.

## Aggiornare il programma

Adesso che il programma sta sulle postazioni, un aggiornamento e' una copia
dentro una cartella tua, dove i permessi ce li hai:

1. scarica il pacchetto nuovo;
2. su ogni postazione, **sostituisci la cartella locale** del programma;
3. **non toccare `inventario_percorso.json`**: e' la riga che dice dove sta
   l'inventario condiviso. Se lo sovrascrivi, ripassa
   `Collega inventario condiviso.bat`.

**I dati non si toccano mai**: stanno sulla share, e nessun aggiornamento del
programma li sfiora. Anche i collegamenti sul desktop restano validi, perche'
puntano allo stesso percorso locale.

## Gli avvisi di sicurezza di Windows

Il pacchetto non contiene nessun eseguibile costruito da noi, il che
toglie di mezzo la causa piu' comune di segnalazione. Resta una sola accortezza,
al momento del download:

**Sblocca lo zip prima di estrarlo.** Tasto destro sul file scaricato >
*Proprieta'* > in fondo alla scheda Generale, spunta **Annulla blocco** >
*Applica*.

Windows marca come "scaricato da Internet" tutto cio' che si estrae da un
archivio marchiato, e quel marchio segue i file. Sbloccando l'archivio prima di
aprirlo, il marchio non passa a niente.

**Se la casella "Annulla blocco" non c'e', non c'e' niente da sbloccare**: vuol
dire che Windows non ha messo il marchio - capita quando il download passa da un
proxy aziendale - e i file estratti sono gia' puliti.

Se in azienda l'esecuzione di programmi non firmati e' regolata da criteri di
sicurezza (AppLocker, Windows Defender Application Control), serve
un'autorizzazione dagli amministratori: un binario firmato dalla Python Software
Foundation, installato in locale, e' pero' il caso piu' facile da far passare.

## Un foglio da importare non e' un inventario

Sono due cose diverse e non vanno mai scambiate:

- **l'inventario** e' `Produzione\Inventario.xlsx` sulla share: il programma lo
  apre e ci scrive dentro;
- **un foglio da importare** e' un file Excel che contiene dispositivi da
  caricare, spesso diviso da righe con il nome della stanza. Si carica con
  *Importa xls...*, da dentro il programma.

Se un foglio da importare viene aperto *come* inventario, le righe separatore
diventano dispositivi e nessun dispositivo ha una stanza. Il programma se ne
accorge e avvisa, ma la regola resta: i fogli si importano, non si aprono.

## Quando qualcosa non torna: la diagnostica

Nella cartella del programma c'e' **`Diagnostica.bat`**. Doppio clic, e scrive
`Diagnostica.txt` li' accanto, aprendolo nel Blocco note. Non modifica niente.

Dentro c'e' quello che serve a capire un problema senza tirare a indovinare: la
versione in uso e da dove parte, quale inventario apre e se puo' scriverci, le
stanze che conosce, cosa legge davvero da un file Excel, la misura della
finestra e della tabella, e dove Windows mette il desktop dell'utente.

Il file contiene percorsi e nomi di stanza, niente di riservato: si manda a chi
assiste e risponde in un colpo solo a domande che altrimenti costano giorni.

## Come viene costruito il pacchetto

Non serve costruirlo: lo fa [GitHub Actions](.github/workflows/build-windows.yml)
su una macchina Windows a ogni versione pubblicata. Il pacchetto e' il Python
incorporabile ufficiale di python.org piu' il programma; la compilazione
controlla la firma di `pythonw.exe`, prova un avvio vero e rifiuta di
pubblicare se resta anche un solo file in sola lettura.

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

Il problema riguarda **solo** l'esecuzione dai sorgenti su Mac: il pacchetto
Windows si porta dietro il suo Python con Tcl/Tk 8.6 e non ne risente.

Il percorso del file dati si determina in quest'ordine: la variabile d'ambiente
`INVENTARIO_FILE`, poi `inventario_percorso.json` accanto al programma o nel
profilo utente - e' quello che scrive *Collega inventario condiviso* - e infine
`Produzione\Inventario.xlsx` accanto al programma.

Se il percorso e' configurato ma non si raggiunge, il programma si ferma e lo
dice: non crea mai un inventario locale al posto di quello condiviso.

## Cosa fa

- **Home** con una scheda per stanza: nome, numero di dispositivi e ripartizione
  per tipo. Cliccando una scheda si apre l'inventario di quella stanza; sotto le
  schede la home mostra comunque l'inventario completo.
- In coda alle stanze c'e' la scheda **Iphone**, della stessa forma ma non e' una
  stanza: e' una scorciatoia che raccoglie automaticamente tutti i telefoni, per
  arrivarci con un clic. Gli iPhone restano registrati nella loro stanza e
  compaiono normalmente anche li'.
- Elenco dei dispositivi con **Asset Tag**, **Tipo**, **Modello/Descrizione**, **Numero di
  serie**, **IMEI**, **Restituito da**, **Stanza**, **Stato**, **In prestito a**,
  **Prestato il** e **Note**, piu' data e autore dell'ultima modifica.
  L'asset tag e' la chiave univoca dell'inventario.
- **Righe colorate**, sempre, sia nell'inventario completo sia dentro le stanze:
  verde chiaro gli iPhone, azzurro i tablet Dell, viola i dispositivi gia'
  rispediti, rosso quelli in prestito. In caso di sovrapposizione vince
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
- **Aggiungi** chiede prima **che cosa** stai inserendo, con una tendina fra i
  tipi configurati, e solo dopo **come**: inserimento manuale o scansione con il
  lettore di codici (vedi piu' sotto). Il tipo scelto decide i campi richiesti e
  cosa si puo' leggere con il lettore. Se manca un dato obbligatorio il
  programma elenca quali e non inserisce nulla finche' non sono compilati.

| Tipo | Campo obbligatorio | Il resto |
| --- | --- | --- |
| Laptop, Tablet, ... | **Asset Tag** | modello, numero di serie, note: si completano poi |
| iPhone | **IMEI** | modello, restituito da, note: si completano poi |

  Obbligatorio e' **solo l'identificativo**: e' l'unica cosa senza la quale il
  dispositivo non esiste in inventario. Il modello e il numero di serie spesso
  non si hanno sottomano nel momento in cui si registra un arrivo, e pretenderli
  significa far rimandare l'inserimento - cioe' perdere la riga. La **stanza**
  non blocca il salvataggio: e' una tendina, e se non viene scelta parte dalla
  prima, cosi' nessun dispositivo resta senza.

  Un iPhone **non ha asset tag**: l'IMEI e' il suo unico identificativo. La
  colonna *Asset Tag* resta vuota per i telefoni nell'elenco, nella stampa e nel
  file dati, e non compare affatto nel loro contenitore.

  Un iPhone **non ha numero di serie e non si presta mai**, oltre a non avere
  asset tag. La regola e' nell'archivio dati: anche importando o modificando una
  scheda, quei campi vengono ripuliti, e il tentativo di prestare un iPhone viene
  rifiutato. Nel contenitore *Iphone* le colonne *Asset Tag*, *Numero di serie*,
  *In prestito a* e *Prestato il* non compaiono affatto: restano nelle altre
  schermate, dove servono a laptop e tablet.

  Scegliendo **iPhone** il modulo si ridisegna: al posto di *Asset Tag* e
  *Numero di serie* compaiono **IMEI** e **Restituito da** (il nome di chi lo ha
  consegnato). Cambiando tipo i valori gia' scritti e le note restano.

  La **stanza di un iPhone e' bloccata**: nella scheda il campo appare gia'
  compilato e non modificabile. La stanza si sceglie una volta per tutte da
  *Impostazioni* > *Stanza degli iPhone* (di serie *Site Services BAU*).

  Per un iPhone l'**IMEI fa da identificativo**. Internamente il programma lo usa
  anche come chiave della riga, cosi' il telefono resta univoco e le operazioni
  funzionano come per gli altri dispositivi, ma la cosa non si vede da nessuna
  parte: nel file dati la colonna *Asset Tag* di un iPhone e' vuota, e alla
  rilettura la chiave viene ricostruita dall'IMEI. Il tipo e' riconosciuto senza
  badare a maiuscole e minuscole: *iPhone*, *Iphone* e *IPHONE* sono la stessa
  cosa.
- **Modifica al volo**, senza aprire nessuna finestra: doppio clic sulla cella
  *Note* o *Modello/Descrizione* per scriverci (`Invio` salva, `Esc` annulla),
  su *Stato* per la tendina. Sul *Tipo* la tendina propone i tipi configurati ma
  si puo' anche scrivere o incollare un valore diverso, che compare poi anche fra
  i filtri. Doppio clic su una qualsiasi altra colonna
  apre la scheda completa.
- **Elimina +** toglie molti dispositivi in una volta: si incollano i codici da
  Excel, si legge che cosa sparisce e da quale stanza, e si conferma scrivendo
  `ACCETTO`.
- Aggiunta, modifica ed eliminazione; spostamento di un dispositivo da una
  stanza all'altra; **esportazione della singola stanza** e **reset completo**
  dell'inventario con copia di sicurezza. Gli iPhone fanno eccezione: non si spostano, e se una
  selezione ne contiene qualcuno gli altri vengono spostati e il programma dice
  quanti telefoni ha lasciato dov'erano. Vale anche in importazione: un iPhone
  con una stanza diversa viene ricondotto alla sua.
- Ricerca libera (asset tag, modello, numero di serie, IMEI, note, stato, nome di
  chi ha restituito il dispositivo e di chi lo ha in prestito) e filtri per stanza
  e per tipo.
- **L'elenco parte dal piu' recente**: l'ultimo dispositivo inserito o modificato
  e' sempre in cima. Cliccando sulle intestazioni si ordina per qualsiasi colonna;
  le colonne con una data si ordinano per data, non alfabeticamente, e partono
  dalla piu' recente.
- **Modello di importazione** scaricabile dal programma, gia' pronto da compilare.
- **Importazione** da file Excel in modalita' *unisci* o *sostituisci*, con la
  possibilita' di dividere per stanza un inventario unico usando le
  righe-separatore (vedi piu' sotto);
  le intestazioni piu' comuni sono riconosciute da sole (per il seriale: *Numero
  di serie*, *Seriale*, *Serial Number*, *S/N*, *Matricola*, *Service Tag*).
- **Esportazione** in `.xlsx` con **tutte** le colonne, in un unico elenco o con
  un foglio per stanza. Quello che esporti si puo' reimportare senza perdere nulla.
- **Stampa** su A4 orizzontale, con intestazioni ripetute a ogni pagina, numero di
  pagina e data; invio diretto alla stampante predefinita di Windows.
- Nomi delle stanze, tipi di dispositivo, stanze con prestito e stanza degli
  iPhone modificabili da *Impostazioni*, validi per tutti gli utenti.
- **Italiano e inglese**: la tendina *Lingua* traduce interfaccia, colonne e
  stati; i file si possono esportare in inglese anche lavorando in italiano.

Le stanze predefinite sono **Site Services BAU**, **Digital Kiosk** e
**Magazzino Disaster Recovery**; i prestiti sono attivi sul Digital Kiosk.

## Due cose che il programma fa da solo

**Un identificativo non si ripete.** Inserendo un dispositivo con un asset tag -
o un IMEI - gia' presente, il programma **non inserisce niente** e dice dove sta
quello che ce l'ha gia', con il modello e la stanza: si capisce subito se e' un
errore di battitura o un dispositivo gia' registrato.

**Ogni cinque dispositivi toccati, ricorda la copia locale.** Aggiunte,
modifiche ed eliminazioni si contano insieme - e contano i *record*, quindi
un'eliminazione in blocco di trenta li conta tutti e trenta. Al quinto il
programma chiede se vuoi salvare una copia sul tuo computer, e si puo' dire di
no: torna a chiedere dopo altri cinque. Salvando una copia il conto riparte.

Serve perche' le copie automatiche stanno sulla cartella di rete accanto ai
dati: coprono l'errore umano, non la cartella che sparisce.

## Eliminare piu' dispositivi in una volta

**`Elimina +`**, accanto a *Elimina*, serve a togliere trenta dispositivi senza
cercarli uno per uno - e senza rinunciare a sapere che cosa sparisce.

1. si **incollano i codici** presi da un foglio Excel, uno per riga. Va bene una
   colonna di asset tag o righe intere: viene letto il primo codice che
   corrisponde a un dispositivo in inventario;
2. si preme **Controlla**. Il programma non tocca niente e mostra:
   - i dispositivi che **verranno eliminati**, raggruppati per stanza, con il
     modello e la segnalazione di quelli **in prestito**;
   - quelli **saltati perche' non si possono eliminare** - un iPhone non ancora
     rispedito, uno in conservazione - con il motivo e la data;
   - quelli **saltati perche' non sono in inventario**, elencati;
3. per procedere si scrive **`ACCETTO`**. Prima di eliminare viene salvata una
   copia di sicurezza, e alla fine il programma dice quanti dispositivi
   restano.

I doppioni contano una volta sola, le righe vuote si ignorano, e gli iPhone
mantengono tutte le loro protezioni.

## Le colonne di ogni stanza

Ogni vista mostra **solo** le colonne che li' possono avere un valore. Una
colonna vuota per costruzione non porta informazione: toglie spazio a quello che
si deve leggere.

L'**inventario completo in home** e' una panoramica: dice che cos'e' un
dispositivo, dov'e' e come sta - asset tag, tipo, stanza, note, stato,
modello, seriale - in una riga che si legge senza scorrere di lato. Le domande piu'
precise si fanno dentro la stanza che le riguarda, e in panoramica ci pensa lo
stato a riassumerle: *In prestito*, *Spedito al servizio telefonia*.

| Dove sei | Che cosa sparisce |
| --- | --- |
| **Home** | IMEI, prestiti, restituito da, spedito il, ultima modifica |
| **Site Services BAU** | Stanza, In prestito a, Prestato il |
| **Digital Kiosk** | Stanza, IMEI, Restituito da, Spedito il |
| **Magazzino Disaster Recovery** | Stanza, prestiti e campi degli iPhone |
| **Contenitore Iphone** | Asset tag, seriale, prestiti, stanza e tipo |

Ogni colonna e' **larga quanto serve** a mostrare per intero quello che contiene:
niente testo tagliato, niente colonne da allargare a mano. Una colonna vuota resta
comunque larga almeno quanto il suo nome. Se l'elenco supera la finestra si scorre
con la barra in basso o con Maiusc + rotella.

Una **riga verticale colorata** divide una colonna dall'altra, e la stessa tinta
compare come barretta nell'intestazione: serve a non perdere la colonna mentre si
scorre un elenco largo. Il colore raggruppa per significato - blu gli
identificativi, verde che cos'e' il dispositivo, viola dov'e', rosso come sta,
ambra il prestito. Il testo resta nero su bianco: il colore sta nelle divisioni e
nelle intestazioni, non dentro le celle.

La colonna **Stanza** sparisce sempre dentro una stanza: sarebbe uguale su ogni
riga, e il nome e' gia' scritto sopra l'elenco. Lo stesso vale per il **tipo**
dentro il contenitore Iphone.

Le colonne dipendono da **come e' configurato l'inventario**, non dai
dispositivi presenti in quel momento: una stanza vuota mostra le stesse colonne
di quando sara' piena, e niente balla mentre si lavora. Attivando i prestiti in
un'altra stanza dalle impostazioni, le due colonne compaiono li'.

### La stessa regola nei file che escono

Un file esportato dice **che cosa abbiamo, dove sta e che cosa c'e' da sapere**.
Quattro colonne: asset tag, tipo, stanza, note - le note viaggiano con il
dispositivo, perche' sono quello che una riga ha di particolare.

Stato, modello, numero di serie, prestiti, IMEI, spedizioni e ultima modifica
restano fuori: servono a chi lavora davanti all'elenco, dentro la stanza che li
riguarda, non a chi riceve il file. La **stampa** invece li porta, perche' e'
fatta per chi lavora.

**Da un'esportazione non si ricostruisce un inventario**, perche' quei campi nel
file non ci sono. Per rimettere in piedi l'inventario per intero si usa *Salva
copia in locale...*, che copia il file vero e si porta dietro tutto.

## Copia e incolla

Funziona **ovunque**, in ogni campo, con la tastiera o con il tasto destro: i
codici arrivano da un foglio Excel e ci tornano, e questo e' il giro che si fa
tutti i giorni.

| Dove | Che cosa si puo' fare |
| --- | --- |
| in **qualsiasi campo di testo** | `Ctrl+C` copia, `Ctrl+V` incolla, `Ctrl+X` taglia, `Ctrl+A` seleziona tutto; il **tasto destro** apre lo stesso menu |
| nelle **celle modificabili al volo** | doppio clic e si incolla come in un qualsiasi programma; `Invio` salva, `Esc` annulla |
| nei **campi in sola lettura** | scrivere no, ma **copiare sempre**: e' il modo piu' rapido per prendere il percorso di un file appena esportato |
| dall'**elenco dei dispositivi** | `Ctrl+C` copia la riga selezionata, incolonnata e pronta per Excel; il tasto destro offre anche *Copia l'identificativo* |
| in **`Elimina +`** | e' fatto apposta per incollare: una colonna di asset tag presa da Excel, una per riga |

Le scorciatoie sono collegate esplicitamente dal programma e non lasciate ai
comportamenti predefiniti di Tk, che cambiano fra sistemi e disposizioni di
tastiera: `Ctrl` e `Cmd` funzionano tutti e due, ovunque.

## Prestiti

Aprendo una stanza elencata fra le *stanze con prestito* (di serie il Digital
Kiosk) l'elenco guadagna la colonna *Prestito*, con su ogni riga un pulsante che
cambia in base allo stato del dispositivo. Fuori da quella stanza la colonna non
esiste:

| Stato | Pulsante | Cosa succede |
| --- | --- | --- |
| Disponibile | **Presta** | chiede il nome della persona e registra nome, data e ora accanto al dispositivo |
| In prestito | **Registra rientro** | chiude il prestito e riporta il dispositivo fra i disponibili |

Finche' il dispositivo e' fuori, la sua riga e' **evidenziata in rosso** e lo
stato risulta **In prestito**; l'evidenziazione si ritrova anche nei file
esportati e stampati. Il nome di chi ha in prestito un dispositivo e' incluso
nella ricerca, quindi basta digitare il cognome per trovare cosa ha in mano.

Tre stati sono automatici e vincono su tutto: **In prestito** mentre c'e' un
prestito in corso, **Da Rispedire** per gli iPhone ancora in casa e **Spedito al
servizio telefonia** per quelli gia' partiti. Registrato il rientro,
il dispositivo torna *Disponibile*. Vedi *Stati* qui sotto per gli altri.

## Spedizione degli iPhone

Nel contenitore **Iphone** ogni telefono non ancora rispedito ha sul proprio rigo
il pulsante **Conferma spedizione**, da premere quando parte davvero verso il
servizio telefonia. La colonna compare solo li': nelle stanze non c'e'.

Alla conferma il programma registra **giorno e ora** nella colonna *Spedito il*,
porta lo stato a *Spedito al servizio telefonia*, colora la riga di **viola** e mostra la frase che
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

## Impostazioni

Il pulsante **Impostazioni**, in alto a destra, apre l'unico posto da cui si
configura il programma. Quello che si decide li' vale **per tutti gli utenti**,
perche' viene salvato in `inventario_impostazioni.json` accanto al file dati -
tranne la lingua, che e' una preferenza personale del singolo computer.

| Campo | A cosa serve |
| --- | --- |
| **Stanze** | l'elenco delle stanze, una per riga. L'ordine e' quello con cui compaiono le schede in home |
| **Tipi di dispositivo** | le voci della tendina *Tipo*, una per riga. Di serie sono *Laptop*, *Tablet* e *Iphone* |
| **Stanze con prestito** | in quali stanze compare la colonna *Prestito*. Devono essere nomi presenti fra le stanze |
| **Stanza degli iPhone** | dove finiscono i telefoni, sempre |
| **Lingua** | Italiano o English. La stessa tendina e' anche nell'intestazione della finestra |

### Creare o rinominare una stanza

Aggiungi una riga nel riquadro **Stanze** e salva: la scheda compare subito in
home, vuota. Per rinominarne una, cambia il testo della riga.

Attenzione: rinominare una stanza **non sposta i dispositivi**, che restano
etichettati con il vecchio nome e compaiono in una scheda a parte. Per portarli
nella stanza rinominata, aprili e usa *Sposta in stanza...*, oppure - piu'
rapido con molti dispositivi - esporta la vecchia stanza, cancella i dispositivi
e reimporta il file dentro la stanza nuova.

Se rinomini la stanza indicata come *Stanza degli iPhone* senza aggiornare quel
campo, il programma se ne accorge e ripiega sulla prima stanza dell'elenco, per
non lasciare i telefoni in una stanza inesistente.

### Aggiungere un tipo di dispositivo

Basta una riga in **Tipi di dispositivo**. Il tipo compare nella tendina di
*Aggiungi* e nel filtro. Un tipo chiamato **iPhone** - scritto come vuoi,
maiuscole e minuscole non contano - attiva da solo tutte le regole dei telefoni:
IMEI al posto di asset tag e seriale, niente prestiti, stanza bloccata,
contenitore in home, spedizione e conservazione.

### Attivare i prestiti in una stanza

Scrivi il nome della stanza nel riquadro **Stanze con prestito**. Se il nome non
corrisponde a nessuna stanza, il salvataggio viene rifiutato con un avviso: e'
il modo per accorgersi di un errore di battitura invece di scoprire piu' tardi
che i pulsanti non compaiono.

### Aggiungere dispositivi

Non passa dalle impostazioni: si usa **Aggiungi** nella barra in alto, che
chiede prima il tipo e poi se inserire a mano o con il lettore di codici. Per
caricarne molti insieme c'e' l'importazione da Excel, e per svuotare tutto prima
di una reimportazione c'e' *Reset inventario*.

### Gli stati

L'elenco degli stati non e' nella finestra: sta alla voce `states` del file
`inventario_impostazioni.json`, che si apre con un editor di testo. Gli stati
automatici - *In prestito*, *Da Rispedire*, *Spedito al servizio telefonia* -
non si toccano.

### Il file delle impostazioni

Viene scritto da solo quando il programma crea un inventario nuovo, accanto al
file dati in `Produzione`, con i valori di partenza: tre stanze, i tre tipi, i
prestiti sul Digital Kiosk, gli iPhone in Site Services BAU e i cinque stati.
Cosi' la configurazione e' visibile e uguale per tutti i tecnici, invece di
dipendere da quello che il programma ha dentro.

Se una voce viene tolta dal file, per quella valgono di nuovo i valori di
partenza. Fa eccezione `loan_rooms`, che puo' restare vuota di proposito: vuol
dire che nessuna stanza gestisce i prestiti.

## Lingua

La tendina **Lingua** sta in due punti: nell'**intestazione** della finestra,
accanto al titolo in alto a destra, e in **Impostazioni**. Passa fra italiano e
inglese e cambia tutto: pulsanti, messaggi, avvisi, nomi delle colonne
nell'elenco e negli stati. Il programma si ridisegna subito, senza riavviare.

La scelta e' **personale del computer**: chi lavora sullo stesso inventario da
un altro PC puo' tenerlo in italiano. I dati nel file restano sempre in italiano
- stati compresi - cosi' due utenti con lingue diverse leggono lo stesso
inventario senza conflitti.

### Esportare in inglese restando in italiano

Nella finestra di esportazione c'e' la casella **Esporta i file in inglese**: e'
spuntata di serie quando l'interfaccia e' in inglese, e si puo' attivare a mano
quando si lavora in italiano ma il file va mandato a qualcuno che legge inglese.

Traduce le intestazioni delle colonne e gli stati. **Non** traduce i nomi delle
stanze e dei tipi, che sono testo tuo. Un file esportato in inglese si
reimporta senza problemi: le intestazioni inglesi sono riconosciute e gli stati
tornano alla forma italiana.

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
prestito** e' *In prestito* e lo stato torna modificabile solo dopo il
rientro; un **iPhone** e' *Da Rispedire* finche' non viene spedito, poi *Spedito al
servizio telefonia*. In entrambi i casi la tendina
appare gia' compilata e bloccata, con la spiegazione accanto, e provando a
cambiarla dall'elenco il programma lo dice nella barra di stato in fondo, senza
aprire finestre.

Lo stato non si scrive mai a mano nel file: se un'importazione porta uno stato
non previsto, viene riportato a *Disponibile*. L'elenco degli stati e' in
`inventario_impostazioni.json`, alla voce `states`. Chi registra prestito e rientro
resta tracciato nelle colonne *Ultima modifica* e *Modificato da*.

Le righe delle stanze senza prestito non hanno alcun pulsante.

## Dove tenere l'inventario vero

**Sulla cartella di rete, in `Produzione\Inventario.xlsx`. Mai dentro la
cartella del programma.**

Il programma sulla postazione contiene dati dimostrativi e file di prova:
l'inventario reale sta altrove, uno solo, dove tutti lo vedono. Cosi' non
finisce in un repository, non lo si sovrascrive con una prova, e nessun
aggiornamento del programma lo sfiora.

Quale inventario aprire e' scritto in `inventario_percorso.json`, accanto al
programma. Per cambiarlo, ripassa `Collega inventario condiviso.bat`.

## Provare subito con dati di esempio

La cartella `Esempio/` contiene un `Inventario.xlsx` con tredici dispositivi
**finti** divisi fra le tre stanze, di cui due gia' in prestito nel Digital
Kiosk. Serve a far vedere il programma a chi lo apre la prima volta: non e'
l'inventario vero, e si rigenera con
`.venv/bin/python tests/genera_esempio.py`.

Per aprirlo senza toccare la configurazione:

```bash
INVENTARIO_FILE=Esempio/Inventario.xlsx python Inventario.py
```

Su Windows, dal Prompt dei comandi nella cartella dell'applicazione:

```bash
set INVENTARIO_FILE=Esempio\Inventario.xlsx && python Inventario.py
```

## Aprire i file su Mac con Numbers

Tutti i file prodotti - inventario, esportazioni, modello, stampa - sono
**.xlsx standard** (Office Open XML). Numbers li apre con un doppio clic: non
serve conversione, e formule non ce ne sono. Lo stesso vale per LibreOffice e
Google Fogli.

Due cose da sapere quando si passa da Numbers:

- Numbers **non importa le tendine** di *Tipo* e *Stato* del modello. I valori
  ammessi restano scritti nel foglio *Istruzioni*: vanno digitati a mano, uguali.
- Numbers salva nel proprio formato **.numbers**, che il programma non legge. Per
  reimportare un file modificato con Numbers usa *Archivio > Esporta con nome >
  Excel*.

Impaginazione di stampa, filtri automatici e righe di intestazione ripetute sono
impostazioni di Excel: Numbers le ignora, ma i dati restano identici.

## Cosa c'e' dove

Sulla **cartella di rete** ci sono solo i dati:

| Percorso | Contenuto |
| --- | --- |
| `Produzione\Inventario.xlsx` | l'inventario; e' gia' il file da consultare, apribile in Excel |
| `Produzione\inventario_impostazioni.json` | stanze, tipi, stanze con prestito, stanza degli iPhone, stati |
| `Produzione\Backup\` | le copie salvate prima di ogni reset e di ogni importazione in sostituzione |
| `.Inventario.xlsx.lock` | presente solo per una frazione di secondo durante un salvataggio |

Sulla **postazione** c'e' solo il programma:

| Percorso | Contenuto |
| --- | --- |
| `Inventario.py`, `inventario\` | il programma |
| `python\` | il Python ufficiale di python.org |
| `inventario_percorso.json` | quale inventario aprire, e la preferenza di lingua |
| `Collaudo\` | i file di prova e le istruzioni di collaudo |

Il programma sulla postazione e' sostituibile in qualsiasi momento senza
conseguenze: tutto quello che conta sta sulla share.

## Provare l'importazione

La cartella **`Collaudo/`** viaggia con il programma: nel pacchetto Windows sta
accanto a `Inventario.py`. Contiene due fogli Excel pronti da importare - uno
regolare con **30 dispositivi, 10 per stanza**, e uno con dentro apposta i casi
che il programma deve segnalare - e le istruzioni passo passo:
[**Come testare l'importazione**](Collaudo/README.md).

Le prove si fanno con i file inclusi, che restano dove sono. Le istruzioni
spiegano anche dove tenere gli **inventari veri** - in una cartella personale sul
proprio computer, mai in quella del programma, che a ogni aggiornamento viene
sostituita - e **come caricare l'inventario definitivo** una volta finito il
collaudo: prima il reset, oppure l'importazione che sostituisce tutto.

I file si rigenerano con `.venv/bin/python tests/genera_file_di_prova.py`.

## Test automatici

```bash
.venv/bin/python tests/run_all.py
```

Diciotto suite che coprono archivio dati e accessi concorrenti, risoluzione del
percorso, schermate e colori, scheda di inserimento, prestiti, spedizioni,
importazione ed esportazione. Girano senza bisogno di una finestra a schermo e
non toccano i dati reali: ognuna si costruisce il proprio inventario in una
cartella temporanea.

Una di queste, `test_collaudo.py`, ripete sui file di `Collaudo/` esattamente
gli scenari descritti nelle istruzioni: se il comportamento cambia, il test
fallisce e le istruzioni vanno riscritte, cosi' non promettono mai cose che non
succedono.
| `.Inventario.xlsx.lock` | presente solo per una frazione di secondo durante un salvataggio |
| `Backup/` | le copie salvate prima di reset e sostituzioni, una per ogni operazione |

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

## Permessi da dare sulla cartella di rete

Un permesso solo, su un oggetto solo.

| Oggetto | Permesso NTFS | Perche' |
| --- | --- | --- |
| la **cartella condivisa** e quello che contiene | **Modifica** | il programma vi crea, sostituisce ed elimina file, non solo scrive dentro `Inventario.xlsx` |

Non serve nessun permesso di esecuzione: dalla share non viene eseguito niente,
il programma sta sulle postazioni.

*Modifica* e non *Scrittura*, perche' ogni salvataggio comporta tre operazioni:

1. crea il file di lock `.Inventario.xlsx.lock`, e poi **lo elimina**;
2. scrive un file temporaneo `Inventario.xlsx.tmp-...`;
3. **sostituisce** `Inventario.xlsx` con il temporaneo.

Dare solo *Scrittura* sul file non basta: mancherebbero la creazione e
l'eliminazione di file nella cartella, e i salvataggi fallirebbero. Anche la
condivisione SMB, non solo NTFS, deve concedere la scrittura: fra i due vince il
piu' restrittivo.

**Utenti in sola consultazione.** Chi ha solo *Lettura* apre il programma e
consulta l'inventario senza problemi; fallisce appena prova a modificare
qualcosa. E' un modo legittimo di dare accesso a chi deve solo guardare.

Se qualcuno tiene `Inventario.xlsx` aperto in Excel, i salvataggi possono
fallire perche' Windows blocca il file: chiudere Excel e riprovare. Per
consultare i dati in Excel senza rischi, usa *Esporta xls...*.

## Scorciatoie

| Tasto | Azione |
| --- | --- |
| `Ctrl+N` | aggiungi dispositivo |
| `Ctrl+F` | vai alla casella di ricerca |
| `Ctrl+P` | stampa la vista corrente |
| `Esc` | torna alla home |
| `F5` | ricarica dalla rete |
| `Canc` | elimina il dispositivo spuntato |
| clic sulla casella | spunta o toglie la spunta alla riga |
| doppio clic sulle note | modifica la nota nell'elenco |
| doppio clic su modello/descrizione | modifica la descrizione nell'elenco |
| doppio clic sullo stato | tendina per cambiare stato nell'elenco |
| doppio clic sul tipo | tendina per cambiare tipo, o si scrive/incolla |
| doppio clic altrove | apre la scheda del dispositivo |
| `Ctrl+C` sull'elenco | copia la riga selezionata, pronta per Excel |
| tasto destro sull'elenco | copia l'identificativo o la riga intera |
| `Ctrl+C` / `Ctrl+V` in un campo | copia e incolla; il tasto destro apre il menu |

## Aggiungere con il lettore di codici a barre

Premendo **Aggiungi** si sceglie prima il tipo di dispositivo, poi fra
inserimento manuale e scansione. Cosa viene letto dipende dal tipo scelto.

**Laptop e tablet** - tre passi, numerati in alto nella finestra:

1. **Scansiona l'asset tag**
2. **Scansiona il numero di serie**
3. **Scrivi il modello** del dispositivo, che sull'etichetta non c'e'

**iPhone** - un passo solo:

1. **Scansiona l'IMEI**, che e' il loro unico identificativo

In entrambi i casi al termine si apre la scheda gia' compilata: restano da
completare i campi che il codice non porta - modello e chi lo ha restituito per
un iPhone, tipo e stanza per gli altri - poi *Salva*. La stanza proposta e'
quella che stai guardando, e per gli iPhone e' sempre la loro.

I lettori di codici a barre si comportano come una tastiera: scrivono nel campo
e confermano da soli, quindi si passa da un codice all'altro senza toccare il
mouse.

### Se il codice non si legge

Ogni passo di scansione ha il pulsante **"Non riesco a scansionare - inserisci a
mano"**: cambia la finestra in scrittura manuale per quel solo campo, senza
perdere i passi gia' fatti e senza uscire dalla procedura. Il campo non puo'
restare vuoto: se si conferma a vuoto, il programma lo dice e resta li'.

## Tornare indietro dopo un errore

Se un'importazione va storta - dispositivi duplicati, la stanza sbagliata, un
file che non era quello - non serve rimediare a mano: si torna alla versione
buona.

**Dalla barra in alto**, il pulsante **Ripristina** propone l'**ultima copia
salvata**, dicendo di quando e' e quanti dispositivi conteneva rispetto a quelli
che ci sono adesso. E' il caso piu' frequente: si annulla l'ultima operazione
distruttiva con due clic.

**Da *Impostazioni* > *Ripristina da una copia...*** si sceglie invece fra tutte
le copie disponibili, elencate dalla piu' recente con data, ora e numero di
dispositivi: serve quando l'errore risale a qualche passaggio prima.

Il ripristino **salva prima lo stato attuale** in una nuova copia, quindi anche
un ripristino sbagliato si annulla. Una copia illeggibile o sparita viene
rifiutata senza toccare l'inventario.

Il ripristino agisce sull'inventario condiviso: quello che si riporta indietro
lo vedono tutti i tecnici.

## La copia in locale

Le copie automatiche stanno sulla cartella di rete, accanto ai dati. Coprono
l'errore umano - un reset di troppo, un'importazione sbagliata - ma **non**
coprono il caso in cui sparisca la cartella di rete, o qualcuno ci cancelli
dentro: in quel caso spariscono anche loro.

Per quello c'e' **`Salva copia in locale...`** nella barra dei comandi. Salva
dove decidi tu - il tuo PC, una chiavetta, fuori dalla rete - una copia
dell'inventario **com'e' in quel secondo**:

- viene presa dal file sulla rete nel momento in cui la chiedi, non da quello
  che il programma aveva letto prima;
- se in quell'istante un altro tecnico sta salvando, il programma aspetta che
  finisca: non ottieni mai un file colto a meta' scrittura;
- accanto ai dati viene salvato anche il file delle impostazioni, con lo stesso
  nome e il suffisso `_impostazioni.json`: da soli i dati non basterebbero a
  rimettere l'inventario com'era;
- **e' un inventario completo, non un estratto**: si apre in Excel, e si
  ricarica con *Ripristina* o con *Importa xls...* in modalita' Sostituisci.

Il nome proposto porta data e ora, `Inventario_2026-08-31_18-30.xlsx`. Vale la
pena farla prima di ogni operazione grossa, e ogni tanto per abitudine.

## Le copie di sicurezza automatiche

Prima di ogni operazione che cancella dati - il **reset** e ogni **importazione
in sostituzione**, sull'intero inventario o su una singola stanza - il programma
duplica il file dati nella cartella **`Backup`**, accanto ai dati sulla rete.

Il nome porta **la data del file salvato**, non quella della copia:
`Inventario_2026-08-31_09-12-45.xlsx`. Cosi' due reset di fila sullo stesso
inventario non producono due file identici, e cercando una versione si guarda a
quando risale il contenuto invece che a quando qualcuno ha premuto un pulsante.
Se una copia con quel nome esiste gia', ne viene aggiunta una numerata.

**Se la copia non riesce, l'operazione si annulla** e non viene toccato niente.
Se la cartella del programma e' in sola lettura, il programma ripiega su una
cartella `Backup` accanto al file dati e poi sul profilo utente: una copia deve
poter essere scritta.

Per recuperare, apri il file di backup con il programma o con Excel: e' un
inventario completo, non un formato speciale. Le copie **non finiscono mai nel
repository**: la cartella ha una propria regola che le esclude.

Vanno svuotate a mano ogni tanto: nessuno le cancella al posto tuo.

## Svuotare l'inventario per ricaricarlo

Il pulsante **Reset inventario**, in alto a destra, serve a ripartire da zero
prima di una reimportazione completa.

Non e' un'operazione che si fa per sbaglio: prima di procedere il programma

1. mostra un avviso con quanti dispositivi verranno eliminati **per tutti gli
   utenti**, e chiede di scrivere per esteso `ELIMINA TUTTO`;
2. **salva una copia** del file dati nella cartella **`Backup`**, dentro quella
   del programma, con la data del file salvato nel nome
   (`Inventario_2026-08-31_09-12-45.xlsx`);
3. solo allora svuota l'inventario.

Se la copia di sicurezza non riesce - cartella piena, permessi mancanti - il
reset viene annullato e non si tocca niente.

**Gli iPhone restano sempre.** Il reset non ne elimina nessuno, in nessuno
stato: non arrivano da un'importazione, quindi cancellarli qui vorrebbe dire
perderli per sempre. Vale anche per quelli spediti da oltre tre mesi, che a mano
si potrebbero eliminare. Il programma dice quanti ne ha mantenuti; se in
inventario ci sono solo iPhone, avverte che non c'e' niente da eliminare.

Dopo il reset si ricarica tutto con *Importa xls...*.

## Importare dentro una sola stanza

Ci si arriva in due modi, che si comportano in modo **identico**:

- da dentro la stanza, con **Importa i dati di questa stanza**, accanto al suo nome;
- dalla pagina principale, con *Importa xls...* e l'opzione **Una sola stanza**.

La regola e' una sola, e dipende da cosa c'e' nel foglio:

| Il foglio | Cosa succede |
| --- | --- |
| **dichiara le stanze** con le righe-separatore | viene caricata solo la sezione della stanza scelta; tutte le altre righe vengono **scartate**, anche se nello stesso file |
| **non dichiara nessuna stanza** | vale la scelta fatta: tutte le righe finiscono nella stanza indicata |
| dichiara stanze, ma **non quella scelta** | non viene importato **niente**, e un avviso spiega di aggiungere la riga mancante, elencando le stanze trovate al suo posto |

Va bene anche la forma breve - `KIOSK` per *Digital Kiosk*. Il riepilogo prima
della conferma dice quante righe entrano, quante ne vengono scartate perche' di
altre stanze, e - quando il foglio non dichiara stanze - che tutte finiranno
nella stanza scelta. Come sempre si sceglie fra *unisci* e *sostituisci*, e la
sostituzione riguarda solo quella stanza.

## Come funziona l'esportazione

Come per l'importazione, *Esporta xls...* apre prima una finestra con due
domande.

**Che cosa** - tutto l'inventario, oppure una sola stanza scelta dalla tendina.

**In che forma** - quando esporti tutto:

| | |
| --- | --- |
| Un unico elenco | tutti i dispositivi in un solo foglio |
| Un foglio per ogni stanza | un solo file, con dentro un foglio per stanza |
| Un file separato per ogni stanza | si sceglie una cartella, ed esce un file per stanza, chiamato `Inventario_<Stanza>_<data>.xlsx` |

Le stanze senza dispositivi non producono ne' fogli ne' file vuoti.

**Ogni foglio dice di chi e'.** Il nome della stanza compare in tre punti: e' il
nome del foglio, e' scritto in testa alla prima riga, e sotto c'e' la data di
esportazione con il numero di dispositivi. La colonna *Stanza* resta sempre nella
tabella. Cosi' un foglio resta riconoscibile anche se viene copiato altrove,
rinominato o stampato.

**Si reimporta tutto.** Qualsiasi cosa produca l'esportazione si puo' ricaricare:
un file con piu' fogli viene letto per intero, e il nome di ogni foglio vale come
la riga-separatore di quella stanza.

Dentro una stanza c'e' anche la scorciatoia **Esporta questa stanza in xls**, che
salta la finestra e produce direttamente il file di quella stanza.

### Mandarlo per e-mail

Ogni esportazione, qualunque forma abbia scelto, finisce con la stessa domanda:
**Invia per e-mail con Outlook**, *Apri il file*, oppure niente.

Scegliendo l'invio si apre un **messaggio nuovo di Outlook con il file gia'
allegato**. Destinatario, oggetto e testo li scrivi tu, e l'invio resta tuo: il
programma prepara il messaggio e si ferma li'. Se l'esportazione ha prodotto piu'
file - un file per stanza - vengono raccolti in un **archivio zip**, perche'
Outlook accetta un allegato solo dalla riga di comando.

Se Outlook non e' installato su quel computer, il programma lo dice e ricorda
che **il file e' stato creato lo stesso**: basta allegarlo a mano.

Non c'e' nessuna voce di menu per l'invio: e' un'opzione in coda
all'esportazione, perche' si manda quello che si e' appena prodotto.

Esporta la stanza **intera**, non quello che stai vedendo: eventuali ricerche o
filtri attivi non la riducono. Per esportare esattamente la vista corrente c'e'
sempre *Esporta xls...* nella barra in alto.

## Il modello di importazione

Per caricare in blocco laptop e tablet parti dal modello gia' pronto, che si
ottiene in due modi:

- dal programma, pulsante **Scarica il modello di importazione** in home, a
  destra del titolo *Inventario completo*: esce nella lingua dell'interfaccia;
- dalla pagina del progetto, o direttamente da
  [`docs/Modello_inventario.xlsx`](docs/Modello_inventario.xlsx) (italiano) e
  [`docs/Import_template.xlsx`](docs/Import_template.xlsx) (inglese).

Ha **le stesse colonne di un file esportato** - *Asset Tag*, *Tipo*, *Stanza*,
*Note* - con le tendine su *Tipo* e *Stanza*, le righe-separatore gia' pronte e
un foglio *Istruzioni*. Cosi' si esporta, si corregge in Excel e si reimporta
senza cambiare formato.

Le colonne che il modello non ha - *Modello/Descrizione*, *Numero di serie*,
*Stato* - **restano importabili**: se il tuo foglio le contiene vengono
riconosciute dal nome e caricate lo stesso.
Il modello generato dal programma rispecchia le stanze e gli stati configurati
in quel momento.

Compilalo e caricalo con *Importa xls...*: non serve aggiungere la colonna
*Stanza*, ci pensano i separatori.

## Gli iPhone restano fuori da import ed export

Gli iPhone si inseriscono **solo a mano** dal programma. Non compaiono nel
modello, vengono ignorati se presenti in un file importato (la finestra di
importazione dice quanti), e non finiscono in nessuna esportazione, nemmeno in
quella divisa per stanza. Una **sostituzione** dell'inventario non li cancella:
vengono mantenuti, perche' non potrebbero essere ricaricati da un file.

Restano invece nella **stampa**, che e' consultazione interna, e ovviamente
nell'inventario a video e nel file dati.

## Come funziona l'importazione

Premendo *Importa xls...* la prima cosa che compare non e' il selettore dei file,
ma la scelta di **che cosa** si carica e **come**.

**Che cosa**

| | |
| --- | --- |
| Tutto l'inventario | il file riguarda l'intero parco dispositivi |
| Una sola stanza | si sceglie la stanza dalla tendina, e comandano i separatori: vedi *Importare dentro una sola stanza* |

**Come**

| | |
| --- | --- |
| Unisci | aggiunge i nuovi e aggiorna quelli gia' presenti con lo stesso asset tag |
| Sostituisci | svuota prima, poi carica solo cio' che c'e' nel file |

Le due scelte si combinano: *sostituisci una sola stanza* rifa' da zero il
contenuto di quella stanza e lascia intatte le altre.

Scelto il file, un riepilogo mostra quante righe sono state lette, cosa e' stato
ignorato e - per una sostituzione - quanti dispositivi verranno eliminati. Fino a
quel momento non e' stato scritto niente: annullando, l'inventario resta com'era.

### Le protezioni sulla sostituzione

- Prima di ogni sostituzione viene **salvata una copia** del file dati nella
  cartella `Backup`. Se la copia non riesce, l'operazione si annulla.
- Per sostituire **tutto l'inventario** bisogna scrivere per esteso
  `ELIMINA TUTTO`: e' l'operazione piu' distruttiva del programma e riguarda i
  dati di tutti. Per una singola stanza basta la conferma, con il numero di
  dispositivi in chiaro.
- **Gli iPhone non vengono mai eliminati**, in nessuna delle due modalita': non
  arrivano da un'importazione, quindi una sostituzione li perderebbe per sempre.

## Come preparare il file Excel

La guida completa, con esempi e tabelle, sta in una pagina a parte:
**[Come preparare il file Excel](https://angelo-tassi.github.io/site-services-inventario/formato-xls.html)**
(sorgente: [`docs/formato-xls.html`](docs/formato-xls.html)). Copre i nomi
riconosciuti per ogni colonna, le righe-separatore per dividere le stanze, gli
errori piu' frequenti e cosa succede a ogni riga.

In breve: una riga di intestazione con i nomi delle colonne, una riga per
dispositivo, e almeno l'**Asset Tag** (o l'**IMEI** per i telefoni). Il resto e'
facoltativo. Partendo dal
[modello](docs/Modello_inventario.xlsx) non serve nemmeno leggerla.

## Se il file ha colonne diverse dalle nostre

L'importazione e' tollerante e non si blocca per un file "sporco". Ecco
esattamente cosa succede.

| Nel file | Cosa fa il programma |
| --- | --- |
| Colonne in piu' (costo, fornitore, centro di costo...) | le ignora e **te le elenca** prima di importare |
| Nomi con maiuscole o spazi diversi | li riconosce lo stesso: `  ASSET TAG `, `tipo`, `MoDeLLo`, `s/n` vanno bene |
| Due colonne per lo stesso dato | usa la prima e segnala la seconda fra quelle ignorate |
| Manca il **modello** | importa lo stesso e ti dice quante righe restano senza |
| Manca l'**asset tag** (o l'IMEI) | si ferma con un errore e non importa niente |
| Un titolo prima della tabella | lo salta e cerca le intestazioni nelle prime 12 righe |
| Piu' fogli | li legge tutti; un foglio intitolato come una stanza vale come separatore |
| Un foglio senza tabella (istruzioni, appunti) | lo ignora |
| Righe vuote | le salta senza contarle |
| Righe senza identificativo | le conta come scartate e va avanti |

Il punto delicato sono le **colonne non riconosciute**: se il tuo file chiama il
modello *Descrizione articolo*, quel dato verrebbe perso in silenzio. Per questo
la finestra di importazione, prima di chiederti conferma, mostra un riquadro con
i nomi delle colonne che non ha capito e ti invita a rinominarle. Rinominare
l'intestazione nel foglio e riprovare e' sufficiente: nessun dato va perso,
perche' finche' non confermi non viene scritto niente.

Il modo piu' sicuro di non incontrare il problema e' partire dal
[modello](docs/Modello_inventario.xlsx), che ha gia' le intestazioni giuste.

## Dividere per stanza un inventario unico

Se hai un solo foglio Excel con tutti i dispositivi e nessuna colonna *Stanza*,
non serve aggiungerla: basta spezzare l'elenco con delle **righe-separatore**.

Una riga con **una sola cella scritta**, contenente il nome di una stanza,
assegna quella stanza a tutte le righe che seguono, fino al separatore
successivo.

| Asset Tag | Tipo | Modello/Descrizione | Numero di serie |
| --- | --- | --- | --- |
| **BAU** | | | |
| IT-0101 | Laptop | Lenovo ThinkPad T14 Gen 4 | PF4A1B2C |
| IT-0104 | Tablet | Dell Latitude 7320 Detachable | 8H2KLM3 |
| **KIOSK** | | | |
| IT-0106 | Laptop | Lenovo ThinkPad T14 Gen 5 | PF5K9M8F |
| **DISASTER** | | | |
| DR-0201 | Laptop | Lenovo ThinkPad T14 Gen 4 | PF4B7T1J |

I tag riconosciuti sono ricavati **dai nomi delle stanze configurate**: vale il
nome completo e ogni singola parola che non sia ambigua. Con le stanze di serie:

| Scrivi | Finisce in |
| --- | --- |
| `BAU`, `SITE`, `SERVICES`, `SITE SERVICES BAU` | Site Services BAU |
| `KIOSK`, `DIGITAL`, `DIGITAL KIOSK` | Digital Kiosk |
| `DISASTER`, `RECOVERY`, `MAGAZZINO`, `MAGAZZINO DISASTER RECOVERY` | Magazzino Disaster Recovery |

Maiuscole, minuscole e due punti finali non contano: `KIOSK`, `kiosk` e
`Kiosk:` sono la stessa cosa. Se rinomini una stanza, i tag si aggiornano da
soli. Una riga con piu' di una cella scritta non e' mai un separatore, quindi un
dispositivo il cui asset tag sia per caso `BAU` non crea confusione.

Il separatore ha la **precedenza sulla colonna Stanza**, se presente. Fanno
eccezione gli iPhone, che finiscono comunque nella loro stanza. Prima di
confermare, la finestra di importazione dice quante righe hanno preso la stanza
dai separatori.

## Creare l'inventario da un file Excel esistente

Non serve inserire i dispositivi a mano: *Importa xls...* accetta un qualsiasi
file Excel la cui **prima riga contenga le intestazioni**. L'unica colonna
obbligatoria e' l'asset tag - o, per gli iPhone, l'IMEI, che ne fa le veci; le
altre sono riconosciute per nome, in italiano o in inglese, per esempio:

| Colonna | Intestazioni riconosciute |
| --- | --- |
| Asset Tag | Asset Tag, Asset, Tag, Etichetta, Inventario |
| Tipo | Tipo, Tipologia, Categoria, Type |
| Modello/Descrizione | Modello, Modello/Descrizione, Model, Descrizione, Dispositivo |
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
Collega inventario condiviso.bat  punta la postazione all'inventario sulla share
Crea collegamento sul desktop.bat mette l'icona sulla scrivania dell'utente
Diagnostica.bat                   raccoglie le informazioni per chi assiste
Avvia Inventario.bat              avvio dai sorgenti su Windows
Avvia Inventario.command          avvio dai sorgenti su macOS / Linux
Inventario.py                     avvio dell'applicazione
inventario/config.py     percorso del file dati e impostazioni condivise
inventario/configura.py  collega la postazione all'inventario condiviso
inventario/diagnostica.py rapporto sullo stato dell'installazione
inventario/store.py      lettura/scrittura del file .xlsx, lock, operazioni
inventario/excel_io.py   esportazione, impaginazione di stampa, invio a stampante
inventario/lingua.py     traduzioni in italiano e inglese
inventario/theme.py      palette, font e stili dell'interfaccia
inventario/ui.py         interfaccia grafica (home, vista stanza, dialoghi)
```

Per aggiungere un campo basta inserirlo in `FIELDS`, `HEADERS`, `HEADER_ALIASES`
e nelle larghezze colonna in `inventario/store.py`, in `PRINT_FIELDS` in
`inventario/excel_io.py`, e nella scheda `ItemDialog` in `inventario/ui.py`.

Un file `Inventario.xlsx` creato con una versione precedente resta leggibile: le
colonne mancanti risultano vuote e vengono aggiunte al primo salvataggio.
