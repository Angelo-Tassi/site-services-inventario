# File di prova per il collaudo

Due fogli Excel pronti da importare, per provare il programma senza toccare
l'inventario vero.

| File | Contenuto |
| --- | --- |
| `Inventario_di_prova.xlsx` | 30 dispositivi, **10 per stanza**, divisi dalle righe-separatore. Tutto regolare. |
| `Inventario_di_prova_con_difetti.xlsx` | Lo stesso impianto, ma con dentro apposta i casi che il programma deve segnalare. |

Si rigenerano con:

```bash
.venv/bin/python tests/genera_file_di_prova.py
```

## Prima di cominciare

Fai le prove su un inventario **finto**, non su quello della cartella di rete.
Il modo piu' rapido e' creare una cartella vuota, copiarci `Inventario.exe` e
aprirlo: al primo avvio propone di creare li' un `Inventario.xlsx` nuovo.

---

# Come testare l'importazione

## 1. Caricamento iniziale

1. *Importa xls...*
2. **Tutto l'inventario** + **Unisci**, poi *Scegli il file*
3. seleziona `Inventario_di_prova.xlsx`
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

**Cosa deve risultare.** Digital Kiosk contiene ancora **10** dispositivi, quelli
elencati sotto la riga `DIGITAL KIOSK`: il foglio dichiara le stanze, quindi
comandano i separatori e le righe delle altre stanze vengono buttate via. Le
altre due stanze restano com'erano, con 10 ciascuna.

Nella cartella dei dati deve essere comparso un file
`Inventario_prima_del_reset_...xlsx`: e' la copia di sicurezza.

### Quando il foglio non dichiara le stanze

Apri `Inventario_di_prova.xlsx`, **cancella tutte e tre le righe-separatore** e
salva con un altro nome. Reimportalo con **Una sola stanza** > `Digital Kiosk`.

Stavolta il riepilogo dice **30 righe valide** e avverte che *il foglio non
dichiara stanze: tutte le righe finiranno in Digital Kiosk*. E' il caso in cui
la scelta dell'utente e' l'unica informazione disponibile.

## 4. Sostituire tutto

1. *Importa xls...* > **Tutto l'inventario** > **Sostituisci**
2. scegli `Inventario_di_prova.xlsx`
3. il riepilogo chiede di scrivere `ELIMINA TUTTO`

Prova prima a confermare **senza** scriverlo, o scrivendo "si": deve rifiutare e
restare aperto. Poi scrivi la frase e conferma: si torna ai 30 dispositivi
distribuiti 10, 10 e 10.

## 5. Le segnalazioni sui file imperfetti

Importa `Inventario_di_prova_con_difetti.xlsx` con **Tutto l'inventario** +
**Unisci**. Il riepilogo deve mostrare, tutte insieme:

- **3 righe valide trovate**
- **1 riga ignorata**: manca l'identificativo
- **3 righe hanno preso la stanza dai separatori**
- **1 iPhone ignorato**: si inseriscono solo a mano
- il riquadro rosso con le **colonne non riconosciute**: `Costo`, `Fornitore`,
  `Centro di costo`
- l'avviso che **1 riga non ha il modello**

Confermando, entrano 3 dispositivi: due in Site Services BAU (di cui uno senza
modello, da completare a mano) e uno in Digital Kiosk. L'iPhone **non** entra.

## 6. Importare dentro una stanza

Questa e' la prova del pulsante **Importa i dati di questa stanza**, che sta
accanto al nome della stanza quando la apri. Deve dare lo **stesso risultato**
dell'opzione *Una sola stanza* del punto 3: e' la stessa regola, raggiunta da
due strade diverse.

1. entra in **Digital Kiosk**
2. premi *Importa i dati di questa stanza*, scegli **Unisci**
3. seleziona `Inventario_di_prova.xlsx`

**Cosa deve risultare.** Il riepilogo dice **10 righe valide** e **20 righe di
altre stanze scartate**. Entrano solo i dieci dispositivi elencati sotto la riga
`DIGITAL KIOSK`: quelli di BAU e del magazzino vengono buttati via, anche se
sono nello stesso foglio.

### La riga della stanza e' obbligatoria

Prendi `Inventario_di_prova.xlsx`, **cancella la riga `DIGITAL KIOSK`** e salva
con un altro nome. Poi riprova a importarlo dentro Digital Kiosk.

Deve comparire un avviso che dice che nel foglio non c'e' nessuna riga che
indichi quella stanza, che **non e' stato importato niente**, e che spiega di
aggiungere una riga vuota con scritto `DIGITAL KIOSK` nella prima cella.
L'avviso elenca anche le stanze che ha trovato al suo posto. Controlla che
l'inventario sia rimasto identico.

Prova anche la forma breve: al posto di `DIGITAL KIOSK` scrivi solo `KIOSK`.
Deve funzionare lo stesso.

## 7. Reset e ricarica

1. *Reset inventario*, scrivi `ELIMINA TUTTO`
2. controlla il messaggio finale: dice quanti eliminati e dove ha salvato la copia
3. reimporta `Inventario_di_prova.xlsx`

Se prima del reset avevi inserito a mano un iPhone, deve **sopravvivere**: gli
iPhone non si eliminano e non si reimportano.

## 8. Esportazione

Con l'inventario carico dal punto 1, prova le tre forme da *Esporta xls...*.

**Un unico elenco.** Un foglio solo con tutti e 30 i dispositivi.

**Un foglio per ogni stanza.** Un file con tre fogli, uno per stanza, 10
dispositivi ciascuno. Apri ogni foglio e controlla che in **A1** ci sia il nome
della stanza, sotto la data di esportazione, e che la colonna *Stanza* ci sia e
riporti sempre la stessa stanza. Era questo il difetto della versione
precedente: il foglio non diceva di chi era.

**Un file separato per ogni stanza.** Scegli una cartella: devono uscire tre
file `Inventario_<Nome_stanza>_<data>.xlsx`, ciascuno con la sua stanza in A1.

Poi la controprova: **reimporta** il file con i tre fogli, con *Tutto
l'inventario* + *Unisci*. Deve leggerli tutti e tre e riconoscere 30
dispositivi, perche' il nome del foglio vale come riga-separatore.

Infine, entra in una stanza e usa la scorciatoia **Esporta questa stanza in
xls**: stesso risultato dell'opzione *Una sola stanza*. In nessuna forma devono
comparire iPhone.

---

## Se qualcosa non torna

Segnala il problema con: quale punto stavi facendo, cosa ti aspettavi, cosa e'
successo, e se puoi la schermata del riepilogo. Su Mac, il diario di avvio e' in
`avvio.log` nella cartella del programma.
