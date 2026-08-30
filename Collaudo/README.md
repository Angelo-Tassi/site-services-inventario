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
3. il riepilogo avverte che verranno prima eliminati **10** dispositivi

**Cosa deve risultare.** Digital Kiosk contiene ora **tutti e 30** i dispositivi
del file, perche' con la stanza scelta ogni riga ci finisce dentro,
indipendentemente dai separatori. Le altre due stanze restano vuote. E' il
comportamento voluto: serve a caricare l'elenco di una stanza da un foglio che
non dichiara la stanza.

Nella cartella dei dati deve essere comparso un file
`Inventario_prima_del_reset_...xlsx`: e' la copia di sicurezza.

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

## 6. Reset e ricarica

1. *Reset inventario*, scrivi `ELIMINA TUTTO`
2. controlla il messaggio finale: dice quanti eliminati e dove ha salvato la copia
3. reimporta `Inventario_di_prova.xlsx`

Se prima del reset avevi inserito a mano un iPhone, deve **sopravvivere**: gli
iPhone non si eliminano e non si reimportano.

## 7. Esportazione di una stanza

Entra in una stanza e premi **Esporta questa stanza in xls**. Il file deve
chiamarsi `Inventario_<Nome_stanza>_<data>.xlsx` e, aperto, avere il nome della
stanza come **nome del foglio** e nella **prima riga**, con sotto la data di
esportazione. Devono esserci solo i dispositivi di quella stanza, e nessun
iPhone.

---

## Se qualcosa non torna

Segnala il problema con: quale punto stavi facendo, cosa ti aspettavi, cosa e'
successo, e se puoi la schermata del riepilogo. Su Mac, il diario di avvio e' in
`avvio.log` nella cartella del programma.
