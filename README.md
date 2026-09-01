# Asta Fantacalcio

Strumento da usare durante l'asta: liste di calciatori divise in fasce, segnalini rapidi e gestione della rosa.
All'apertura scegli **Fantacalcio 300** o **Fantacalcio 500**: sono due app separate, con rosa, preferiti e piani
che non si mescolano. Pensato per iPad in verticale, funziona anche offline.

Il valore di mercato (FVM) del file Excel è su 1000 crediti: serve a ordinare i giocatori e compare
sulla card, quotato sui crediti della versione scelta (300 o 500) e arrotondato all'intero.

## Come si usa

- All'inizio (e dal bottone **F300** / **F500** a destra): scegli la versione.
- In alto: **Rosa** apre la squadra, ricerca per nome o squadra, crediti rimasti a destra della ricerca.
- In basso: lo strumento attivo. Lo scegli una volta e resta attivo fino al cambio.

| Strumento | Effetto sul giocatore |
| --------- | --------------------- |
| Altri | preso da un'altra squadra (rosso) |
| Mio | preso da te: chiede i crediti e lo mette in rosa (verde) |
| Nascondi | lo toglie dalla lista (recuperabile con "Mostra nascosti") |
| Piace | lo evidenzia in giallo |
| Rischio | badge `R` per infortunati o incognite |
| Libero | azzera i segni e lo rimuove dalla rosa |

Composizione della rosa:

- **Fantacalcio 300:** 300 crediti, 2 portieri, 9 difensori, 9 centrocampisti, 6 attaccanti.
- **Fantacalcio 500:** 500 crediti, 3 portieri, 8 difensori, 8 centrocampisti, 6 attaccanti.

Tutto lo stato resta salvato nel browser, quindi puoi chiudere e riaprire senza perdere l'asta.

## Piani di spesa e valori

Da **Piani** puoi creare più ripartizioni dei crediti della versione scelta, scegliere quella attiva e
compilarla in crediti o percentuale.
Durante l'asta i tab mostrano quanto resta per ruolo; nella rosa trovi piano, speso e margine con una barra di avanzamento.
Tocca la testata di una fascia per salvarne il valore minimo e massimo: il range compare anche quando assegni un giocatore.
Il reset dell'asta cancella segni e rosa, ma conserva piani e valutazioni di quella versione.

## Fasce

- **Portieri:** il migliore di ogni squadra, tutte e 20 rappresentate, senza fasce.
- **Difensori e centrocampisti (300):** 9 fasce da 8 + una fascia bonus da 10.
- **Difensori e centrocampisti (500):** 8 fasce da 8; i giocatori della nona fascia restano in bonus insieme agli altri.
- **Attaccanti:** 6 fasce da 8 + una fascia bonus da 10.

Le fasce sono da 8 perché la lega ha 8 squadre: la prima fascia contiene i giocatori che, in teoria, si spartiscono
le 8 squadre, e così via.

## Aggiornare le liste con un nuovo Excel

Serve Python con `openpyxl` (`pip install openpyxl`).

1. Metti il nuovo file `.xlsx` (stessa struttura, foglio `Tutti`) nella cartella.
2. Esegui `python build_liste.py` — oppure `python build_liste.py percorso.xlsx` per indicarlo a mano.
3. Ricarica la pagina sull'iPad.

Lo script riscrive solo il blocco dati dentro `asta.html`, tra i marker `PLAYERS_DATA_START` e `PLAYERS_DATA_END`:
grafica, strumenti e rosa restano intatti. I segni già messi restano sui giocatori ancora presenti, quelli
scomparsi dal listone escono dalla rosa.

## File

| File | A cosa serve |
| ---- | ------------ |
| `asta.html` | l'app: unico file necessario per usarla |
| `build_liste.py` | rigenera le liste da un Excel |
| `make_icone.py` | rigenera le icone per la schermata Home |
| `sw.js`, `manifest.webmanifest`, `icon-*.png` | funzionamento offline e icona quando è pubblicata online |

## Aggiungerla alla schermata Home (iPad/iPhone)

Apri il link con **Safari**, poi **Condividi** → **Aggiungi alla schermata Home**. Si apre a tutto schermo,
senza barre del browser, e dopo la prima apertura funziona anche senza rete.
