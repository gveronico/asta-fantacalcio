# Asta Fantacalcio

Strumento da usare durante l'asta: liste di calciatori divise in fasce, segnalini rapidi e gestione della rosa
da 300 crediti. Pensato per iPad in verticale, funziona anche offline.

Il valore di mercato (FVM) del file Excel serve solo a ordinare i giocatori: **non compare mai** nella pagina,
nemmeno nel codice sorgente.

## Come si usa

- In alto: tab dei ruoli, ricerca per nome o squadra, crediti rimasti (tocca per aprire la rosa).
- In basso: lo strumento attivo. Lo scegli una volta e resta attivo fino al cambio.

| Strumento | Effetto sul giocatore |
| --------- | --------------------- |
| Altri | preso da un'altra squadra (rosso) |
| Mio | preso da te: chiede i crediti e lo mette in rosa (verde) |
| Nascondi | lo toglie dalla lista (recuperabile con "Mostra nascosti") |
| Piace | lo evidenzia in giallo |
| Rischio | badge `R` per infortunati o incognite |
| Libero | azzera i segni e lo rimuove dalla rosa |

Composizione della rosa: 3 portieri, 9 difensori, 9 centrocampisti, 6 attaccanti. Tutto lo stato resta salvato
nel browser, quindi puoi chiudere e riaprire senza perdere l'asta.

## Fasce

- **Portieri:** il migliore di ogni squadra, tutte e 20 rappresentate, senza fasce.
- **Difensori e centrocampisti:** 9 fasce da 8 + una fascia bonus da 10.
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
