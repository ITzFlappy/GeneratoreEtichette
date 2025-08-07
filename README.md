
# Generatore Etichette - Streamlit App

Questa app genera automaticamente slide PowerPoint a partire da:
- Un file `.pptx` template (con segnaposto come `<colonna_1>`, `<qta>`, ecc.)
- Un file `.xlsx` o `.xls` con i dati

Ogni riga dell'Excel genera una nuova slide.

## Esempio di utilizzo

1. Vai su https://generatoreetichette.streamlit.app/
2. Carica il powerpoint contenente il template e il file Excel con i dati
3. Scarica la presentazione generata

## Requisiti
- I segnaposto nel template devono essere scritti come: `<nome_colonna_excel>`
- Il file Excel deve avere intestazioni compatibili

Funziona anche con immagini, linee, griglie e layout verticale.
