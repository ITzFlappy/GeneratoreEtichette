
# Generatore Slide Bennet - Streamlit App

Questa app genera automaticamente slide PowerPoint a partire da:
- Un file `.pptx` template (con segnaposto come `<colonna_1>`, `<qta>`, ecc.)
- Un file `.xlsx` o `.xls` con i dati

Ogni riga dell'Excel genera una nuova slide.

## Esempio di utilizzo

1. Vai su https://streamlit.io/cloud
2. Collega il tuo GitHub
3. Importa questo repository
4. Carica `template.pptx` e il file Excel
5. Scarica la presentazione generata

## Requisiti
- I segnaposto nel template devono essere scritti come: `<nome_colonna_excel>`
- Il file Excel deve avere intestazioni compatibili

Funziona anche con immagini, linee, griglie e layout verticale.
