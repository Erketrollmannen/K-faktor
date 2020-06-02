# Kfaktorlogg

Enkelt script for å raskere kunne hente ut K-faktorer og få disse over i Excel.

## Virkemåte

Programmet består av to script. Sortering av PDFer generert av flowcomputer og tallflytting frå Excel til Excel. 

### Sortering av PDF

pdfSortV2 sorterer PDFen ut frå målestasjon, måleløp og oljetype. Den oppretter og sorterer i følgende mapper:

- MSA\_1
- MSA\_14
- MSB\_1
- MSB\_14

pdfSortV2 ser etter PDFer i den mappa det blir kjørt frå. 

### Flytting og tallknusing.

Før autoxl kan kjøres må man semi manuellt lage excel ark frå PDFene. Dette gjeres i Adobe Acrobat. Autoxl ser etter "Merget" celle. Den dataden du vil ha med må skilles av med å merge celler. 

Når autoxl flytter data sjekker den om datoen allerede eksisterer i Excelarket det sender dataen til. Dersom dataen allerede eksisterer hoppes den over. 

Dei nye "oppdaterte" excelarkene blir putta i mappa "out". Dei må manuellt flyttes herfrå til G-disk. Kan være lurt å ta ein manuall spotsjekk av generert excel ark. 

## Biblioteker i bruk.

### PDFminer

For å lese PDF brukes [pdfminer](pypi.org/project/pdfminer)

### openpyxl 

For å lese og srive til excel ark brukes [openpyxl](https://openpyxl.readthedocs.io/en/stable/#)

