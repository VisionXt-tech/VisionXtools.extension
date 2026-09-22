# Changelog VisionXtools

Il pulsante **About > Version** legge questo file: la prima sezione e' la versione installata.
Formato di ogni sezione: `## <versione> - <data>`, seguita dall'elenco delle novita'.

## 1.1.0 - 22/09/2026

- Center Room Tags: il punto di riferimento della stanza e il tag vengono centrati sulla forma della stanza (baricentro dell'area, fori esclusi) invece che sul centro della bounding box.
- Center Room Tags: sulle stanze a L o a U, dove il baricentro cade fuori dalla stanza, viene usato il punto interno piu' lontano dai muri.
- Center Room Tags: risolto l'errore che interrompeva il comando sulle stanze non visibili nella vista attiva.
- Center Room Tags: i tag delle stanze nei modelli collegati vengono saltati invece di bloccare il comando.
- About: nuovo pulsante Version, che mostra versione, data e novita' della copia installata.
- Installer: `Installa_VisionXtools.bat` ora aggiorna anche le installazioni gia' presenti, mettendo da parte la copia precedente.

## 1.0.0 - 16/09/2026

- Primo rilascio distribuibile dell'estensione.
- Installer per Windows a un clic (`Installa_VisionXtools.bat`).
- Guida di installazione nel README.
- Licenza PolyForm Internal Use 1.0.0.
- Tab AnasL4 archiviata e nascosta dalla barra.
