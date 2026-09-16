# VisionXtools.extension

Estensione pyRevit di VisionXt: aggiunge a Revit i tab **VisionXtools** e **AnasL4**.

**Prerequisito:** pyRevit già installato (il tab *pyRevit* compare in Revit).

## Installazione (consigliata): doppio clic

1. Apri [`Installa_VisionXtools.bat`](Installa_VisionXtools.bat) e scaricalo con il pulsante
   **Download raw file** (l'icona con la freccia in alto a destra).
2. Doppio clic sul file scaricato. Se Windows avvisa che l'autore non è verificabile,
   scegli **Ulteriori informazioni** → **Esegui comunque**.
3. Aspetta il messaggio **INSTALLAZIONE COMPLETATA** e chiudi la finestra nera.
4. Apri Revit: compaiono i tab **VisionXtools** e **AnasL4**.
   Se Revit era già aperto: tab **pyRevit** → **Reload**.

Il file controlla che pyRevit ci sia, scarica l'estensione e verifica che sia arrivata tutta.

## Installazione da riga di comando (alternativa)

1. Chiudi Revit.
2. Apri il **Prompt dei comandi** (tasto Windows → scrivi `cmd` → Invio).
3. Incolla questo comando e premi Invio:

   ```
   pyrevit extend ui VisionXtools https://github.com/VisionXt-tech/VisionXtools.extension.git --branch=main
   ```

   Se alla fine compare `Error: Failed to save config to "C:\ProgramData\pyRevit\pyRevit_config.ini" ... Access ... denied`,
   **ignoralo**: succede quando pyRevit è installato per tutti gli utenti, ma l'estensione è stata installata lo stesso.

4. Apri Revit: compaiono i tab **VisionXtools** e **AnasL4**.

Non serve installare Git: pyRevit scarica l'estensione da solo in
`%APPDATA%\pyRevit\Extensions\VisionXtools.extension`.

## Aggiornamento

In Revit: tab **VisionXtools** → pannello **About** → **Update**, oppure tab **pyRevit** → **Update**.

## Installazione manuale (se il comando `pyrevit` non funziona)

1. Su GitHub: **Code** → **Download ZIP**, poi estrai lo ZIP.
2. Rinomina la cartella estratta da `VisionXtools.extension-main` a **`VisionXtools.extension`**.
   Il nome deve finire con `.extension`, altrimenti pyRevit la ignora.
3. Spostala in una cartella fissa, per esempio `C:\pyRevitExtensions\VisionXtools.extension`.
4. In Revit: tab **pyRevit** → **Settings** → **Custom Extension Directories** → **Add folder** → scegli `C:\pyRevitExtensions` (la cartella che *contiene* l'estensione, non l'estensione stessa).
5. **Save Settings and Reload**.

Con questo metodo il pulsante **Update** non aggiorna l'estensione: per aggiornarla ripeti i passi 1–3.

## Se i tab non compaiono

- Tab **pyRevit** → **Reload**.
- Controlla che esista la cartella `%APPDATA%\pyRevit\Extensions\VisionXtools.extension`
  (incolla il percorso nella barra di Esplora file) e che contenga `VisionXtools.tab` e `AnasL4.tab`.
- Nell'installazione manuale, controlla di non avere una cartella dentro l'altra
  (`VisionXtools.extension\VisionXtools.extension-main\...`): i file `VisionXtools.tab` e `AnasL4.tab`
  devono stare direttamente dentro `VisionXtools.extension`.
- Non cambiare il motore Python in pyRevit: gli strumenti girano sul motore predefinito.

## Disinstallazione

Con Revit chiuso, cancella la cartella `%APPDATA%\pyRevit\Extensions\VisionXtools.extension`
(o quella scelta nell'installazione manuale).
