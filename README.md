# WinCC Unified List Converter

Webversion des Excel-basierten *Unified Migration Tool (v2)*. Die App bereitet Excel-Exporte aus dem TIA Portal für WinCC Unified auf und verwendet dabei exakt die Regeln der getesteten VBA-Makros. Die Verarbeitung läuft vollständig im Browser: Es gibt keinen Server, kein Konto und keine Übertragung der Dateien.

| Konverter | Ehemals (Excel-Button) | Eingabe | Ergebnis |
|---|---|---|---|
| HMI-Variablen | HMI TagList Konverter | `.xlsx` | `HMITagList_MIGUNI.xlsx` |
| Textlisten · Legacy | HMI TextList Konverter (alt) | `.xlsx`, `.xlsm` | `<Name>_MIGUNI.xlsx` |
| Textlisten · Standard | HMI TextList Konverter (neu) | `.xlsx`, `.xlsm` | `<Name>.xlsx` |
| HMI-Meldungen | HMI-Alarm Konverter | `.xlsx`, `.xlsm` | `Unified_HMI-Alarms.xlsx` |

Dokumentation:
- [docs/ANALYSE_UND_PLAN.md](docs/ANALYSE_UND_PLAN.md) – Analyse der VBA-Makros und alle Umwandlungsregeln
- [docs/ABWEICHUNGEN.md](docs/ABWEICHUNGEN.md) – bekannte Unterschiede zum Original und Teststand

## Lokal starten

Voraussetzung: Node.js 20 oder neuer.

```bash
npm install
npm run dev        # Entwicklungsserver auf http://localhost:5173
npm test           # automatisierte Tests
npm run build      # statischer Build nach dist/
npm run preview    # Build lokal ansehen
```

## Aufbau

```
src/xlsx/        XLSX lesen und Zellen direkt im XML ändern (Formatierung bleibt erhalten), neue Mappe schreiben
src/vba/         Nachbildung der VBA/Excel-Funktionen (Trim, Like, CStr, Text-Interpretation …)
src/converters/  die vier Konverter (je eine Datei, Kommentar nennt das VBA-Original)
src/main.ts      Oberfläche
tests/           Unit-Tests und Beispieldateien
```

Abhängigkeiten zur Laufzeit: nur [JSZip](https://stuk.github.io/jszip/). Die Bibliotheken SheetJS (Community Edition) und ExcelJS kamen nicht infrage, weil sie beim Speichern Formatierungen bzw. Teile der Datei verwerfen. Die Original-Makros verändern dagegen nur Zellwerte.

## Veröffentlichung über GitHub Pages

Die App ist eine statische Seite mit relativen Pfaden (`base: './'`). Sie funktioniert deshalb auch unter `https://<user>.github.io/<repo>/`.

```bash
npm run deploy     # baut die App und pusht dist/ in den Branch gh-pages
```

Einmalig in GitHub unter **Settings → Pages**: *Source: Deploy from a branch*, Branch `gh-pages`, Ordner `/ (root)` auswählen.

Für eine Veröffentlichung per GitHub Actions wird ein Token mit `workflow`-Berechtigung benötigt (`gh auth refresh -s workflow`).

## Vergleich mit dem Original

Damit sich die Webversion gegen das Original prüfen lässt, lege Paare aus Eingabedatei und Ausgabe des Excel-Tools in `tests/fixtures/` ab und ergänze sie in `tests/unit/fixtures.test.ts`. Wird `OUT_DIR=<ordner> npm test` gesetzt, speichern die Tests die erzeugten Dateien zum manuellen Vergleich.
