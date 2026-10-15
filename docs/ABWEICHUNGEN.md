# Abweichungen zum Excel-Original & offene Prüfpunkte

Referenz: `Unified_Migration_Tool.xlsm` (VBA-Module 1, 2, 3, 7, 8 + Hilfsfunktionen). Details der Regeln: [ANALYSE_UND_PLAN.md](ANALYSE_UND_PLAN.md).

## Umfang

| Original-Button | Webversion | Portiert aus |
|---|---|---|
| HMI TagList Konverter (ALT) | **HMI-Variablen** | `Modul1.SonderzeichenMitKnopf` |
| HMI TextList Konverter (alt) | **Textlisten · Klassisch** | `Modul2.Clean_TextList_Sheet` |
| HMI TextList Konverter (neu) | **Textlisten · Aktuell** | `Modul3.ReplaceSpecialCharsInWorkbook` |
| HMI-Alarm Konverter | **HMI-Meldungen** | `Modul8.ConvertAlarmList_Unified` (+ `ExtractRefs`, `ExtractTagByRef`, `ExtractLengthByRef`) |
| BETA AZU | **Meldetexte umbrechen** (Beta) | `Modul7.FormatHMITextInFile` / `FormatHMITextWithErrorCode` – im Original war der Button versehentlich mit Modul1 verknüpft |

Nicht portiert (kein Button, unvollständig): `Excel_Bereinigen` (Modul4/5), `ConvertAlarmLists2` (Modul9–11).

## Bewusste Abweichungen

| Thema | Original | Webversion | Grund |
|---|---|---|---|
| `.xls` | Dateifilter erlaubt `.xls` (Modul3, Modul8) | wird mit Hinweis abgelehnt | Binärformat; TIA Portal exportiert `.xlsx` |
| Fehlerwerte (`#NV` …) in zu lesenden Zellen | Laufzeitfehler 13, Abbruch | Abbruch mit Meldung inkl. Zelle | gleiche Wirkung, verständlichere Meldung |
| Fehlendes Trennzeichen in FieldInfo / `ref="` | Laufzeitfehler 5, Abbruch | Abbruch mit Meldung | dito |
| Blattschutz (HMI-Variablen, Textlisten · Klassisch, Meldetexte umbrechen) | Laufzeitfehler 1004 | Abbruch mit Meldung | dito |
| Blattschutz (Textlisten · Aktuell) | Fehlerhandler meldet „Blatt nicht gefunden“ und macht weiter | Hinweis „schreibgeschützt, übersprungen“ und weiter | gleiche Wirkung, korrekte Meldung |
| Gemeinsame Formeln/Matrixformeln in Zielzellen | werden überschrieben | Abbruch mit Meldung | würde die Datei ohne Excel-Neuberechnung beschädigen |
| `.xlsm` bei Textlisten · Klassisch | Speichern-Vorschlag `Name.xlsm` als `.xlsx`-Format | `Name_MIGUNI.xlsx` | Original-Vorschlag führt zu Formatkonflikt |
| Abbruch-/Speicherdialoge | Excel-Dialoge | Download-Schaltfläche | nur Oberfläche |
| Meldetexte umbrechen: Speichern | Makro speichert nicht, Mappe bleibt offen | Download unter gleichem Namen und Format (`.xlsm` bleibt mit Makros erhalten) | entspricht „Speichern“ durch den Benutzer |
| Meldetexte umbrechen: `Rows.AutoFit` | Excel berechnet die Höhe sofort | gespeicherte Zeilenhöhe wird entfernt, Excel berechnet sie beim Öffnen | Höhe hängt von Schrift/Spaltenbreite ab; in Excel geprüft |

## Nicht zweifelsfrei nachbildbares Excel-Verhalten

Excel interpretiert Text, der per VBA in eine Zelle geschrieben wird (`Range.Value = "…"`), wie eine Benutzereingabe.

- **Umgesetzt:** reine Zahlen (`123`, `007`, `1E5`, `1.5`) werden zu Zahlen – außer in Zellen mit Textformat (`@`). Leerer Text erzeugt eine leere Zelle.
- **Als Text belassen und als Hinweis gemeldet:** Text mit `=` am Anfang, Prozentwerte, Zahlen mit Komma, Datums-/Uhrzeit-Muster, `(5)`, `TRUE/FALSE/WAHR/FALSCH`.
- **Textlisten · Klassisch:** Zellen mit Kommazahlen, Datumswerten oder Wahrheitswerten werden nicht verändert (Hinweis), weil `CStr` + Rückschreiben in deutschem Excel gebietsschemaabhängig ist.
- **HMI-Variablen:** Zahlen/Datumswerte in Spalte A werden wie `CStr` in de-DE umgewandelt (`1,5` → `1_5`, Datum → `01_05_2025`), Wahrheitswerte als `Wahr`/`Falsch` (Hinweis).
- **HMI-Meldungen:** Datumswerte in A–G werden als Zahl übernommen (Excel würde ein Datumsformat setzen; Hinweis).
- **Textlisten · Aktuell:** Formel- und Fehlerzellen werden nicht verändert (Hinweis); Zahlen bleiben unverändert (in deutschem Excel enthält der Formeltext einer Zahl kein `.`).

## Technische Unterschiede ohne fachliche Auswirkung

- Ausgabe ist wertgleich, nicht byte-identisch: Nur geänderte Zellen werden neu geschrieben, alle anderen Teile der Datei bleiben unverändert. Excel würde beim Speichern die ganze Datei neu erzeugen.
- Neue Texte werden in die vorhandene Shared-String-Tabelle angehängt; nicht mehr benutzte Einträge bleiben stehen (Excel räumt sie beim nächsten Speichern auf).
- Wird eine Formel überschrieben, wird `calcChain.xml` entfernt; Excel baut sie beim Öffnen neu auf.
- Neue Alarm-Mappe: Standardschrift „Aptos Narrow 11“, keine Theme-Datei.

## Durchgeführte Prüfungen (Stand 15.10.2026)

- 35 automatisierte Tests (`npm test`) für alle Regeln und Sonderfälle aus der VBA-Analyse – bestanden.
- Beispieldateien, die unabhängig mit openpyxl erzeugt wurden (`tests/fixtures`), wurden konvertiert; die Ergebnisse wurden mit openpyxl gelesen und in **Microsoft Excel für Mac** ohne Reparaturmeldung geöffnet (bei *Meldetexte umbrechen* inkl. Prüfung von Zeilenumbruch und automatischer Zeilenhöhe).
- Alle fünf Funktionen im Browser getestet (Auswahl, Drag & Drop, Konvertierung, Download, Fehlermeldung bei falschem Dateityp, Handy-Breite).

## Noch offen

- **Vergleich mit echten Original-Ausgaben:** Bitte je Konverter einen echten TIA-Export und die vom Original-Tool (Windows, deutsches Excel) erzeugte Datei bereitstellen. Erst damit lässt sich das Verhalten von Excel bei der Text-Interpretation (siehe oben) endgültig bestätigen.
- Import der Ergebnisse in TIA Portal / WinCC Unified wurde nicht getestet.
