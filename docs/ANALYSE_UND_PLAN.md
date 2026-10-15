# WinCC Unified List Converter – Analysebericht & Implementierungsplan

Stand: 15.10.2026 · Phase A (Analyse) und Phase B (Entwurf) · **noch keine Implementierung**

Quellen (unverändert, nur Kopien analysiert):
- `~/Downloads/Unified_Migration_Tool.xlsm` – VBA-Projekt mit 11 Modulen, 2 UserForms, 1 Arbeitsblatt
- `~/Downloads/WinCC_Unified_Panels_Migration.docx` – Migrationsanleitung (Stand 05.05.2025, Autor HAPL), beschreibt die Tools in Kap. „Migrations-Hilfetool“

---

## 1. Bestandsaufnahme

### 1.1 Arbeitsmappe

| Element | Befund |
|---|---|
| Blätter | nur `Tabelle1` (UI-Blatt, geschützt, Titel „WinCC Unified Panel Migration Hilfstool by HAPL“, Hinweistext, Logo) |
| Daten im Tool selbst | keine – alle Makros öffnen eine **externe** Datei |
| Schaltflächen | 5 Formular-Buttons (siehe 1.2) |
| UserForm1, UserForm2 | leer (kein Code, keine Verwendung) |
| `DieseArbeitsmappe`, `Tabelle1`, `Modul6` | leer |
| Externe Referenzen | `CreateObject("VBScript.RegExp")` (Modul2, nur unter Windows verfügbar) |
| Definierter Name `_xleta.N` | XLM-Artefakt (`#NAME?`), ohne Funktion |

### 1.2 Schaltflächen → Makros (aus `vmlDrawing1.vml`)

| Button-Beschriftung | Makro | Modul |
|---|---|---|
| HMI TagList Konverter (ALT) | `SonderzeichenMitKnopf` | Modul1 |
| HMI TextList Konverter (alt) | `Clean_TextList_File_SelectAndSave` → `Clean_TextList_Sheet` | Modul2 |
| HMI TextList Konverter (neu) | `Button_DateiAuswaehlen_Bearbeiten_Speichern` → `ReplaceSpecialCharsInWorkbook` → `ReplaceInSheetColumn` | Modul3 |
| HMI-Alarm Konverter | `ConvertAlarmList_Unified` → `ExtractRefs`, `ExtractTagByRef`, `ExtractLengthByRef` | Modul8 |
| BETA AZU | **`SonderzeichenMitKnopf`** (identisch mit TagList-Konverter!) | Modul1 |

### 1.3 Makros ohne Button (nicht über die Oberfläche erreichbar)

| Makro | Modul | Zustand |
|---|---|---|
| `FormatHMITextInFile` / `FormatHMITextWithErrorCode` | Modul7 | inhaltlich die in der Doku beschriebene **BETA-AZU**-Funktion (Zeilenumbruch mit `¶`). Kein Button verknüpft, **speichert die Datei nicht** (lässt sie geöffnet). |
| `ConvertAlarmLists2` + `ConvertAllRefs` + `ExtractRefID/Tag/DisplayType/Length` | Modul9/10/11 | Alternative Alarm-Konvertierung mit zwei Eingabedateien (alt + bereits migrierte Unified-Liste). Kein Button. |
| `Excel_Bereinigen` | Modul4 (+ `GetValueBetween` Modul5) | ersetzt `<field ref="n" />` durch `<Variable : Länge. Tag>`; zeigt am Ende nur einen Speichern-Dialog, **speichert aber nicht**. Kein Button. |

→ Diese gelten nicht als „getestete Referenz“ und werden **ohne deine Freigabe nicht portiert** (siehe Frage 1 in Abschnitt 6).

---

## 2. Konvertierungsregeln (VBA ist maßgeblich)

Allgemein: Die drei Listen-Konverter (TagList, TextList alt/neu) **ändern die geöffnete Originaldatei und speichern sie unter neuem Namen**. Alles, was sie nicht anfassen (andere Blätter, Formatierung, Spaltenbreiten, Validierungen …), bleibt erhalten. Der Alarm-Konverter erzeugt dagegen eine **neue** Arbeitsmappe.

„Letzte Zeile“ = `Cells(Rows.Count, Spalte).End(xlUp).Row`: letzte nicht-leere Zelle der genannten Spalte (Formeln mit Ergebnis `""` zählen als nicht leer, reine Formatierung nicht). Ist die Spalte leer, ergibt sich Zeile 1.

### 2.1 HMI TagList Konverter – `SonderzeichenMitKnopf`

| Aspekt | Regel |
|---|---|
| Eingabe | nur `*.xlsx` |
| Blatt | **erstes** Blatt (`Sheets(1)`, Name egal; TIA-Export: „Hmi Tags“) |
| Bereich | Spalte A von **A1** (Kopfzeile wird mitverarbeitet!) bis letzte Zeile in A, inkl. leerer Zellen dazwischen |
| Regel | jedes Zeichen (UTF-16-Codeeinheit), das nicht `[A-Za-z0-9]` ist (`Like`, binärer Vergleich), wird durch `_` ersetzt. Umlaute, ß, Leerzeichen, `-`, `.`, `[`, `]` usw. → `_`. Länge bleibt gleich. |
| Zahlen/Datum in A | werden per `CStr` (Gebietsschema de-DE) in Text gewandelt, dann ersetzt (z. B. 1,5 → `1_5`) |
| Formeln in A | werden durch den (ersetzten) Ergebniswert überschrieben |
| Fehlerwerte in A (`#N/A` …) | VBA bricht mit Laufzeitfehler 13 ab → Web: verständliche Fehlermeldung, kein Ergebnis |
| Rückschreiben | `cell.Value = temp` → Excel interpretiert den String wie eine Eingabe (siehe Risiko R1) |
| Ausgabe | gleiche Arbeitsmappe, Dateiname-Vorschlag `HMITagList_MIGUNI.xlsx` |

### 2.2 HMI TextList Konverter (alt) – `Clean_TextList_Sheet`

| Aspekt | Regel |
|---|---|
| Eingabe | `*.xlsx; *.xlsm` |
| Blatt | `TextListEntry`; fehlt es → Meldung „Arbeitsblatt 'TextListEntry' nicht gefunden!“, Abbruch ohne Speichern |
| Zeilen | 2 … letzte Zeile in **Spalte E** |
| Spalten | E–J (5–10) werden bearbeitet; I (9) = Fallback-Text, K (11) = Feldinfo |

Pro Zeile (vor der Spaltenschleife):
- `fallbackText = Trim(I)`, `kText = Trim(K)` (VBA-`Trim` entfernt nur Leerzeichen U+0020)
- Wenn `kText` nicht leer: Regex `TextList\s*=\s*([^;>]+)` (ignoriere Groß/Klein), erster Treffer, Gruppe 1 → `Trim` → `textListVal`; sonst `""`

Pro Zelle E…J (`txt = Trim(Zelle)`), Reihenfolge der Prüfungen:
1. **Nur Spalte E:** wenn `txt` (klein) weder `<no value>` noch `<field ref="0" />` ist → Zelle **unverändert lassen** (kein Rückschreiben).
2. `txt` = `<field ref="0" />` (ignoriere Groß/Klein, exakt): → `textListVal`, falls leer → `—` (Geviertstrich U+2014)
3. `txt` = `<no value>` (ignoriere Groß/Klein, exakt):
   - K leer → `fallbackText`, sofern dieser nicht leer und nicht `<no value>` ist; sonst `—`
   - K nicht leer → `textListVal`, falls leer → `—`
4. Sonst: alle `<…>`-Tags per Regex `<[^>]+>` entfernen, danach `WorksheetFunction.Trim` (führende/abschließende Leerzeichen weg, innere Mehrfach-Leerzeichen → eins; nur U+0020). Ergebnis leer → `—`. **Wird immer zurückgeschrieben**, d. h. auch leere Zellen in F–J werden zu `—`, und unveränderte Texte werden getrimmt.

Ausgabe: gleiche Arbeitsmappe als `.xlsx`, Name-Vorschlag `<Quelle>.xlsx` → `<Quelle>_MIGUNI.xlsx`.

### 2.3 HMI TextList Konverter (neu) – `ReplaceSpecialCharsInWorkbook`

| Blatt | Spalte | ab Zeile | bis |
|---|---|---|---|
| `TextList` | A | 2 | letzte Zeile in A |
| `TextListEntry` | B | 2 | letzte Zeile in B |

Ersetzungen (`Range.Replace`, Teilstring, in dieser Reihenfolge): `#`→`_`, `/`→`_`, `.`→`_`, `[`→`(`, `]`→`)`.
- Fehlt ein Blatt: Hinweis „Blatt '…' wurde nicht gefunden und wurde übersprungen.“, Verarbeitung läuft weiter (Datei wird trotzdem gespeichert).
- Ist die letzte Zeile < 2: nichts tun.
- Eingabe: `*.xlsx; *.xlsm; *.xls`; Ausgabe als `.xlsx`, Name-Vorschlag = Name der Quelldatei.

### 2.4 HMI-Alarm Konverter – `ConvertAlarmList_Unified`

Eingabe: `*.xls*`, erstes Blatt. Ausgabe: **neue** Arbeitsmappe, ein Blatt `DiscreteAlarms`, Name-Vorschlag `Unified_HMI-Alarms.xlsx`.

Zeile 1 = feste Kopfzeile mit 44 Spalten (A … AR): `ID`, `Name`, `Alarm text [de-DE], Alarm text 1`, `FieldInfo [Alarm text 1]`, `Class`, `Trigger tag`, `Trigger bit`, `Trigger mode`, `Acknowledgement tag`, `Acknowledgement bit`, `PLC acknowledgement tag`, `PLC acknowledgement bit`, `Priority`, `Info text [de-DE], Info text`, `Additional text 1 [de-DE], Alarm text 2`, `FieldInfo [Alarm text 2]`, … `Additional text 9 [de-DE], Alarm text 10`, `FieldInfo [Alarm text 10]`, `Alarm parameter 1` … `Alarm parameter 10`, `Area`, `Origin`.

Für jede Zeile 2 … letzte Zeile in **Spalte A der Quelle** (auch leere Zwischenzeilen):

| Spalte(n) | Endwert |
|---|---|
| A–C, E–G (1–3, 5–7) | Werte aus der Quelle (gleiche Spalte) |
| D (4) | neue FieldInfo (siehe unten), sonst leer |
| H (8) | `On rising edge` |
| I (9), K (11) | `<No value>` |
| J, L, M (10, 12, 13) | Zahl `0` |
| N, P, R, T, V, X, Z, AB, AD, AF (14, 16, …, 32) | `<No value>` |
| O, Q, S, U, W, Y, AA, AC, AE (15, 17, …, 31) | leer |
| AG (33) „Alarm parameter 1“ | Tag des **letzten** gefundenen Feldes, sonst leer |
| AH–AP (34–42) | `<No value>` |
| AQ (43) „Area“ | `HMI_RT_2::Alarming` |
| AR (44) „Origin“ | leer |

Hinweis: Der Code setzt die Spalten 16–32 (gerade) zuerst auf leer und überschreibt sie danach in der Schleife `14 To 33 Step 2` mit `<No value>`. Laut Kopfzeile sind das die „FieldInfo“-Spalten, während die „Additional text“-Spalten leer bleiben. Da das Tool getestet ist, wird **genau dieses Endergebnis** übernommen (keine Korrektur).

FieldInfo-Verarbeitung (alter Alarmtext C, alte FieldInfo D):
1. `ExtractRefs(C)`: alle Vorkommen von `ref="` (Groß/Klein beachtet); ID = Text bis zum nächsten `"`. Fehlt das schließende `"` → VBA-Laufzeitfehler 5.
2. Pro ID: `Tag` = Text nach dem ersten `Tag = ` **ab** der ersten Fundstelle von `id = <ID>` bis zum nächsten `;`, getrimmt.
   - Teilstring-Suche: `id = 1` passt auch auf `id = 10` (wird originalgetreu übernommen).
   - Hat der Block selbst kein `Tag = `, wird das Tag eines späteren Blocks gefunden (originalgetreu).
   - `Tag = ` ohne folgendes `;` → VBA-Laufzeitfehler 5.
3. Nur wenn Tag ≠ "": `Length` analog (`Length = … ;`), Fallback `3` wenn `id = …` oder `Length = ` nicht gefunden. Dann anhängen:
   `<ref id = {ID}; type = AlarmParameterWithOrWithoutCommonTextList; Parameter = Parameter 1; Tag = {Tag}; DisplayType = Decimal; Length = {Length}; Precision = 0; Alignment = Right; ZeroPadding = False;>` + CR LF
4. D = `Trim(Ergebnis)` – VBA-`Trim` entfernt **keine** Zeilenumbrüche, d. h. der Wert endet mit CR LF (originalgetreu).

### 2.5 Nur UI-bezogen (wird durch die Web-Oberfläche ersetzt)

Datei-/Speicherdialoge, `MsgBox`, `ScreenUpdating/EnableEvents/DisplayAlerts`, Blattschutz und Buttons auf `Tabelle1`, Abbruch-Prüfung `= "Falsch"` (Hinweis: zeigt, dass das Tool in **deutschem Excel** verwendet wird – relevant für R1/R3).

---

## 3. VBA ↔ JavaScript: Unterschiede, die nachgebildet werden müssen

| VBA/Excel | Verhalten | Umsetzung im Web |
|---|---|---|
| `Trim` | nur U+0020 | eigene Funktion, nicht `String.trim()` |
| `WorksheetFunction.Trim` | U+0020, innen zusammenfassen | eigene Funktion |
| `Like "[A-Za-z0-9]"` | binär, pro UTF-16-Einheit | Code-Unit-Schleife |
| `VBScript.RegExp` `\s` | nur `[ \t\n\r\f\v]` | `\s` explizit ersetzen (JS-`\s` matcht auch NBSP u. a.) |
| `InStr` ohne Vergleichsangabe | Groß/Klein beachtet | `indexOf` |
| `LCase` | Unicode | `toLowerCase()` (ohne Locale) |
| `CStr(Zahl/Datum)` | Gebietsschema de-DE | nachgebildet für Zahl/Datum/Boolean, dokumentiert |
| `Range.Value = String` | Excel interpretiert Text wie Eingabe | siehe R1 |
| leere Zelle `.Value` | `Empty` → `""` | Reader liefert `""` |
| Fehlerwert `.Value` + String-Operation | Laufzeitfehler 13 | Abbruch mit Meldung (Zeile/Spalte) |
| Laufzeitfehler 5 (`Mid` mit negativer Länge) | Abbruch | Abbruch mit Meldung statt stiller Korrektur |

---

## 4. Risiken & offene Punkte

- **R1 – Typumwandlung beim Rückschreiben (höchstes Risiko).** `cell.Value = "…"` lässt Excel den Text interpretieren: `"123"` → Zahl, `"1E5"` → 100000, `"(5)"` → −5, evtl. Datum/Uhrzeit/Wahrheitswert, `=`-Anfang → Formel – außer die Zelle ist als Text (`@`) formatiert. Das tritt bei TagList, TextList alt und Alarm (kopierte Werte A–G, Tag in AG) auf. Plan: eigenes Modul `excelCoerce`, das das Excel-Verhalten für die realistisch vorkommenden Fälle nachbildet (ganze Zahlen, Dezimalzahlen, Exponent, Text-Format-Zellen); **Feinabstimmung nur mit Referenzdateien** aus dem Original-Tool. Nicht belegbare Fälle werden dokumentiert und nicht geraten.
- **R2 – `Range.Replace` (TextList neu).** Arbeitet auf dem Formel-/Eingabetext der Zelle und interpretiert das Ergebnis neu (z. B. `[5]` → `(5)` → −5). Formeln würden im Formeltext ersetzt. Für reine Text-Namen unkritisch; Randfälle werden dokumentiert.
- **R3 – Gebietsschema.** Zahl→Text erfolgt in de-DE (Komma). Ich nehme de-DE als Referenz an.
- **R4 – Formaterhalt.** SheetJS (Community) verliert beim Schreiben Formatierung, ExcelJS verwirft Teile (z. B. VML, Erweiterungen). Deshalb: **direktes Bearbeiten der OOXML-Dateien im ZIP** – nur die betroffenen `<c>`-Knoten werden geändert, alles andere bleibt byte-identisch.
- **R5 – `.xls` (BIFF8).** Wird von TIA Portal nicht exportiert; Formaterhalt ohne Excel nicht möglich. Vorschlag: `.xls` mit klarer Meldung ablehnen (abweichend von den Dateifiltern in Modul3/8, dokumentiert).
- **R6 – `.xlsm`-Eingabe.** VBA speichert als `.xlsx` (Makros werden entfernt). Web entfernt `vbaProject.bin` und passt Content-Types an.
- **R7 – Referenztests.** Das Original kann hier nicht automatisiert laufen: Modul2 benötigt `VBScript.RegExp` (gibt es auf dem Mac nicht), die Dialoge sind interaktiv. Ohne Paare aus Eingabe- und Original-Ausgabedatei kann Phase D nur gegen die dokumentierten Regeln testen, nicht gegen das Original.
- **Nicht vorhanden:** keine Beispiel-Exportdateien (HMI-Tags, Textlisten, Alarme), keine Referenzausgaben, keine Info, welche TIA-Version die Exporte erzeugt.

---

## 5. Technischer Entwurf & Implementierungsplan

### Architektur
- **Vite + TypeScript**, statischer Build nach `dist/` mit `base: './'` (funktioniert unter `https://<user>.github.io/<repo>/`)
- **JSZip** (ZIP lesen/schreiben) + Browser-`DOMParser`/`XMLSerializer` (in Tests: `@xmldom/xmldom`); keine weiteren Laufzeit-Abhängigkeiten, kein Backend
- **Vitest** für Unit- und Vergleichstests; GitHub-Actions-Workflow für Pages (Veröffentlichung erst nach deiner Freigabe)

```
src/
  xlsx/       Workbook-Reader (Shared Strings, Rich Text, Inline-Strings, Zahlen, Bool, Fehler, Formeln),
              Zell-Patcher (in-place), Minimal-Writer (neue Mappe), lastRowInColumn (End(xlUp))
  vba/        vbaTrim, worksheetTrim, likeAlnum, vbsRegex, cstrDe, excelCoerce
  converters/ tagList.ts, textListOld.ts, textListNew.ts, alarmList.ts   (reine Funktionen: Workbook → Ergebnis + Hinweise)
  ui/         Konverter-Auswahl (4 Typen wie die Buttons), Drag & Drop, Status, Fehler, Download, „Weitere Datei“
tests/
  unit/       eine Testgruppe je VBA-Regel/Sonderfall aus Abschnitt 2
  golden/     Eingabe + Original-Ausgabe → Webausgabe vergleichen
scripts/compare-xlsx.ts   Vergleich: Blätter, Reihenfolge, Zellwerte, Typen, Formate betroffener Zellen
```

### Phasen
1. **B/C – Kern:** XLSX-Reader/Patcher/Writer + VBA-Hilfsfunktionen mit Unit-Tests
2. **C – Konverter:** die 4 Button-Funktionen 1:1 nach Abschnitt 2, inkl. Fehlerfällen
3. **C – Oberfläche:** eine Seite, Auswahl des Konvertertyps, Upload/Drag & Drop, Konvertieren, Status, Hinweise (z. B. „Blatt übersprungen“), Download mit dem Originaldateinamen-Vorschlag
4. **D – Validierung:** Vergleichsskript gegen deine Referenzdateien; jede Abweichung analysieren, Webversion anpassen; verbleibende Unterschiede in `docs/ABWEICHUNGEN.md`
5. **E – Deployment:** Pages-Workflow, README (Installation, `npm run dev`, Tests, Veröffentlichung), Prüfung des Builds unter Unterpfad

### Bekannte, unvermeidbare Unterschiede
- Ausgabe ist wertgleich, aber nicht byte-identisch (Excel schreibt beim Speichern die gesamte Datei neu).
- Neue Alarm-Mappe: Standardschrift/Theme der Excel-Version, mit der das Original lief, kann abweichen (für den Import irrelevant).

---

## 6. Offene Entscheidungen (benötige ich von dir)

1. **Umfang:** Nur die 4 per Button erreichbaren Konverter portieren (Empfehlung)? Der Button „BETA AZU“ ruft aktuell den TagList-Konverter auf – vermutlich ein Verknüpfungsfehler; die eigentliche AZU-Logik (Modul7) ist ungetestet und speichert nicht. Soll AZU weggelassen oder als „Beta“ nach Modul7 umgesetzt werden?
2. **Referenzdateien:** Bitte je Konverter mindestens eine echte Eingabedatei und die vom Original-Tool (Windows, deutsches Excel) erzeugte Ausgabe bereitstellen – ideal mit Randfällen (Umlaute, Zahlen als Namen, leere Zellen, mehrere Feldreferenzen).
