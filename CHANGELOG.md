# Changelog

Alle nennenswerten Änderungen an diesem Projekt werden hier dokumentiert.
Das Format orientiert sich an [Keep a Changelog](https://keepachangelog.com/de/1.1.0/),
die Versionierung folgt [Semantic Versioning](https://semver.org/lang/de/).

## [1.4.0] - 2026-11-02

### Neu

- **Variablen-Ereignisse vollautomatisch aus dem Openness-Export:** ZIP (manifest.json + Tags/*.xml) oder einzelne
  Variablen-XML hochladen – das Skript erscheint direkt auf der Seite und lässt sich mit **Skript kopieren** übernehmen.
  - Alle Ereignisse pro Variable (Wertänderung, Überschreitung, Unterschreitung) mit ihren kompletten Funktionslisten.
  - Umgesetzt: SetTag, LinearScaling, InvertLinearScaling, SetBit, ResetBit, InvertBit, IncreaseValue, DecreaseValue,
    ActivateScreen, ActivateScreenByNumber (mit Bildnummern-Tabelle); andere Funktionen als `// TODO` mit Hinweis.
  - Verhalten wie in Advanced: Funktionslisten laufen nur beim Ereignis; Grenzwerte kommen aus der Variable.
    Nach dem Start werden Berechnungen einmal ausgeführt, Bildwechsel nicht.
  - Innerhalb einer Funktionsliste wird ein gerade berechneter Wert direkt weiterverwendet, statt ihn zurückzulesen.
  - Variablennamen werden wie bei „HMI-Variablen“ ins Unified-Format übersetzt und gegen den Export geprüft.

## [1.3.0] - 2026-11-02

### Neu

- **Variablen-Ereignisse:** Ersatz für die Variablen-Ereignisse aus WinCC Advanced (Wertänderung, Bei Überschreitung,
  Bei Unterschreitung). Der Excel-Export des TIA Portals enthält diese Ereignisse nicht, deshalb in zwei Schritten:
  1. HMI-Variablen-Export hochladen → Vorlage `TagEvents.xlsx` mit allen Variablen; Variablen mit Grenzwert sind als
     Über-/Unterschreitung vorbelegt.
  2. Ereignisse eintragen (Trigger, Ereignis, Aktion, Ziel, Wert/Formel) und hochladen → Skript `TagEvent.js`.
  - Formeln mit `{Trigger}` und `{Variablenname}`, z. B. `10000 * {PlcNo} + {Trigger}`.
  - Grenzwerte werden aus dem Export übernommen und im Skript geprüft.
  - Bildwechsel und Skripte laufen nur, wenn das Ereignis tatsächlich eintritt – nicht bei jedem Zyklus.
  - Alle Variablennamen werden gegen den Export geprüft; Fehler werden mit Zelle gemeldet.

## [1.2.0] - 2026-11-02

### Geändert

- **Oberfläche verfeinert** (gleicher Aufbau, gleiche Farben):
  - Funktionskarten breiter und besser lesbar, ausgewählte Karte mit roter Tönung, Rahmen und Häkchen.
  - Upload mit klaren Zuständen: Ziehen, geladen (Name, Typ, Größe), ungültig (Meldung direkt am Upload-Bereich),
    Datei ersetzen oder entfernen.
  - Konvertieren-Button mit Symbol und eindeutigem Lade- und Gesperrt-Zustand.
  - Ergebnisbereich mit Ergebnisdatei, Anzahl der Hinweise und Download.
  - Anleitung mit besser lesbaren Tabs, einheitlichen Tastenkappen und klar abgesetztem Regelkasten.
  - Sichtbare Tastatur-Fokusrahmen, keine Schrift unter 12 px, geprüft bei 1440, 1024, 768 und 390 px Breite.
- Der Release-Workflow veröffentlicht die Webseite jetzt automatisch mit, damit Release und Live-Seite nicht auseinanderlaufen.
- **Textlisten · Klassisch** und **Textlisten · Aktuell** sind zu einer Funktion **Textlisten** zusammengelegt.
  Sie erkennt am Exportformat selbst, was zu tun ist: Mit Spalte `FieldInfos` werden Namen bereinigt und die
  FieldInfos ins Unified-Format umgewandelt, ohne diese Spalte läuft die klassische Platzhalter-Bereinigung.

### Behoben

- Die klassische Bereinigung hat bei Exporten mit weniger als sechs Sprachen die Spalte `FieldInfos` überschrieben.
  Sie bricht jetzt mit einer Erklärung ab.
- Wird für eine Datei die falsche Funktion gewählt (z. B. eine Textliste bei „HMI-Meldungen“), meldet der Converter
  das und nennt die passende Funktion, statt ein unbrauchbares Ergebnis zu erzeugen.

## [1.1.0] - 2026-11-02

### Neu

- **Textlisten · Aktuell** wandelt jetzt auch die Spalte `FieldInfos` ins Unified-Format um
  (`CommonTextList` → `ILCommonTagDisplayFormatWithTextList`, `CommonTagDisplayFormat` → `ILCommonTagDisplayFormat`).
  Variablennamen werden dabei wie bei den HMI-Variablen bereinigt. Damit entspricht ein Textlisten-Export
  aus dem klassischen Projekt dem Format eines Unified-Textlisten-Exports.

### Behoben

- Exportdateien direkt aus dem TIA Portal (XML mit UTF-8-BOM) wurden als „beschädigt“ abgelehnt.

## [1.0.0] - 2026-10-19

Erste stabile Version des WinCC Unified List Converters – eine Webversion des
Unified Migration Tools, die vollständig im Browser läuft.

### Funktionen

- **HMI-Variablen** – ersetzt Sonderzeichen in Variablennamen durch `_` (`HMITagList_MIGUNI.xlsx`)
- **Textlisten · Klassisch** – bereinigt Platzhalter (`<No value>`, `<field ref="0" />`) und Tags, füllt leere Sprachvarianten
- **Textlisten · Aktuell** – ersetzt `# / .` und `[ ]` in den Namen von Textlisten und Einträgen
- **HMI-Meldungen** – erzeugt die Unified-Alarmliste `DiscreteAlarms` (44 Spalten) inkl. neuer FieldInfo
- **Meldetexte umbrechen** – teilt Meldetexte in Fehlercode und Zeilen à max. 80 Zeichen (`¶`)

### Highlights

- 100 % clientseitig: kein Server, kein Upload, kein Tracking
- Formattreue Bearbeitung direkt auf XML-Ebene – nur betroffene Zellen werden geändert
- Hinweise mit Zellangabe bei Werten, die Excel anders interpretieren könnte
- Integrierte Anleitung zu jeder Funktion
- Responsives Design im hellen und dunklen Modus (PRH Design Kit)
- Automatisierte Tests für alle Umwandlungsregeln

[1.4.0]: https://github.com/paul-rafael/wincc-unified-list-converter/releases/tag/v1.4.0
[1.3.0]: https://github.com/paul-rafael/wincc-unified-list-converter/releases/tag/v1.3.0
[1.2.0]: https://github.com/paul-rafael/wincc-unified-list-converter/releases/tag/v1.2.0
[1.1.0]: https://github.com/paul-rafael/wincc-unified-list-converter/releases/tag/v1.1.0
[1.0.0]: https://github.com/paul-rafael/wincc-unified-list-converter/releases/tag/v1.0.0
