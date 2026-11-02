# Changelog

Alle nennenswerten Änderungen an diesem Projekt werden hier dokumentiert.
Das Format orientiert sich an [Keep a Changelog](https://keepachangelog.com/de/1.1.0/),
die Versionierung folgt [Semantic Versioning](https://semver.org/lang/de/).

## [Unreleased]

### Geändert

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

[1.1.0]: https://github.com/paul-rafael/wincc-unified-list-converter/releases/tag/v1.1.0
[1.0.0]: https://github.com/paul-rafael/wincc-unified-list-converter/releases/tag/v1.0.0
