# Changelog

Alle nennenswerten Änderungen an diesem Projekt werden hier dokumentiert.
Das Format orientiert sich an [Keep a Changelog](https://keepachangelog.com/de/1.1.0/),
die Versionierung folgt [Semantic Versioning](https://semver.org/lang/de/).

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

[1.0.0]: https://github.com/paul-rafael/wincc-unified-list-converter/releases/tag/v1.0.0
