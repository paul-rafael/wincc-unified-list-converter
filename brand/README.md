# PRH Design Kit

Gemeinsamer Stil für alle Projekte von Paul-Rafael Hasan-Man: elegant und zurückhaltend, Schwarz/Weiß mit dunkelroten Akzenten. Heller und dunkler Modus funktionieren automatisch.

| Datei | Zweck |
|---|---|
| `DESIGN.md` | **Design-Richtlinie** – alle Regeln und Werte in einer Datei, zum Übertragen auf andere Projekte |
| `brand.css` | Farben, Schrift und Bausteine (Klassen mit Präfix `pr-`) |
| `logo.svg` | Signatur als Vektorgrafik, Farbe folgt dem Text (`currentColor`) |
| `favicon.svg` | Signatur für den Browser-Tab (schwarz im hellen, weiß im dunklen Modus) |
| `template.html` | Startvorlage für ein neues Projekt |
| `preview.html` | Übersicht aller Farben und Bausteine (im Browser öffnen) |

## In einem neuen Projekt verwenden

1. Den Ordner `brand/` in das Projekt kopieren.
2. `template.html` als `index.html` übernehmen – oder diese beiden Zeilen in den `<head>` einer bestehenden Seite einfügen:

```html
<link rel="icon" type="image/svg+xml" href="brand/favicon.svg" />
<link rel="stylesheet" href="brand/brand.css" />
```

3. Die Signatur in die Fußzeile setzen:

```html
<a class="pr-signature" href="https://github.com/paul-rafael">
  <span class="pr-logo" aria-hidden="true"></span>
  Paul-Rafael Hasan-Man
</a>
```

## Farben

Immer über die Variablen verwenden, nie Farbwerte direkt eintragen – dann stimmen heller und dunkler Modus automatisch.

| Variable | Verwendung |
|---|---|
| `--pr-bg`, `--pr-surface`, `--pr-surface-2` | Hintergrund, Karten, leicht abgesetzte Flächen |
| `--pr-text`, `--pr-text-2`, `--pr-text-3` | Text: Haupttext, Fließtext, Nebeninfos |
| `--pr-border`, `--pr-border-strong` | Rahmen und Trennlinien |
| `--pr-red`, `--pr-red-hover` | Hauptaktionen (Buttons), Auswahl |
| `--pr-red-text`, `--pr-red-soft` | dunkelrote Schrift, zarter roter Hintergrund |
| `--pr-ink`, `--pr-ink-soft` | Zweitakzent Schwarz (im Dunkelmodus Weiß): Nummern, Symbole, Info-Hinweise, Zweit-Buttons |

**Faustregel:** Dunkelrot für das, was der Nutzer *tun* soll, Schwarz für das, was er *wissen* soll. Beides sparsam einsetzen – der Großteil der Seite bleibt Schwarz/Weiß.

## Bausteine

| Klasse | Element |
|---|---|
| `pr-header`, `pr-brand`, `pr-nav` | Kopfzeile mit Logo und Navigation |
| `pr-hero`, `pr-eyebrow`, `pr-title`, `pr-lead` | Überschriftenbereich |
| `pr-btn` + `pr-btn-primary` / `-secondary` / `-ghost` | Buttons |
| `pr-panel`, `pr-card`, `pr-card-select` | Flächen und auswählbare Karten |
| `pr-note`, `pr-note-accent`, `pr-alert`, `pr-success` | Hinweise und Meldungen |
| `pr-tabs`, `pr-tab` | Reiter |
| `pr-input`, `pr-select`, `pr-textarea`, `pr-dropzone` | Formulare und Datei-Upload |
| `pr-table`, `pr-chip`, `pr-badge`, `pr-number`, `pr-icon-box` | Tabellen und kleine Elemente |
| `pr-footer`, `pr-signature`, `pr-logo` | Fußzeile mit Signatur |

Alle Beispiele siehe `preview.html`.

## Dunkelmodus erzwingen

Standardmäßig folgt die Seite der Systemeinstellung. Fest einstellen mit:

```html
<html lang="de" data-theme="dark">
```
