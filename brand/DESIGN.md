# Design-Richtlinie · PRH Design Kit

**Version 1.0 · Paul-Rafael Hasan-Man**

Diese Datei beschreibt das einheitliche Design aller meiner Projekte. Sie ist die verbindliche Referenz: Wer (Mensch oder KI-Assistent) an einem meiner Projekte arbeitet, hält sich an diese Regeln. Alle Werte stehen direkt in dieser Datei, damit sie auch ohne die übrigen Kit-Dateien nutzbar ist.

---

## 1. Grundhaltung

**Elegant, zurückhaltend, professionell.** Die Seite wirkt wie ein hochwertiges Werkzeug, nicht wie eine Werbeseite.

- Schwarz/Weiß bildet die Basis. Farbe ist die Ausnahme, nicht die Regel.
- **Dunkelrot** ist der einzige Farbakzent. **Schwarz** (im Dunkelmodus Weiß) ist der ruhige Zweitakzent.
- Feine Linien (1 px), viel Weißraum, ruhige Typografie mit mittlerem Schriftgewicht (600 statt 800).
- Keine grellen Farben, keine Farbverläufe, keine farbigen Hintergrundbänder, keine dekorativen Animationen.
- Inhalt vor Dekoration: Jedes Element muss einen praktischen Nutzen haben.

---

## 2. Farben

Farben **immer über Variablen** verwenden, nie direkt als Hex-Wert in Komponenten. Dann stimmen heller und dunkler Modus automatisch.

### Heller Modus (Standard)

| Variable | Wert | Verwendung |
|---|---|---|
| `--pr-bg` | `#ffffff` | Seitenhintergrund |
| `--pr-surface` | `#ffffff` | Karten, Panels |
| `--pr-surface-2` | `#f8f8f8` | leicht abgesetzte Flächen (Dropzone, Info-Kästen) |
| `--pr-border` | `#e7e7e7` | Rahmen, Trennlinien |
| `--pr-border-strong` | `#d2d2d2` | betonte Rahmen, gestrichelte Linien |
| `--pr-text` | `#121212` | Überschriften, Haupttext |
| `--pr-text-2` | `#4b4b4b` | Fließtext, Beschreibungen |
| `--pr-text-3` | `#8a8a8a` | Nebeninfos, Fußzeile, Platzhalter |
| `--pr-red` | `#8a0b18` | **Akzent:** Hauptbuttons, Auswahlrahmen |
| `--pr-red-hover` | `#6d0812` | Hover von Hauptbuttons |
| `--pr-red-text` | `#8a0b18` | dunkelrote Schrift (Eyebrow, Links, Hervorhebung im Titel) |
| `--pr-red-soft` | `#f9ecee` | zarter roter Hintergrund (Fehler, Drag-Over) |
| `--pr-red-ring` | `rgba(138, 11, 24, 0.16)` | Auswahl-Schein um Karten |
| `--pr-ink` | `#121212` | **Zweitakzent:** Nummern, Symbole, Info-Linien, Fokus |
| `--pr-ink-soft` | `#f4f4f4` | Hintergrund von Info-Hinweisen und Badges |
| `--pr-on-ink` | `#ffffff` | Schrift auf schwarzen Buttons |
| `--pr-code-bg` | `#f3f3f3` | Hintergrund von Code |

### Dunkler Modus

Wird automatisch über `prefers-color-scheme: dark` aktiv; erzwingen mit `<html data-theme="dark">`.

| Variable | Wert |
|---|---|
| `--pr-bg` | `#0a0a0a` |
| `--pr-surface` | `#111111` |
| `--pr-surface-2` | `#161616` |
| `--pr-border` | `#232323` |
| `--pr-border-strong` | `#343434` |
| `--pr-text` | `#f0f0f0` |
| `--pr-text-2` | `#b5b5b5` |
| `--pr-text-3` | `#7f7f7f` |
| `--pr-red` | `#9a0e1b` |
| `--pr-red-hover` | `#b31220` |
| `--pr-red-text` | `#c7232f` |
| `--pr-red-soft` | `#1b0608` |
| `--pr-red-ring` | `rgba(199, 35, 47, 0.25)` |
| `--pr-ink` | `#f0f0f0` |
| `--pr-ink-soft` | `#171717` |
| `--pr-on-ink` | `#0a0a0a` |
| `--pr-code-bg` | `#1a1a1a` |

### Farbregeln

- **Dunkelrot = was der Nutzer *tun* soll** (Hauptbutton, Auswahl, wichtige Hervorhebung).
- **Schwarz/Weiß-Akzent = was der Nutzer *wissen* soll** (Nummern, Symbole, Info-Hinweise).
- Pro Ansicht **ein** dunkelroter Hauptbutton. Weitere Aktionen als Ghost- oder schwarze Buttons.
- Grautöne sind **neutral** – kein Blau-, Grün- oder Braunstich.
- **Nicht verwenden:** helles/grelles Rot, Navy oder andere Blautöne, Grün für Erfolg, Gelb/Orange für Warnungen. Erfolg wird neutral (Grau/Schwarz) dargestellt, Fehler und Warnungen in Dunkelrot.
- Hintergründe sind einfarbig: Weiß bzw. Schwarz. Keine farbigen Bänder hinter Kopfzeile, Hero oder Fußzeile.

---

## 3. Typografie

Systemschriften – keine Webfonts (schnell, datenschutzfreundlich, offline nutzbar).

| Zweck | Schrift |
|---|---|
| Text | `ui-sans-serif, system-ui, -apple-system, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif` |
| Code, Dateiendungen, Zellbezüge | `ui-monospace, 'SF Mono', 'Cascadia Code', Menlo, Consolas, monospace` |

| Element | Größe | Gewicht | Sonstiges |
|---|---|---|---|
| Seitentitel (h1) | `clamp(30px, 4.4vw, 46px)` | 600 | `letter-spacing: -0.03em`, ein Wort darf dunkelrot sein |
| Abschnittstitel (h2) | `clamp(24px, 3.2vw, 32px)` | 600 | `letter-spacing: -0.025em` |
| Untertitel (h3) | 16 px | 600 | |
| Eyebrow (Kategorie über dem Titel) | 12 px | 600 | GROSSBUCHSTABEN, `letter-spacing: 0.16em`, dunkelrot |
| Einleitung (Lead) | 17 px | 400 | `--pr-text-2` |
| Fließtext | 15–16 px | 400 | Zeilenhöhe 1.6 |
| Kleintext, Labels | 13–14.5 px | 400–600 | |
| Fußzeile | 13 px | 400 | `--pr-text-3` |

- Überschriften nie fetter als 600, nie in Großbuchstaben (außer Eyebrow und Tabellenköpfe).
- Hervorhebung im Fließtext mit `<strong>` (600) in `--pr-text`, nicht mit Farbe.

---

## 4. Form, Abstände, Layout

| Token | Wert |
|---|---|
| `--pr-radius-sm` | 6 px (Hinweise) |
| `--pr-radius` | 10 px (Karten, Dropzone) |
| `--pr-radius-lg` | 14 px (große Panels) |
| Buttons, Eingabefelder | 8 px Radius |
| Chips, Badges | vollständig rund (999 px) |
| `--pr-container` | max. 1120 px Inhaltsbreite |
| Seitenrand | 20 px (≤ 560 px Breite: 16 px) |
| `--pr-header-h` | 60 px |

- Abstände im 4-px-Raster: typisch 8 / 12 / 16 / 22 / 32 px; zwischen großen Abschnitten 80–96 px.
- Rahmen immer 1 px. Ausnahmen: Akzentlinien 2 px (Tabs, Hinweise).
- Schatten nur sehr zart und nur im hellen Modus: `0 1px 2px rgba(18,18,18,.04), 0 10px 30px rgba(18,18,18,.05)`. Im dunklen Modus kein Schatten.
- Mobil (≤ 560 px): Raster einspaltig, Buttons volle Breite, Hero linksbündig, kein horizontales Scrollen.

---

## 5. Logo & Signatur

Das Logo ist meine Signatur „PR“ (`logo.svg`, Seitenverhältnis 624 : 387). Es ist **immer dezent**.

| Ort | Darstellung |
|---|---|
| Browser-Tab (Favicon) | feine Signatur **ohne Kachel**, schwarz im hellen / weiß im dunklen Modus (`favicon.svg`) |
| Kopfzeile links | feine Signatur, 24 px hoch, in Textfarbe, vor dem Projektnamen |
| Fußzeile | feine Signatur 22 px hoch, 55 % Deckkraft, daneben klein „Paul-Rafael Hasan-Man“ bzw. „Entwickelt von Paul-Rafael Hasan-Man“; beim Überfahren 90 % |

- Das Logo übernimmt immer die Textfarbe (`currentColor` bzw. CSS-Maske) – nie farbig, nie mit Hintergrundkachel, nie verdickt.
- Nie groß im Inhaltsbereich, nicht als Wasserzeichen, nicht verzerrt.
- Einbindung als CSS-Maske, damit es sich automatisch an hell/dunkel anpasst:

```html
<span class="pr-logo" aria-hidden="true"></span>
```

```css
.pr-logo {
  display: inline-block; height: 1em; aspect-ratio: 624 / 387;
  background: currentColor;
  -webkit-mask: url('logo.svg') no-repeat center / contain;
  mask: url('logo.svg') no-repeat center / contain;
}
```

---

## 6. Symbole

- Linien-Symbole (Outline) im Stil von Lucide: 24×24-Raster, `stroke-width: 1.75`, runde Enden, keine Füllung.
- Größe 16–20 px; in Symbolkästen (38×38 px, 1 px Rahmen, Radius 9 px) in `--pr-ink`.
- Ausgewählte Elemente: Symbol und Rahmen dunkelrot.
- Keine Emojis, keine bunten Icons.

---

## 7. Bausteine

Alle Klassen tragen das Präfix `pr-`, damit sie mit Projektcode nicht kollidieren. Vollständige Umsetzung in `brand.css`.

| Baustein | Gestaltung |
|---|---|
| **Kopfzeile** `pr-header` | sticky, Hintergrund = Seitenhintergrund mit 88 % Deckkraft + Unschärfe, 1 px Linie unten. Links Logo + Projektname (Name 600, Zusatz in `--pr-text-3`), rechts schlichte Textlinks. Keine Badges. |
| **Hero** `pr-hero` | zentriert, max. 780 px: Eyebrow → Titel → Lead. Oben 80 px, unten 48 px Abstand. |
| **Hauptbutton** `pr-btn-primary` | 44 px hoch, dunkelrot, weiße Schrift 15 px/600, Radius 8 px |
| **Zweitbutton** `pr-btn-secondary` | schwarz (dunkel: weiß), invertierte Schrift |
| **Ghost-Button** `pr-btn-ghost` | transparent, 1 px Rahmen `--pr-border-strong`, beim Überfahren Rahmen in Textfarbe |
| **Deaktiviert** | 35 % Deckkraft |
| **Panel** `pr-panel` | Fläche, 1 px Rahmen, Radius 14 px, zarter Schatten, Innenabstand 32 px |
| **Karte** `pr-card` | Fläche, 1 px Rahmen, Radius 10 px, Innenabstand 22 px |
| **Auswahlkarte** `pr-card-select` | ausgewählt: dunkelroter Rahmen + 3 px Schein `--pr-red-ring`, kleines rundes Häkchen oben rechts |
| **Nummer** `pr-number` | Kreis 24 px, 1 px Rahmen und Ziffer in `--pr-ink` |
| **Schrittnummern im Text** | „01 / 02 / 03“ in 12 px, 600, gesperrt |
| **Chip** `pr-chip` | Monospace 11 px, runder 1 px Rahmen, `--pr-text-3` (z. B. Dateiendungen) |
| **Badge** `pr-badge` | 12 px/600 auf `--pr-ink-soft` |
| **Info-Hinweis** `pr-note` | 2 px Linie links in `--pr-ink`, Hintergrund `--pr-ink-soft` |
| **Wichtiger Hinweis** `pr-note-accent` | 2 px Linie links dunkelrot, Hintergrund `--pr-red-soft` |
| **Fehler** `pr-alert` | Rahmen dunkelrot (45 %), Hintergrund `--pr-red-soft`, rundes dunkelrotes Symbol |
| **Erfolg** `pr-success` | neutral: Fläche `--pr-surface-2`, 1 px Rahmen, Häkchen im Kreis mit dunkelrotem Rand |
| **Tabs** `pr-tabs` | Textreiter mit Linie unten; aktiv: Textfarbe + 2 px dunkelrote Unterstreichung |
| **Eingabefeld** `pr-input` | 1 px Rahmen, Radius 8 px; Fokus: Rahmen `--pr-ink` + 3 px Schein |
| **Dropzone** `pr-dropzone` | gestrichelter 1 px Rahmen, `--pr-surface-2`; beim Ziehen dunkelroter Rahmen + `--pr-red-soft` |
| **Tabelle** `pr-table` | Kopf 12 px Großbuchstaben in `--pr-text-3`, Zeilen mit 1 px Linien, keine Zebra-Streifen |
| **Fußzeile** `pr-footer` | 1 px Linie oben, 13 px in `--pr-text-3`; links Projektname, rechts Links; ganz unten die Signatur |

---

## 8. Bewegung

- Übergänge nur für Farbe, Rahmen und Deckkraft: `0.2s`.
- Einzige Animation: Lade-Kreisel im Button während einer Verarbeitung.
- `prefers-reduced-motion` respektieren (Übergänge aus, kein weiches Scrollen).

---

## 9. Texte & Inhalte

- Sprache Deutsch, Ansprache mit **du**, kurze und klare Sätze.
- Werkzeuge erklären sich selbst: kurze Beschreibung oben, ausführliche **Anleitung** weiter unten auf derselben Seite (Wofür? · So gehst du vor · Was genau passiert).
- **Keine** Firmen-, Kunden- oder Bibliotheksnamen auf der Website (Produktnamen der Zielsoftware wie „WinCC Unified“ oder „TIA Portal“ sind in Ordnung).
- **Keine** „BETA“-Kennzeichnungen, keine Marketing-Badges wie „Lokal im Browser“ in der Kopfzeile, keine künstlichen Statistiken.
- Fehlermeldungen sagen, **was** passiert ist und **was** der Nutzer tun kann.
- Datenschutzhinweis, falls Dateien verarbeitet werden: klein, unter dem Werkzeug, mit Schild-Symbol.
- Fußzeile enthält immer klein meinen Namen mit Signatur.

---

## 10. Barrierefreiheit

- Kontrast für Fließtext mindestens 4.5 : 1 (die Farbwerte oben erfüllen das).
- Sichtbarer Fokus: 2 px Kontur in `--pr-ink`, 2 px Abstand.
- Alle Bedienelemente per Tastatur erreichbar; Auswahlkarten als echte Radio-Buttons.
- Dekorative Symbole und das Logo mit `aria-hidden="true"`.

---

## 11. GitHub-Auftritt

- README beginnt zentriert mit: Titel · ein Satz Beschreibung · Badges · Link „… öffnen →“ · Screenshot (heller Modus).
- Badge-Farben: Dunkelrot `8a0b18` für den Haupt-Badge (z. B. Live-Demo), Schwarz `0a0a0a` für die übrigen; Stil `for-the-badge`.
- Danach: kurzer Einleitungstext → Funktionstabelle → Highlights → Nutzung → Entwicklung → Projektstruktur → Dokumentation.
- Ganz unten klein: `<sub>Entwickelt von **Paul-Rafael Hasan-Man**</sub>`.
- Repo-Beschreibung als ein sachlicher Satz, Homepage-Link gesetzt, passende Topics.

---

## 12. Neues Projekt – Checkliste

1. Ordner `brand/` (mit `brand.css`, `logo.svg`, `favicon.svg`) ins Projekt kopieren.
2. Im `<head>` einbinden:
   ```html
   <link rel="icon" type="image/svg+xml" href="brand/favicon.svg" />
   <link rel="stylesheet" href="brand/brand.css" />
   <meta name="author" content="Paul-Rafael Hasan-Man" />
   ```
3. Grundgerüst aus `template.html` übernehmen (Kopfzeile · Hero · Panel · Fußzeile mit Signatur).
4. Eigene Styles nur mit `var(--pr-…)`-Variablen schreiben, keine neuen Farben einführen.
5. Hell **und** dunkel prüfen, außerdem 375 px Breite.
6. README nach Abschnitt 11 gestalten.

Ohne die Kit-Dateien: Den CSS-Block aus Abschnitt 13 als `brand.css` anlegen; Logo und Favicon aus einem bestehenden Projekt übernehmen.

---

## 13. Design-Tokens zum Kopieren

```css
:root {
  --pr-bg: #ffffff;
  --pr-surface: #ffffff;
  --pr-surface-2: #f8f8f8;
  --pr-border: #e7e7e7;
  --pr-border-strong: #d2d2d2;
  --pr-text: #121212;
  --pr-text-2: #4b4b4b;
  --pr-text-3: #8a8a8a;
  --pr-red: #8a0b18;
  --pr-red-hover: #6d0812;
  --pr-red-text: #8a0b18;
  --pr-red-soft: #f9ecee;
  --pr-red-ring: rgba(138, 11, 24, 0.16);
  --pr-ink: #121212;
  --pr-ink-soft: #f4f4f4;
  --pr-ink-ring: rgba(18, 18, 18, 0.14);
  --pr-on-ink: #ffffff;
  --pr-on-accent: #ffffff;
  --pr-code-bg: #f3f3f3;
  --pr-shadow: 0 1px 2px rgba(18, 18, 18, 0.04), 0 10px 30px rgba(18, 18, 18, 0.05);
  --pr-font: ui-sans-serif, system-ui, -apple-system, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
  --pr-mono: ui-monospace, 'SF Mono', 'Cascadia Code', Menlo, Consolas, monospace;
  --pr-radius-sm: 6px;
  --pr-radius: 10px;
  --pr-radius-lg: 14px;
  --pr-container: 1120px;
  --pr-header-h: 60px;
  color-scheme: light;
}

@media (prefers-color-scheme: dark) {
  :root:not([data-theme='light']) {
    --pr-bg: #0a0a0a;
    --pr-surface: #111111;
    --pr-surface-2: #161616;
    --pr-border: #232323;
    --pr-border-strong: #343434;
    --pr-text: #f0f0f0;
    --pr-text-2: #b5b5b5;
    --pr-text-3: #7f7f7f;
    --pr-red: #9a0e1b;
    --pr-red-hover: #b31220;
    --pr-red-text: #c7232f;
    --pr-red-soft: #1b0608;
    --pr-red-ring: rgba(199, 35, 47, 0.25);
    --pr-ink: #f0f0f0;
    --pr-ink-soft: #171717;
    --pr-ink-ring: rgba(240, 240, 240, 0.18);
    --pr-on-ink: #0a0a0a;
    --pr-code-bg: #1a1a1a;
    --pr-shadow: none;
    color-scheme: dark;
  }
}
```

---

## Hinweis für KI-Assistenten

Wenn du in einem Projekt arbeitest, das diese Datei enthält:

- Halte dich bei jeder Oberfläche an diese Richtlinie. Führe keine neuen Farben, Schriften oder Effekte ein.
- Verwende die `pr-`-Bausteine aus `brand.css`, bevor du eigene Komponenten baust.
- Zeige Designänderungen zuerst (Screenshot hell und dunkel) und veröffentliche sie erst nach Freigabe.
