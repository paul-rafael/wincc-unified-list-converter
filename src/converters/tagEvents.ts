/**
 * Variablen-Ereignisse – replacement for WinCC Advanced tag events (not part of the Excel tool).
 *
 * The TIA Portal Excel export of HMI tags does not contain tag events, so this works in two steps:
 * 1. HMI tag export ("Hmi Tags") → template TagEvents.xlsx: sheet "Ereignisse" pre-filled with every tag
 *    that has a limit (candidates for "Überschreitung"/"Unterschreitung"), sheet "Variablen" for the name check.
 * 2. Filled-in template → TagEvent.js: one function per event group, called from TagEvent(), meant to be run
 *    cyclically in WinCC Unified.
 */
import { ConversionError } from '../errors';
import { buildWorkbookSheets } from '../xlsx/newWorkbook';
import { cellRef } from '../xlsx/ref';
import type { CellValue, Sheet, WriteValue } from '../xlsx/workbook';
import { Workbook } from '../xlsx/workbook';
import { assertExportKind, assertExtension, ConversionResult, headerTexts, NoticeLog } from './common';
import { convertAdvancedEvents } from './tagEventsXml';

export const TEMPLATE_NAME = 'TagEvents.xlsx';
export const SCRIPT_NAME = 'TagEvent.js';
export const EVENTS_SHEET = 'Ereignisse';
export const TAGS_SHEET = 'Variablen';

export const EVENTS = ['Wertänderung', 'Überschreitung', 'Unterschreitung'] as const;
export const ACTIONS = ['Variable setzen', 'Bild wechseln', 'Skript'] as const;
type EventKind = (typeof EVENTS)[number];
type ActionKind = (typeof ACTIONS)[number];

const EVENT_HEADER = ['Funktion', 'Trigger-Variable', 'Ereignis', 'Grenzwert', 'Aktion', 'Ziel', 'Wert / Formel', 'Kommentar'];
const TAG_HEADER = ['Name', 'DataType', 'Grenzwert oben', 'Grenzwert unten', 'Erfassung'];
const TEMPLATE_ROWS = 2000;
const INDENT = '    ';

const text = (v: CellValue): string => {
  switch (v.kind) {
    case 'string':
      return v.text.trim();
    case 'number':
      return String(v.value);
    case 'bool':
      return v.value ? 'true' : 'false';
    default:
      return '';
  }
};

const noValue = (s: string) => s === '' || s.toLowerCase() === '<no value>';

/** JavaScript identifier for a tag or function name ("12V_CCS_HmiAxis_StatRunNo" → "HmiAxis_StatRunNo"). */
export function functionName(tag: string): string {
  let id = tag.replace(/[^A-Za-z0-9_]/g, '_');
  // TIA projects often prefix tags with "<no>V_<Connection>_"; the rest reads better as a function name.
  const stripped = id.replace(/^\d+[A-Za-z]*_[A-Za-z0-9]+_+/, '');
  if (/^[A-Za-z]/.test(stripped)) id = stripped;
  return /^[A-Za-z_$]/.test(id) ? id : `Ev_${id}`;
}

/** Short local variable name for a tag ("12V_CCS_ModStat_PlcNo" → "plcNo"). */
function localName(tag: string, used: Set<string>): string {
  const last = tag.split(/[^A-Za-z0-9]+/).filter(Boolean).pop() ?? 'wert';
  let base = /^\d/.test(last) ? `v${last}` : last.charAt(0).toLowerCase() + last.slice(1);
  if (RESERVED.has(base)) base = `${base}Wert`;
  let name = base;
  for (let n = 2; used.has(name); n++) name = `${base}${n}`;
  used.add(name);
  return name;
}

const RESERVED = new Set(['break', 'case', 'catch', 'class', 'const', 'continue', 'default', 'delete', 'do', 'else', 'export', 'false', 'for', 'function', 'if', 'in', 'let', 'new', 'null', 'return', 'switch', 'this', 'throw', 'true', 'try', 'typeof', 'var', 'void', 'while', 'with', 'trigger', 'changed']);

const js = (s: string) => JSON.stringify(s);
const readTag = (tag: string) => `HMIRuntime.Tags.SysFct.GetTagValue(${js(tag)})`;

// ------------------------------------------------------------------ step 1: template

export async function buildEventTemplate(wb: Workbook): Promise<{ data: Uint8Array; candidates: number; tags: number }> {
  const ws = await wb.sheet(0);
  const headers = headerTexts(ws).map((h) => h.toLowerCase());
  const col = (name: string) => headers.indexOf(name.toLowerCase()) + 1;
  const cName = col('Name');
  if (cName !== 1 || !col('DataType')) {
    throw new ConversionError('Die Datei ist kein HMI-Variablen-Export (erwartet: Blatt mit den Spalten „Name“ und „DataType“). Bitte die HMI-Variablen aus dem TIA Portal exportieren.');
  }
  const get = (r: number, name: string) => (col(name) ? text(ws.getCell(r, col(name)).value) : '');

  const s = (t: string): WriteValue => ({ kind: 'string', text: t });
  const limitValue = (type: string, value: string): WriteValue | null => {
    if (type === '' || type === 'None' || noValue(value)) return null;
    const n = Number(value);
    return type === 'Constant' && value !== '' && Number.isFinite(n) ? { kind: 'number', value: n } : s(value);
  };

  const tagGrid = new Map<number, Map<number, WriteValue>>([[1, new Map(TAG_HEADER.map((h, i) => [i + 1, s(h)]))]]);
  const eventGrid = new Map<number, Map<number, WriteValue>>([[1, new Map(EVENT_HEADER.map((h, i) => [i + 1, s(h)]))]]);
  let tagRow = 1;
  let eventRow = 1;
  const last = ws.lastRowInColumn(1);
  for (let r = 2; r <= last; r++) {
    const name = get(r, 'Name');
    if (!name) continue;
    const upper = limitValue(get(r, 'Limit Upper 2 Type'), get(r, 'Limit Upper 2'));
    const lower = limitValue(get(r, 'Limit Lower 2 Type'), get(r, 'Limit Lower 2'));
    const row = new Map<number, WriteValue>([[1, s(name)], [2, s(get(r, 'DataType'))], [5, s(get(r, 'Acquisition mode'))]]);
    if (upper) row.set(3, upper);
    if (lower) row.set(4, lower);
    tagGrid.set(++tagRow, row);
    for (const [event, limit] of [['Überschreitung', upper], ['Unterschreitung', lower]] as const) {
      if (!limit) continue;
      eventGrid.set(++eventRow, new Map([[1, s(functionName(name))], [2, s(name)], [3, s(event)], [4, limit]]));
    }
  }

  const help = [
    'Variablen-Ereignisse für WinCC Unified',
    '',
    'Der Excel-Export des TIA Portals enthält keine Variablen-Ereignisse. Trage sie deshalb im Blatt „Ereignisse“ ein und lade diese Datei danach wieder im Converter hoch (Funktion „Variablen-Ereignisse“). Daraus entsteht das Skript TagEvent.js.',
    '',
    'Vorbelegt sind alle Variablen mit Grenzwert – Kandidaten für „Überschreitung“/„Unterschreitung“. Zeilen ohne Aktion werden ignoriert, nicht benötigte Zeilen einfach löschen.',
    '',
    'Spalten im Blatt „Ereignisse“',
    'Funktion – Name der Skript-Funktion. Zeilen mit gleicher Funktion landen in derselben Funktion (mehrere Trigger und Aktionen möglich).',
    'Trigger-Variable – Variable, an der das Ereignis hing.',
    'Ereignis – Wertänderung, Überschreitung oder Unterschreitung.',
    'Grenzwert – nur bei Über-/Unterschreitung: Zahl oder Name einer Variable.',
    'Aktion – Variable setzen, Bild wechseln oder Skript.',
    'Ziel – bei „Variable setzen“ die Zielvariable, bei „Bild wechseln“ der Bildname.',
    'Wert / Formel – Zahl, true/false, "Text" oder Formel. {Trigger} steht für den Wert der Trigger-Variable, {Variablenname} für den Wert einer anderen Variable.',
    '    Beispiele: 1010   |   {Trigger} + 12   |   10000 * {12V_CCS_ModStat_PlcNo} + {12V_CCS_ModStat_RunNo}',
    '    Bei „Skript“: JavaScript-Code, ebenfalls mit {Trigger} und {Variablenname}.',
    '',
    'Verhalten im Skript',
    'TagEvent() wird zyklisch aufgerufen. Wertänderung: die Aktion läuft bei jedem Aufruf. Überschreitung/Unterschreitung: die Aktion läuft bei jedem Aufruf, solange der Wert über bzw. unter dem Grenzwert liegt.',
    'Alle Variablennamen werden gegen das Blatt „Variablen“ (aus dem Export) geprüft.',
  ];
  const helpGrid = new Map<number, Map<number, WriteValue>>(help.map((line, i) => [i + 1, new Map([[1, s(line)]])]));

  const data = await buildWorkbookSheets([
    {
      name: EVENTS_SHEET,
      grid: eventGrid,
      header: true,
      widths: [30, 46, 18, 16, 18, 46, 48, 30],
      lists: [
        { sqref: `C2:C${TEMPLATE_ROWS}`, list: [...EVENTS] },
        { sqref: `E2:E${TEMPLATE_ROWS}`, list: [...ACTIONS] },
      ],
    },
    { name: TAGS_SHEET, grid: tagGrid, header: true, widths: [52, 14, 18, 18, 22] },
    { name: 'Anleitung', grid: helpGrid, widths: [150] },
  ]);
  return { data, candidates: eventRow - 1, tags: tagRow - 1 };
}

// ------------------------------------------------------------------ step 2: script

interface EventRow {
  row: number;
  fn: string;
  trigger: string;
  event: EventKind;
  limit: string;
  action: ActionKind;
  target: string;
  value: string;
  comment: string;
}

const where = (r: number, c: number) => `${EVENTS_SHEET}!${cellRef(r, c)}`;

function readRows(ws: Sheet, log: NoticeLog): EventRow[] {
  const rows: EventRow[] = [];
  let withoutAction = 0;
  const last = Math.max(ws.lastRowInColumn(2), ws.lastRowInColumn(5));
  for (let r = 2; r <= last; r++) {
    const cell = (c: number) => text(ws.getCell(r, c).value);
    const trigger = cell(2);
    const actionText = cell(5);
    if (!trigger && !actionText) continue;
    if (!actionText) {
      withoutAction++;
      continue;
    }
    if (!trigger) {
      log.warn('Zeile ohne Trigger-Variable übersprungen.', where(r, 2));
      continue;
    }
    const event = EVENTS.find((e) => e.toLowerCase() === cell(3).toLowerCase());
    if (!event) {
      log.warn(`Unbekanntes Ereignis „${cell(3)}“ – erlaubt: ${EVENTS.join(', ')}. Zeile übersprungen.`, where(r, 3));
      continue;
    }
    const action = ACTIONS.find((a) => a.toLowerCase() === actionText.toLowerCase());
    if (!action) {
      log.warn(`Unbekannte Aktion „${actionText}“ – erlaubt: ${ACTIONS.join(', ')}. Zeile übersprungen.`, where(r, 5));
      continue;
    }
    const limit = cell(4);
    if (event !== 'Wertänderung' && limit === '') {
      log.warn(`„${event}“ braucht einen Grenzwert. Zeile übersprungen.`, where(r, 4));
      continue;
    }
    if (action !== 'Skript' && cell(6) === '') {
      log.warn(`„${action}“ braucht ein Ziel. Zeile übersprungen.`, where(r, 6));
      continue;
    }
    if (action === 'Variable setzen' && cell(7) === '') {
      log.warn('„Variable setzen“ braucht einen Wert oder eine Formel. Zeile übersprungen.', where(r, 7));
      continue;
    }
    const fn = cell(1) === '' ? functionName(trigger) : cell(1);
    if (!/^[A-Za-z_$][A-Za-z0-9_$]*$/.test(fn) || RESERVED.has(fn)) {
      log.warn(`„${fn}“ ist kein gültiger Funktionsname (nur Buchstaben, Ziffern, _; nicht mit Ziffer beginnen). Zeile übersprungen.`, where(r, 1));
      continue;
    }
    rows.push({ row: r, fn, trigger, event, limit, action, target: cell(6), value: cell(7), comment: cell(8) });
  }
  if (withoutAction) log.info(`${withoutAction} Zeilen ohne Aktion wurden ignoriert (z. B. vorbelegte Grenzwert-Variablen ohne Ereignis).`);
  return rows;
}

const REF_RE = /\{([^{}]+)\}/g;
const VALUE_RE = /^(?:true|false|"(?:[^"\\]|\\.)*")$/;

/** A formula may only consist of numbers, operators, brackets, Math.* and the tag values it references. */
function isFormula(expr: string, locals: Iterable<string>): boolean {
  const names = new Set(locals);
  const rest = expr
    .replace(/Math\.[A-Za-z]+/g, ' ')
    .replace(/[A-Za-z_$][A-Za-z0-9_$]*/g, (id) => (names.has(id) ? ' ' : '#'))
    .replace(/\d+(?:\.\d+)?/g, ' ');
  return /^[\s+\-*/%(),]*$/.test(rest) && /\S/.test(expr);
}

interface FnBuild {
  name: string;
  locals: Map<string, string>;
  used: Set<string>;
  triggers: Map<string, Set<EventKind>>;
  blocks: { cond: string; lines: string[] }[];
  comments: string[];
  /** Triggers whose value change starts a screen change or script (edge detected via TagEventLast). */
  changeTriggers: Set<string>;
}

const CHANGED = '\u0000changed';

function local(fn: FnBuild, tag: string): string {
  let name = fn.locals.get(tag);
  if (!name) {
    name = localName(tag, fn.used);
    fn.locals.set(tag, name);
  }
  return name;
}

/** Replaces {Trigger} and {Tag} with local variables and collects the referenced tags. */
function substitute(src: string, fn: FnBuild, trigger: string, refs: string[]): string {
  return src.replace(REF_RE, (_, raw: string) => {
    const tag = raw.trim().toLowerCase() === 'trigger' ? trigger : raw.trim();
    refs.push(tag);
    return local(fn, tag);
  });
}

export function buildEventScript(rows: EventRow[], knownTags: Set<string> | null, log: NoticeLog, source: string): { script: string; functions: number; actions: number } {
  const fns = new Map<string, FnBuild>();
  let actions = 0;
  let edgeDetection = false;

  const check = (tag: string, r: number, c: number) => {
    if (knownTags && !knownTags.has(tag)) log.warn(`Variable „${tag}“ steht nicht im HMI-Variablen-Export. Bitte Schreibweise prüfen.`, where(r, c));
  };

  for (const e of rows) {
    let fn = fns.get(e.fn);
    if (!fn) {
      fn = { name: e.fn, locals: new Map(), used: new Set(), triggers: new Map(), blocks: [], comments: [], changeTriggers: new Set() };
      fns.set(e.fn, fn);
    }
    check(e.trigger, e.row, 2);
    if (!fn.triggers.has(e.trigger)) fn.triggers.set(e.trigger, new Set());
    fn.triggers.get(e.trigger)!.add(e.event);
    if (e.comment && !fn.comments.includes(e.comment)) fn.comments.push(e.comment);

    let cond = '';
    if (e.event !== 'Wertänderung') {
      const limitNum = Number(e.limit.replace(',', '.'));
      let limit: string;
      if (Number.isFinite(limitNum) && e.limit !== '') limit = String(limitNum);
      else {
        check(e.limit, e.row, 4);
        limit = local(fn, e.limit);
      }
      cond = `${local(fn, e.trigger)} ${e.event === 'Überschreitung' ? '>' : '<'} ${limit}`;
    }
    // Setting a tag every cycle is harmless; a screen change or script must only run when the event occurs.
    if (e.action !== 'Variable setzen') {
      edgeDetection = true;
      if (cond) cond = `TagEventChanged(${js(cond.replace(local(fn, e.trigger), e.trigger))}, ${cond}) && ${cond}`;
      else {
        local(fn, e.trigger);
        fn.changeTriggers.add(e.trigger);
        cond = CHANGED;
      }
    }

    const refs: string[] = [];
    let line: string;
    if (e.action === 'Bild wechseln') {
      line = `HMIRuntime.UI.SysFct.ChangeScreen(${js(e.target)}, ".");`;
    } else if (e.action === 'Skript') {
      line = substitute(e.value, fn, e.trigger, refs);
    } else {
      check(e.target, e.row, 6);
      // German decimal comma ("1,5") becomes a point; "Math.max(a, b)" keeps its comma.
      const expr = substitute(e.value.replace(/(\d),(\d)/g, '$1.$2'), fn, e.trigger, refs);
      if (!VALUE_RE.test(expr) && !isFormula(expr, fn.locals.values())) {
        log.warn(`Wert/Formel „${e.value}“ ist ungültig (erlaubt: Zahl, true/false, "Text", Rechnung mit {Variable}). Zeile übersprungen.`, where(e.row, 7));
        continue;
      }
      line = `HMIRuntime.Tags.SysFct.SetTagValue(${js(e.target)}, ${expr.trim()});`;
    }
    for (const t of refs) check(t, e.row, 7);

    // Several triggers of one function usually share the same action: write it once.
    const block = fn.blocks.find((b) => b.cond === cond);
    if (block?.lines.includes(line)) continue;
    const lastBlock = fn.blocks[fn.blocks.length - 1];
    if (lastBlock && lastBlock.cond === cond) lastBlock.lines.push(line);
    else fn.blocks.push({ cond, lines: [line] });
    actions++;
  }

  const used = Array.from(fns.values()).filter((f) => f.blocks.length);
  const out: string[] = [];
  out.push(`// ${SCRIPT_NAME} – erzeugt mit dem WinCC Unified List Converter aus „${source}“`);
  out.push('// Ersatz für WinCC-Advanced-Variablenereignisse. TagEvent() zyklisch aufrufen.');
  out.push('// Wertänderung: Aktion läuft bei jedem Aufruf. Über-/Unterschreitung: bei jedem Aufruf, solange der Grenzwert über- bzw. unterschritten ist.');
  if (edgeDetection) {
    out.push('');
    out.push('// Globale Definition – in WinCC Unified in den Bereich „Globale Definitionen“ des Skriptmoduls übernehmen.');
    out.push('// Merkt sich die letzten Werte, damit Bildwechsel und Skripte nur beim Ereignis selbst laufen.');
    out.push('var TagEventLast = {};');
    out.push('');
  }
  out.push('export function TagEvent(parameter1, parameter2) {');
  for (const f of used) out.push(`${INDENT}${f.name}();`);
  if (edgeDetection) {
    out.push('');
    out.push(`${INDENT}// true, wenn sich der Wert seit dem letzten Aufruf geändert hat (beim ersten Aufruf: false)`);
    out.push(`${INDENT}function TagEventChanged(key, value) {`);
    out.push(`${INDENT}${INDENT}var changed = key in TagEventLast && TagEventLast[key] !== value;`);
    out.push(`${INDENT}${INDENT}TagEventLast[key] = value;`);
    out.push(`${INDENT}${INDENT}return changed;`);
    out.push(`${INDENT}}`);
  }
  for (const f of used) {
    out.push('');
    const trig = Array.from(f.triggers.entries())
      .map(([t, ev]) => `${t} (${Array.from(ev).join(', ')})`)
      .join(' / ');
    out.push(`${INDENT}// ${f.name} – Trigger: ${trig}`);
    for (const c of f.comments) out.push(`${INDENT}// ${c}`);
    out.push(`${INDENT}function ${f.name}() {`);
    for (const [tag, name] of f.locals) out.push(`${INDENT}${INDENT}var ${name} = ${readTag(tag)};`);
    if (f.changeTriggers.size) {
      // "|" instead of "||": every trigger has to update its stored value.
      const checks = Array.from(f.changeTriggers).map((t) => `TagEventChanged(${js(t)}, ${f.locals.get(t)})`);
      out.push(`${INDENT}${INDENT}var changed = ${checks.join(' | ')};`);
    }
    for (const b of f.blocks) {
      const pad = b.cond ? INDENT.repeat(3) : INDENT.repeat(2);
      if (b.cond) out.push(`${INDENT}${INDENT}if (${b.cond === CHANGED ? 'changed' : b.cond}) {`);
      for (const l of b.lines) for (const part of l.split(/\r?\n/)) out.push(`${pad}${part}`);
      if (b.cond) out.push(`${INDENT}${INDENT}}`);
    }
    out.push(`${INDENT}}`);
  }
  out.push('}');
  out.push('');
  if (edgeDetection) {
    log.info('Bildwechsel und Skripte laufen nur, wenn das Ereignis eintritt. Dafür steht oben im Skript die Zeile „var TagEventLast = {};“ – sie gehört in die globalen Definitionen des Skriptmoduls.');
  }
  return { script: out.join('\n'), functions: used.length, actions };
}

// ------------------------------------------------------------------ entry point

export async function convertTagEvents(input: ArrayBuffer | Uint8Array, fileName: string): Promise<ConversionResult> {
  assertExtension(fileName, ['.zip', '.xml', '.xlsx', '.xlsm']);
  // Openness export: the events are in the files, the script is generated directly.
  if (/\.(zip|xml)$/i.test(fileName)) return convertAdvancedEvents(input, fileName);
  const wb = await Workbook.load(input);
  const log = new NoticeLog();

  const idx = wb.sheetIndexByName(EVENTS_SHEET);
  if (idx < 0) {
    await assertExportKind(wb, 'tags');
    const { data, candidates, tags } = await buildEventTemplate(wb);
    return {
      data,
      fileName: TEMPLATE_NAME,
      notices: log.items,
      summary:
        `Vorlage erstellt: ${tags} Variablen übernommen, ${candidates} Grenzwert-Ereignisse vorbelegt. ` +
        'Ereignisse im Blatt „Ereignisse“ eintragen und die Datei hier wieder hochladen – daraus entsteht TagEvent.js.',
    };
  }

  const rows = readRows(await wb.sheet(idx), log);
  if (!rows.length) {
    throw new ConversionError('Das Blatt „Ereignisse“ enthält keine Zeile mit Aktion. Bitte mindestens ein Ereignis mit Aktion eintragen.');
  }
  let known: Set<string> | null = null;
  const tagIdx = wb.sheetIndexByName(TAGS_SHEET);
  if (tagIdx >= 0) {
    const ts = await wb.sheet(tagIdx);
    known = new Set();
    for (let r = 2; r <= ts.lastRowInColumn(1); r++) {
      const n = text(ts.getCell(r, 1).value);
      if (n) known.add(n);
    }
  } else {
    log.info('Blatt „Variablen“ fehlt – Variablennamen wurden nicht geprüft.');
  }

  const { script, functions, actions } = buildEventScript(rows, known, log, fileName);
  if (!actions) throw new ConversionError('Keine Zeile konnte umgesetzt werden. Bitte die Hinweise prüfen.');
  return {
    data: new TextEncoder().encode(script),
    fileName: SCRIPT_NAME,
    text: script,
    notices: log.items,
    summary: `Skript erstellt: ${functions} Funktionen mit ${actions} Aktionen aus ${rows.length} Zeilen.`,
  };
}
