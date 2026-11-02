/**
 * Variablen-Ereignisse aus dem Openness-Export – reads the HMI tags of a WinCC Advanced/Comfort device as
 * exported by TIA Portal Openness (one SimaticML file per tag, "Hmi.Tag.Tag" with "Hmi.Event.Event" children)
 * and generates the Unified script that replaces their events.
 *
 * Accepted input: the export ZIP (manifest.json + Tags/*.xml) or a single tag XML file.
 *
 * Advanced behaviour that is emulated (TagEvent() is called cyclically):
 * - ChangeValue: the function list runs when the value has changed since the last call.
 * - AboveUpperLimit / BelowLowerLimit: the function list runs when the limit is crossed.
 * - First call after runtime start: tag functions (SetTag, LinearScaling, …) run once, so derived tags are
 *   initialised; screen changes do not.
 */
import JSZip from 'jszip';
import { ConversionError } from '../errors';
import { replaceNonAlnum } from '../vba/vba';
import { parseXml } from '../xlsx/xml';
import { ConversionResult, NoticeLog } from './common';
import { functionName, SCRIPT_NAME } from './tagEvents';

export type ParamValue = { kind: 'self' } | { kind: 'tag'; name: string } | { kind: 'const'; type: string; value: string } | { kind: 'empty' };

export interface FunctionEntry {
  name: string;
  type: string;
  params: Map<string, ParamValue>;
}

export interface TagEventDef {
  kind: string;
  entries: FunctionEntry[];
}

export interface AdvancedTag {
  /** Name as in the Advanced project, e.g. "12V_CCS\HmiSc\ScDbAdrDatNw". */
  name: string;
  dataType: string;
  limitUpper: ParamValue;
  limitLower: ParamValue;
  events: TagEventDef[];
  source: string;
}

const EVENT_LABEL: Record<string, string> = {
  ChangeValue: 'Wertänderung',
  AboveUpperLimit: 'Überschreitung',
  BelowLowerLimit: 'Unterschreitung',
};

/** Unified tag name: the HMI tag migration replaces every character outside [A-Za-z0-9] by "_". */
export function unifiedName(advancedName: string): string {
  return replaceNonAlnum(advancedName.trim().replace(/^"(.*)"$/, '$1'));
}

// ------------------------------------------------------------------ reading

const kids = (el: Element | null, name: string): Element[] =>
  el ? (Array.from(el.childNodes).filter((n) => n.nodeType === 1 && (n as Element).nodeName === name) as Element[]) : [];
const kid = (el: Element | null, name: string): Element | null => kids(el, name)[0] ?? null;
const textOf = (el: Element | null): string => (el?.textContent ?? '').trim();
const path = (el: Element | null, ...names: string[]): Element | null => names.reduce<Element | null>((e, n) => kid(e, n), el);

function paramValue(param: Element): ParamValue {
  const value = path(param, 'AttributeList', 'Value');
  if (value) return { kind: 'const', type: value.getAttribute('Type') ?? '', value: textOf(value) };
  const link = path(param, 'LinkList', 'Value');
  if (link) {
    const name = textOf(kid(link, 'Name'));
    if (name) return { kind: 'tag', name: name.replace(/^"(.*)"$/, '$1') };
    if ((link.getAttribute('TargetID') ?? '').startsWith('#')) return { kind: 'self' };
  }
  return { kind: 'empty' };
}

function limitOf(tag: Element, name: string): ParamValue {
  const attr = textOf(path(tag, 'AttributeList', name));
  if (attr !== '') return { kind: 'const', type: 'System.Double', value: attr };
  const link = path(tag, 'LinkList', name);
  const linked = textOf(kid(link, 'Name'));
  return linked ? { kind: 'tag', name: linked.replace(/^"(.*)"$/, '$1') } : { kind: 'empty' };
}

export function parseTagXml(text: string, source: string): AdvancedTag | null {
  const doc = parseXml(text, source);
  const tag = doc.getElementsByTagName('Hmi.Tag.Tag')[0];
  if (!tag) return null;
  const events: TagEventDef[] = [];
  for (const ev of kids(kid(tag, 'ObjectList'), 'Hmi.Event.Event')) {
    const entries: FunctionEntry[] = [];
    for (const entry of Array.from(ev.getElementsByTagName('Hmi.Event.FunctionListEntry'))) {
      const params = new Map<string, ParamValue>();
      for (const p of kids(kid(entry, 'ObjectList'), 'Hmi.Event.FunctionListEntryParameter')) {
        params.set(textOf(path(p, 'AttributeList', 'Name')), paramValue(p));
      }
      entries.push({ name: textOf(path(entry, 'AttributeList', 'Name')), type: textOf(path(entry, 'AttributeList', 'Type')), params });
    }
    events.push({ kind: textOf(path(ev, 'AttributeList', 'Name')), entries });
  }
  return {
    name: textOf(path(tag, 'AttributeList', 'Name')),
    dataType: textOf(path(tag, 'LinkList', 'DataType', 'Name')),
    limitUpper: limitOf(tag, 'LimitUpper2'),
    limitLower: limitOf(tag, 'LimitLower2'),
    events,
    source,
  };
}

/** Reads all tag files of an Openness export ZIP or a single XML file. */
export async function readAdvancedTags(input: ArrayBuffer | Uint8Array, fileName: string, log: NoticeLog): Promise<AdvancedTag[]> {
  const files: { name: string; text: string }[] = [];
  if (fileName.toLowerCase().endsWith('.xml')) {
    files.push({ name: fileName, text: new TextDecoder().decode(input) });
  } else {
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(input);
    } catch {
      throw new ConversionError('Die ZIP-Datei kann nicht gelesen werden. Bitte den Export erneut erstellen.');
    }
    for (const entry of Object.values(zip.files)) {
      if (!entry.dir && /\.xml$/i.test(entry.name)) files.push({ name: entry.name, text: await entry.async('string') });
    }
    const manifest = zip.file('manifest.json');
    if (manifest) {
      try {
        const m = JSON.parse(await manifest.async('string')) as { complete?: boolean; failedExports?: number; tagCount?: number };
        if (m.complete === false || (m.failedExports ?? 0) > 0) {
          log.warn(`Der Export ist laut manifest.json unvollständig (${m.failedExports ?? '?'} Variablen fehlgeschlagen). Ereignisse dieser Variablen fehlen im Skript.`);
        }
      } catch {
        log.warn('manifest.json konnte nicht gelesen werden – Vollständigkeit des Exports nicht geprüft.');
      }
    }
  }
  const tags: AdvancedTag[] = [];
  for (const f of files) {
    try {
      const t = parseTagXml(f.text, f.name);
      if (t) tags.push(t);
    } catch {
      log.warn(`Datei „${f.name}“ ist kein gültiges XML und wurde übersprungen.`);
    }
  }
  if (!tags.length) {
    throw new ConversionError('Keine HMI-Variablen gefunden. Erwartet wird der Openness-Export der HMI-Variablen (ZIP mit Tags/*.xml oder einzelne Variablen-XML).');
  }
  return tags;
}

// ------------------------------------------------------------------ script

const INDENT = '    ';
const js = (s: string) => JSON.stringify(s);
const get = (unified: string) => `HMIRuntime.Tags.SysFct.GetTagValue(${js(unified)})`;

function num(v: string): string | null {
  const n = Number(v.replace(',', '.'));
  return v.trim() !== '' && Number.isFinite(n) ? String(n) : null;
}

interface Ctx {
  self: string;
  known: Set<string>;
  log: NoticeLog;
  where: string;
  screensByNumber: boolean;
  /**
   * Tags written earlier in the same function list. Advanced reads the new value right away; in Unified a
   * read directly after SetTagValue is not guaranteed to return it, so the computed value is used instead.
   */
  written: Map<string, string>;
}

const wrap = (e: string) => (/^[\w.$-]+$/.test(e) || /^HMIRuntime\.Tags\.SysFct\.GetTagValue\("[^"]*"\)$/.test(e) ? e : `(${e})`);

/** JavaScript expression for a parameter, or null if it cannot be expressed. */
function expr(p: ParamValue | undefined, ctx: Ctx, dataType = ''): string | null {
  if (!p || p.kind === 'empty') return null;
  if (p.kind === 'self') return ctx.written.has(ctx.self) ? wrap(ctx.written.get(ctx.self)!) : 'value';
  if (p.kind === 'tag') {
    const u = unifiedName(p.name);
    if (!ctx.known.has(u)) ctx.log.warn(`Variable „${p.name}“ ist nicht im Export enthalten – Name im Skript bitte prüfen („${u}“).`, ctx.where);
    if (ctx.written.has(u)) return wrap(ctx.written.get(u)!);
    return u === ctx.self ? 'value' : get(u);
  }
  if (/Double|Int|Single|Decimal/.test(p.type)) {
    const n = num(p.value);
    if (n === null) return null;
    return dataType === 'Bool' ? String(n !== '0') : n;
  }
  if (/Boolean/.test(p.type)) return p.value.toLowerCase() === 'true' ? 'true' : 'false';
  return js(p.value);
}

/** Y = a * X + b, written as readable as the hand-written script ("value + 12", "10000 * plcNo + runNo"). */
function linear(a: string, x: string, b: string): string {
  let out = a === '1' ? x : a === '0' ? '' : `${a} * ${x}`;
  if (b !== '0') {
    if (out === '') out = b;
    else if (b.startsWith('-')) out = `${out} - ${b.slice(1)}`;
    else out = `${out} + ${b}`;
  }
  return out === '' ? '0' : out;
}

interface Statement {
  code: string;
  /** Screen changes only run on a real event, not on the first call after runtime start. */
  onEventOnly: boolean;
}

function setTag(target: string, value: string, ctx: Ctx): Statement {
  ctx.written.set(target, value);
  return { code: `HMIRuntime.Tags.SysFct.SetTagValue(${js(target)}, ${value});`, onEventOnly: false };
}

function statement(e: FunctionEntry, ctx: Ctx, dataTypes: Map<string, string>): Statement | null {
  const p = (n: string) => e.params.get(n);
  const targetName = (n: string) => {
    const t = p(n);
    if (t?.kind === 'self') return ctx.self;
    if (t?.kind === 'tag') {
      const u = unifiedName(t.name);
      if (!ctx.known.has(u)) ctx.log.warn(`Variable „${t.name}“ ist nicht im Export enthalten – Name im Skript bitte prüfen („${u}“).`, ctx.where);
      return u;
    }
    return null;
  };
  const fail = (why: string): Statement => {
    ctx.log.warn(`${e.name}: ${why} – als Kommentar ins Skript übernommen, bitte manuell umsetzen.`, ctx.where);
    const args = Array.from(e.params.entries())
      .map(([k, v]) => `${k} = ${v.kind === 'tag' ? v.name : v.kind === 'const' ? v.value : v.kind === 'self' ? '(diese Variable)' : ''}`)
      .join(', ');
    return { code: `// TODO ${e.name}(${args})`, onEventOnly: false };
  };
  if (e.type && e.type !== 'SystemFunction') return fail(`Typ „${e.type}“ (z. B. VB-Skript) wird nicht automatisch übersetzt`);

  switch (e.name) {
    case 'SetTag': {
      const target = targetName('Tag');
      const value = expr(p('Value'), ctx, target ? dataTypes.get(target) : '');
      if (!target || value === null) return fail('Ziel oder Wert fehlt');
      return setTag(target, value, ctx);
    }
    case 'LinearScaling':
    case 'InvertLinearScaling': {
      const target = targetName(e.name === 'LinearScaling' ? 'Y' : 'X');
      const x = expr(p(e.name === 'LinearScaling' ? 'X' : 'Y'), ctx);
      const a = expr(p('a'), ctx);
      const b = expr(p('b'), ctx);
      if (!target || x === null || a === null || b === null) return fail('Parameter fehlen');
      // InvertLinearScaling: X = (Y - b) / a
      const formula = e.name === 'LinearScaling' ? linear(a, x, b) : `(${x} - ${b}) / ${wrap(a)}`;
      return setTag(target, formula, ctx);
    }
    case 'SetBit':
    case 'ResetBit':
    case 'InvertBit': {
      const target = targetName('Tag');
      if (!target) return fail('Variable fehlt');
      const current = expr({ kind: 'tag', name: target }, ctx) ?? get(target);
      const value = e.name === 'SetBit' ? 'true' : e.name === 'ResetBit' ? 'false' : `!${current}`;
      return setTag(target, value, ctx);
    }
    case 'IncreaseValue':
    case 'DecreaseValue': {
      const target = targetName('Tag');
      const value = expr(p('Value'), ctx);
      if (!target || value === null) return fail('Variable oder Wert fehlt');
      const current = expr({ kind: 'tag', name: target }, ctx) ?? get(target);
      return setTag(target, `${current} ${e.name === 'IncreaseValue' ? '+' : '-'} ${wrap(value)}`, ctx);
    }
    case 'ActivateScreen': {
      const screen = p('Screen name');
      const name = screen?.kind === 'tag' ? screen.name : screen?.kind === 'const' ? screen.value : '';
      if (!name) return fail('Bildname fehlt');
      return { code: `HMIRuntime.UI.SysFct.ChangeScreen(${js(name)}, ".");`, onEventOnly: true };
    }
    case 'ActivateScreenByNumber': {
      const number = expr(p('Screen number'), ctx);
      if (number === null) return fail('Bildnummer fehlt');
      ctx.screensByNumber = true;
      return { code: `if (ev === 2 && TagEventScreens[${number}]) HMIRuntime.UI.SysFct.ChangeScreen(TagEventScreens[${number}], ".");`, onEventOnly: false };
    }
    default:
      return fail('Systemfunktion ohne automatisches Gegenstück');
  }
}

export interface ScriptStats {
  tags: number;
  withEvents: number;
  handlers: number;
  statements: number;
  todos: number;
}

export function buildScriptFromTags(tags: AdvancedTag[], log: NoticeLog, source: string): { script: string; stats: ScriptStats } {
  const known = new Set(tags.map((t) => unifiedName(t.name)));
  const dataTypes = new Map(tags.map((t) => [unifiedName(t.name), t.dataType]));
  const usedNames = new Set<string>();
  const fns: string[][] = [];
  const calls: string[] = [];
  const stats: ScriptStats = { tags: tags.length, withEvents: 0, handlers: 0, statements: 0, todos: 0 };
  let screensByNumber = false;

  const sorted = [...tags].sort((a, b) => a.name.localeCompare(b.name, 'de'));
  for (const tag of sorted) {
    const events = tag.events.filter((e) => e.entries.length);
    if (!events.length) continue;
    stats.withEvents++;
    const self = unifiedName(tag.name);
    for (const ev of events) {
      const ctx: Ctx = { self, known, log, where: tag.source.split('/').pop() ?? tag.source, screensByNumber: false, written: new Map() };
      const label = EVENT_LABEL[ev.kind];
      if (!label) {
        log.warn(`Ereignis „${ev.kind}“ an „${tag.name}“ wird nicht unterstützt und fehlt im Skript.`, ctx.where);
        continue;
      }
      let condition = '';
      if (ev.kind !== 'ChangeValue') {
        const limit = expr(ev.kind === 'AboveUpperLimit' ? tag.limitUpper : tag.limitLower, ctx);
        if (limit === null) {
          log.warn(`„${label}“ an „${tag.name}“ hat keinen Grenzwert und fehlt im Skript.`, ctx.where);
          continue;
        }
        condition = `value ${ev.kind === 'AboveUpperLimit' ? '>' : '<'} ${limit}`;
      }
      const statements = ev.entries.map((e) => statement(e, ctx, dataTypes)).filter((s): s is Statement => !!s);
      screensByNumber ||= ctx.screensByNumber;

      let fnName = functionName(self) + (ev.kind === 'ChangeValue' ? '' : ev.kind === 'AboveUpperLimit' ? '_Ueber' : '_Unter');
      for (let n = 2; usedNames.has(fnName); n++) fnName = `${fnName}_${n}`;
      usedNames.add(fnName);
      calls.push(`${INDENT}${fnName}();`);

      const key = condition ? `${self} ${condition.replace('value ', '')}` : self;
      const lines = [`${INDENT}// ${tag.name} – ${label}${condition ? ` (${condition.replace('value', 'Wert')})` : ''}`, `${INDENT}function ${fnName}() {`];
      lines.push(`${INDENT}${INDENT}var value = ${get(self)};`);
      lines.push(`${INDENT}${INDENT}var ev = TagEventCheck(${js(key)}, ${condition || 'value'});`);
      lines.push(`${INDENT}${INDENT}if (!ev${condition ? ` || !(${condition})` : ''}) return;`);
      for (const s of statements) {
        if (s.code.startsWith('// TODO')) stats.todos++;
        lines.push(`${INDENT}${INDENT}${s.onEventOnly ? `if (ev === 2) ${s.code}` : s.code}`);
        stats.statements++;
      }
      lines.push(`${INDENT}}`);
      fns.push(lines);
      stats.handlers++;
    }
  }

  const out: string[] = [
    `// ${SCRIPT_NAME} – erzeugt mit dem WinCC Unified List Converter aus „${source}“`,
    `// Ersatz für die Variablen-Ereignisse aus WinCC Advanced: ${stats.withEvents} Variablen, ${stats.handlers} Ereignisse.`,
    '// TagEvent() zyklisch aufrufen, z. B. über eine Aufgabe im Scheduler.',
    '',
    '// Globale Definition – in WinCC Unified in den Bereich „Globale Definitionen“ des Skriptmoduls übernehmen.',
    '// Merkt sich die letzten Werte, damit jedes Ereignis nur beim Eintreten ausgeführt wird.',
    'var TagEventLast = {};',
    '',
    'export function TagEvent(parameter1, parameter2) {',
  ];
  if (screensByNumber) {
    out.push(`${INDENT}// Bildnummern aus WinCC Advanced – Unified kennt keine Bildnummern. Bitte Nummer und Bildname eintragen, z. B. 1010: "Uebersicht".`);
    out.push(`${INDENT}var TagEventScreens = {`);
    out.push(`${INDENT}};`);
    out.push('');
    log.warn('ActivateScreenByNumber: Unified kennt keine Bildnummern. Im Skript oben in „TagEventScreens“ die Zuordnung Bildnummer → Bildname eintragen.');
  }
  out.push(...calls, '');
  out.push(`${INDENT}// 0 = nichts zu tun, 1 = erster Aufruf nach dem Start, 2 = Ereignis eingetreten`);
  out.push(`${INDENT}function TagEventCheck(key, state) {`);
  out.push(`${INDENT}${INDENT}var first = !(key in TagEventLast);`);
  out.push(`${INDENT}${INDENT}var changed = !first && TagEventLast[key] !== state;`);
  out.push(`${INDENT}${INDENT}TagEventLast[key] = state;`);
  out.push(`${INDENT}${INDENT}return first ? 1 : changed ? 2 : 0;`);
  out.push(`${INDENT}}`);
  for (const f of fns) out.push('', ...f);
  out.push('}', '');
  return { script: out.join('\n'), stats };
}

export async function convertAdvancedEvents(input: ArrayBuffer | Uint8Array, fileName: string): Promise<ConversionResult> {
  const log = new NoticeLog();
  const tags = await readAdvancedTags(input, fileName, log);
  const { script, stats } = buildScriptFromTags(tags, log, fileName);
  if (!stats.handlers) {
    throw new ConversionError(`In den ${tags.length} Variablen ist kein Ereignis mit Funktionsliste hinterlegt – es gibt nichts umzusetzen.`);
  }
  return {
    data: new TextEncoder().encode(script),
    fileName: SCRIPT_NAME,
    text: script,
    notices: log.items,
    summary:
      `Skript erstellt: ${stats.handlers} Ereignisse an ${stats.withEvents} von ${stats.tags} Variablen, ${stats.statements} Anweisungen` +
      (stats.todos ? `, davon ${stats.todos} zum manuellen Nacharbeiten.` : '.'),
  };
}
