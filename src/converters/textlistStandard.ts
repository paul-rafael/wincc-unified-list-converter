/**
 * Textlisten (Standard) – port of Modul3.ReplaceSpecialCharsInWorkbook (Button "HMI TextList Konverter (neu)").
 *
 * Sheet "TextList" column A and sheet "TextListEntry" column B, each from row 2 to its last used row:
 * "#", "/", "." → "_", "[" → "(", "]" → ")" (Range.Replace, part match, in this order).
 *
 * Extension (not part of the macro): the "FieldInfos" column of "TextListEntry" is converted from the
 * Comfort/Advanced format (CommonTextList, CommonTagDisplayFormat) to the Unified format
 * (ILCommonTagDisplayFormatWithTextList, ILCommonTagDisplayFormat), as found in Unified text list exports.
 */
import { replaceNonAlnum } from '../vba/vba';
import { Workbook } from '../xlsx/workbook';
import { assertExportKind, assertExtension, ConversionResult, NoticeLog, stemOf, where, writeText } from './common';

const TARGETS: { sheet: string; col: number }[] = [
  { sheet: 'TextList', col: 1 },
  { sheet: 'TextListEntry', col: 2 },
];

const REPLACEMENTS: [string, string][] = [
  ['#', '_'],
  ['/', '_'],
  ['.', '_'],
  ['[', '('],
  [']', ')'],
];

export function replaceSpecialChars(s: string): string {
  let out = s;
  for (const [what, by] of REPLACEMENTS) out = out.split(what).join(by);
  return out;
}

const REF_RE = /<ref\b[^>]*>/g;
const ATTR_RE = /([A-Za-z]+)\s*=\s*("(?:[^"]|"")*"|[^;>]*?)\s*;/g;

/** Unified display defaults for a text list field (Unified exports use these for migrated fields). */
const TEXTLIST_LENGTH = 5;
const MAX_HEADER_COLS = 200;

function parseRef(ref: string): Map<string, string> {
  const attrs = new Map<string, string>();
  for (const m of ref.matchAll(ATTR_RE)) {
    let v = m[2];
    if (v.length >= 2 && v.startsWith('"') && v.endsWith('"')) v = v.slice(1, -1).split('""').join('"');
    attrs.set(m[1].toLowerCase(), v);
  }
  return attrs;
}

/** HMI tag names in Unified: every character outside [A-Za-z0-9] becomes "_" (same rule as "HMI-Variablen"). */
export function unifiedTagName(tag: string): string {
  return replaceNonAlnum(tag);
}

/**
 * Converts one `<ref …>` from the classic to the Unified format.
 * Returns the ref unchanged if it is already in Unified format, null if its format is unknown.
 */
export function convertFieldInfoRef(ref: string): string | null {
  const a = parseRef(ref);
  const id = a.get('id') ?? '0';
  const type = a.get('type') ?? '';
  const tag = unifiedTagName(a.get('tag') ?? '');
  if (type.startsWith('IL')) return ref;
  if (type === 'CommonTextList') {
    const list = replaceSpecialChars(a.get('textlist') ?? '');
    return `<ref id = ${id}; type = ILCommonTagDisplayFormatWithTextList; TextList = ${list}; Tag = ${tag}; DisplayType = Textlist; Length = ${TEXTLIST_LENGTH}; Precision = 0; Alignment = Right; ZeroPadding = False;>`;
  }
  if (type === 'CommonTagDisplayFormat' && (a.get('displaytype') ?? '') === 'Text') {
    const format = a.get('displayformat') ?? '';
    if (!/^\?+$/.test(format)) return null;
    return `<ref id = ${id}; type = ILCommonTagDisplayFormat; Tag = ${tag}; DisplayType = Text; Length = ${format.length}; Precision = 0; Alignment = Left; ZeroPadding = False;>`;
  }
  return null;
}

/** Converts every ref of a FieldInfos cell; `unknown` collects refs that were left unchanged. */
export function convertFieldInfos(text: string, unknown: string[] = []): string {
  return text.replace(REF_RE, (ref) => {
    const next = convertFieldInfoRef(ref);
    if (next === null) unknown.push(ref);
    return next ?? ref;
  });
}

/** InitialFileName = wb.Name; the file is always saved as .xlsx. */
export function standardOutputName(fileName: string): string {
  return `${stemOf(fileName)}.xlsx`;
}

export async function convertTextlistStandard(input: ArrayBuffer | Uint8Array, fileName: string): Promise<ConversionResult> {
  assertExtension(fileName, ['.xlsx', '.xlsm']);
  const wb = await Workbook.load(input);
  await assertExportKind(wb, 'textlist');
  const log = new NoticeLog();
  const parts: string[] = [];

  for (const { sheet, col } of TARGETS) {
    const idx = wb.sheetIndexByName(sheet);
    if (idx < 0) {
      log.warn(`Blatt „${sheet}“ wurde nicht gefunden und wurde übersprungen.`);
      continue;
    }
    const ws = await wb.sheet(idx);
    if (ws.isProtected) {
      // The macro's error handler catches this and continues with the next sheet.
      log.warn(`Blatt „${ws.name}“ ist schreibgeschützt und wurde übersprungen.`);
      continue;
    }
    const lastRow = ws.lastRowInColumn(col);
    let changed = 0;
    for (let r = 2; r <= lastRow; r++) {
      const cell = ws.getCell(r, col);
      if (cell.hasFormula) {
        log.warn('Formelzelle wurde nicht verändert. Bitte manuell prüfen.', where(ws, r, col));
        continue;
      }
      if (cell.value.kind === 'error') {
        log.warn(`Fehlerwert ${cell.value.text} wurde nicht verändert. Bitte manuell prüfen.`, where(ws, r, col));
        continue;
      }
      if (cell.value.kind !== 'string') continue;
      const next = replaceSpecialChars(cell.value.text);
      if (next === cell.value.text) continue;
      if (writeText(ws, r, col, next, log)) changed++;
    }
    parts.push(`„${ws.name}“: ${changed} geändert`);
  }

  const fieldInfos = await convertFieldInfoColumn(wb, log);
  if (fieldInfos > 0) parts.push(`FieldInfos: ${fieldInfos} ins Unified-Format umgewandelt`);

  await wb.stripMacros();
  const data = await wb.save();
  return {
    data,
    fileName: standardOutputName(fileName),
    notices: log.items,
    summary: parts.length ? `Namen bereinigt – ${parts.join(', ')}.` : 'Keines der Blätter „TextList“ / „TextListEntry“ wurde gefunden.',
  };
}

async function convertFieldInfoColumn(wb: Workbook, log: NoticeLog): Promise<number> {
  const idx = wb.sheetIndexByName('TextListEntry');
  if (idx < 0) return 0;
  const ws = await wb.sheet(idx);
  if (ws.isProtected) return 0;
  // TIA exports place FieldInfos after the language columns, so its position depends on the project.
  let col = 0;
  for (let c = 1; c <= MAX_HEADER_COLS && !col; c++) {
    const h = ws.getCell(1, c).value;
    if (h.kind === 'string' && h.text.trim().toLowerCase() === 'fieldinfos') col = c;
  }
  if (!col) return 0;

  let changed = 0;
  for (let r = 2; r <= ws.lastRowInColumn(col); r++) {
    const cell = ws.getCell(r, col);
    if (cell.hasFormula || cell.value.kind !== 'string') continue;
    const unknown: string[] = [];
    const next = convertFieldInfos(cell.value.text, unknown);
    for (const ref of unknown) log.warn(`FieldInfo „${ref}“ hat ein unbekanntes Format und wurde nicht verändert. Bitte manuell prüfen.`, where(ws, r, col));
    if (next !== cell.value.text && writeText(ws, r, col, next, log)) changed++;
  }
  return changed;
}
