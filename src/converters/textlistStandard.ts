/**
 * Textlisten (Standard) – port of Modul3.ReplaceSpecialCharsInWorkbook (Button "HMI TextList Konverter (neu)").
 *
 * Sheet "TextList" column A and sheet "TextListEntry" column B, each from row 2 to its last used row:
 * "#", "/", "." → "_", "[" → "(", "]" → ")" (Range.Replace, part match, in this order).
 */
import { Workbook } from '../xlsx/workbook';
import { assertExtension, ConversionResult, NoticeLog, stemOf, where, writeText } from './common';

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

/** InitialFileName = wb.Name; the file is always saved as .xlsx. */
export function standardOutputName(fileName: string): string {
  return `${stemOf(fileName)}.xlsx`;
}

export async function convertTextlistStandard(input: ArrayBuffer | Uint8Array, fileName: string): Promise<ConversionResult> {
  assertExtension(fileName, ['.xlsx', '.xlsm']);
  const wb = await Workbook.load(input);
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

  await wb.stripMacros();
  const data = await wb.save();
  return {
    data,
    fileName: standardOutputName(fileName),
    notices: log.items,
    summary: parts.length ? `Namen bereinigt – ${parts.join(', ')}.` : 'Keines der Blätter „TextList“ / „TextListEntry“ wurde gefunden.',
  };
}
