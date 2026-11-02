import { ConversionError } from '../errors';
import { coerceForCell } from '../vba/vba';
import { cellRef } from '../xlsx/ref';
import type { CellValue, Sheet, Workbook } from '../xlsx/workbook';

export interface Notice {
  level: 'info' | 'warn';
  text: string;
  cell?: string;
}

export interface ConversionResult {
  data: Uint8Array;
  fileName: string;
  notices: Notice[];
  summary: string;
}

export type ConverterId = 'tags' | 'tagEvents' | 'textlists' | 'alarms' | 'lineBreaks';

export class NoticeLog {
  readonly items: Notice[] = [];
  warn(text: string, cell?: string): void {
    this.items.push({ level: 'warn', text, cell });
  }
  info(text: string, cell?: string): void {
    this.items.push({ level: 'info', text, cell });
  }
}

export function extensionOf(fileName: string): string {
  const i = fileName.lastIndexOf('.');
  return i < 0 ? '' : fileName.slice(i).toLowerCase();
}

export function stemOf(fileName: string): string {
  const i = fileName.lastIndexOf('.');
  return i < 0 ? fileName : fileName.slice(0, i);
}

export function assertExtension(fileName: string, allowed: string[]): void {
  const ext = extensionOf(fileName);
  if (ext === '.xls') {
    throw new ConversionError('Alte .xls-Dateien werden nicht unterstützt. Bitte die Datei in Excel als .xlsx speichern oder direkt aus dem TIA Portal als .xlsx exportieren.');
  }
  if (!allowed.includes(ext)) {
    throw new ConversionError(`Dateityp „${ext || 'ohne Endung'}“ wird hier nicht unterstützt. Erlaubt: ${allowed.join(', ')}.`);
  }
}

export function sameValue(a: CellValue, b: CellValue): boolean {
  if (a.kind !== b.kind) return false;
  switch (a.kind) {
    case 'empty':
      return true;
    case 'string':
      return a.text === (b as typeof a).text;
    case 'number':
    case 'bool':
      return a.value === (b as typeof a).value;
    case 'error':
      return a.text === (b as typeof a).text;
  }
}

/** `cell.Value = text` on an existing sheet, including Excel's text interpretation. Returns true if the value changed. */
export function writeText(sheet: Sheet, row: number, col: number, text: string, log: NoticeLog): boolean {
  const before = sheet.getCell(row, col);
  const { value, warning } = coerceForCell(text, sheet.isTextFormatted(row, col));
  if (warning) log.warn(warning, `${sheet.name}!${cellRef(row, col)}`);
  const changed = before.hasFormula || !sameValue(before.value, value);
  sheet.setCell(row, col, value);
  return changed;
}

export function where(sheet: Sheet, row: number, col: number): string {
  return `${sheet.name}!${cellRef(row, col)}`;
}

export function protectedSheetError(name: string): ConversionError {
  return new ConversionError(`Das Blatt „${name}“ ist schreibgeschützt. Bitte den Blattschutz in Excel aufheben und die Datei erneut hochladen.`);
}

/** Kind of TIA Portal export, recognised from sheet names and column headers. */
export type ExportKind = 'textlist' | 'alarms' | 'tags' | 'events';

const KIND_HINT: Record<ExportKind, string> = {
  textlist: 'ein Textlisten-Export – bitte „Textlisten“ wählen',
  alarms: 'ein Meldungs-Export – bitte „HMI-Meldungen“ oder „Meldetexte umbrechen“ wählen',
  tags: 'ein Variablen-Export – bitte „HMI-Variablen“ oder „Variablen-Ereignisse“ wählen',
  events: 'eine Ereignis-Tabelle – bitte „Variablen-Ereignisse“ wählen',
};

/** Header texts of the first row of a sheet (only the leading columns, where TIA puts its fixed headers). */
export function headerTexts(sheet: Sheet, maxCols = 60): string[] {
  const out: string[] = [];
  for (let c = 1; c <= maxCols; c++) {
    const v = sheet.getCell(1, c).value;
    out.push(v.kind === 'string' ? v.text.trim() : '');
  }
  return out;
}

export async function detectExportKind(wb: Workbook): Promise<ExportKind | null> {
  if (wb.sheetIndexByName('Ereignisse') >= 0) return 'events';
  if (wb.sheetIndexByName('TextList') >= 0 && wb.sheetIndexByName('TextListEntry') >= 0) return 'textlist';
  const headers = headerTexts(await wb.sheet(0)).map((h) => h.toLowerCase());
  if (headers.some((h) => h.startsWith('alarm text'))) return 'alarms';
  if (headers.includes('plc tag') || headers.includes('datatype') || headers.includes('data type')) return 'tags';
  return null;
}

/**
 * Stops with a clear message when the file is recognisably a different kind of export than the
 * selected function expects. Unrecognised files are processed as before (like the original macros).
 */
export async function assertExportKind(wb: Workbook, expected: ExportKind): Promise<void> {
  const kind = await detectExportKind(wb);
  if (kind && kind !== expected) throw new ConversionError(`Diese Datei ist ${KIND_HINT[kind]}.`);
}
