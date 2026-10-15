import { ConversionError } from '../errors';
import { coerceForCell } from '../vba/vba';
import { cellRef } from '../xlsx/ref';
import type { CellValue, Sheet } from '../xlsx/workbook';

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

export type ConverterId = 'tags' | 'textlistLegacy' | 'textlistStandard' | 'alarms';

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
