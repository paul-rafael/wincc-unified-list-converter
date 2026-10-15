/**
 * Textlisten (Legacy) – port of Modul2.Clean_TextList_Sheet (Button "HMI TextList Konverter (alt)").
 *
 * Sheet "TextListEntry", rows 2..last used row of column E, columns E–J.
 * Column I is the fallback text, column K holds the field info with "TextList = …".
 */
import { ConversionError } from '../errors';
import { cellToVbaString, VBS_WS, vbaTrim, worksheetTrim } from '../vba/vba';
import { Workbook } from '../xlsx/workbook';
import { assertExtension, ConversionResult, NoticeLog, protectedSheetError, stemOf, where, writeText } from './common';

const DASH = '—'; // "—" as written by the macro
const NO_VALUE = '<no value>';
const FIELD_REF_0 = '<field ref="0" />';
const TEXTLIST_RE = new RegExp(`TextList${VBS_WS}*=${VBS_WS}*([^;>]+)`, 'i');
const TAG_RE = /<[^>]+>/g;

/** Replace(srcPath, ".xlsx", "_MIGUNI.xlsx"); for .xlsm the result is saved as .xlsx. */
export function legacyOutputName(fileName: string): string {
  if (fileName.includes('.xlsx')) return fileName.split('.xlsx').join('_MIGUNI.xlsx');
  return `${stemOf(fileName)}_MIGUNI.xlsx`;
}

/** The cell rule of Clean_TextList_Sheet; returns null where the macro leaves the cell untouched. */
export function legacyCellValue(col: number, txt: string, kText: string, fallbackText: string, textListVal: string): string | null {
  const lower = txt.toLowerCase();
  if (col === 5 && lower !== NO_VALUE && lower !== FIELD_REF_0) return null;
  if (lower === FIELD_REF_0) return textListVal !== '' ? textListVal : DASH;
  if (lower === NO_VALUE) {
    if (kText === '') return fallbackText.toLowerCase() !== NO_VALUE && fallbackText !== '' ? fallbackText : DASH;
    return textListVal !== '' ? textListVal : DASH;
  }
  const cleaned = worksheetTrim(txt.replace(TAG_RE, ''));
  return cleaned === '' ? DASH : cleaned;
}

export function extractTextListValue(kText: string): string {
  if (kText === '') return '';
  const m = TEXTLIST_RE.exec(kText);
  return m ? vbaTrim(m[1]) : '';
}

export async function convertTextlistLegacy(input: ArrayBuffer | Uint8Array, fileName: string): Promise<ConversionResult> {
  assertExtension(fileName, ['.xlsx', '.xlsm']);
  const wb = await Workbook.load(input);
  const idx = wb.sheetIndexByName('TextListEntry');
  if (idx < 0) throw new ConversionError('Arbeitsblatt „TextListEntry“ nicht gefunden! Bitte eine Textlisten-Exportdatei aus dem TIA Portal verwenden.');
  const ws = await wb.sheet(idx);
  if (ws.isProtected) throw protectedSheetError(ws.name);
  const log = new NoticeLog();

  const text = (r: number, c: number) => cellToVbaString(ws.getCell(r, c).value, ws.isDateFormatted(r, c), wb.date1904, where(ws, r, c));

  const lastRow = ws.lastRowInColumn(5);
  let changed = 0;
  for (let r = 2; r <= lastRow; r++) {
    const fallbackText = vbaTrim(text(r, 9).text);
    const kText = vbaTrim(text(r, 11).text);
    const textListVal = extractTextListValue(kText);

    for (let c = 5; c <= 10; c++) {
      const src = text(r, c);
      const next = legacyCellValue(c, vbaTrim(src.text), kText, fallbackText, textListVal);
      if (next === null) continue;
      if (!src.certain) {
        log.warn(`Zelle enthält eine Zahl, ein Datum oder einen Wahrheitswert („${src.text}“) und wurde unverändert gelassen. Bitte manuell prüfen.`, where(ws, r, c));
        continue;
      }
      if (writeText(ws, r, c, next, log)) changed++;
    }
  }

  await wb.stripMacros();
  const data = await wb.save();
  return {
    data,
    fileName: legacyOutputName(fileName),
    notices: log.items,
    summary: `Blatt „TextListEntry“: ${Math.max(0, lastRow - 1)} Zeilen verarbeitet, ${changed} Zellen geändert.`,
  };
}
