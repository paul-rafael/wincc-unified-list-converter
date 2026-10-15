/**
 * HMI-Variablen bereinigen – port of Modul1.SonderzeichenMitKnopf (Button "HMI TagList Konverter").
 *
 * First sheet, column A from row 1 (header included) to the last used row of column A:
 * every character that is not [A-Za-z0-9] becomes "_". The workbook is otherwise unchanged.
 */
import { cellToVbaString, replaceNonAlnum } from '../vba/vba';
import { Workbook } from '../xlsx/workbook';
import { assertExtension, ConversionResult, NoticeLog, protectedSheetError, where, writeText } from './common';

export const TAGS_OUTPUT_NAME = 'HMITagList_MIGUNI.xlsx';

export async function convertTags(input: ArrayBuffer | Uint8Array, fileName: string): Promise<ConversionResult> {
  assertExtension(fileName, ['.xlsx']);
  const wb = await Workbook.load(input);
  const ws = await wb.sheet(0);
  if (ws.isProtected) throw protectedSheetError(ws.name);
  const log = new NoticeLog();

  const lastRow = ws.lastRowInColumn(1);
  let changed = 0;
  for (let r = 1; r <= lastRow; r++) {
    const cell = ws.getCell(r, 1);
    const { text, certain } = cellToVbaString(cell.value, ws.isDateFormatted(r, 1), wb.date1904, where(ws, r, 1));
    if (!certain) {
      log.warn(`Wert „${text}“ (keine Textzelle) wurde wie im Original in Text umgewandelt; Ergebnis bitte prüfen.`, where(ws, r, 1));
    }
    if (writeText(ws, r, 1, replaceNonAlnum(text), log)) changed++;
  }

  const data = await wb.save();
  return {
    data,
    fileName: TAGS_OUTPUT_NAME,
    notices: log.items,
    summary: `Blatt „${ws.name}“, Spalte A: ${lastRow} Zeilen geprüft, ${changed} Namen geändert.`,
  };
}
