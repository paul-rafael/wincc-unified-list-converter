/**
 * Meldetexte umbrechen (Beta) – port of Modul7.FormatHMITextInFile / FormatHMITextWithErrorCode
 * (Button "BETA AZU", automatischer Zeilenumbruch).
 *
 * First sheet, rows 2..last used row of column A: the alarm text in column C is split into
 * "error code ¶ line 2 ¶ line 3" (max. 80 characters per line, forced breaks before keywords),
 * the cell gets "Zeilenumbruch" and the row height is fitted automatically.
 */
import { cellToVbaString, inStr, vbaMid, vbaTrim } from '../vba/vba';
import { Workbook } from '../xlsx/workbook';
import { assertExtension, ConversionResult, NoticeLog, protectedSheetError, where, writeText } from './common';

export const LINE_BREAK = '¶'; // "¶" – line break marker for the HMI
const BREAK_WORDS = ['Position:', 'Geschw.:', 'X:', 'Y:', 'Z:'];

export function formatHmiText(alarmText: string, maxLen = 80): string {
  // VBA Replace is case-sensitive and replaces all occurrences.
  alarmText = vbaTrim(alarmText).split('MFS_CON').join('MFS-CON');

  // Error code = everything up to the first "_" (the "_" itself is dropped).
  let pos = inStr(1, alarmText, '_');
  if (pos === 0) pos = alarmText.length + 1;
  const firstLine = alarmText.slice(0, pos - 1);
  const restText = vbaMid(alarmText, pos + 1);

  // Split(…, " ") keeps empty entries; Split("") returns no entries at all.
  const words = restText === '' ? [] : restText.split(' ');
  let currentLine = '';
  let secondLine = '';
  let thirdLine = '';

  for (const word of words) {
    // `word Like bw & "*"`: binary prefix match (none of the keywords contains Like wildcards).
    const doBreak = BREAK_WORDS.some((bw) => word.startsWith(bw));
    if (doBreak || `${currentLine} ${word}`.length > maxLen) {
      if (secondLine === '') secondLine = vbaTrim(currentLine);
      else thirdLine = vbaTrim(currentLine); // a later break overwrites line 3 (original behaviour)
      currentLine = word;
    } else {
      currentLine = currentLine === '' ? word : `${currentLine} ${word}`;
    }
  }

  if (secondLine === '') secondLine = vbaTrim(currentLine);
  else if (thirdLine === '') thirdLine = vbaTrim(currentLine);
  else thirdLine = vbaTrim(`${thirdLine} ${currentLine}`);

  let out = firstLine;
  if (secondLine !== '') out += LINE_BREAK + secondLine;
  if (thirdLine !== '') out += LINE_BREAK + thirdLine;
  return out;
}

export async function convertLineBreaks(input: ArrayBuffer | Uint8Array, fileName: string): Promise<ConversionResult> {
  assertExtension(fileName, ['.xlsx', '.xlsm']);
  const wb = await Workbook.load(input);
  const ws = await wb.sheet(0);
  if (ws.isProtected) throw protectedSheetError(ws.name);
  const log = new NoticeLog();

  const lastRow = ws.lastRowInColumn(1);
  let changed = 0;
  for (let i = 2; i <= lastRow; i++) {
    const src = cellToVbaString(ws.getCell(i, 3).value, ws.isDateFormatted(i, 3), wb.date1904, where(ws, i, 3));
    if (!src.certain) log.warn(`Wert „${src.text}“ ist kein Text und wurde wie im Original in Text umgewandelt.`, where(ws, i, 3));
    if (writeText(ws, i, 3, formatHmiText(src.text, 80), log)) changed++;
    ws.setWrapText(i, 3);
    ws.autoFitRow(i);
  }

  // The macro edits the opened workbook and leaves saving to the user, so name and format stay the same.
  const data = await wb.save();
  return {
    data,
    fileName,
    notices: log.items,
    summary: `Blatt „${ws.name}“, Spalte C: ${Math.max(0, lastRow - 1)} Meldetexte verarbeitet, ${changed} umgebrochen.`,
  };
}
