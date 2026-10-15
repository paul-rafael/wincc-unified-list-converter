/**
 * HMI-Meldungen – port of Modul8.ConvertAlarmList_Unified (Button "HMI-Alarm Konverter").
 *
 * Reads the first sheet of an exported alarm list and writes a new workbook with one sheet
 * "DiscreteAlarms" in the WinCC Unified column layout.
 */
import { cellToVbaString, coerceForCell, inStr, vbaMid, vbaTrim } from '../vba/vba';
import { cellRef } from '../xlsx/ref';
import { buildWorkbook } from '../xlsx/newWorkbook';
import type { CellValue, WriteValue } from '../xlsx/workbook';
import { Workbook } from '../xlsx/workbook';
import { assertExtension, ConversionResult, NoticeLog, where } from './common';

export const ALARMS_OUTPUT_NAME = 'Unified_HMI-Alarms.xlsx';
export const ALARMS_SHEET_NAME = 'DiscreteAlarms';

export const ALARM_HEADER: string[] = [
  'ID', 'Name', 'Alarm text [de-DE], Alarm text 1', 'FieldInfo [Alarm text 1]',
  'Class', 'Trigger tag', 'Trigger bit', 'Trigger mode',
  'Acknowledgement tag', 'Acknowledgement bit',
  'PLC acknowledgement tag', 'PLC acknowledgement bit',
  'Priority', 'Info text [de-DE], Info text',
  'Additional text 1 [de-DE], Alarm text 2', 'FieldInfo [Alarm text 2]',
  'Additional text 2 [de-DE], Alarm text 3', 'FieldInfo [Alarm text 3]',
  'Additional text 3 [de-DE], Alarm text 4', 'FieldInfo [Alarm text 4]',
  'Additional text 4 [de-DE], Alarm text 5', 'FieldInfo [Alarm text 5]',
  'Additional text 5 [de-DE], Alarm text 6', 'FieldInfo [Alarm text 6]',
  'Additional text 6 [de-DE], Alarm text 7', 'FieldInfo [Alarm text 7]',
  'Additional text 7 [de-DE], Alarm text 8', 'FieldInfo [Alarm text 8]',
  'Additional text 8 [de-DE], Alarm text 9', 'FieldInfo [Alarm text 9]',
  'Additional text 9 [de-DE], Alarm text 10', 'FieldInfo [Alarm text 10]',
  'Alarm parameter 1', 'Alarm parameter 2', 'Alarm parameter 3',
  'Alarm parameter 4', 'Alarm parameter 5', 'Alarm parameter 6',
  'Alarm parameter 7', 'Alarm parameter 8', 'Alarm parameter 9',
  'Alarm parameter 10', 'Area', 'Origin',
];

/** ExtractRefs: every `ref="…"` in the alarm text, in order (binary compare). */
export function extractRefs(text: string, context = ''): string[] {
  const out: string[] = [];
  let pos = 1;
  while (inStr(pos, text, 'ref="') > 0) {
    pos = inStr(pos, text, 'ref="') + 5;
    out.push(vbaMid(text, pos, inStr(pos, text, '"') - pos, context));
    pos = pos + 1;
  }
  return out;
}

/** ExtractTagByRef: "Tag = …;" after the first occurrence of "id = <refID>". */
export function extractTagByRef(fieldInfo: string, refID: string, context = ''): string {
  let startPos = inStr(1, fieldInfo, `id = ${refID}`);
  if (startPos > 0) {
    startPos = inStr(startPos, fieldInfo, 'Tag = ');
    if (startPos > 0) {
      startPos += 6;
      const endPos = inStr(startPos, fieldInfo, ';');
      return vbaTrim(vbaMid(fieldInfo, startPos, endPos - startPos, context));
    }
  }
  return '';
}

/** ExtractLengthByRef: "Length = …;" after the first occurrence of "id = <refID>", fallback "3". */
export function extractLengthByRef(fieldInfo: string, refID: string, context = ''): string {
  let startPos = inStr(1, fieldInfo, `id = ${refID}`);
  if (startPos > 0) {
    startPos = inStr(startPos, fieldInfo, 'Length = ');
    if (startPos > 0) {
      startPos += 'Length = '.length;
      const endPos = inStr(startPos, fieldInfo, ';');
      return vbaTrim(vbaMid(fieldInfo, startPos, endPos - startPos, context));
    }
  }
  return '3';
}

export interface FieldInfoResult {
  fieldInfo: string;
  /** Tag of the last reference with a tag (goes to "Alarm parameter 1"), '' if none. */
  parameter1: string;
}

export function buildFieldInfo(alarmText: string, fieldInfoOld: string, context = ''): FieldInfoResult {
  let newFieldInfo = '';
  let parameter1 = '';
  for (const refID of extractRefs(alarmText, context)) {
    const tagValue = extractTagByRef(fieldInfoOld, refID, context);
    if (tagValue !== '') {
      const paramLength = extractLengthByRef(fieldInfoOld, refID, context);
      newFieldInfo +=
        `<ref id = ${refID}; type = AlarmParameterWithOrWithoutCommonTextList; ` +
        'Parameter = Parameter 1; ' +
        `Tag = ${tagValue}; DisplayType = Decimal; Length = ${paramLength}; Precision = 0; ` +
        'Alignment = Right; ZeroPadding = False;>\r\n';
      parameter1 = tagValue;
    }
  }
  // VBA Trim removes spaces only, so the trailing CR LF stays.
  return { fieldInfo: vbaTrim(newFieldInfo), parameter1 };
}

export async function convertAlarms(input: ArrayBuffer | Uint8Array, fileName: string): Promise<ConversionResult> {
  assertExtension(fileName, ['.xlsx', '.xlsm']);
  const wb = await Workbook.load(input);
  const ws = await wb.sheet(0);
  const log = new NoticeLog();
  const grid = new Map<number, Map<number, WriteValue>>();

  const header = new Map<number, WriteValue>();
  ALARM_HEADER.forEach((h, i) => header.set(i + 1, { kind: 'string', text: h }));
  grid.set(1, header);

  const put = (row: Map<number, WriteValue>, r: number, c: number, text: string) => {
    const { value, warning } = coerceForCell(text, false);
    if (warning) log.warn(warning, `${ALARMS_SHEET_NAME}!${cellRef(r, c)}`);
    row.set(c, value);
  };

  const lastRow = ws.lastRowInColumn(1);
  let withParams = 0;
  for (let i = 2; i <= lastRow; i++) {
    const row = new Map<number, WriteValue>();
    grid.set(i, row);

    // A–G copied as values (H is overwritten below).
    for (let c = 1; c <= 7; c++) {
      const v: CellValue = ws.getCell(i, c).value;
      if (v.kind === 'string') put(row, i, c, v.text);
      else row.set(c, v);
      if (v.kind === 'number' && ws.isDateFormatted(i, c)) {
        log.warn('Datumswert wurde als Zahl übernommen (Excel würde ein Datumsformat setzen).', `${ALARMS_SHEET_NAME}!${cellRef(i, c)}`);
      }
    }
    const alarmText = cellToVbaString(ws.getCell(i, 3).value, ws.isDateFormatted(i, 3), wb.date1904, where(ws, i, 3)).text;
    const fieldInfoOld = cellToVbaString(ws.getCell(i, 4).value, ws.isDateFormatted(i, 4), wb.date1904, where(ws, i, 4)).text;

    row.set(8, { kind: 'string', text: 'On rising edge' });
    row.set(9, { kind: 'string', text: '<No value>' });
    row.set(11, { kind: 'string', text: '<No value>' });
    row.set(10, { kind: 'number', value: 0 });
    row.set(12, { kind: 'number', value: 0 });
    row.set(13, { kind: 'number', value: 0 });
    for (const c of [16, 18, 20, 22, 24, 26, 28, 30, 32]) row.set(c, { kind: 'empty' });
    for (let c = 14; c <= 33; c += 2) row.set(c, { kind: 'string', text: '<No value>' });
    for (let c = 34; c <= 42; c++) row.set(c, { kind: 'string', text: '<No value>' });
    row.set(43, { kind: 'string', text: 'HMI_RT_2::Alarming' });

    const { fieldInfo, parameter1 } = buildFieldInfo(alarmText, fieldInfoOld, `Zeile ${i}`);
    if (parameter1 !== '') {
      put(row, i, 33, parameter1);
      withParams++;
    }
    put(row, i, 4, fieldInfo);
  }

  const data = await buildWorkbook(ALARMS_SHEET_NAME, grid);
  return {
    data,
    fileName: ALARMS_OUTPUT_NAME,
    notices: log.items,
    summary: `${Math.max(0, lastRow - 1)} Meldungen konvertiert, davon ${withParams} mit Variablen im Meldetext.`,
  };
}
