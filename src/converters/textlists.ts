/**
 * Textlisten – one entry point for both text list functions of the Excel tool.
 *
 * The export format decides which one applies:
 * - "TextListEntry" has a "FieldInfos" column (current TIA exports): names are cleaned and the field
 *   references are kept and converted to the Unified format (textlistStandard).
 * - otherwise (older exports): placeholders are replaced by fixed texts and empty languages are
 *   filled, as the macro "HMI TextList Konverter (alt)" does (textlistLegacy).
 */
import { Workbook } from '../xlsx/workbook';
import { assertExtension, ConversionResult, headerTexts } from './common';
import { convertTextlistLegacy } from './textlistLegacy';
import { convertTextlistStandard } from './textlistStandard';

export type TextlistMode = 'unified' | 'legacy';

export async function detectTextlistMode(wb: Workbook): Promise<TextlistMode> {
  const idx = wb.sheetIndexByName('TextListEntry');
  if (idx < 0) return 'unified';
  const headers = headerTexts(await wb.sheet(idx), 200);
  return headers.some((h) => h.toLowerCase() === 'fieldinfos') ? 'unified' : 'legacy';
}

export async function convertTextlists(input: ArrayBuffer | Uint8Array, fileName: string): Promise<ConversionResult> {
  assertExtension(fileName, ['.xlsx', '.xlsm']);
  const mode = await detectTextlistMode(await Workbook.load(input));
  if (mode === 'legacy') {
    const res = await convertTextlistLegacy(input, fileName);
    return { ...res, summary: `Älteres Exportformat erkannt (ohne Spalte „FieldInfos“) – Platzhalter bereinigt. ${res.summary}` };
  }
  const res = await convertTextlistStandard(input, fileName);
  return { ...res, summary: `Aktuelles Exportformat erkannt – ${res.summary}` };
}
