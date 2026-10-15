import JSZip from 'jszip';
import { Workbook, type CellValue } from '../src/xlsx/workbook';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

export type Cell = string | number | boolean | null | { f: string; v?: string | number } | { err: string } | { s: number; v?: string | number } | { rich: string[] };

export interface SheetSpec {
  name: string;
  rows: Record<number, Record<string, Cell>>;
  extra?: string; // XML inserted after sheetData (e.g. sheetProtection)
  cols?: string;
}

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * Builds an Excel-like .xlsx (shared strings, styles with a Text and a Date format, rich text)
 * so the converters are tested against the structures Excel/TIA Portal write.
 * Style 0 = General, 1 = Text ("@"), 2 = date (numFmtId 14), 3 = bold font.
 */
export async function makeXlsx(sheets: SheetSpec[], opts: { macro?: boolean } = {}): Promise<Uint8Array> {
  const sst: string[] = [];
  const sstXml: string[] = [];
  const si = (s: string) => {
    sst.push(s);
    sstXml.push(`<si><t xml:space="preserve">${esc(s)}</t></si>`);
    return sst.length - 1;
  };
  const zip = new JSZip();
  sheets.forEach((sh, n) => {
    const rows = Object.keys(sh.rows).map(Number).sort((a, b) => a - b).map((r) => {
      const cells = Object.entries(sh.rows[r]).map(([col, v]) => {
        const ref = `${col}${r}`;
        if (v === null) return `<c r="${ref}" s="3"/>`;
        if (typeof v === 'string') return `<c r="${ref}" t="s"><v>${si(v)}</v></c>`;
        if (typeof v === 'number') return `<c r="${ref}"><v>${v}</v></c>`;
        if (typeof v === 'boolean') return `<c r="${ref}" t="b"><v>${v ? 1 : 0}</v></c>`;
        if ('err' in v) return `<c r="${ref}" t="e"><v>${v.err}</v></c>`;
        if ('rich' in v) {
          sst.push(v.rich.join(''));
          sstXml.push(`<si>${v.rich.map((t, i) => `<r>${i ? '<rPr><b/></rPr>' : ''}<t xml:space="preserve">${esc(t)}</t></r>`).join('')}</si>`);
          return `<c r="${ref}" t="s"><v>${sst.length - 1}</v></c>`;
        }
        if ('f' in v) {
          const val = v.v === undefined ? '' : typeof v.v === 'number' ? `<v>${v.v}</v>` : `<v>${esc(v.v)}</v>`;
          return `<c r="${ref}"${typeof v.v === 'string' ? ' t="str"' : ''}><f>${esc(v.f)}</f>${val}</c>`;
        }
        const val = v.v === undefined ? '' : typeof v.v === 'number' ? `<v>${v.v}</v>` : `<v>${si(v.v)}</v>`;
        return `<c r="${ref}" s="${v.s}"${typeof v.v === 'string' ? ' t="s"' : ''}>${val}</c>`;
      });
      return `<row r="${r}">${cells.join('')}</row>`;
    });
    zip.file(
      `xl/worksheets/sheet${n + 1}.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1"/>${sh.cols ?? ''}<sheetData>${rows.join('')}</sheetData>${sh.extra ?? ''}</worksheet>`,
    );
  });
  const wbType = opts.macro ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${opts.macro ? '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>' : ''}<Override PartName="/xl/workbook.xml" ContentType="${wbType}"/>${sheets.map((_, n) => `<Override PartName="/xl/worksheets/sheet${n + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`,
  );
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, n) => `<sheet name="${esc(s.name)}" sheetId="${n + 1}" r:id="rId${n + 10}"/>`).join('')}</sheets></workbook>`,
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, n) => `<Relationship Id="rId${n + 10}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n + 1}.xml"/>`).join('')}<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>${opts.macro ? '<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>' : ''}</Relationships>`,
  );
  zip.file(
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${NS}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>`,
  );
  zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS}" count="${sst.length}" uniqueCount="${sst.length}">${sstXml.join('')}</sst>`);
  if (opts.macro) zip.file('xl/vbaProject.bin', new Uint8Array([1, 2, 3]));
  return zip.generateAsync({ type: 'uint8array' });
}

/** Reads a converted workbook back as a plain grid of values. */
export async function readSheet(data: Uint8Array, sheet: number | string = 0): Promise<(r: number, c: string) => CellValue> {
  const wb = await Workbook.load(data);
  const idx = typeof sheet === 'number' ? sheet : wb.sheetIndexByName(sheet);
  const ws = await wb.sheet(idx);
  return (r, c) => {
    let col = 0;
    for (const ch of c) col = col * 26 + ch.charCodeAt(0) - 64;
    return ws.getCell(r, col).value;
  };
}

export const str = (text: string): CellValue => ({ kind: 'string', text });
export const num = (value: number): CellValue => ({ kind: 'number', value });
export const empty: CellValue = { kind: 'empty' };
