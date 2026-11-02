import JSZip from 'jszip';
import type { WriteValue } from './workbook';
import { numberToXml } from './workbook';
import { cellRef } from './ref';
import { encodeXString, escapeXmlAttr, escapeXmlText, needsPreserve, XML_DECL } from './xml';

export interface NewSheet {
  name: string;
  grid: Map<number, Map<number, WriteValue>>;
  /** Column widths in characters, starting at column A. */
  widths?: number[];
  /** Bold first row, frozen below it. */
  header?: boolean;
  /** Drop-down lists, e.g. { sqref: 'C2:C1000', list: ['a', 'b'] }. */
  lists?: { sqref: string; list: string[] }[];
}

/**
 * Writes a new single-sheet workbook, equivalent to Workbooks.Add + SaveAs xlOpenXMLWorkbook
 * (default styles, all values in "General" format).
 */
export async function buildWorkbook(sheetName: string, grid: Map<number, Map<number, WriteValue>>): Promise<Uint8Array> {
  return buildWorkbookSheets([{ name: sheetName, grid }]);
}

/** Same as buildWorkbook, with several sheets and optional header style, widths and drop-down lists. */
export async function buildWorkbookSheets(sheets: NewSheet[]): Promise<Uint8Array> {
  const sst: string[] = [];
  const sstIndex = new Map<string, number>();
  let sstRefs = 0;
  const sheetXmls = sheets.map((sheet, n) => {
    const { xml, refs } = sheetXml(sheet, n === 0, sst, sstIndex);
    sstRefs += refs;
    return xml;
  });

  const sstXml =
    XML_DECL +
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sstRefs}" uniqueCount="${sst.length}">` +
    sst
      .map((s) => {
        const enc = escapeXmlText(encodeXString(s));
        return needsPreserve(s) ? `<si><t xml:space="preserve">${enc}</t></si>` : `<si><t>${enc}</t></si>`;
      })
      .join('') +
    `</sst>`;

  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    XML_DECL +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      sheets.map((_, n) => `<Override PartName="/xl/worksheets/sheet${n + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
      `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` +
      `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
      `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
      `</Types>`,
  );
  zip.file(
    '_rels/.rels',
    XML_DECL +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>` +
      `</Relationships>`,
  );
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  zip.file(
    'docProps/core.xml',
    XML_DECL +
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
      `</cp:coreProperties>`,
  );
  zip.file(
    'docProps/app.xml',
    XML_DECL +
      `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>`,
  );
  zip.file(
    'xl/workbook.xml',
    XML_DECL +
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<bookViews><workbookView/></bookViews><sheets>` +
      sheets.map((sh, n) => `<sheet name="${escapeXmlAttr(sh.name)}" sheetId="${n + 1}" r:id="rId${n === 0 ? 1 : n + 3}"/>`).join('') +
      `</sheets></workbook>`,
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    XML_DECL +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` +
      sheets
        .slice(1)
        .map((_, i) => `<Relationship Id="rId${i + 4}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 2}.xml"/>`)
        .join('') +
      `</Relationships>`,
  );
  zip.file(
    'xl/styles.xml',
    XML_DECL +
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font>` +
      `<font><b/><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font></fonts>` +
      `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
      `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Standard" xfId="0" builtinId="0"/></cellStyles>` +
      `<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>` +
      `</styleSheet>`,
  );
  zip.file('xl/sharedStrings.xml', sstXml);
  sheetXmls.forEach((xml, n) => zip.file(`xl/worksheets/sheet${n + 1}.xml`, xml));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

function sheetXml(sheet: NewSheet, selected: boolean, sst: string[], sstIndex: Map<string, number>): { xml: string; refs: number } {
  let refs = 0;
  const rowsXml: string[] = [];
  let maxRow = 1, maxCol = 1;
  const bold = (row: number) => (sheet.header && row === 1 ? ' s="1"' : '');
  for (const row of Array.from(sheet.grid.keys()).sort((a, b) => a - b)) {
    const cells = sheet.grid.get(row)!;
    const cols = Array.from(cells.keys()).sort((a, b) => a - b);
    const parts: string[] = [];
    for (const col of cols) {
      const v = cells.get(col)!;
      const r = cellRef(row, col);
      switch (v.kind) {
        case 'empty':
          continue;
        case 'string': {
          let idx = sstIndex.get(v.text);
          if (idx === undefined) {
            idx = sst.length;
            sst.push(v.text);
            sstIndex.set(v.text, idx);
          }
          refs++;
          parts.push(`<c r="${r}"${bold(row)} t="s"><v>${idx}</v></c>`);
          break;
        }
        case 'number':
          parts.push(`<c r="${r}"><v>${numberToXml(v.value)}</v></c>`);
          break;
        case 'bool':
          parts.push(`<c r="${r}" t="b"><v>${v.value ? 1 : 0}</v></c>`);
          break;
        case 'error':
          parts.push(`<c r="${r}" t="e"><v>${escapeXmlText(v.text)}</v></c>`);
          break;
      }
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    }
    if (parts.length) rowsXml.push(`<row r="${row}">${parts.join('')}</row>`);
  }

  const dim = maxRow === 1 && maxCol === 1 ? 'A1' : `A1:${cellRef(maxRow, maxCol)}`;
  const view = sheet.header
    ? `<sheetView${selected ? ' tabSelected="1"' : ''} workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView>`
    : `<sheetView${selected ? ' tabSelected="1"' : ''} workbookViewId="0"/>`;
  const cols = sheet.widths?.length
    ? `<cols>${sheet.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>`
    : '';
  const lists = sheet.lists?.length
    ? `<dataValidations count="${sheet.lists.length}">` +
      sheet.lists
        .map((l) => `<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="${l.sqref}"><formula1>${escapeXmlText(`"${l.list.join(',')}"`)}</formula1></dataValidation>`)
        .join('') +
      `</dataValidations>`
    : '';
  const xml =
    XML_DECL +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<dimension ref="${dim}"/><sheetViews>${view}</sheetViews>` +
    `<sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${rowsXml.join('')}</sheetData>${lists}` +
    `<pageMargins left="0.7" right="0.7" top="0.787401575" bottom="0.787401575" header="0.3" footer="0.3"/></worksheet>`;
  return { xml, refs };
}

