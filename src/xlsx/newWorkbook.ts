import JSZip from 'jszip';
import type { WriteValue } from './workbook';
import { numberToXml } from './workbook';
import { cellRef } from './ref';
import { encodeXString, escapeXmlAttr, escapeXmlText, needsPreserve, XML_DECL } from './xml';

/**
 * Writes a new single-sheet workbook, equivalent to Workbooks.Add + SaveAs xlOpenXMLWorkbook
 * (default styles, all values in "General" format).
 */
export async function buildWorkbook(sheetName: string, grid: Map<number, Map<number, WriteValue>>): Promise<Uint8Array> {
  const sst: string[] = [];
  const sstIndex = new Map<string, number>();
  let sstRefs = 0;
  const rowsXml: string[] = [];
  let maxRow = 1, maxCol = 1;

  for (const row of Array.from(grid.keys()).sort((a, b) => a - b)) {
    const cells = grid.get(row)!;
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
          sstRefs++;
          parts.push(`<c r="${r}" t="s"><v>${idx}</v></c>`);
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
  const sheetXml =
    XML_DECL +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<dimension ref="${dim}"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="15"/><sheetData>${rowsXml.join('')}</sheetData>` +
    `<pageMargins left="0.7" right="0.7" top="0.787401575" bottom="0.787401575" header="0.3" footer="0.3"/></worksheet>`;

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
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
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
      `<bookViews><workbookView/></bookViews><sheets><sheet name="${escapeXmlAttr(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    XML_DECL +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` +
      `</Relationships>`,
  );
  zip.file(
    'xl/styles.xml',
    XML_DECL +
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font></fonts>` +
      `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
      `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Standard" xfId="0" builtinId="0"/></cellStyles>` +
      `<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>` +
      `</styleSheet>`,
  );
  zip.file('xl/sharedStrings.xml', sstXml);
  zip.file('xl/worksheets/sheet1.xml', sheetXml);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}
