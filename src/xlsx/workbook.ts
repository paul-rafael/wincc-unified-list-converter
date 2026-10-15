import JSZip from 'jszip';
import { ConversionError } from '../errors';
import { cellRef, parseRange, parseRef } from './ref';
import {
  NS_CT,
  NS_MAIN,
  NS_PKG_REL,
  NS_R,
  childElements,
  createLike,
  decodeXString,
  encodeXString,
  firstChild,
  parseXml,
  relsPathFor,
  resolveTarget,
  richText,
  serializeXml,
} from './xml';

/**
 * In-place editor for .xlsx/.xlsm packages.
 *
 * Only the XML nodes of cells that are written are touched; every other part of the
 * package (styles, other sheets, drawings, validations, …) is kept byte for byte.
 * This mirrors the VBA tools, which open the source workbook, change cell values and
 * save it under a new name.
 */

export type CellValue =
  | { kind: 'empty' }
  | { kind: 'string'; text: string }
  | { kind: 'number'; value: number }
  | { kind: 'bool'; value: boolean }
  | { kind: 'error'; text: string };

export interface CellInfo {
  value: CellValue;
  hasFormula: boolean;
  styleIndex: number | null;
}

const EMPTY: CellValue = { kind: 'empty' };
const STRICT_NS = 'http://purl.oclc.org/ooxml/';
const REL_WORKSHEET = '/worksheet';
const REL_SST = '/sharedStrings';
const REL_STYLES = '/styles';
const REL_CALCCHAIN = '/calcChain';
const REL_VBA = '/vbaProject';
const REL_TABLE = '/table';
const CT_SST = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
const CT_WB_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const CT_PRINTER = 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings';

interface SheetEntry {
  name: string;
  path: string | null; // null for chart sheets / dialog sheets
}

interface XfInfo {
  numFmtId: number;
  formatCode: string | null;
}

export class Workbook {
  readonly sheets: SheetEntry[] = [];
  date1904 = false;
  isMacroEnabled = false;

  private sst: string[] = [];
  private sstIndex = new Map<string, number>();
  private sstDoc: Document | null = null;
  private sstPath: string | null = null;
  private sstDirty = false;
  private xfs: XfInfo[] = [];
  private stylesDoc: Document | null = null;
  private stylesPath: string | null = null;
  private wrapStyleCache = new Map<number, number>();
  private openSheets = new Map<string, Sheet>();
  private dirtyDocs = new Map<string, Document>();
  formulasRemoved = false;

  private constructor(
    private zip: JSZip,
    private wbPath: string,
    private wbRels: Document,
  ) {}

  static async load(data: ArrayBuffer | Uint8Array): Promise<Workbook> {
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(data);
    } catch {
      throw new ConversionError(
        'Die Datei ist keine gültige Excel-Datei (.xlsx/.xlsm). Passwortgeschützte oder alte .xls-Dateien werden nicht unterstützt.',
      );
    }
    const rootRelsText = await zip.file('_rels/.rels')?.async('string');
    if (!rootRelsText) throw new ConversionError('Die Datei ist keine gültige Excel-Datei (Paketstruktur fehlt).');
    const rootRels = parseXml(rootRelsText, '_rels/.rels');
    let wbTarget: string | null = null;
    for (const rel of Array.from(rootRels.getElementsByTagNameNS(NS_PKG_REL, 'Relationship'))) {
      const type = rel.getAttribute('Type') ?? '';
      if (type.startsWith(STRICT_NS)) {
        throw new ConversionError('Dateien im Format „Strict Open XML“ werden nicht unterstützt. Bitte als normale Excel-Arbeitsmappe speichern.');
      }
      if (type.endsWith('/officeDocument')) wbTarget = resolveTarget('', rel.getAttribute('Target') ?? '');
    }
    if (!wbTarget || !zip.file(wbTarget)) throw new ConversionError('Die Datei enthält keine Excel-Arbeitsmappe.');

    const wbRelsPath = relsPathFor(wbTarget);
    const wbRelsText = await zip.file(wbRelsPath)?.async('string');
    if (!wbRelsText) throw new ConversionError('Die Arbeitsmappe ist beschädigt (Beziehungen fehlen).');
    const wb = new Workbook(zip, wbTarget, parseXml(wbRelsText, wbRelsPath));
    await wb.init();
    return wb;
  }

  private rels(): Element[] {
    return Array.from(this.wbRels.getElementsByTagNameNS(NS_PKG_REL, 'Relationship'));
  }

  private relTargetByType(suffix: string): string | null {
    const rel = this.rels().find((r) => (r.getAttribute('Type') ?? '').endsWith(suffix));
    return rel ? resolveTarget(this.wbPath, rel.getAttribute('Target') ?? '') : null;
  }

  private async init(): Promise<void> {
    const wbDoc = parseXml(await this.text(this.wbPath), this.wbPath);
    const pr = wbDoc.getElementsByTagNameNS(NS_MAIN, 'workbookPr')[0];
    const d1904 = pr?.getAttribute('date1904');
    this.date1904 = d1904 === '1' || d1904 === 'true';

    const relById = new Map<string, Element>();
    for (const r of this.rels()) relById.set(r.getAttribute('Id') ?? '', r);
    for (const s of Array.from(wbDoc.getElementsByTagNameNS(NS_MAIN, 'sheet'))) {
      const rel = relById.get(s.getAttributeNS(NS_R, 'id') ?? '');
      const isWorksheet = !!rel && (rel.getAttribute('Type') ?? '').endsWith(REL_WORKSHEET);
      this.sheets.push({
        name: s.getAttribute('name') ?? '',
        path: isWorksheet ? resolveTarget(this.wbPath, rel!.getAttribute('Target') ?? '') : null,
      });
    }

    const ct = await this.text('[Content_Types].xml');
    this.isMacroEnabled = ct.includes('macroEnabled.main+xml') || this.relTargetByType(REL_VBA) !== null;

    this.sstPath = this.relTargetByType(REL_SST);
    if (this.sstPath && this.zip.file(this.sstPath)) {
      this.sstDoc = parseXml(await this.text(this.sstPath), this.sstPath);
      const sis = childElements(this.sstDoc.documentElement, 'si');
      sis.forEach((si, i) => {
        const text = richText(si);
        this.sst.push(text);
        const ts = childElements(si, 't');
        if (ts.length === 1 && si.childNodes.length === 1 && !this.sstIndex.has(text)) this.sstIndex.set(text, i);
      });
    }

    const stylesPath = this.relTargetByType(REL_STYLES);
    if (stylesPath && this.zip.file(stylesPath)) {
      const st = parseXml(await this.text(stylesPath), stylesPath);
      this.stylesDoc = st;
      this.stylesPath = stylesPath;
      const custom = new Map<number, string>();
      for (const nf of Array.from(st.getElementsByTagNameNS(NS_MAIN, 'numFmt'))) {
        custom.set(Number(nf.getAttribute('numFmtId')), nf.getAttribute('formatCode') ?? '');
      }
      const cellXfs = st.getElementsByTagNameNS(NS_MAIN, 'cellXfs')[0];
      if (cellXfs) {
        for (const xf of childElements(cellXfs, 'xf')) {
          const id = Number(xf.getAttribute('numFmtId') ?? '0');
          this.xfs.push({ numFmtId: id, formatCode: custom.get(id) ?? null });
        }
      }
    }
  }

  private async text(path: string): Promise<string> {
    const f = this.zip.file(path);
    if (!f) throw new ConversionError(`Die Arbeitsmappe ist beschädigt (${path} fehlt).`);
    return f.async('string');
  }

  sharedString(i: number): string {
    const s = this.sst[i];
    if (s === undefined) throw new ConversionError(`Die Arbeitsmappe ist beschädigt (Textindex ${i} fehlt).`);
    return s;
  }

  /** Returns the index of `text` in the shared string table, appending it if needed. */
  addSharedString(text: string): number {
    const known = this.sstIndex.get(text);
    if (known !== undefined) return known;
    if (!this.sstDoc) this.createSst();
    const root = this.sstDoc!.documentElement;
    const si = createLike(this.sstDoc!, root, 'si');
    const t = createLike(this.sstDoc!, root, 't');
    if (/^\s|\s$|[\n\r\t]/.test(text)) t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    t.appendChild(this.sstDoc!.createTextNode(encodeXString(text)));
    si.appendChild(t);
    root.appendChild(si);
    const idx = this.sst.length;
    this.sst.push(text);
    this.sstIndex.set(text, idx);
    this.sstDirty = true;
    return idx;
  }

  private createSst(): void {
    this.sstPath = resolveTarget(this.wbPath, 'sharedStrings.xml');
    this.sstDoc = parseXml(`<sst xmlns="${NS_MAIN}" count="0" uniqueCount="0"></sst>`, this.sstPath);
    const root = this.wbRels.documentElement;
    const ids = new Set(this.rels().map((r) => r.getAttribute('Id')));
    let n = 1;
    while (ids.has(`rId${n}`)) n++;
    const rel = this.wbRels.createElementNS(NS_PKG_REL, 'Relationship');
    rel.setAttribute('Id', `rId${n}`);
    rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings');
    rel.setAttribute('Target', 'sharedStrings.xml');
    root.appendChild(rel);
    this.dirtyDocs.set(relsPathFor(this.wbPath), this.wbRels);
    this.pendingContentType = this.sstPath;
  }

  private pendingContentType: string | null = null;

  xf(styleIndex: number | null): XfInfo {
    return this.xfs[styleIndex ?? 0] ?? { numFmtId: 0, formatCode: null };
  }

  isTextStyle(styleIndex: number | null): boolean {
    const xf = this.xf(styleIndex);
    return xf.numFmtId === 49 || xf.formatCode === '@';
  }

  isDateStyle(styleIndex: number | null): boolean {
    const { numFmtId: id, formatCode } = this.xf(styleIndex);
    if ((id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58)) return true;
    if (!formatCode) return false;
    const section = formatCode.split(';')[0];
    const cleaned = section
      .replace(/"[^"]*"/g, '')
      .replace(/\\./g, '')
      .replace(/_.|\*./g, '')
      .replace(/\[(h+|m+|s+)\]/gi, 'h')
      .replace(/\[[^\]]*\]/g, '')
      .replace(/General/gi, '');
    return /[dmyhs]/i.test(cleaned);
  }

  /**
   * Index of a cell format equal to `styleIndex` but with "Zeilenumbruch" (wrapText) switched on,
   * which is what `Range.WrapText = True` does. A new <xf> is appended when needed.
   */
  wrapStyle(styleIndex: number | null): number {
    const base = styleIndex ?? 0;
    const cached = this.wrapStyleCache.get(base);
    if (cached !== undefined) return cached;
    if (!this.stylesDoc || !this.stylesPath) {
      throw new ConversionError('Die Arbeitsmappe enthält keine Formatvorlagen (styles.xml); der Zeilenumbruch kann nicht gesetzt werden.');
    }
    const cellXfs = this.stylesDoc.getElementsByTagNameNS(NS_MAIN, 'cellXfs')[0];
    const xfList = cellXfs ? childElements(cellXfs, 'xf') : [];
    const xf = xfList[base];
    if (!xf) throw new ConversionError(`Die Arbeitsmappe ist beschädigt (Zellformat ${base} fehlt).`);
    const align = firstChild(xf, 'alignment');
    if (align && (align.getAttribute('wrapText') === '1' || align.getAttribute('wrapText') === 'true')) {
      this.wrapStyleCache.set(base, base);
      return base;
    }
    const clone = xf.cloneNode(true) as Element;
    clone.setAttribute('applyAlignment', '1');
    let a = firstChild(clone, 'alignment');
    if (!a) {
      a = createLike(this.stylesDoc, clone, 'alignment');
      clone.insertBefore(a, clone.firstChild);
    }
    a.setAttribute('wrapText', '1');
    cellXfs.appendChild(clone);
    const idx = xfList.length;
    cellXfs.setAttribute('count', String(idx + 1));
    this.xfs.push({ ...this.xf(base) });
    this.wrapStyleCache.set(base, idx);
    this.dirtyDocs.set(this.stylesPath, this.stylesDoc);
    return idx;
  }

  sheetIndexByName(name: string): number {
    const lower = name.toLowerCase();
    return this.sheets.findIndex((s) => s.name.toLowerCase() === lower);
  }

  async sheet(index: number): Promise<Sheet> {
    const entry = this.sheets[index];
    if (!entry) throw new ConversionError('Die Arbeitsmappe enthält kein Arbeitsblatt.');
    if (!entry.path) throw new ConversionError(`Das Blatt „${entry.name}“ ist kein Tabellenblatt (z. B. Diagrammblatt).`);
    const cached = this.openSheets.get(entry.path);
    if (cached) return cached;
    const doc = parseXml(await this.text(entry.path), entry.path);
    const tables = await this.loadTables(entry.path);
    const sheet = new Sheet(this, entry.name, entry.path, doc, tables);
    this.openSheets.set(entry.path, sheet);
    return sheet;
  }

  private async loadTables(sheetPath: string): Promise<TablePart[]> {
    const relsPath = relsPathFor(sheetPath);
    const f = this.zip.file(relsPath);
    if (!f) return [];
    const rels = parseXml(await f.async('string'), relsPath);
    const out: TablePart[] = [];
    for (const r of Array.from(rels.getElementsByTagNameNS(NS_PKG_REL, 'Relationship'))) {
      if (!(r.getAttribute('Type') ?? '').endsWith(REL_TABLE)) continue;
      const path = resolveTarget(sheetPath, r.getAttribute('Target') ?? '');
      if (!this.zip.file(path)) continue;
      const doc = parseXml(await this.text(path), path);
      const root = doc.documentElement;
      const ref = root.getAttribute('ref');
      if (!ref) continue;
      out.push({ path, doc, range: parseRange(ref), headerRows: Number(root.getAttribute('headerRowCount') ?? '1') });
    }
    return out;
  }

  markDirty(path: string, doc: Document): void {
    this.dirtyDocs.set(path, doc);
  }

  /** Converts a macro-enabled package into a plain .xlsx (what SaveAs xlOpenXMLWorkbook does). */
  async stripMacros(): Promise<void> {
    if (!this.isMacroEnabled) return;
    const ctDoc = parseXml(await this.text('[Content_Types].xml'), '[Content_Types].xml');
    for (const o of Array.from(ctDoc.getElementsByTagNameNS(NS_CT, 'Override'))) {
      const type = o.getAttribute('ContentType') ?? '';
      if (type.includes('macroEnabled.main+xml')) o.setAttribute('ContentType', CT_WB_XLSX);
      else if (type.includes('vbaProject') || type.includes('vbaProjectSignature')) o.parentNode!.removeChild(o);
    }
    for (const d of Array.from(ctDoc.getElementsByTagNameNS(NS_CT, 'Default'))) {
      if ((d.getAttribute('ContentType') ?? '').includes('vbaProject')) d.setAttribute('ContentType', CT_PRINTER);
    }
    this.dirtyDocs.set('[Content_Types].xml', ctDoc);
    for (const rel of this.rels()) {
      if ((rel.getAttribute('Type') ?? '').endsWith(REL_VBA)) {
        const target = resolveTarget(this.wbPath, rel.getAttribute('Target') ?? '');
        this.zip.remove(target);
        this.zip.remove(relsPathFor(target));
        rel.parentNode!.removeChild(rel);
      }
    }
    for (const name of Object.keys(this.zip.files)) {
      if (/vbaProjectSignature.*\.bin$/i.test(name)) this.zip.remove(name);
    }
    this.dirtyDocs.set(relsPathFor(this.wbPath), this.wbRels);
    this.isMacroEnabled = false;
  }

  async save(): Promise<Uint8Array> {
    for (const sheet of this.openSheets.values()) sheet.flush();
    if (this.sstDirty && this.sstDoc && this.sstPath) {
      const root = this.sstDoc.documentElement;
      root.removeAttribute('count');
      root.setAttribute('uniqueCount', String(this.sst.length));
      this.dirtyDocs.set(this.sstPath, this.sstDoc);
    }
    if (this.pendingContentType || this.formulasRemoved) {
      const ctPath = '[Content_Types].xml';
      const ctDoc = this.dirtyDocs.get(ctPath) ?? parseXml(await this.text(ctPath), ctPath);
      const root = ctDoc.documentElement;
      if (this.pendingContentType) {
        const o = ctDoc.createElementNS(NS_CT, 'Override');
        o.setAttribute('PartName', `/${this.pendingContentType}`);
        o.setAttribute('ContentType', CT_SST);
        root.appendChild(o);
      }
      if (this.formulasRemoved) {
        // A stale calculation chain makes Excel report a damaged file; Excel rebuilds it on load.
        const chain = this.relTargetByType(REL_CALCCHAIN);
        if (chain) {
          this.zip.remove(chain);
          for (const rel of this.rels()) {
            if ((rel.getAttribute('Type') ?? '').endsWith(REL_CALCCHAIN)) rel.parentNode!.removeChild(rel);
          }
          for (const o of Array.from(ctDoc.getElementsByTagNameNS(NS_CT, 'Override'))) {
            if (o.getAttribute('PartName') === `/${chain}`) o.parentNode!.removeChild(o);
          }
          this.dirtyDocs.set(relsPathFor(this.wbPath), this.wbRels);
        }
      }
      this.dirtyDocs.set(ctPath, ctDoc);
    }
    for (const [path, doc] of this.dirtyDocs) this.zip.file(path, serializeXml(doc));
    return this.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }
}

interface TablePart {
  path: string;
  doc: Document;
  range: { r1: number; c1: number; r2: number; c2: number };
  headerRows: number;
}

export type WriteValue = CellValue;

export class Sheet {
  readonly isProtected: boolean;
  private sheetData: Element;
  private rows = new Map<number, Element>();
  private cells = new Map<number, Map<number, Element>>();
  private colStyles: { min: number; max: number; style: number }[] = [];
  private touchedRows = new Set<number>();
  private dirty = false;

  constructor(
    private wb: Workbook,
    readonly name: string,
    private path: string,
    private doc: Document,
    private tables: TablePart[],
  ) {
    const root = doc.documentElement;
    const sd = firstChild(root, 'sheetData');
    if (!sd) throw new ConversionError(`Das Blatt „${name}“ ist beschädigt (keine Zelldaten).`);
    this.sheetData = sd;
    const prot = firstChild(root, 'sheetProtection');
    this.isProtected = !!prot && (prot.getAttribute('sheet') === '1' || prot.getAttribute('sheet') === 'true');
    const cols = firstChild(root, 'cols');
    if (cols) {
      for (const c of childElements(cols, 'col')) {
        if (c.hasAttribute('style')) {
          this.colStyles.push({ min: Number(c.getAttribute('min')), max: Number(c.getAttribute('max')), style: Number(c.getAttribute('style')) });
        }
      }
    }
    let rowNo = 0;
    for (const row of childElements(sd, 'row')) {
      const r = row.getAttribute('r');
      rowNo = r ? parseInt(r, 10) : rowNo + 1;
      if (!r) row.setAttribute('r', String(rowNo));
      this.rows.set(rowNo, row);
      const map = new Map<number, Element>();
      let colNo = 0;
      for (const c of childElements(row, 'c')) {
        const ref = c.getAttribute('r');
        colNo = ref ? parseRef(ref).col : colNo + 1;
        if (!ref) c.setAttribute('r', cellRef(rowNo, colNo));
        map.set(colNo, c);
      }
      this.cells.set(rowNo, map);
    }
  }

  private cellEl(row: number, col: number): Element | undefined {
    return this.cells.get(row)?.get(col);
  }

  getCell(row: number, col: number): CellInfo {
    const c = this.cellEl(row, col);
    if (!c) return { value: EMPTY, hasFormula: false, styleIndex: this.inheritedStyle(row, col) };
    const s = c.getAttribute('s');
    const styleIndex = s !== null ? Number(s) : null;
    const hasFormula = !!firstChild(c, 'f');
    const t = c.getAttribute('t') ?? 'n';
    const vEl = firstChild(c, 'v');
    const v = vEl ? vEl.textContent ?? '' : null;
    let value: CellValue = EMPTY;
    switch (t) {
      case 's':
        if (v !== null && v !== '') value = { kind: 'string', text: this.wb.sharedString(Number(v)) };
        break;
      case 'str':
        value = { kind: 'string', text: decodeXString(v ?? '') };
        break;
      case 'inlineStr': {
        const is = firstChild(c, 'is');
        if (is) value = { kind: 'string', text: richText(is) };
        break;
      }
      case 'b':
        if (v !== null) value = { kind: 'bool', value: v === '1' || v === 'true' };
        break;
      case 'e':
        if (v !== null) value = { kind: 'error', text: v };
        break;
      case 'd':
        if (v) value = { kind: 'number', value: isoToSerial(v, this.wb.date1904) };
        break;
      default:
        if (v !== null && v !== '') value = { kind: 'number', value: Number(v) };
    }
    return { value, hasFormula, styleIndex };
  }

  /** Style a new cell gets in Excel: row style (customFormat) first, then column style. */
  private inheritedStyle(row: number, col: number): number | null {
    const r = this.rows.get(row);
    if (r && (r.getAttribute('customFormat') === '1' || r.getAttribute('customFormat') === 'true') && r.hasAttribute('s')) {
      return Number(r.getAttribute('s'));
    }
    const cs = this.colStyles.find((c) => col >= c.min && col <= c.max);
    return cs ? cs.style : null;
  }

  isTextFormatted(row: number, col: number): boolean {
    return this.wb.isTextStyle(this.getCell(row, col).styleIndex);
  }

  isDateFormatted(row: number, col: number): boolean {
    return this.wb.isDateStyle(this.getCell(row, col).styleIndex);
  }

  /** Equivalent of Cells(Rows.Count, col).End(xlUp).Row. */
  lastRowInColumn(col: number): number {
    const rowNos = Array.from(this.cells.keys()).sort((a, b) => b - a);
    for (const r of rowNos) {
      const c = this.cells.get(r)!.get(col);
      if (c && (firstChild(c, 'f') || firstChild(c, 'v') || firstChild(c, 'is'))) return r;
    }
    return 1;
  }

  setCell(row: number, col: number, val: WriteValue): void {
    let c = this.cellEl(row, col);
    if (!c) {
      if (val.kind === 'empty') return;
      c = this.createCell(row, col);
    }
    const f = firstChild(c, 'f');
    if (f) {
      const ft = f.getAttribute('t');
      const ref = f.getAttribute('ref');
      if (ft === 'shared' && ref && ref.includes(':')) {
        throw new ConversionError(
          `Zelle ${this.name}!${cellRef(row, col)} enthält eine gemeinsam genutzte Formel, die nicht überschrieben werden kann. Bitte Formeln vorher durch Werte ersetzen.`,
        );
      }
      if (ft === 'array' && ref && ref.includes(':')) {
        throw new ConversionError(`Zelle ${this.name}!${cellRef(row, col)} ist Teil einer Matrixformel. Bitte Formeln vorher durch Werte ersetzen.`);
      }
      c.removeChild(f);
      this.wb.formulasRemoved = true;
    }
    for (const name of ['v', 'is']) {
      const el = firstChild(c, name);
      if (el) c.removeChild(el);
    }
    c.removeAttribute('t');
    c.removeAttribute('cm');
    c.removeAttribute('vm');
    const ext = firstChild(c, 'extLst');
    const addV = (text: string) => {
      const v = createLike(this.doc, c!, 'v');
      v.appendChild(this.doc.createTextNode(text));
      c!.insertBefore(v, ext);
    };
    switch (val.kind) {
      case 'string':
        c.setAttribute('t', 's');
        addV(String(this.wb.addSharedString(val.text)));
        break;
      case 'number':
        addV(numberToXml(val.value));
        break;
      case 'bool':
        c.setAttribute('t', 'b');
        addV(val.value ? '1' : '0');
        break;
      case 'error':
        c.setAttribute('t', 'e');
        addV(val.text);
        break;
      case 'empty':
        if (!c.hasAttribute('s') && !ext) {
          c.parentNode!.removeChild(c);
          this.cells.get(row)!.delete(col);
        }
        break;
    }
    this.syncTableHeader(row, col, val);
    this.dirty = true;
  }

  private syncTableHeader(row: number, col: number, val: WriteValue): void {
    for (const t of this.tables) {
      const { r1, c1, c2 } = t.range;
      if (t.headerRows < 1 || row !== r1 || col < c1 || col > c2) continue;
      const tc = t.doc.getElementsByTagNameNS(NS_MAIN, 'tableColumn')[col - c1];
      if (!tc) continue;
      const name = val.kind === 'string' ? val.text : val.kind === 'number' ? String(val.value) : val.kind === 'empty' ? `Spalte${col - c1 + 1}` : '';
      tc.setAttribute('name', name);
      this.wb.markDirty(t.path, t.doc);
    }
  }

  private createCell(row: number, col: number): Element {
    let r = this.rows.get(row);
    if (!r) {
      r = createLike(this.doc, this.sheetData, 'row');
      r.setAttribute('r', String(row));
      const after = Array.from(this.rows.keys()).filter((k) => k > row).sort((a, b) => a - b)[0];
      this.sheetData.insertBefore(r, after !== undefined ? this.rows.get(after)! : null);
      this.rows.set(row, r);
      this.cells.set(row, new Map());
    }
    const style = this.inheritedStyle(row, col);
    const c = createLike(this.doc, this.sheetData, 'c');
    c.setAttribute('r', cellRef(row, col));
    if (style !== null && style !== 0) c.setAttribute('s', String(style));
    const map = this.cells.get(row)!;
    const next = Array.from(map.keys()).filter((k) => k > col).sort((a, b) => a - b)[0];
    r.insertBefore(c, next !== undefined ? map.get(next)! : (firstChild(r, 'extLst') ?? null));
    map.set(col, c);
    this.touchedRows.add(row);
    return c;
  }

  /** Range.WrapText = True for one cell (creates the cell if it does not exist yet). */
  setWrapText(row: number, col: number): void {
    let c = this.cellEl(row, col);
    if (!c) c = this.createCell(row, col);
    const s = c.getAttribute('s');
    const idx = this.wb.wrapStyle(s !== null ? Number(s) : null);
    if (idx === 0) c.removeAttribute('s');
    else c.setAttribute('s', String(idx));
    this.dirty = true;
  }

  /**
   * Rows(i).AutoFit: the height is computed by Excel from the content, so the stored height is
   * removed and Excel sizes the row automatically when the file is opened.
   */
  autoFitRow(row: number): void {
    const r = this.rows.get(row);
    if (!r) return;
    r.removeAttribute('ht');
    r.removeAttribute('customHeight');
    this.dirty = true;
  }

  /** Writes the sheet back: refreshes <dimension> and row spans of rows that got new cells. */
  flush(): void {
    if (!this.dirty) return;
    for (const row of this.touchedRows) {
      const r = this.rows.get(row)!;
      const cols = Array.from(this.cells.get(row)!.keys());
      if (r.hasAttribute('spans') && cols.length) r.setAttribute('spans', `${Math.min(...cols)}:${Math.max(...cols)}`);
    }
    let r1 = Infinity, r2 = -Infinity, c1 = Infinity, c2 = -Infinity;
    for (const [row, map] of this.cells) {
      for (const col of map.keys()) {
        r1 = Math.min(r1, row); r2 = Math.max(r2, row);
        c1 = Math.min(c1, col); c2 = Math.max(c2, col);
      }
    }
    const dim = firstChild(this.doc.documentElement, 'dimension');
    if (dim && r1 !== Infinity) {
      const a = cellRef(r1, c1), b = cellRef(r2, c2);
      dim.setAttribute('ref', a === b ? a : `${a}:${b}`);
    }
    this.wb.markDirty(this.path, this.doc);
  }
}

export function numberToXml(n: number): string {
  if (!Number.isFinite(n)) throw new ConversionError('Ungültiger Zahlenwert.');
  return String(n).replace('e', 'E');
}

function isoToSerial(iso: string, date1904: boolean): number {
  const ms = Date.parse(/[zZ]|[+-]\d\d:?\d\d$/.test(iso) ? iso : `${iso}Z`);
  if (Number.isNaN(ms)) return 0;
  const serial = ms / 86400000 + 25569;
  return date1904 ? serial - 1462 : serial;
}
