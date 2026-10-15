import { ConversionError } from '../errors';

export const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
export const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export function parseXml(text: string, partName: string): Document {
  let doc: Document;
  try {
    doc = new DOMParser().parseFromString(text, 'application/xml');
  } catch (e) {
    throw new ConversionError(`Die Datei ist beschädigt (XML-Fehler in ${partName}).`);
  }
  if (doc.getElementsByTagName('parsererror').length > 0 || !doc.documentElement) {
    throw new ConversionError(`Die Datei ist beschädigt (XML-Fehler in ${partName}).`);
  }
  return doc;
}

export function serializeXml(doc: Document): string {
  const s = new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, '');
  return XML_DECL + s;
}

/** Direct element children with the given local name (namespace-aware, any prefix). */
export function childElements(parent: Element, localName: string, ns = NS_MAIN): Element[] {
  const out: Element[] = [];
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1) {
      const el = n as Element;
      if (el.localName === localName && el.namespaceURI === ns) out.push(el);
    }
  }
  return out;
}

export function firstChild(parent: Element, localName: string, ns = NS_MAIN): Element | null {
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1) {
      const el = n as Element;
      if (el.localName === localName && el.namespaceURI === ns) return el;
    }
  }
  return null;
}

/** Creates an element in the same namespace/prefix style as `like` uses. */
export function createLike(doc: Document, like: Element, localName: string): Element {
  const prefix = like.prefix;
  return doc.createElementNS(like.namespaceURI, prefix ? `${prefix}:${localName}` : localName);
}

/** Decodes the OOXML ST_Xstring escape `_xHHHH_`. */
export function decodeXString(s: string): string {
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)));
}

/** Encodes text for <t>/<v> the way Excel does (CR and other control chars as _xHHHH_). */
export function encodeXString(s: string): string {
  return s
    .replace(/_(x[0-9A-Fa-f]{4}_)/g, '_x005F_$1')
    .replace(/[\u0000-\u0008\u000B\u000C\u000D\u000E-\u001F￾￿]/g, (c) =>
      `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`,
    );
}

export function escapeXmlText(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export function escapeXmlAttr(s: string): string {
  return escapeXmlText(s).replace(/"/g, '&quot;');
}

export function needsPreserve(s: string): boolean {
  return /^[\s]|[\s]$|[\n\r\t]/.test(s);
}

/** Text of an <si> or <is> element: all <t> runs, phonetic runs (<rPh>) excluded. */
export function richText(el: Element): string {
  let out = '';
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const c = n as Element;
    if (c.namespaceURI !== NS_MAIN) continue;
    if (c.localName === 't') out += c.textContent ?? '';
    else if (c.localName === 'r') {
      const t = firstChild(c, 't');
      if (t) out += t.textContent ?? '';
    }
  }
  return decodeXString(out);
}

/** Resolves a relationship target relative to the part that owns the .rels file. */
export function resolveTarget(sourcePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const base = sourcePart.includes('/') ? sourcePart.slice(0, sourcePart.lastIndexOf('/')).split('/') : [];
  for (const seg of target.split('/')) {
    if (seg === '..') base.pop();
    else if (seg !== '.' && seg !== '') base.push(seg);
  }
  return base.join('/');
}

export function relsPathFor(part: string): string {
  const i = part.lastIndexOf('/');
  return i < 0 ? `_rels/${part}.rels` : `${part.slice(0, i)}/_rels/${part.slice(i + 1)}.rels`;
}
