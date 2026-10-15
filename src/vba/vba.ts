/**
 * Re-implementations of the VBA / Excel built-ins the original macros rely on.
 * Each function documents where it differs from the obvious JavaScript equivalent.
 */
import { ConversionError } from '../errors';
import type { CellValue, WriteValue } from '../xlsx/workbook';

/** VBA Trim: removes leading/trailing U+0020 only (not tabs, line breaks or NBSP). */
export function vbaTrim(s: string): string {
  return s.replace(/^ +| +$/g, '');
}

/** WorksheetFunction.Trim: VBA Trim plus collapsing inner runs of U+0020 to a single space. */
export function worksheetTrim(s: string): string {
  return vbaTrim(s.replace(/ {2,}/g, ' '));
}

/** Each UTF-16 code unit not matching `Like "[A-Za-z0-9]"` (binary compare) becomes "_". */
export function replaceNonAlnum(s: string): string {
  return s.replace(/[^A-Za-z0-9]/g, '_');
}

/** VBScript.RegExp `\s` is ASCII whitespace only; JavaScript's `\s` also matches NBSP etc. */
export const VBS_WS = '[ \\t\\n\\r\\f\\v]';

/** Mid(s, start, length) with VBA's run-time error 5 for a negative length. */
export function vbaMid(s: string, start: number, length?: number, context = ''): string {
  if (length !== undefined && length < 0) {
    throw new ConversionError(`Ungültiger Aufbau${context ? ` (${context})` : ''}: erwartetes Trennzeichen fehlt. Das Original-Makro bricht hier mit Laufzeitfehler 5 ab.`);
  }
  return length === undefined ? s.slice(start - 1) : s.slice(start - 1, start - 1 + length);
}

/** InStr([start,] haystack, needle) – binary compare, 1-based, 0 = not found. */
export function inStr(start: number, hay: string, needle: string): number {
  if (start > hay.length + 1) return 0;
  return hay.indexOf(needle, start - 1) + 1;
}

/** CStr(Double) in a German locale: 15 significant digits, scientific below 1E-04 and from 1E+15. */
export function cstrNumberDe(n: number): string {
  if (n === 0) return '0';
  const sign = n < 0 ? '-' : '';
  const [mant, expStr] = Math.abs(n).toExponential(14).split('e');
  const exp = parseInt(expStr, 10);
  const digits = mant.replace('.', '').replace(/0+$/, '') || '0';
  if (exp >= 15 || exp <= -5) {
    const m = digits.length > 1 ? `${digits[0]},${digits.slice(1)}` : digits;
    const e = Math.abs(exp).toString().padStart(2, '0');
    return `${sign}${m}E${exp < 0 ? '-' : '+'}${e}`;
  }
  let out: string;
  if (exp < 0) out = `0,${'0'.repeat(-exp - 1)}${digits}`;
  else if (digits.length <= exp + 1) out = digits + '0'.repeat(exp + 1 - digits.length);
  else out = `${digits.slice(0, exp + 1)},${digits.slice(exp + 1)}`;
  return sign + out;
}

/** CStr(Date) in a German locale: "dd.mm.yyyy", "hh:nn:ss" or both. Input is an Excel serial. */
export function cstrDateDe(serial: number, date1904: boolean): string {
  let s = date1904 ? serial + 1462 : serial;
  // Excel serials below 61 include the fictitious 29.02.1900; VBA dates do not.
  if (!date1904 && s < 61 && s >= 1) s += 1;
  let days = Math.floor(s);
  let secs = Math.round((s - days) * 86400);
  if (secs >= 86400) { days += 1; secs -= 86400; }
  const d = new Date(Date.UTC(1899, 11, 30) + days * 86400000);
  const p = (x: number) => String(x).padStart(2, '0');
  const date = `${p(d.getUTCDate())}.${p(d.getUTCMonth() + 1)}.${d.getUTCFullYear()}`;
  const time = `${p(Math.floor(secs / 3600))}:${p(Math.floor((secs % 3600) / 60))}:${p(secs % 60)}`;
  if (days === 0 && secs !== 0) return time;
  return secs === 0 ? date : `${date} ${time}`;
}

export interface VbaText {
  text: string;
  /** false when the conversion depends on locale/type details that could not be verified. */
  certain: boolean;
}

/**
 * Converts a cell value the way `Dim s As String: s = cell.Value` does.
 * Error values raise "type mismatch" (run-time error 13) in VBA; here a ConversionError.
 */
export function cellToVbaString(v: CellValue, isDate: boolean, date1904: boolean, where: string): VbaText {
  switch (v.kind) {
    case 'empty':
      return { text: '', certain: true };
    case 'string':
      return { text: v.text, certain: true };
    case 'number':
      if (isDate) return { text: cstrDateDe(v.value, date1904), certain: false };
      return { text: cstrNumberDe(v.value), certain: Number.isInteger(v.value) && Math.abs(v.value) < 1e15 };
    case 'bool':
      return { text: v.value ? 'Wahr' : 'Falsch', certain: false };
    case 'error':
      throw new ConversionError(
        `Zelle ${where} enthält den Fehlerwert ${v.text}. Das Original-Makro bricht hier mit „Typen unverträglich“ (Laufzeitfehler 13) ab. Bitte den Fehler in der Quelldatei beheben.`,
      );
  }
}

export interface Coerced {
  value: WriteValue;
  /** Set when Excel might interpret the text differently than this implementation does. */
  warning?: string;
}

const NUMBER_RE = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;

/**
 * What `Range.Value = "text"` stores. Excel parses the string like a typed entry unless the
 * cell is formatted as Text ("@"). Implemented: plain numbers (VBA parses with "." as decimal
 * separator). Everything else is kept as text; patterns Excel might parse otherwise are flagged.
 */
export function coerceForCell(text: string, isTextFormat: boolean): Coerced {
  if (text === '') return { value: { kind: 'empty' } };
  if (isTextFormat) return { value: { kind: 'string', text } };
  if (NUMBER_RE.test(text)) {
    const n = Number(text);
    if (Number.isFinite(n)) {
      const value = Number(n.toPrecision(15));
      const warning = /[.eE]/.test(text) ? `„${text}“ wurde wie in Excel als Zahl ${String(value)} übernommen.` : undefined;
      return { value: { kind: 'number', value }, warning };
    }
  }
  return { value: { kind: 'string', text }, warning: riskyText(text) };
}

function riskyText(t: string): string | undefined {
  if (t.startsWith('=')) return `„${t}“ beginnt mit „=“; Excel würde daraus eine Formel machen. Als Text übernommen.`;
  if (/^[+-]?\d[\d.,]*%$/.test(t)) return `„${t}“ könnte von Excel als Prozentwert interpretiert werden. Als Text übernommen.`;
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t) || /^[+-]?\d+,\d+$/.test(t)) return `„${t}“ könnte von Excel als Zahl interpretiert werden. Als Text übernommen.`;
  if (/^\d{1,4}[./-]\d{1,2}([./-]\d{1,4})?$/.test(t)) return `„${t}“ könnte von Excel als Datum interpretiert werden. Als Text übernommen.`;
  if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(t)) return `„${t}“ könnte von Excel als Uhrzeit interpretiert werden. Als Text übernommen.`;
  if (/^\(\d+([.,]\d+)?\)$/.test(t)) return `„${t}“ könnte von Excel als negative Zahl interpretiert werden. Als Text übernommen.`;
  if (/^(true|false|wahr|falsch)$/i.test(t)) return `„${t}“ könnte von Excel als Wahrheitswert interpretiert werden. Als Text übernommen.`;
  return undefined;
}
