import { describe, expect, it } from 'vitest';
import { coerceForCell, cstrDateDe, cstrNumberDe, replaceNonAlnum, vbaTrim, worksheetTrim } from '../../src/vba/vba';

describe('Trim variants', () => {
  it('VBA Trim removes only U+0020', () => {
    expect(vbaTrim('  a b  ')).toBe('a b');
    expect(vbaTrim('\ta\t')).toBe('\ta\t');
    expect(vbaTrim(' a ')).toBe(' a ');
    expect(vbaTrim('x\r\n')).toBe('x\r\n');
  });
  it('WorksheetFunction.Trim collapses inner spaces', () => {
    expect(worksheetTrim('  a   b  c ')).toBe('a b c');
    expect(worksheetTrim('a  b')).toBe('a  b');
  });
});

describe('Like "[A-Za-z0-9]" replacement', () => {
  it('replaces every other code unit with "_"', () => {
    expect(replaceNonAlnum('Motor[1].Drehzahl#Ist')).toBe('Motor_1__Drehzahl_Ist');
    expect(replaceNonAlnum('Größe Ü')).toBe('Gr__e__');
    expect(replaceNonAlnum('a😀b')).toBe('a__b'); // surrogate pair = two code units, like VBA
    expect(replaceNonAlnum('')).toBe('');
  });
});

describe('CStr in de-DE', () => {
  it('formats doubles like VBA', () => {
    expect(cstrNumberDe(123)).toBe('123');
    expect(cstrNumberDe(1.5)).toBe('1,5');
    expect(cstrNumberDe(-3)).toBe('-3');
    expect(cstrNumberDe(0.0001)).toBe('0,0001');
    expect(cstrNumberDe(0.00001)).toBe('1E-05');
    expect(cstrNumberDe(999999999999999)).toBe('999999999999999');
    expect(cstrNumberDe(1e15)).toBe('1E+15');
    expect(cstrNumberDe(1.25e20)).toBe('1,25E+20');
    expect(cstrNumberDe(0.1 + 0.2)).toBe('0,3');
  });
  it('formats dates like VBA', () => {
    expect(cstrDateDe(45778, false)).toBe('01.05.2025');
    expect(cstrDateDe(45778.5, false)).toBe('01.05.2025 12:00:00');
    expect(cstrDateDe(0.75, false)).toBe('18:00:00');
  });
});

describe('Range.Value = String coercion', () => {
  it('turns plain numbers into numbers', () => {
    expect(coerceForCell('123', false).value).toEqual({ kind: 'number', value: 123 });
    expect(coerceForCell('007', false).value).toEqual({ kind: 'number', value: 7 });
    expect(coerceForCell('1E5', false).value).toEqual({ kind: 'number', value: 100000 });
  });
  it('keeps text in Text-formatted cells', () => {
    expect(coerceForCell('123', true).value).toEqual({ kind: 'string', text: '123' });
  });
  it('writes "" as an empty cell', () => {
    expect(coerceForCell('', false).value).toEqual({ kind: 'empty' });
  });
  it('keeps everything else as text and flags risky patterns', () => {
    expect(coerceForCell('Motor_1', false)).toEqual({ value: { kind: 'string', text: 'Motor_1' } });
    expect(coerceForCell('1,5', false).warning).toBeDefined();
    expect(coerceForCell('01.05.2025', false).warning).toBeDefined();
    expect(coerceForCell('=A1', false).value).toEqual({ kind: 'string', text: '=A1' });
    expect(coerceForCell('—', false).warning).toBeUndefined();
  });
});
