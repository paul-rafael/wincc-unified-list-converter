import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
  ALARM_HEADER,
  buildFieldInfo,
  convertAlarms,
  extractLengthByRef,
  extractRefs,
  extractTagByRef,
} from '../../src/converters/alarms';
import { convertTags } from '../../src/converters/tags';
import { convertTextlistLegacy, extractTextListValue, legacyCellValue, legacyOutputName } from '../../src/converters/textlistLegacy';
import { convertTextlistStandard, replaceSpecialChars, standardOutputName } from '../../src/converters/textlistStandard';
import { empty, makeXlsx, num, readSheet, str } from '../helpers';

const DASH = '—';

describe('HMI-Variablen (SonderzeichenMitKnopf)', () => {
  it('replaces special characters in column A including the header', async () => {
    const input = await makeXlsx([
      {
        name: 'Hmi Tags',
        rows: {
          1: { A: 'Name', B: 'Path' },
          2: { A: 'Motor[1].Ist', B: 'Default tag table' },
          3: { A: 'Füllstand #2', B: 'x.y' },
          4: { A: 'Plain1' },
          6: { A: 'Last/Tag' },
        },
      },
      { name: 'Multiplexing', rows: { 1: { A: 'a.b' } } },
    ]);
    const res = await convertTags(input, 'HMITags.xlsx');
    expect(res.fileName).toBe('HMITagList_MIGUNI.xlsx');
    const get = await readSheet(res.data, 0);
    expect(get(1, 'A')).toEqual(str('Name'));
    expect(get(2, 'A')).toEqual(str('Motor_1__Ist'));
    expect(get(3, 'A')).toEqual(str('F_llstand__2'));
    expect(get(4, 'A')).toEqual(str('Plain1'));
    expect(get(5, 'A')).toEqual(empty);
    expect(get(6, 'A')).toEqual(str('Last_Tag'));
    expect(get(3, 'B')).toEqual(str('x.y')); // other columns untouched
    const other = await readSheet(res.data, 1);
    expect(other(1, 'A')).toEqual(str('a.b')); // other sheets untouched
  });

  it('applies CStr to numbers and Excel coercion to numeric text', async () => {
    const input = await makeXlsx([
      { name: 'S', rows: { 1: { A: 'Name' }, 2: { A: 1.5 }, 3: { A: '123' }, 4: { A: { s: 1, v: '0042' } }, 5: { A: -7 } } },
    ]);
    const get = await readSheet((await convertTags(input, 'a.xlsx')).data);
    expect(get(2, 'A')).toEqual(str('1_5'));
    expect(get(3, 'A')).toEqual(num(123));
    expect(get(4, 'A')).toEqual(str('0042')); // Text format keeps leading zeros
    expect(get(5, 'A')).toEqual(str('_7'));
  });

  it('aborts on error values like VBA (type mismatch)', async () => {
    const input = await makeXlsx([{ name: 'S', rows: { 1: { A: 'Name' }, 2: { A: { err: '#N/A' } } } }]);
    await expect(convertTags(input, 'a.xlsx')).rejects.toThrow(/Fehlerwert/);
  });

  it('rejects other file types', async () => {
    const input = await makeXlsx([{ name: 'S', rows: { 1: { A: 'x' } } }]);
    await expect(convertTags(input, 'a.xlsm')).rejects.toThrow(/nicht unterstützt/);
    await expect(convertTags(input, 'a.xls')).rejects.toThrow(/\.xls/);
  });

  it('refuses protected sheets', async () => {
    const input = await makeXlsx([{ name: 'S', rows: { 1: { A: 'x' } }, extra: '<sheetProtection sheet="1"/>' }]);
    await expect(convertTags(input, 'a.xlsx')).rejects.toThrow(/schreibgeschützt/);
  });

  it('keeps cell styles and untouched package parts', async () => {
    const input = await makeXlsx([{ name: 'S', rows: { 1: { A: { s: 3, v: 'Na me' } } } }]);
    const res = await convertTags(input, 'a.xlsx');
    const inZip = await JSZip.loadAsync(input);
    const outZip = await JSZip.loadAsync(res.data);
    expect(await outZip.file('xl/styles.xml')!.async('string')).toBe(await inZip.file('xl/styles.xml')!.async('string'));
    const sheet = await outZip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet).toMatch(/<c r="A1" s="3" t="s">/);
    expect((await readSheet(res.data))(1, 'A')).toEqual(str('Na_me'));
  });

  it('replaces formulas by their converted value and drops the calc chain', async () => {
    const input = await makeXlsx([{ name: 'S', rows: { 1: { A: { f: '"a"&"-b"', v: 'a-b' } } } }]);
    const res = await convertTags(input, 'a.xlsx');
    const sheet = await (await JSZip.loadAsync(res.data)).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet).not.toContain('<f>');
    expect((await readSheet(res.data))(1, 'A')).toEqual(str('a_b'));
  });
});

describe('Textlisten · Legacy (Clean_TextList_Sheet)', () => {
  it('cell rule matches the VBA branches', () => {
    // column E: only <No value> / <field ref="0" /> are replaced
    expect(legacyCellValue(5, 'Hallo <b>Welt</b>', '', '', '')).toBeNull();
    expect(legacyCellValue(5, '<No value>', '', 'Fallback', '')).toBe('Fallback');
    expect(legacyCellValue(5, '<FIELD REF="0" />', 'x', '', 'Liste')).toBe('Liste');
    // <field ref="0" />
    expect(legacyCellValue(6, '<field ref="0" />', 'k', 'f', '')).toBe(DASH);
    // <No value> with K empty → fallback unless it is <No value> / empty
    expect(legacyCellValue(6, '<no value>', '', 'Text I', '')).toBe('Text I');
    expect(legacyCellValue(6, '<no value>', '', '<No Value>', '')).toBe(DASH);
    expect(legacyCellValue(6, '<no value>', '', '', '')).toBe(DASH);
    // <No value> with K set → TextList value or dash
    expect(legacyCellValue(6, '<No value>', 'abc', 'Text I', 'TL')).toBe('TL');
    expect(legacyCellValue(6, '<No value>', 'abc', 'Text I', '')).toBe(DASH);
    // tags removed, WorksheetFunction.Trim, empty → dash
    expect(legacyCellValue(7, 'Motor  <field ref="1" />  läuft', '', '', '')).toBe('Motor läuft');
    expect(legacyCellValue(8, '', '', '', '')).toBe(DASH);
    expect(legacyCellValue(8, '<a><b>', '', '', '')).toBe(DASH);
  });

  it('extracts TextList values with VBScript whitespace semantics', () => {
    expect(extractTextListValue('<field ref="0" TextList = Betriebsart ; x>')).toBe('Betriebsart');
    expect(extractTextListValue('textlist=Abc>')).toBe('Abc');
    expect(extractTextListValue('TextList = Abc;')).toBe(''); // NBSP is not \s in VBScript
    expect(extractTextListValue('')).toBe('');
  });

  it('output name follows Replace(path, ".xlsx", "_MIGUNI.xlsx")', () => {
    expect(legacyOutputName('TextLists.xlsx')).toBe('TextLists_MIGUNI.xlsx');
    expect(legacyOutputName('TextLists.xlsm')).toBe('TextLists_MIGUNI.xlsx');
  });

  it('converts a TextListEntry sheet', async () => {
    const input = await makeXlsx([
      { name: 'TextList', rows: { 1: { A: 'Name' }, 2: { A: 'Liste.1' } } },
      {
        name: 'TextListEntry',
        rows: {
          1: { A: 'Name', E: 'Text [de-DE]', F: 'Text [en-US]', I: 'I', K: 'K' },
          2: { E: 'Ein', F: '<No value>', G: '  Hallo   <b>Welt</b> ', I: 'Fallback', J: 5 },
          3: { E: '<No value>', F: '<field ref="0" />', K: 'TextList = Modi;' },
          4: { E: 'Zwei', H: 1.5 },
        },
      },
    ]);
    const res = await convertTextlistLegacy(input, 'TL.xlsx');
    const get = await readSheet(res.data, 'TextListEntry');
    expect(get(2, 'E')).toEqual(str('Ein'));
    expect(get(2, 'F')).toEqual(str('Fallback'));
    expect(get(2, 'G')).toEqual(str('Hallo Welt'));
    expect(get(2, 'H')).toEqual(str(DASH));
    expect(get(2, 'I')).toEqual(str('Fallback'));
    expect(get(2, 'J')).toEqual(num(5));
    expect(get(3, 'E')).toEqual(str('Modi'));
    expect(get(3, 'F')).toEqual(str('Modi'));
    expect(get(3, 'G')).toEqual(str(DASH));
    expect(get(3, 'K')).toEqual(str('TextList = Modi;')); // K is read, never written
    expect(get(4, 'H')).toEqual(num(1.5)); // non-integer number left as is, with a notice
    expect(res.notices.some((n) => n.cell === 'TextListEntry!H4')).toBe(true);
    const tl = await readSheet(res.data, 'TextList');
    expect(tl(2, 'A')).toEqual(str('Liste.1'));
  });

  it('stops when TextListEntry is missing', async () => {
    const input = await makeXlsx([{ name: 'Other', rows: { 1: { A: 'x' } } }]);
    await expect(convertTextlistLegacy(input, 'a.xlsx')).rejects.toThrow(/TextListEntry/);
  });

  it('turns .xlsm into a macro-free .xlsx', async () => {
    const input = await makeXlsx([{ name: 'TextListEntry', rows: { 1: { E: 'h' }, 2: { E: 'x' } } }], { macro: true });
    const res = await convertTextlistLegacy(input, 'a.xlsm');
    const zip = await JSZip.loadAsync(res.data);
    expect(zip.file('xl/vbaProject.bin')).toBeNull();
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    expect(ct).not.toContain('macroEnabled');
    expect(await zip.file('xl/_rels/workbook.xml.rels')!.async('string')).not.toContain('vbaProject');
  });
});

describe('Textlisten · Standard (ReplaceSpecialCharsInWorkbook)', () => {
  it('replaces in the documented order', () => {
    expect(replaceSpecialChars('A#B/C.D[1]')).toBe('A_B_C_D(1)');
    expect(standardOutputName('Lists.xlsm')).toBe('Lists.xlsx');
  });

  it('touches only TextList!A and TextListEntry!B from row 2', async () => {
    const input = await makeXlsx([
      { name: 'TextList', rows: { 1: { A: 'Na.me' }, 2: { A: 'List.A' }, 3: { A: 'Liste[2]', B: 'x.y' } } },
      { name: 'TextListEntry', rows: { 1: { B: 'H.1' }, 2: { A: 'a.b', B: 'Entry/1' }, 3: { B: 7 } } },
    ]);
    const res = await convertTextlistStandard(input, 'Lists.xlsx');
    expect(res.fileName).toBe('Lists.xlsx');
    const tl = await readSheet(res.data, 'TextList');
    expect(tl(1, 'A')).toEqual(str('Na.me'));
    expect(tl(2, 'A')).toEqual(str('List_A'));
    expect(tl(3, 'A')).toEqual(str('Liste(2)'));
    expect(tl(3, 'B')).toEqual(str('x.y'));
    const tle = await readSheet(res.data, 'TextListEntry');
    expect(tle(1, 'B')).toEqual(str('H.1'));
    expect(tle(2, 'A')).toEqual(str('a.b'));
    expect(tle(2, 'B')).toEqual(str('Entry_1'));
    expect(tle(3, 'B')).toEqual(num(7));
  });

  it('skips a missing sheet with a notice and still produces a file', async () => {
    const input = await makeXlsx([{ name: 'TextList', rows: { 2: { A: 'a.b' } } }]);
    const res = await convertTextlistStandard(input, 'L.xlsx');
    expect(res.notices[0].text).toMatch(/TextListEntry/);
    expect((await readSheet(res.data, 'TextList'))(2, 'A')).toEqual(str('a_b'));
  });
});

describe('HMI-Meldungen (ConvertAlarmList_Unified)', () => {
  const oldInfo =
    '<ref id = 1; type = AlarmTag; Tag = Motor_Ist; DisplayType = Decimal; Length = 5; Precision = 0;>' +
    '<ref id = 2; type = AlarmTag; Tag = Druck; Length = 4;>';

  it('extracts refs, tags and lengths like the VBA helpers', () => {
    expect(extractRefs('Ist <field ref="1" /> bar <field ref="2" />')).toEqual(['1', '2']);
    expect(extractRefs('kein Feld')).toEqual([]);
    expect(extractRefs('ref="1 /> x"')).toEqual(['1 /> x']);
    expect(() => extractRefs('ref="1')).toThrow(/Laufzeitfehler 5/);
    expect(extractTagByRef(oldInfo, '2')).toBe('Druck');
    expect(extractTagByRef(oldInfo, '9')).toBe('');
    expect(extractLengthByRef(oldInfo, '1')).toBe('5');
    expect(extractLengthByRef('<ref id = 3; Tag = X;>', '3')).toBe('3');
    // substring match: "id = 1" also finds "id = 10" (original behaviour)
    expect(extractTagByRef('<ref id = 10; Tag = Zehn;><ref id = 1; Tag = Eins;>', '1')).toBe('Zehn');
  });

  it('builds the Unified field info with trailing CR LF', () => {
    const r = buildFieldInfo('A <field ref="1" /> B <field ref="2" />', oldInfo);
    expect(r.parameter1).toBe('Druck');
    expect(r.fieldInfo).toBe(
      '<ref id = 1; type = AlarmParameterWithOrWithoutCommonTextList; Parameter = Parameter 1; Tag = Motor_Ist; DisplayType = Decimal; Length = 5; Precision = 0; Alignment = Right; ZeroPadding = False;>\r\n' +
        '<ref id = 2; type = AlarmParameterWithOrWithoutCommonTextList; Parameter = Parameter 1; Tag = Druck; DisplayType = Decimal; Length = 4; Precision = 0; Alignment = Right; ZeroPadding = False;>\r\n',
    );
    expect(buildFieldInfo('ohne', oldInfo)).toEqual({ fieldInfo: '', parameter1: '' });
  });

  it('writes the DiscreteAlarms workbook', async () => {
    const input = await makeXlsx([
      {
        name: 'DiscreteAlarms',
        rows: {
          1: { A: 'ID', B: 'Name', C: 'Alarm text', D: 'FieldInfo', E: 'Class', F: 'Trigger tag', G: 'Trigger bit', H: 'Mode' },
          2: { A: 1, B: 'Alarm_1', C: 'Motor <field ref="1" /> zu heiß', D: oldInfo, E: 'Errors', F: 'Stoerwort_1', G: 0, H: 'x' },
          3: { A: 2, B: 'Alarm_2', C: 'Ohne Feld', E: 'Warnings', F: 'Stoerwort_1', G: 1 },
        },
      },
    ]);
    const res = await convertAlarms(input, 'HMIAlarms.xlsx');
    expect(res.fileName).toBe('Unified_HMI-Alarms.xlsx');
    const get = await readSheet(res.data, 'DiscreteAlarms');
    const col = (n: number) => {
      let s = '';
      while (n > 0) { s = String.fromCharCode(65 + ((n - 1) % 26)) + s; n = Math.floor((n - 1) / 26); }
      return s;
    };
    ALARM_HEADER.forEach((h, i) => expect(get(1, col(i + 1))).toEqual(str(h)));
    expect(ALARM_HEADER).toHaveLength(44);

    expect(get(2, 'A')).toEqual(num(1));
    expect(get(2, 'B')).toEqual(str('Alarm_1'));
    expect(get(2, 'C')).toEqual(str('Motor <field ref="1" /> zu heiß'));
    expect(get(2, 'D')).toEqual(str(buildFieldInfo('<field ref="1" />', oldInfo).fieldInfo));
    expect(get(2, 'E')).toEqual(str('Errors'));
    expect(get(2, 'F')).toEqual(str('Stoerwort_1'));
    expect(get(2, 'G')).toEqual(num(0));
    expect(get(2, 'H')).toEqual(str('On rising edge'));
    expect(get(2, 'I')).toEqual(str('<No value>'));
    expect(get(2, 'J')).toEqual(num(0));
    expect(get(2, 'K')).toEqual(str('<No value>'));
    expect(get(2, 'L')).toEqual(num(0));
    expect(get(2, 'M')).toEqual(num(0));
    for (let c = 14; c <= 32; c += 2) expect(get(2, col(c))).toEqual(str('<No value>'));
    for (let c = 15; c <= 31; c += 2) expect(get(2, col(c))).toEqual(empty);
    expect(get(2, 'AG')).toEqual(str('Motor_Ist'));
    for (let c = 34; c <= 42; c++) expect(get(2, col(c))).toEqual(str('<No value>'));
    expect(get(2, 'AQ')).toEqual(str('HMI_RT_2::Alarming'));
    expect(get(2, 'AR')).toEqual(empty);

    expect(get(3, 'D')).toEqual(empty);
    expect(get(3, 'AG')).toEqual(empty);

    // CR is stored as _x000D_ like Excel does
    const sst = await (await JSZip.loadAsync(res.data)).file('xl/sharedStrings.xml')!.async('string');
    expect(sst).toContain('_x000D_\n');
  });
});

describe('Meldetexte umbrechen (FormatHMITextWithErrorCode)', () => {
  it('splits error code and wraps at 80 characters', async () => {
    const { formatHmiText } = await import('../../src/converters/lineBreaks');
    expect(formatHmiText('E101_Motor laeuft nicht')).toBe('E101¶Motor laeuft nicht');
    expect(formatHmiText('MFS_CON_12_Text')).toBe('MFS-CON¶12_Text');
    expect(formatHmiText('OhneUnterstrich')).toBe('OhneUnterstrich');
    expect(formatHmiText('')).toBe('');
    expect(formatHmiText('  E1_a  ')).toBe('E1¶a');
    expect(formatHmiText('E2_Achse steht Position: 100 Geschw.: 5')).toBe('E2¶Achse steht¶Position: 100 Geschw.: 5');
    const long = 'E3_' + Array.from({ length: 30 }, (_, i) => `Wort${i}`).join(' ');
    const [code, l2, l3] = formatHmiText(long).split('¶');
    expect(code).toBe('E3');
    expect(l2.length).toBeLessThanOrEqual(80);
    expect(l3.startsWith('Wort')).toBe(true);
    // first word is a keyword: line 2 stays empty and the next break fills it
    expect(formatHmiText('E4_X: 1 Y: 2')).toBe('E4¶X: 1¶Y: 2');
  });

  it('writes column C with wrap text and auto row height', async () => {
    const { convertLineBreaks } = await import('../../src/converters/lineBreaks');
    const input = await makeXlsx([
      { name: 'DiscreteAlarms', rows: { 1: { A: 'ID', C: 'Text' }, 2: { A: 1, C: 'E1_Motor heiss' }, 3: { A: 2, C: { s: 3, v: 'E2_Druck' } } } },
    ]);
    const res = await convertLineBreaks(input, 'Alarms.xlsx');
    expect(res.fileName).toBe('Alarms.xlsx');
    const get = await readSheet(res.data);
    expect(get(1, 'C')).toEqual(str('Text'));
    expect(get(2, 'C')).toEqual(str('E1¶Motor heiss'));
    expect(get(3, 'C')).toEqual(str('E2¶Druck'));
    const zip = await JSZip.loadAsync(res.data);
    const styles = await zip.file('xl/styles.xml')!.async('string');
    expect(styles).toContain('<cellXfs count="6">');
    expect((styles.match(/wrapText="1"/g) ?? []).length).toBe(2); // General + bold, each cloned once
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet).toMatch(/<c r="C2"(?=[^>]*s="4")[^>]*>/);
    expect(sheet).toMatch(/<c r="C3"(?=[^>]*s="5")[^>]*>/);
  });
});
