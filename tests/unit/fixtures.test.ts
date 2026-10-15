import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { CONVERTERS } from '../../src/converters';
import { num, readSheet, str } from '../helpers';

// Sample exports written by openpyxl (an independent XLSX writer). Set OUT_DIR to keep the outputs.
const dir = join(__dirname, '..', 'fixtures');
const load = (name: string) => new Uint8Array(readFileSync(join(dir, name)));
const run = async (id: string, name: string) => {
  const res = await CONVERTERS.find((c) => c.id === id)!.run(load(name).buffer as ArrayBuffer, name);
  if (process.env.OUT_DIR) {
    mkdirSync(process.env.OUT_DIR, { recursive: true });
    writeFileSync(join(process.env.OUT_DIR, `${id}__${res.fileName}`), res.data);
  }
  return res;
};

describe('fixtures written by openpyxl', () => {
  it('HMI-Variablen', async () => {
    const res = await run('tags', 'HMITags.xlsx');
    const get = await readSheet(res.data, 'Hmi Tags');
    expect(get(2, 'A')).toEqual(str('Motor_1__Ist'));
    expect(get(3, 'A')).toEqual(str('F_llstand__2'));
    expect(get(2, 'D')).toEqual(str('DB1.Motor[1].Ist'));
  });

  it('Textlisten · Legacy', async () => {
    const res = await run('textlistLegacy', 'TextLists.xlsx');
    const get = await readSheet(res.data, 'TextListEntry');
    expect(get(2, 'F')).toEqual(str('Automatik'));
    expect(get(2, 'G')).toEqual(str('Auto matic'));
    expect(get(2, 'H')).toEqual(str('—'));
    expect(get(3, 'E')).toEqual(str('Modi_Liste'));
    expect(get(3, 'F')).toEqual(str('Modi_Liste'));
  });

  it('Textlisten · Standard', async () => {
    const res = await run('textlistStandard', 'TextLists.xlsx');
    expect((await readSheet(res.data, 'TextList'))(2, 'A')).toEqual(str('Betriebs_Art'));
    expect((await readSheet(res.data, 'TextListEntry'))(2, 'B')).toEqual(str('Betriebs_Art'));
  });

  it('HMI-Meldungen', async () => {
    const res = await run('alarms', 'HMIAlarms.xlsx');
    const get = await readSheet(res.data, 'DiscreteAlarms');
    expect(get(2, 'A')).toEqual(num(1));
    expect(get(2, 'AG')).toEqual(str('Motor_1_Temp'));
    expect(get(3, 'AG')).toEqual(str('Druck_Soll'));
    expect(get(4, 'D')).toEqual({ kind: 'empty' });
  });
});
