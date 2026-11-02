import { describe, expect, it } from 'vitest';
import { convertTagEvents, functionName } from '../../src/converters/tagEvents';
import { makeXlsx, readSheet, str } from '../helpers';

const tagHeader = {
  A: 'Name', B: 'Path', C: 'Connection', D: 'PLC tag', E: 'DataType', F: 'Acquisition mode',
  G: 'Limit Upper 2 Type', H: 'Limit Upper 2', I: 'Limit Lower 2 Type', J: 'Limit Lower 2',
};
const tagExport = () =>
  makeXlsx([
    {
      name: 'Hmi Tags',
      rows: {
        1: tagHeader,
        2: { A: '12V_CCS_Pu_Adr', E: 'Int', F: 'Cyclic in operation', G: 'None', H: '<No Value>', I: 'None', J: '<No Value>' },
        3: { A: '12V_CCS_Lacs_0', E: 'Int', F: 'Cyclic in operation', G: 'Constant', H: '8', I: 'None', J: '<No Value>' },
        4: { A: '12V_CCS_RunNo', E: 'Int', F: 'Continuous', G: 'Constant', H: '2047', I: 'Constant', J: '1000' },
      },
    },
  ]);

/** Runs the generated script against a fake HMIRuntime. */
function runScript(script: string, tags: Record<string, unknown>, calls = 1) {
  const writes: Record<string, unknown> = {};
  const screens: string[] = [];
  const HMIRuntime = {
    Tags: { SysFct: { GetTagValue: (n: string) => tags[n] ?? 0, SetTagValue: (n: string, v: unknown) => { writes[n] = v; tags[n] = v; } } },
    UI: { SysFct: { ChangeScreen: (s: string) => screens.push(s) } },
  };
  const run = new Function('HMIRuntime', `${script.replace('export function', 'function')}\nreturn TagEvent;`)(HMIRuntime) as () => void;
  return { run, writes, screens, call: () => { for (let i = 0; i < calls; i++) run(); } };
}

describe('Variablen-Ereignisse', () => {
  it('derives readable function names', () => {
    expect(functionName('12V_CCS_HmiAxis_StatRunNo')).toBe('HmiAxis_StatRunNo');
    expect(functionName('Motor.Ist')).toBe('Motor_Ist');
    expect(functionName('1_Tag')).toBe('Ev_1_Tag');
  });

  it('builds a template from the HMI tag export with the limit tags pre-filled', async () => {
    const res = await convertTagEvents(await tagExport(), 'HMI_Tags.xlsx');
    expect(res.fileName).toBe('TagEvents.xlsx');
    expect(res.summary).toMatch(/3 Variablen übernommen, 3 Grenzwert-Ereignisse/);
    const ev = await readSheet(res.data, 'Ereignisse');
    expect(ev(1, 'B')).toEqual(str('Trigger-Variable'));
    expect(ev(2, 'B')).toEqual(str('12V_CCS_Lacs_0'));
    expect(ev(2, 'C')).toEqual(str('Überschreitung'));
    expect(ev(2, 'D')).toEqual({ kind: 'number', value: 8 });
    expect(ev(4, 'C')).toEqual(str('Unterschreitung'));
    expect((await readSheet(res.data, 'Variablen'))(4, 'A')).toEqual(str('12V_CCS_RunNo'));
  });

  const filled = (rows: Record<number, Record<string, string | number>>) =>
    makeXlsx([
      { name: 'Ereignisse', rows: { 1: { A: 'Funktion', B: 'Trigger-Variable', C: 'Ereignis', D: 'Grenzwert', E: 'Aktion', F: 'Ziel', G: 'Wert / Formel' }, ...rows } },
      { name: 'Variablen', rows: { 1: { A: 'Name' }, 2: { A: 'Pu_Adr' }, 3: { A: 'Pu_Pl1' }, 4: { A: 'Lacs_0' }, 5: { A: 'Pic_1' }, 6: { A: 'Plc' }, 7: { A: 'ModNo' }, 8: { A: 'Err' } } },
    ]);

  it('generates a script that sets tags, respects limits and changes screens only on change', async () => {
    const input = await filled({
      2: { A: 'Pu', B: 'Pu_Adr', C: 'Wertänderung', E: 'Variable setzen', F: 'Pu_Pl1', G: '{Trigger} + 78' },
      3: { A: 'Mod', B: 'Pu_Adr', C: 'Wertänderung', E: 'Variable setzen', F: 'ModNo', G: '10000 * {Plc} + {Trigger}' },
      4: { A: 'Lacs', B: 'Lacs_0', C: 'Überschreitung', D: 8, E: 'Variable setzen', F: 'Pic_1', G: 1011 },
      5: { A: 'Err', B: 'Err', C: 'Wertänderung', E: 'Bild wechseln', F: '2_error_0' },
      6: { B: 'Lacs_0', C: 'Überschreitung', D: 8 },
    });
    const res = await convertTagEvents(input, 'TagEvents.xlsx');
    expect(res.fileName).toBe('TagEvent.js');
    expect(res.summary).toMatch(/4 Funktionen mit 4 Aktionen/);
    const script = new TextDecoder().decode(res.data);
    expect(script).toContain('var TagEventLast = {};');

    const tags: Record<string, unknown> = { Pu_Adr: 100, Plc: 2, Lacs_0: 8, Err: 1 };
    const sim = runScript(script, tags);
    sim.call();
    expect(sim.writes).toEqual({ Pu_Pl1: 178, ModNo: 20100 });
    expect(sim.screens).toEqual([]);
    tags.Lacs_0 = 9;
    tags.Err = 2;
    sim.call();
    expect(sim.writes.Pic_1).toBe(1011);
    expect(sim.screens).toEqual(['2_error_0']);
    sim.call();
    expect(sim.screens).toHaveLength(1);
  });

  it('reports unknown tags, invalid formulas and unknown events with their cell', async () => {
    const input = await filled({
      2: { B: 'Pu_Adr', C: 'Wertänderung', E: 'Variable setzen', F: 'Tippfehler', G: '{Trigger} + 1' },
      3: { B: 'Pu_Adr', C: 'Wertänderung', E: 'Variable setzen', F: 'Pu_Pl1', G: 'alert(1)' },
      4: { B: 'Pu_Adr', C: 'Bei Druck', E: 'Variable setzen', F: 'Pu_Pl1', G: '1' },
    });
    const res = await convertTagEvents(input, 'TagEvents.xlsx');
    const cells = res.notices.map((n) => n.cell);
    expect(cells).toEqual(expect.arrayContaining(['Ereignisse!F2', 'Ereignisse!G3', 'Ereignisse!C4']));
    expect(new TextDecoder().decode(res.data)).not.toContain('alert');
  });

  it('stops when no row has an action', async () => {
    await expect(convertTagEvents(await filled({ 2: { B: 'Pu_Adr', C: 'Überschreitung', D: 1 } }), 'T.xlsx')).rejects.toThrow(/keine Zeile mit Aktion/);
  });

  it('is refused by the other functions', async () => {
    const { convertTags } = await import('../../src/converters/tags');
    await expect(convertTags(await filled({}), 'T.xlsx')).rejects.toThrow(/Ereignis-Tabelle/);
  });
});
