import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { convertTagEvents } from '../../src/converters/tagEvents';
import { parseTagXml, unifiedName } from '../../src/converters/tagEventsXml';

type P = [name: string, value: string | number | '#self' | { tag: string }];
const param = ([name, v]: P) =>
  `<Hmi.Event.FunctionListEntryParameter ID="1" CompositionName="Parameters"><AttributeList><Name>${name}</Name>` +
  (typeof v === 'number' ? `<Value Type="System.Double">${v}</Value></AttributeList>` : typeof v === 'string' && v !== '#self' ? `<Value Type="System.String">${v}</Value></AttributeList>` : '</AttributeList>') +
  (v === '#self' ? '<LinkList><Value TargetID="#0" /></LinkList>' : typeof v === 'object' ? `<LinkList><Value TargetID="@OpenLink"><Name>${v.tag}</Name></Value></LinkList>` : '') +
  '</Hmi.Event.FunctionListEntryParameter>';
const entry = (fn: string, ...params: P[]) =>
  `<Hmi.Event.FunctionListEntry ID="1" CompositionName="FunctionListEntries"><AttributeList><Name>${fn}</Name><Type>SystemFunction</Type></AttributeList><ObjectList>${params.map(param).join('')}</ObjectList></Hmi.Event.FunctionListEntry>`;
const event = (kind: string, ...entries: string[]) =>
  `<Hmi.Event.Event ID="1" CompositionName="Events"><AttributeList><Name>${kind}</Name></AttributeList><ObjectList><Hmi.Event.FunctionListEventHandler ID="2" CompositionName="EventHandler"><ObjectList>${entries.join('')}</ObjectList></Hmi.Event.FunctionListEventHandler></ObjectList></Hmi.Event.Event>`;
const tagXml = (name: string, events: string[] = [], extra = '', dataType = 'Int') =>
  '﻿<?xml version="1.0" encoding="utf-8"?><Document><Engineering version="V20" />' +
  `<Hmi.Tag.Tag ID="0"><AttributeList><Name>${name}</Name>${extra}</AttributeList>` +
  `<LinkList><DataType TargetID="@OpenLink"><Name>${dataType}</Name></DataType></LinkList>` +
  `<ObjectList><MultilingualText ID="5" CompositionName="Comment" />${events.join('')}</ObjectList></Hmi.Tag.Tag></Document>`;

async function exportZip(files: Record<string, string>, manifest: object = { complete: true, failedExports: 0 }) {
  const zip = new JSZip();
  zip.file('manifest.json', JSON.stringify(manifest));
  for (const [n, t] of Object.entries(files)) zip.file(`Tags/${n}.xml`, t);
  return zip.generateAsync({ type: 'uint8array' });
}

function sim(script: string) {
  const tags: Record<string, unknown> = {};
  const screens: string[] = [];
  const HMIRuntime = {
    Tags: { SysFct: { GetTagValue: (n: string) => tags[n] ?? 0, SetTagValue: (n: string, v: unknown) => { tags[n] = v; } } },
    UI: { SysFct: { ChangeScreen: (s: string) => screens.push(s) } },
  };
  const run = new Function('HMIRuntime', `${script.replace('export function', 'function')}\nreturn TagEvent;`)(HMIRuntime) as () => void;
  return { tags, screens, run };
}

const files = {
  adr: tagXml('Line\\Sc\\DbAdr', [event('ChangeValue', entry('LinearScaling', ['Y', { tag: 'Line\\Sc\\AdrA' }], ['a', 1], ['X', '#self'], ['b', 12]), entry('LinearScaling', ['Y', { tag: 'Line\\Sc\\AdrB' }], ['a', 2], ['X', { tag: 'Line\\Sc\\AdrA' }], ['b', -4]))]),
  adrA: tagXml('Line\\Sc\\AdrA'),
  adrB: tagXml('Line\\Sc\\AdrB'),
  lacs: tagXml('Line\\Lacs_0', [event('AboveUpperLimit', entry('SetTag', ['Tag', { tag: '"Line\\Pic.No"' }], ['Value', 1011]), entry('ActivateScreen', ['Screen name', { tag: '2_error_0' }], ['Object number', 0]))], '<LimitUpper2>8</LimitUpper2>'),
  pic: tagXml('Line\\Pic.No'),
  vis: tagXml('Line\\Vis', [], '', 'Bool'),
  mux: tagXml('Line\\Mux', [event('ChangeValue', entry('SetBit', ['Tag', { tag: 'Line\\Vis' }]), entry('ActivateScreenByNumber', ['Screen number', '#self'], ['Object number', 0]), entry('StopRuntime'))]),
};

describe('Variablen-Ereignisse aus dem Openness-Export', () => {
  it('maps Advanced names to Unified names', () => {
    expect(unifiedName('12V_CCS\\HmiSc\\ScDbAdrDatNw')).toBe('12V_CCS_HmiSc_ScDbAdrDatNw');
    expect(unifiedName('"01V_Framework\\_HL._MUX_DB"')).toBe('01V_Framework__HL__MUX_DB');
  });

  it('reads events, function lists, parameters and limits', () => {
    const t = parseTagXml(files.lacs, 'lacs.xml')!;
    expect(t.name).toBe('Line\\Lacs_0');
    expect(t.limitUpper).toEqual({ kind: 'const', type: 'System.Double', value: '8' });
    expect(t.events[0].kind).toBe('AboveUpperLimit');
    expect(t.events[0].entries.map((e) => e.name)).toEqual(['SetTag', 'ActivateScreen']);
    expect(t.events[0].entries[0].params.get('Tag')).toEqual({ kind: 'tag', name: 'Line\\Pic.No' });
  });

  it('generates a script that behaves like the Advanced events', async () => {
    const res = await convertTagEvents(await exportZip(files), 'Export.zip');
    expect(res.fileName).toBe('TagEvent.js');
    expect(res.text).toBe(new TextDecoder().decode(res.data));
    expect(res.summary).toMatch(/3 Ereignisse an 3 von 7 Variablen, 7 Anweisungen, davon 1 zum manuellen/);
    const script = res.text!;
    expect(script).toContain('SetTagValue("Line_Sc_AdrA", value + 12)');
    // the value computed in the same function list is used instead of reading the tag back
    expect(script).toContain('SetTagValue("Line_Sc_AdrB", 2 * (value + 12) - 4)');
    expect(script).toContain('// TODO StopRuntime()');

    const s = sim(script);
    s.tags.Line_Sc_DbAdr = 100;
    s.tags.Line_Lacs_0 = 5;
    s.run();
    expect(s.tags.Line_Sc_AdrA).toBe(112); // first call initialises derived tags
    expect(s.tags.Line_Sc_AdrB).toBe(220);
    expect(s.tags.Line_Vis).toBe(true); // SetBit on a Bool tag
    expect(s.screens).toEqual([]);
    s.tags.Line_Sc_AdrA = 0;
    s.run();
    expect(s.tags.Line_Sc_AdrA).toBe(0); // no change, no event
    s.tags.Line_Lacs_0 = 9;
    s.run();
    expect(s.tags.Line_Pic_No).toBe(1011);
    expect(s.screens).toEqual(['2_error_0']);
    s.tags.Line_Lacs_0 = 10;
    s.run();
    expect(s.screens).toHaveLength(1); // still above the limit: no new event

    expect(res.notices.some((n) => /StopRuntime/.test(n.text))).toBe(true);
    expect(res.notices.some((n) => /Bildnummern/.test(n.text))).toBe(true);
  });

  it('warns about an incomplete export and unknown tags', async () => {
    const zip = await exportZip({ adr: files.adr }, { complete: false, failedExports: 3 });
    const res = await convertTagEvents(zip, 'Export.zip');
    expect(res.notices.some((n) => /unvollständig \(3/.test(n.text))).toBe(true);
    expect(res.notices.some((n) => /Line\\Sc\\AdrA/.test(n.text))).toBe(true);
  });

  it('accepts a single tag XML and stops when nothing has an event', async () => {
    const one = await convertTagEvents(new TextEncoder().encode(files.lacs), 'lacs.xml');
    expect(one.summary).toMatch(/1 Ereignisse an 1 von 1/);
    await expect(convertTagEvents(await exportZip({ a: files.adrA }), 'E.zip')).rejects.toThrow(/kein Ereignis/);
    await expect(convertTagEvents(await exportZip({}), 'E.zip')).rejects.toThrow(/Keine HMI-Variablen/);
  });
});
