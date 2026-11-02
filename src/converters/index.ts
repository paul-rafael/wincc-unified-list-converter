import { convertAlarms } from './alarms';
import { convertLineBreaks } from './lineBreaks';
import type { ConversionResult, ConverterId } from './common';
import { convertTagEvents } from './tagEvents';
import { convertTags } from './tags';
import { convertTextlists } from './textlists';

export type { ConversionResult, ConverterId, Notice } from './common';

export interface ConverterDef {
  id: ConverterId;
  title: string;
  subtitle: string;
  /** Name of the button in the original Excel tool. */
  legacyName: string;
  accept: string[];
  run: (data: ArrayBuffer, fileName: string) => Promise<ConversionResult>;
}

export const CONVERTERS: ConverterDef[] = [
  {
    id: 'tags',
    title: 'HMI-Variablen',
    subtitle: 'Sonderzeichen in Variablennamen ersetzen',
    legacyName: 'HMI TagList Konverter',
    accept: ['.xlsx'],
    run: convertTags,
  },
  {
    id: 'tagEvents',
    title: 'Variablen-Ereignisse',
    subtitle: 'Ereignis-Skript aus dem Openness-Export',
    legacyName: 'neu, ohne Vorbild im Excel-Tool',
    accept: ['.zip', '.xml', '.xlsx'],
    run: convertTagEvents,
  },
  {
    id: 'textlists',
    title: 'Textlisten',
    subtitle: 'Namen bereinigen, FieldInfos ins Unified-Format',
    legacyName: 'HMI TextList Konverter (alt) und (neu)',
    accept: ['.xlsx', '.xlsm'],
    run: convertTextlists,
  },
  {
    id: 'alarms',
    title: 'HMI-Meldungen',
    subtitle: 'Alarmliste ins Unified-Format',
    legacyName: 'HMI-Alarm Konverter',
    accept: ['.xlsx', '.xlsm'],
    run: convertAlarms,
  },
  {
    id: 'lineBreaks',
    title: 'Meldetexte umbrechen',
    subtitle: 'Automatischer Zeilenumbruch für lange Meldungen',
    legacyName: 'AZU',
    accept: ['.xlsx', '.xlsm'],
    run: convertLineBreaks,
  },
];
