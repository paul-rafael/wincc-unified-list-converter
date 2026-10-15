import { convertAlarms } from './alarms';
import { convertLineBreaks } from './lineBreaks';
import type { ConversionResult, ConverterId } from './common';
import { convertTags } from './tags';
import { convertTextlistLegacy } from './textlistLegacy';
import { convertTextlistStandard } from './textlistStandard';

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
    id: 'textlistLegacy',
    title: 'Textlisten · Klassisch',
    subtitle: 'Platzhalter bereinigen, leere Sprachen füllen',
    legacyName: 'HMI TextList Konverter (alt)',
    accept: ['.xlsx', '.xlsm'],
    run: convertTextlistLegacy,
  },
  {
    id: 'textlistStandard',
    title: 'Textlisten · Aktuell',
    subtitle: 'Sonderzeichen in Textlisten-Namen ersetzen',
    legacyName: 'HMI TextList Konverter (neu)',
    accept: ['.xlsx', '.xlsm'],
    run: convertTextlistStandard,
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
