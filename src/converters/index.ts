import { convertAlarms } from './alarms';
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
    title: 'Textlisten · Legacy',
    subtitle: 'CICON und ältere Commander HMIs',
    legacyName: 'HMI TextList Konverter (alt)',
    accept: ['.xlsx', '.xlsm'],
    run: convertTextlistLegacy,
  },
  {
    id: 'textlistStandard',
    title: 'Textlisten · Standard',
    subtitle: 'Neuere Commander HMIs',
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
];
