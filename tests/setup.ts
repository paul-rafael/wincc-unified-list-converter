import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

// The converters use the browser's DOMParser/XMLSerializer; Node gets xmldom.
Object.assign(globalThis, { DOMParser, XMLSerializer });
