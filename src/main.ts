import { CONVERTERS, type ConverterDef, type ConversionResult } from './converters';
import { ConversionError } from './errors';

const ICONS: Record<string, string> = {
  tags: '<path d="M4 7h16M4 12h10M4 17h7"/><path d="m15 15 2 2 4-4"/>',
  textlistLegacy: '<rect x="4" y="4" width="16" height="16" rx="2"/><path d="M8 9h8M8 13h5"/><path d="M8 17h2"/>',
  textlistStandard: '<path d="M8 6h12M8 12h12M8 18h12"/><path d="M4 6h.01M4 12h.01M4 18h.01"/>',
  alarms: '<path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.94 1.94 0 0 0 3.4 0"/>',
  lineBreaks: '<path d="M4 6h16M4 12h13a3 3 0 0 1 0 6h-4"/><path d="m15 16-2 2 2 2"/><path d="M4 18h5"/>',
};

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

const typesEl = $('types');
const tabsEl = $('guide-tabs');
const dropzone = $('dropzone');
const fileInput = $<HTMLInputElement>('file-input');
const fileRow = $('file-row');
const convertBtn = $<HTMLButtonElement>('convert');
const statusEl = $('status');
const resultEl = $('result');
const errorEl = $('error');
const downloadEl = $<HTMLAnchorElement>('download');

let selected: ConverterDef = CONVERTERS[0];
let file: File | null = null;
let downloadUrl: string | null = null;
let busy = false;

function svg(paths: string): string {
  return `<svg viewBox="0 0 24 24" aria-hidden="true">${paths}</svg>`;
}

function renderTypes(): void {
  typesEl.innerHTML = '';
  tabsEl.innerHTML = '';
  for (const c of CONVERTERS) {
    const label = document.createElement('label');
    label.className = 'type';
    label.innerHTML = `
      <input type="radio" name="type" value="${c.id}" ${c.id === selected.id ? 'checked' : ''} />
      <span class="type-icon">${svg(ICONS[c.id])}</span>
      <span class="type-body">
        <span class="type-title">${c.title}</span>
        <span class="type-sub">${c.subtitle}</span>
        <span class="type-meta">${c.accept.map((a) => `<span class="chip">${a}</span>`).join('')}</span>
      </span>
      <span class="type-check">${svg('<path d="M20 6 9 17l-5-5"/>')}</span>`;
    label.title = `Entspricht „${c.legacyName}“ im Excel-Tool`;
    label.querySelector('input')!.addEventListener('change', () => select(c.id));
    typesEl.appendChild(label);

    const tab = document.createElement('button');
    tab.type = 'button';
    tab.className = 'tab';
    tab.setAttribute('role', 'tab');
    tab.dataset.id = c.id;
    tab.innerHTML = `${svg(ICONS[c.id])}<span>${c.title}</span>`;
    tab.addEventListener('click', () => showGuide(c.id));
    tabsEl.appendChild(tab);
  }
}

function showGuide(id: string): void {
  for (const tab of Array.from(tabsEl.children) as HTMLElement[]) {
    const active = tab.dataset.id === id;
    tab.classList.toggle('active', active);
    tab.setAttribute('aria-selected', String(active));
    if (active) tabsEl.scrollLeft = tab.offsetLeft - tabsEl.offsetLeft - 8;
  }
  for (const p of Array.from(document.querySelectorAll<HTMLElement>('[data-guide]'))) p.hidden = p.dataset.guide !== id;
}

function select(id: string): void {
  selected = CONVERTERS.find((c) => c.id === id) ?? CONVERTERS[0];
  fileInput.accept = selected.accept.join(',');
  $('accept-hint').textContent = `Unterstützt: ${selected.accept.join(', ')}`;
  showGuide(selected.id);
  clearOutcome();
  validateFile();
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1).replace('.', ',')} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1).replace('.', ',')} MB`;
}

function setFile(f: File | null): void {
  file = f;
  clearOutcome();
  if (f) {
    $('file-name').textContent = f.name;
    $('file-size').textContent = formatSize(f.size);
  }
  fileRow.hidden = !f;
  dropzone.hidden = !!f;
  validateFile();
}

function validateFile(): void {
  if (!file) {
    convertBtn.disabled = true;
    setStatus('');
    return;
  }
  const ext = file.name.slice(file.name.lastIndexOf('.')).toLowerCase();
  if (!selected.accept.includes(ext)) {
    convertBtn.disabled = true;
    setStatus(
      ext === '.xls'
        ? 'Alte .xls-Dateien werden nicht unterstützt – bitte in Excel als .xlsx speichern.'
        : `„${selected.title}“ erwartet ${selected.accept.join(' oder ')}.`,
      'warn',
    );
    return;
  }
  convertBtn.disabled = busy;
  setStatus('Bereit zum Konvertieren.');
}

function setTypesDisabled(disabled: boolean): void {
  for (const i of Array.from(typesEl.querySelectorAll('input'))) i.disabled = disabled;
  fileRow.querySelector('button')!.disabled = disabled;
}

function setStatus(text: string, tone: 'info' | 'warn' = 'info'): void {
  statusEl.textContent = text;
  statusEl.dataset.tone = tone;
}

function clearOutcome(): void {
  resultEl.hidden = true;
  errorEl.hidden = true;
  if (downloadUrl) URL.revokeObjectURL(downloadUrl);
  downloadUrl = null;
}

async function convert(): Promise<void> {
  if (!file || busy) return;
  busy = true;
  setTypesDisabled(true);
  clearOutcome();
  convertBtn.disabled = true;
  convertBtn.classList.add('loading');
  setStatus('Datei wird gelesen …');
  try {
    const data = await file.arrayBuffer();
    setStatus('Konvertiere …');
    await new Promise((r) => setTimeout(r, 0));
    const result = await selected.run(data, file.name);
    showResult(result);
    setStatus('');
  } catch (e) {
    const msg =
      e instanceof ConversionError
        ? e.message
        : `Unerwarteter Fehler: ${e instanceof Error ? e.message : String(e)}. Bitte prüfe, ob die Datei in Excel fehlerfrei geöffnet werden kann.`;
    $('error-text').textContent = msg;
    errorEl.hidden = false;
    setStatus('');
    console.error(e);
  } finally {
    busy = false;
    setTypesDisabled(false);
    convertBtn.classList.remove('loading');
    convertBtn.disabled = !file;
  }
}

function showResult(r: ConversionResult): void {
  const type = r.fileName.toLowerCase().endsWith('.xlsm')
    ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
    : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const blob = new Blob([r.data as BlobPart], { type });
  downloadUrl = URL.createObjectURL(blob);
  downloadEl.href = downloadUrl;
  downloadEl.download = r.fileName;
  $('download-label').textContent = `${r.fileName} herunterladen`;
  $('result-summary').textContent = r.summary;

  const notices = $('notices') as HTMLDetailsElement;
  const list = $('notices-list');
  list.innerHTML = '';
  notices.hidden = r.notices.length === 0;
  notices.open = false;
  const warnCount = r.notices.filter((n) => n.level === 'warn').length;
  $('notices-summary').textContent = `${r.notices.length} ${r.notices.length === 1 ? 'Hinweis' : 'Hinweise'} anzeigen${warnCount ? ' – bitte prüfen' : ''}`;
  for (const n of r.notices.slice(0, 500)) {
    const li = document.createElement('li');
    li.dataset.level = n.level;
    if (n.cell) {
      const c = document.createElement('code');
      c.textContent = n.cell;
      li.appendChild(c);
    }
    li.appendChild(document.createTextNode(n.text));
    list.appendChild(li);
  }
  if (r.notices.length > 500) {
    const li = document.createElement('li');
    li.textContent = `… und ${r.notices.length - 500} weitere.`;
    list.appendChild(li);
  }
  resultEl.hidden = false;
  resultEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// File selection & drag and drop
fileInput.addEventListener('change', () => {
  setFile(fileInput.files?.[0] ?? null);
  fileInput.value = '';
});
dropzone.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    fileInput.click();
  }
});
let dragDepth = 0;
window.addEventListener('dragenter', (e) => {
  if (!e.dataTransfer?.types.includes('Files')) return;
  dragDepth++;
  document.body.classList.add('dragging');
});
window.addEventListener('dragleave', () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if (!dragDepth) document.body.classList.remove('dragging');
});
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => {
  e.preventDefault();
  dragDepth = 0;
  document.body.classList.remove('dragging');
  const files = e.dataTransfer?.files;
  if (!files || !files.length || busy) return;
  if (files.length > 1) setStatus('Bitte nur eine Datei gleichzeitig ablegen – die erste wurde übernommen.', 'warn');
  setFile(files[0]);
  if (files.length > 1) setStatus('Bitte nur eine Datei gleichzeitig ablegen – die erste wurde übernommen.', 'warn');
});

$('file-remove').addEventListener('click', () => setFile(null));
convertBtn.addEventListener('click', convert);
$('again').addEventListener('click', () => {
  setFile(null);
  document.getElementById('converter')!.scrollIntoView({ behavior: 'smooth' });
});

renderTypes();
select(selected.id);
