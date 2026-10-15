export function colToLetters(col: number): string {
  let s = '';
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

export function lettersToCol(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function cellRef(row: number, col: number): string {
  return `${colToLetters(col)}${row}`;
}

export function parseRef(ref: string): { row: number; col: number } {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref);
  if (!m) throw new Error(`Ungültige Zellreferenz: ${ref}`);
  return { col: lettersToCol(m[1]), row: parseInt(m[2], 10) };
}

export function parseRange(ref: string): { r1: number; c1: number; r2: number; c2: number } {
  const [a, b] = ref.split(':');
  const p = parseRef(a);
  const q = b ? parseRef(b) : p;
  return { r1: p.row, c1: p.col, r2: q.row, c2: q.col };
}
