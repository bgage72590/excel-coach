/** A1 address helpers. Rows and columns are 1-based. */

export interface CellRef {
  row: number;
  col: number;
}

export interface RangeRef {
  start: CellRef;
  end: CellRef;
}

export function colToNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numberToCol(n: number): string {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Drops a sheet prefix such as `'Coach-x'!` or `Sheet1!`. */
export function stripSheet(address: string): string {
  const bang = address.lastIndexOf('!');
  return bang >= 0 ? address.slice(bang + 1) : address;
}

export function parseCell(a1: string): CellRef {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(stripSheet(a1).trim());
  if (!m) throw new Error(`Not a cell address: ${a1}`);
  return { col: colToNumber(m[1]), row: Number(m[2]) };
}

export function parseRange(address: string): RangeRef {
  const [a, b] = stripSheet(address).split(':');
  const start = parseCell(a);
  const end = b ? parseCell(b) : start;
  return {
    start: { row: Math.min(start.row, end.row), col: Math.min(start.col, end.col) },
    end: { row: Math.max(start.row, end.row), col: Math.max(start.col, end.col) },
  };
}

export function cellAddress(ref: CellRef): string {
  return `${numberToCol(ref.col)}${ref.row}`;
}

export function rangeAddress(start: CellRef, rows: number, cols: number): string {
  const end = { row: start.row + rows - 1, col: start.col + cols - 1 };
  return rows === 1 && cols === 1 ? cellAddress(start) : `${cellAddress(start)}:${cellAddress(end)}`;
}

export function rangeSize(address: string): { rows: number; cols: number } {
  const r = parseRange(address);
  return { rows: r.end.row - r.start.row + 1, cols: r.end.col - r.start.col + 1 };
}

/** Address of the cell at (rowOffset, colOffset) inside a range. */
export function offsetCell(address: string, rowOffset: number, colOffset: number): string {
  const r = parseRange(address);
  return cellAddress({ row: r.start.row + rowOffset, col: r.start.col + colOffset });
}

/** Quotes a sheet name for use in an address when needed. */
export function sheetRef(sheet: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
}
