// Utility for cell address parsing and conversion

/**
 * Convert a 0-based column index to a column label (0 -> A, 25 -> Z, 26 -> AA, etc.)
 */
export function colIndexToLabel(index) {
  let label = '';
  let n = index;
  while (n >= 0) {
    label = String.fromCharCode((n % 26) + 65) + label;
    n = Math.floor(n / 26) - 1;
  }
  return label;
}

/**
 * Convert a column label to a 0-based index (A -> 0, Z -> 25, AA -> 26, etc.)
 */
export function labelToColIndex(label) {
  let index = 0;
  for (let i = 0; i < label.length; i++) {
    index = index * 26 + (label.charCodeAt(i) - 64);
  }
  return index - 1;
}

/**
 * Parse a cell address string like "A1", "$A$1", "AB123" into {col, row, absCol, absRow}
 */
export function parseAddress(addr) {
  const match = addr.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/i);
  if (!match) return null;
  return {
    col: labelToColIndex(match[2].toUpperCase()),
    row: parseInt(match[4], 10) - 1,
    absCol: match[1] === '$',
    absRow: match[3] === '$',
  };
}

/**
 * Convert col, row (0-based) to a cell address string like "A1"
 */
export function toAddress(col, row) {
  return colIndexToLabel(col) + (row + 1);
}

/**
 * Adjust a cell reference when copying/pasting with an offset
 */
export function adjustReference(addr, colOffset, rowOffset) {
  const parsed = parseAddress(addr);
  if (!parsed) return addr;
  const newCol = parsed.absCol ? parsed.col : parsed.col + colOffset;
  const newRow = parsed.absRow ? parsed.row : parsed.row + rowOffset;
  if (newCol < 0 || newRow < 0) return '#REF!';
  const prefix = (parsed.absCol ? '$' : '') + colIndexToLabel(newCol) + (parsed.absRow ? '$' : '');
  return prefix + (newRow + 1);
}
