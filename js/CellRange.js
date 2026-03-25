// Utility for working with cell ranges (e.g., "A1:B5")

import { parseAddress, toAddress } from './CellAddress.js';

/**
 * Parse a range string like "A1:B5" into {startCol, startRow, endCol, endRow}
 */
export function parseRange(rangeStr) {
  const parts = rangeStr.split(':');
  if (parts.length === 1) {
    const addr = parseAddress(parts[0]);
    if (!addr) return null;
    return { startCol: addr.col, startRow: addr.row, endCol: addr.col, endRow: addr.row };
  }
  const start = parseAddress(parts[0]);
  const end = parseAddress(parts[1]);
  if (!start || !end) return null;
  return {
    startCol: Math.min(start.col, end.col),
    startRow: Math.min(start.row, end.row),
    endCol: Math.max(start.col, end.col),
    endRow: Math.max(start.row, end.row),
  };
}

/**
 * Iterate over all cell addresses in a range
 */
export function* iterateRange(range) {
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let col = range.startCol; col <= range.endCol; col++) {
      yield { col, row, address: toAddress(col, row) };
    }
  }
}

/**
 * Normalize a range so start <= end
 */
export function normalizeRange(range) {
  return {
    startCol: Math.min(range.startCol, range.endCol),
    startRow: Math.min(range.startRow, range.endRow),
    endCol: Math.max(range.startCol, range.endCol),
    endRow: Math.max(range.startRow, range.endRow),
  };
}
