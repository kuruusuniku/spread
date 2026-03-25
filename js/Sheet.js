// Sheet data model - one sheet within a workbook

import { DEFAULTS } from './constants.js';
import { createCell } from './Cell.js';
import { toAddress } from './CellAddress.js';

export class Sheet {
  constructor(name = 'Sheet1') {
    this.name = name;
    this.cells = new Map();          // key: "A1" -> Cell object
    this.colWidths = new Map();      // col index -> pixel width
    this.rowHeights = new Map();     // row index -> pixel height
    this.maxCols = DEFAULTS.MAX_COLS;
    this.maxRows = DEFAULTS.MAX_ROWS;
    this.freezeRows = 0;
    this.freezeCols = 0;
    // Cached cumulative positions
    this._colPositions = null;
    this._rowPositions = null;
    this._colPosCount = 0;
    this._rowPosCount = 0;
  }

  getCell(col, row) {
    return this.cells.get(toAddress(col, row)) || null;
  }

  getCellByAddress(address) {
    return this.cells.get(address) || null;
  }

  setCell(col, row, cell) {
    const addr = toAddress(col, row);
    if (cell === null || (cell.rawValue === null && cell.formula === null &&
        !cell.format.bold && !cell.format.italic && !cell.format.textColor &&
        !cell.format.bgColor && !cell.format.align && !cell.format.fontSize)) {
      this.cells.delete(addr);
    } else {
      this.cells.set(addr, cell);
    }
  }

  setCellByAddress(address, cell) {
    this.cells.set(address, cell);
  }

  getOrCreateCell(col, row) {
    const addr = toAddress(col, row);
    let cell = this.cells.get(addr);
    if (!cell) {
      cell = createCell();
      this.cells.set(addr, cell);
    }
    return cell;
  }

  getColWidth(col) {
    return this.colWidths.get(col) || DEFAULTS.COL_WIDTH;
  }

  setColWidth(col, width) {
    this.colWidths.set(col, Math.max(20, width));
    this._colPositions = null; // invalidate cache
  }

  getRowHeight(row) {
    return this.rowHeights.get(row) || DEFAULTS.ROW_HEIGHT;
  }

  setRowHeight(row, height) {
    this.rowHeights.set(row, Math.max(8, height));
    this._rowPositions = null; // invalidate cache
  }

  /**
   * Get cumulative column positions (left edge of each column)
   */
  getColPositions(count) {
    if (this._colPositions && this._colPosCount >= count) return this._colPositions;
    const positions = new Float64Array(count + 1);
    positions[0] = 0;
    for (let i = 0; i < count; i++) {
      positions[i + 1] = positions[i] + this.getColWidth(i);
    }
    this._colPositions = positions;
    this._colPosCount = count;
    return positions;
  }

  /**
   * Get cumulative row positions (top edge of each row)
   */
  getRowPositions(count) {
    if (this._rowPositions && this._rowPosCount >= count) return this._rowPositions;
    const positions = new Float64Array(count + 1);
    positions[0] = 0;
    for (let i = 0; i < count; i++) {
      positions[i + 1] = positions[i] + this.getRowHeight(i);
    }
    this._rowPositions = positions;
    this._rowPosCount = count;
    return positions;
  }

  invalidatePositionCache() {
    this._colPositions = null;
    this._rowPositions = null;
  }

  /**
   * Total width of all columns up to a given count
   */
  getTotalWidth(colCount) {
    const pos = this.getColPositions(colCount);
    return pos[colCount];
  }

  /**
   * Total height of all rows up to a given count
   */
  getTotalHeight(rowCount) {
    const pos = this.getRowPositions(rowCount);
    return pos[rowCount];
  }
}
