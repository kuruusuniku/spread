// Selection model - tracks active cell and selected ranges

export class SelectionModel {
  constructor() {
    this.activeCell = { col: 0, row: 0 };
    this.ranges = [{ startCol: 0, startRow: 0, endCol: 0, endRow: 0 }];
    this.isSelecting = false;
  }

  setActiveCell(col, row) {
    this.activeCell = { col, row };
    this.ranges = [{ startCol: col, startRow: row, endCol: col, endRow: row }];
  }

  startSelection(col, row, extend = false) {
    this.isSelecting = true;
    if (!extend) {
      this.activeCell = { col, row };
      this.ranges = [{ startCol: col, startRow: row, endCol: col, endRow: row }];
    } else {
      // Extend from active cell
      const range = this.ranges[this.ranges.length - 1];
      range.endCol = col;
      range.endRow = row;
    }
  }

  updateSelection(col, row) {
    if (!this.isSelecting) return;
    const range = this.ranges[this.ranges.length - 1];
    range.endCol = col;
    range.endRow = row;
  }

  endSelection() {
    this.isSelecting = false;
  }

  /**
   * Get the normalized primary selection range
   */
  getPrimaryRange() {
    const r = this.ranges[0];
    return {
      startCol: Math.min(r.startCol, r.endCol),
      startRow: Math.min(r.startRow, r.endRow),
      endCol: Math.max(r.startCol, r.endCol),
      endRow: Math.max(r.startRow, r.endRow),
    };
  }

  /**
   * Check if a cell is within the selection
   */
  isCellSelected(col, row) {
    for (const r of this.ranges) {
      const minC = Math.min(r.startCol, r.endCol);
      const maxC = Math.max(r.startCol, r.endCol);
      const minR = Math.min(r.startRow, r.endRow);
      const maxR = Math.max(r.startRow, r.endRow);
      if (col >= minC && col <= maxC && row >= minR && row <= maxR) return true;
    }
    return false;
  }

  /**
   * Move active cell by offset, clamped to sheet bounds
   */
  move(dCol, dRow, maxCol, maxRow, extend = false) {
    const newCol = Math.max(0, Math.min(maxCol - 1, this.activeCell.col + dCol));
    const newRow = Math.max(0, Math.min(maxRow - 1, this.activeCell.row + dRow));
    if (extend) {
      const range = this.ranges[0];
      range.endCol = newCol;
      range.endRow = newRow;
    } else {
      this.setActiveCell(newCol, newRow);
    }
    // Return new active position for scrolling
    return extend ? { col: newCol, row: newRow } : this.activeCell;
  }

  moveToCell(col, row) {
    this.setActiveCell(col, row);
  }
}
