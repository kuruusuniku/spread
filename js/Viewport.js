// Viewport - manages scroll state, visible range, and coordinate translation

import { DEFAULTS } from './constants.js';

export class Viewport {
  constructor(sheet) {
    this.sheet = sheet;
    this.scrollX = 0;
    this.scrollY = 0;
    this.canvasWidth = 0;
    this.canvasHeight = 0;
    this.headerWidth = DEFAULTS.HEADER_COL_WIDTH;
    this.headerHeight = DEFAULTS.HEADER_ROW_HEIGHT;
  }

  setSize(width, height) {
    this.canvasWidth = width;
    this.canvasHeight = height;
  }

  /**
   * Get the width of frozen columns in pixels
   */
  get frozenWidth() {
    let w = 0;
    for (let i = 0; i < this.sheet.freezeCols; i++) {
      w += this.sheet.getColWidth(i);
    }
    return w;
  }

  /**
   * Get the height of frozen rows in pixels
   */
  get frozenHeight() {
    let h = 0;
    for (let i = 0; i < this.sheet.freezeRows; i++) {
      h += this.sheet.getRowHeight(i);
    }
    return h;
  }

  /**
   * Get the visible column range (scrollable area)
   */
  getVisibleCols() {
    const startX = this.scrollX;
    const endX = startX + this.canvasWidth - this.headerWidth - this.frozenWidth;
    const colPositions = this.sheet.getColPositions(this.sheet.maxCols);

    let start = this.sheet.freezeCols;
    // Binary search for start column
    let lo = this.sheet.freezeCols, hi = this.sheet.maxCols - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (colPositions[mid + 1] <= startX + colPositions[this.sheet.freezeCols]) {
        lo = mid + 1;
      } else {
        hi = mid;
      }
    }
    start = lo;

    // Find end column
    let end = start;
    const maxX = startX + colPositions[this.sheet.freezeCols] + endX;
    while (end < this.sheet.maxCols && colPositions[end] < maxX) {
      end++;
    }

    return { start, end: Math.min(end + 1, this.sheet.maxCols) };
  }

  /**
   * Get the visible row range (scrollable area)
   */
  getVisibleRows() {
    const startY = this.scrollY;
    const endY = startY + this.canvasHeight - this.headerHeight - this.frozenHeight;
    const rowPositions = this.sheet.getRowPositions(this.sheet.maxRows);

    let start = this.sheet.freezeRows;
    // Binary search for start row
    let lo = this.sheet.freezeRows, hi = this.sheet.maxRows - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (rowPositions[mid + 1] <= startY + rowPositions[this.sheet.freezeRows]) {
        lo = mid + 1;
      } else {
        hi = mid;
      }
    }
    start = lo;

    // Find end row
    let end = start;
    const maxY = startY + rowPositions[this.sheet.freezeRows] + endY;
    while (end < this.sheet.maxRows && rowPositions[end] < maxY) {
      end++;
    }

    return { start, end: Math.min(end + 1, this.sheet.maxRows) };
  }

  /**
   * Get the X position of a column on the canvas
   */
  getColX(col) {
    const colPositions = this.sheet.getColPositions(col + 1);
    if (col < this.sheet.freezeCols) {
      return this.headerWidth + colPositions[col];
    }
    return this.headerWidth + this.frozenWidth + colPositions[col] - colPositions[this.sheet.freezeCols] - this.scrollX;
  }

  /**
   * Get the Y position of a row on the canvas
   */
  getRowY(row) {
    const rowPositions = this.sheet.getRowPositions(row + 1);
    if (row < this.sheet.freezeRows) {
      return this.headerHeight + rowPositions[row];
    }
    return this.headerHeight + this.frozenHeight + rowPositions[row] - rowPositions[this.sheet.freezeRows] - this.scrollY;
  }

  /**
   * Convert canvas pixel coordinates to cell col, row
   */
  hitTest(canvasX, canvasY) {
    if (canvasX < this.headerWidth || canvasY < this.headerHeight) {
      return this.hitTestHeader(canvasX, canvasY);
    }

    const colPositions = this.sheet.getColPositions(this.sheet.maxCols);
    const rowPositions = this.sheet.getRowPositions(this.sheet.maxRows);

    let col = -1, row = -1;

    // Determine column
    const relX = canvasX - this.headerWidth;
    if (relX < this.frozenWidth) {
      // In frozen column area
      for (let c = 0; c < this.sheet.freezeCols; c++) {
        if (relX >= colPositions[c] && relX < colPositions[c + 1]) {
          col = c;
          break;
        }
      }
    } else {
      // In scrollable area
      const absX = relX - this.frozenWidth + this.scrollX + colPositions[this.sheet.freezeCols];
      // Binary search
      let lo = this.sheet.freezeCols, hi = this.sheet.maxCols - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        if (absX >= colPositions[mid] && absX < colPositions[mid + 1]) {
          col = mid;
          break;
        } else if (absX < colPositions[mid]) {
          hi = mid - 1;
        } else {
          lo = mid + 1;
        }
      }
      if (col === -1 && lo < this.sheet.maxCols) col = lo;
    }

    // Determine row
    const relY = canvasY - this.headerHeight;
    if (relY < this.frozenHeight) {
      // In frozen row area
      for (let r = 0; r < this.sheet.freezeRows; r++) {
        if (relY >= rowPositions[r] && relY < rowPositions[r + 1]) {
          row = r;
          break;
        }
      }
    } else {
      // In scrollable area
      const absY = relY - this.frozenHeight + this.scrollY + rowPositions[this.sheet.freezeRows];
      // Binary search
      let lo = this.sheet.freezeRows, hi = this.sheet.maxRows - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        if (absY >= rowPositions[mid] && absY < rowPositions[mid + 1]) {
          row = mid;
          break;
        } else if (absY < rowPositions[mid]) {
          hi = mid - 1;
        } else {
          lo = mid + 1;
        }
      }
      if (row === -1 && lo < this.sheet.maxRows) row = lo;
    }

    return { col: Math.max(0, col), row: Math.max(0, row), area: 'cell' };
  }

  /**
   * Hit test for header area
   */
  hitTestHeader(canvasX, canvasY) {
    const colPositions = this.sheet.getColPositions(this.sheet.maxCols);
    const rowPositions = this.sheet.getRowPositions(this.sheet.maxRows);

    if (canvasY < this.headerHeight && canvasX >= this.headerWidth) {
      // Column header
      const relX = canvasX - this.headerWidth;
      let col = -1;
      if (relX < this.frozenWidth) {
        for (let c = 0; c < this.sheet.freezeCols; c++) {
          if (relX >= colPositions[c] && relX < colPositions[c + 1]) { col = c; break; }
        }
      } else {
        const absX = relX - this.frozenWidth + this.scrollX + colPositions[this.sheet.freezeCols];
        let lo = this.sheet.freezeCols, hi = this.sheet.maxCols - 1;
        while (lo <= hi) {
          const mid = (lo + hi) >> 1;
          if (absX >= colPositions[mid] && absX < colPositions[mid + 1]) { col = mid; break; }
          else if (absX < colPositions[mid]) hi = mid - 1;
          else lo = mid + 1;
        }
      }

      // Check if near border for resize
      const colX = col >= 0 ? this.getColX(col) : -1;
      const colW = col >= 0 ? this.sheet.getColWidth(col) : 0;
      const nearRightBorder = col >= 0 && Math.abs(canvasX - (colX + colW)) < 5;

      return { col: Math.max(0, col), row: -1, area: 'colHeader', resize: nearRightBorder };
    }

    if (canvasX < this.headerWidth && canvasY >= this.headerHeight) {
      // Row header
      const relY = canvasY - this.headerHeight;
      let row = -1;
      if (relY < this.frozenHeight) {
        for (let r = 0; r < this.sheet.freezeRows; r++) {
          if (relY >= rowPositions[r] && relY < rowPositions[r + 1]) { row = r; break; }
        }
      } else {
        const absY = relY - this.frozenHeight + this.scrollY + rowPositions[this.sheet.freezeRows];
        let lo = this.sheet.freezeRows, hi = this.sheet.maxRows - 1;
        while (lo <= hi) {
          const mid = (lo + hi) >> 1;
          if (absY >= rowPositions[mid] && absY < rowPositions[mid + 1]) { row = mid; break; }
          else if (absY < rowPositions[mid]) hi = mid - 1;
          else lo = mid + 1;
        }
      }

      // Check if near border for resize
      const rowY = row >= 0 ? this.getRowY(row) : -1;
      const rowH = row >= 0 ? this.sheet.getRowHeight(row) : 0;
      const nearBottomBorder = row >= 0 && Math.abs(canvasY - (rowY + rowH)) < 5;

      return { col: -1, row: Math.max(0, row), area: 'rowHeader', resize: nearBottomBorder };
    }

    return { col: -1, row: -1, area: 'corner' };
  }

  /**
   * Scroll to make a cell visible
   */
  scrollToCell(col, row) {
    const colPositions = this.sheet.getColPositions(col + 2);
    const rowPositions = this.sheet.getRowPositions(row + 2);

    // Only scroll for non-frozen cells
    if (col >= this.sheet.freezeCols) {
      const cellLeft = colPositions[col] - colPositions[this.sheet.freezeCols];
      const cellRight = cellLeft + this.sheet.getColWidth(col);
      const viewWidth = this.canvasWidth - this.headerWidth - this.frozenWidth;

      if (cellLeft < this.scrollX) {
        this.scrollX = cellLeft;
      } else if (cellRight > this.scrollX + viewWidth) {
        this.scrollX = cellRight - viewWidth;
      }
    }

    if (row >= this.sheet.freezeRows) {
      const cellTop = rowPositions[row] - rowPositions[this.sheet.freezeRows];
      const cellBottom = cellTop + this.sheet.getRowHeight(row);
      const viewHeight = this.canvasHeight - this.headerHeight - this.frozenHeight;

      if (cellTop < this.scrollY) {
        this.scrollY = cellTop;
      } else if (cellBottom > this.scrollY + viewHeight) {
        this.scrollY = cellBottom - viewHeight;
      }
    }

    this.clampScroll();
  }

  clampScroll() {
    const totalW = this.sheet.getTotalWidth(this.sheet.maxCols) -
                   (this.sheet.freezeCols > 0 ? this.sheet.getColPositions(this.sheet.freezeCols + 1)[this.sheet.freezeCols] : 0);
    const totalH = this.sheet.getTotalHeight(this.sheet.maxRows) -
                   (this.sheet.freezeRows > 0 ? this.sheet.getRowPositions(this.sheet.freezeRows + 1)[this.sheet.freezeRows] : 0);
    const viewW = this.canvasWidth - this.headerWidth - this.frozenWidth;
    const viewH = this.canvasHeight - this.headerHeight - this.frozenHeight;

    this.scrollX = Math.max(0, Math.min(this.scrollX, totalW - viewW));
    this.scrollY = Math.max(0, Math.min(this.scrollY, totalH - viewH));
  }
}
