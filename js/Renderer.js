// Canvas rendering engine

import { DEFAULTS, COLORS } from './constants.js';
import { colIndexToLabel, toAddress } from './CellAddress.js';

export class Renderer {
  constructor(canvas, viewport, sheet, selectionModel) {
    this.canvas = canvas;
    this.ctx = canvas.getContext('2d');
    this.viewport = viewport;
    this.sheet = sheet;
    this.selection = selectionModel;
    this._dirty = true;
    this._dpr = window.devicePixelRatio || 1;
    this.copyRange = null; // marching ants for copy
    this._marchOffset = 0;
    this._marchTimer = null;

    this._setupHiDPI();
    this._startRenderLoop();
  }

  _setupHiDPI() {
    const rect = this.canvas.parentElement.getBoundingClientRect();
    this.canvas.style.width = rect.width + 'px';
    this.canvas.style.height = rect.height + 'px';
    this.canvas.width = rect.width * this._dpr;
    this.canvas.height = rect.height * this._dpr;
    this.ctx.scale(this._dpr, this._dpr);
    this.viewport.setSize(rect.width, rect.height);
  }

  resize() {
    this._dpr = window.devicePixelRatio || 1;
    this._setupHiDPI();
    this.requestRender();
  }

  requestRender() {
    this._dirty = true;
  }

  _startRenderLoop() {
    const loop = () => {
      if (this._dirty) {
        this._dirty = false;
        this.render();
      }
      requestAnimationFrame(loop);
    };
    requestAnimationFrame(loop);
  }

  updateSheet(sheet) {
    this.sheet = sheet;
    this.viewport.sheet = sheet;
    this.requestRender();
  }

  render() {
    const ctx = this.ctx;
    const vp = this.viewport;
    const w = vp.canvasWidth;
    const h = vp.canvasHeight;

    // Clear
    ctx.clearRect(0, 0, w, h);

    // Get visible ranges
    const visibleCols = vp.getVisibleCols();
    const visibleRows = vp.getVisibleRows();

    // Draw cells (background, grid, text)
    this._drawCells(ctx, vp, visibleCols, visibleRows);

    // Draw frozen cells if any
    if (this.sheet.freezeCols > 0 || this.sheet.freezeRows > 0) {
      this._drawFrozenCells(ctx, vp);
    }

    // Draw selection
    this._drawSelection(ctx, vp);

    // Draw copy border (marching ants)
    if (this.copyRange) {
      this._drawCopyBorder(ctx, vp);
    }

    // Draw headers
    this._drawColumnHeaders(ctx, vp, visibleCols);
    this._drawRowHeaders(ctx, vp, visibleRows);

    // Draw corner
    ctx.fillStyle = COLORS.HEADER_BG;
    ctx.fillRect(0, 0, vp.headerWidth, vp.headerHeight);
    ctx.strokeStyle = COLORS.HEADER_BORDER;
    ctx.lineWidth = 1;
    ctx.strokeRect(0, 0, vp.headerWidth, vp.headerHeight);

    // Draw freeze line
    if (this.sheet.freezeCols > 0) {
      const x = vp.headerWidth + vp.frozenWidth;
      ctx.strokeStyle = COLORS.FREEZE_LINE;
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, h);
      ctx.stroke();
    }
    if (this.sheet.freezeRows > 0) {
      const y = vp.headerHeight + vp.frozenHeight;
      ctx.strokeStyle = COLORS.FREEZE_LINE;
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.moveTo(0, y);
      ctx.lineTo(w, y);
      ctx.stroke();
    }
  }

  _drawCells(ctx, vp, visibleCols, visibleRows) {
    const sheet = this.sheet;
    ctx.save();

    // Clip to cell area (excluding headers)
    ctx.beginPath();
    ctx.rect(vp.headerWidth + vp.frozenWidth, vp.headerHeight + vp.frozenHeight,
             vp.canvasWidth - vp.headerWidth - vp.frozenWidth,
             vp.canvasHeight - vp.headerHeight - vp.frozenHeight);
    ctx.clip();

    for (let row = visibleRows.start; row < visibleRows.end; row++) {
      const y = vp.getRowY(row);
      const rowH = sheet.getRowHeight(row);
      if (y + rowH < vp.headerHeight + vp.frozenHeight) continue;
      if (y > vp.canvasHeight) break;

      for (let col = visibleCols.start; col < visibleCols.end; col++) {
        const x = vp.getColX(col);
        const colW = sheet.getColWidth(col);
        if (x + colW < vp.headerWidth + vp.frozenWidth) continue;
        if (x > vp.canvasWidth) break;

        this._drawCell(ctx, col, row, x, y, colW, rowH);
      }
    }

    ctx.restore();
  }

  _drawFrozenCells(ctx, vp) {
    const sheet = this.sheet;

    // Frozen columns (scrollable rows)
    if (sheet.freezeCols > 0) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(vp.headerWidth, vp.headerHeight + vp.frozenHeight,
               vp.frozenWidth, vp.canvasHeight - vp.headerHeight - vp.frozenHeight);
      ctx.clip();

      const visibleRows = vp.getVisibleRows();
      for (let row = visibleRows.start; row < visibleRows.end; row++) {
        const y = vp.getRowY(row);
        const rowH = sheet.getRowHeight(row);
        for (let col = 0; col < sheet.freezeCols; col++) {
          const x = vp.getColX(col);
          const colW = sheet.getColWidth(col);
          this._drawCell(ctx, col, row, x, y, colW, rowH);
        }
      }
      ctx.restore();
    }

    // Frozen rows (scrollable columns)
    if (sheet.freezeRows > 0) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(vp.headerWidth + vp.frozenWidth, vp.headerHeight,
               vp.canvasWidth - vp.headerWidth - vp.frozenWidth, vp.frozenHeight);
      ctx.clip();

      const visibleCols = vp.getVisibleCols();
      for (let row = 0; row < sheet.freezeRows; row++) {
        const y = vp.getRowY(row);
        const rowH = sheet.getRowHeight(row);
        for (let col = visibleCols.start; col < visibleCols.end; col++) {
          const x = vp.getColX(col);
          const colW = sheet.getColWidth(col);
          this._drawCell(ctx, col, row, x, y, colW, rowH);
        }
      }
      ctx.restore();
    }

    // Top-left frozen corner
    if (sheet.freezeCols > 0 && sheet.freezeRows > 0) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(vp.headerWidth, vp.headerHeight, vp.frozenWidth, vp.frozenHeight);
      ctx.clip();

      for (let row = 0; row < sheet.freezeRows; row++) {
        const y = vp.getRowY(row);
        const rowH = sheet.getRowHeight(row);
        for (let col = 0; col < sheet.freezeCols; col++) {
          const x = vp.getColX(col);
          const colW = sheet.getColWidth(col);
          this._drawCell(ctx, col, row, x, y, colW, rowH);
        }
      }
      ctx.restore();
    }
  }

  _drawCell(ctx, col, row, x, y, w, h) {
    const cell = this.sheet.getCell(col, row);

    // Background
    const bgColor = cell?.format?.bgColor || COLORS.CELL_BG;
    if (bgColor !== COLORS.CELL_BG) {
      ctx.fillStyle = bgColor;
      ctx.fillRect(x, y, w, h);
    }

    // Grid lines
    ctx.strokeStyle = COLORS.GRID_LINE;
    ctx.lineWidth = 0.5;
    ctx.strokeRect(x + 0.5, y + 0.5, w, h);

    // Cell text
    if (cell && (cell.computedValue !== null && cell.computedValue !== undefined && cell.computedValue !== '')) {
      const fmt = cell.format || {};
      const fontSize = fmt.fontSize || DEFAULTS.FONT_SIZE;
      const fontStyle = (fmt.bold ? 'bold ' : '') + (fmt.italic ? 'italic ' : '');
      ctx.font = `${fontStyle}${fontSize}px ${DEFAULTS.FONT_FAMILY}`;
      ctx.fillStyle = fmt.textColor || COLORS.CELL_TEXT;

      const text = this._formatValue(cell.computedValue);
      const padding = DEFAULTS.CELL_PADDING;

      // Determine alignment
      let align = fmt.align;
      if (!align) {
        align = typeof cell.computedValue === 'number' ? 'right' : 'left';
      }

      ctx.textBaseline = 'middle';
      ctx.save();
      ctx.beginPath();
      ctx.rect(x + 1, y + 1, w - 2, h - 2);
      ctx.clip();

      if (align === 'right') {
        ctx.textAlign = 'right';
        ctx.fillText(text, x + w - padding, y + h / 2);
      } else if (align === 'center') {
        ctx.textAlign = 'center';
        ctx.fillText(text, x + w / 2, y + h / 2);
      } else {
        ctx.textAlign = 'left';
        ctx.fillText(text, x + padding, y + h / 2);
      }
      ctx.restore();
    }
  }

  _formatValue(value) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'number') {
      if (Number.isNaN(value)) return '#NaN';
      if (!Number.isFinite(value)) return '#INF';
      // Format with reasonable precision
      if (Number.isInteger(value)) return String(value);
      return String(Math.round(value * 1e10) / 1e10);
    }
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    return String(value);
  }

  _drawSelection(ctx, vp) {
    const sel = this.selection;
    const range = sel.getPrimaryRange();

    // Draw range fill
    const startX = vp.getColX(range.startCol);
    const startY = vp.getRowY(range.startRow);
    let endX = vp.getColX(range.endCol) + this.sheet.getColWidth(range.endCol);
    let endY = vp.getRowY(range.endRow) + this.sheet.getRowHeight(range.endRow);

    // Range fill
    if (range.startCol !== range.endCol || range.startRow !== range.endRow) {
      ctx.fillStyle = COLORS.SELECTION_FILL;
      ctx.fillRect(startX, startY, endX - startX, endY - startY);
    }

    // Range border
    ctx.strokeStyle = COLORS.SELECTION_BORDER;
    ctx.lineWidth = 2;
    ctx.strokeRect(startX, startY, endX - startX, endY - startY);

    // Active cell border (thicker)
    const acX = vp.getColX(sel.activeCell.col);
    const acY = vp.getRowY(sel.activeCell.row);
    const acW = this.sheet.getColWidth(sel.activeCell.col);
    const acH = this.sheet.getRowHeight(sel.activeCell.row);

    // White background for active cell
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(acX + 1, acY + 1, acW - 2, acH - 2);

    // Re-draw cell content for active cell
    const cell = this.sheet.getCell(sel.activeCell.col, sel.activeCell.row);
    if (cell) {
      this._drawCell(ctx, sel.activeCell.col, sel.activeCell.row, acX, acY, acW, acH);
    }

    ctx.strokeStyle = COLORS.ACTIVE_CELL_BORDER;
    ctx.lineWidth = 2;
    ctx.strokeRect(acX, acY, acW, acH);

    // Small fill handle at bottom-right of active cell
    ctx.fillStyle = COLORS.ACTIVE_CELL_BORDER;
    ctx.fillRect(acX + acW - 4, acY + acH - 4, 6, 6);
  }

  _drawCopyBorder(ctx, vp) {
    const range = this.copyRange;
    const startX = vp.getColX(range.startCol);
    const startY = vp.getRowY(range.startRow);
    const endX = vp.getColX(range.endCol) + this.sheet.getColWidth(range.endCol);
    const endY = vp.getRowY(range.endRow) + this.sheet.getRowHeight(range.endRow);

    ctx.save();
    ctx.strokeStyle = COLORS.COPY_BORDER;
    ctx.lineWidth = 2;
    ctx.setLineDash([6, 4]);
    ctx.lineDashOffset = -this._marchOffset;
    ctx.strokeRect(startX, startY, endX - startX, endY - startY);
    ctx.restore();
  }

  startMarchingAnts() {
    this.stopMarchingAnts();
    this._marchTimer = setInterval(() => {
      this._marchOffset = (this._marchOffset + 1) % 10;
      this.requestRender();
    }, 100);
  }

  stopMarchingAnts() {
    if (this._marchTimer) {
      clearInterval(this._marchTimer);
      this._marchTimer = null;
    }
    this.copyRange = null;
    this._marchOffset = 0;
  }

  _drawColumnHeaders(ctx, vp, visibleCols) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(vp.headerWidth, 0, vp.canvasWidth - vp.headerWidth, vp.headerHeight);
    ctx.clip();

    const sel = this.selection;
    const range = sel.getPrimaryRange();

    // Draw frozen column headers
    for (let col = 0; col < this.sheet.freezeCols; col++) {
      this._drawColHeader(ctx, vp, col, range);
    }

    // Draw scrollable column headers
    for (let col = visibleCols.start; col < visibleCols.end; col++) {
      this._drawColHeader(ctx, vp, col, range);
    }

    ctx.restore();
  }

  _drawColHeader(ctx, vp, col, range) {
    const x = vp.getColX(col);
    const w = this.sheet.getColWidth(col);
    const selected = col >= range.startCol && col <= range.endCol;

    ctx.fillStyle = selected ? COLORS.HEADER_SELECTED_BG : COLORS.HEADER_BG;
    ctx.fillRect(x, 0, w, vp.headerHeight);

    ctx.strokeStyle = COLORS.HEADER_BORDER;
    ctx.lineWidth = 0.5;
    ctx.strokeRect(x + 0.5, 0.5, w, vp.headerHeight);

    ctx.fillStyle = COLORS.HEADER_TEXT;
    ctx.font = `bold 12px ${DEFAULTS.FONT_FAMILY}`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(colIndexToLabel(col), x + w / 2, vp.headerHeight / 2);
  }

  _drawRowHeaders(ctx, vp, visibleRows) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, vp.headerHeight, vp.headerWidth, vp.canvasHeight - vp.headerHeight);
    ctx.clip();

    const sel = this.selection;
    const range = sel.getPrimaryRange();

    // Draw frozen row headers
    for (let row = 0; row < this.sheet.freezeRows; row++) {
      this._drawRowHeader(ctx, vp, row, range);
    }

    // Draw scrollable row headers
    for (let row = visibleRows.start; row < visibleRows.end; row++) {
      this._drawRowHeader(ctx, vp, row, range);
    }

    ctx.restore();
  }

  _drawRowHeader(ctx, vp, row, range) {
    const y = vp.getRowY(row);
    const h = this.sheet.getRowHeight(row);
    const selected = row >= range.startRow && row <= range.endRow;

    ctx.fillStyle = selected ? COLORS.HEADER_SELECTED_BG : COLORS.HEADER_BG;
    ctx.fillRect(0, y, vp.headerWidth, h);

    ctx.strokeStyle = COLORS.HEADER_BORDER;
    ctx.lineWidth = 0.5;
    ctx.strokeRect(0.5, y + 0.5, vp.headerWidth, h);

    ctx.fillStyle = COLORS.HEADER_TEXT;
    ctx.font = `bold 12px ${DEFAULTS.FONT_FAMILY}`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(String(row + 1), vp.headerWidth / 2, y + h / 2);
  }
}
