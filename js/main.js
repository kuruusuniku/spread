// Main application - wires everything together

import { Workbook } from './Workbook.js';
import { SelectionModel } from './SelectionModel.js';
import { CommandHistory, SetCellValueCommand, FormatCellsCommand, ResizeCommand } from './CommandHistory.js';
import { FormulaEngine } from './FormulaEngine.js';
import { Viewport } from './Viewport.js';
import { Renderer } from './Renderer.js';
import { CellEditor } from './CellEditor.js';
import { FormulaBar } from './FormulaBar.js';
import { Toolbar } from './Toolbar.js';
import { ContextMenu } from './ContextMenu.js';
import { SheetTabs } from './SheetTabs.js';
import { Clipboard } from './Clipboard.js';
import { EventRouter } from './EventRouter.js';
import { toAddress } from './CellAddress.js';
import { createDefaultFormat, cloneFormat } from './Cell.js';

class SpreadsheetApp {
  constructor() {
    this.workbook = new Workbook();
    this.selection = new SelectionModel();
    this.commandHistory = new CommandHistory();
    this.formulaEngine = new FormulaEngine();
    this.clipboard = new Clipboard();

    const canvas = document.getElementById('spreadsheet-canvas');
    const container = document.getElementById('spreadsheet-container');

    this.viewport = new Viewport(this.sheet);
    this.renderer = new Renderer(canvas, this.viewport, this.sheet, this.selection);
    this.cellEditor = new CellEditor(container);

    this.formulaBar = new FormulaBar(
      document.getElementById('formula-bar'),
      (value) => this._commitFormulaBar(value)
    );

    this.toolbar = new Toolbar(document.getElementById('toolbar'), this);
    this.contextMenu = new ContextMenu(this);
    this.sheetTabs = new SheetTabs(document.getElementById('sheet-tabs'), this);
    this.eventRouter = new EventRouter(canvas, this);

    // Initial state
    this.sheetTabs.update(this.workbook.sheets, this.workbook.activeSheetIndex);
    this.onSelectionChange();

    // Handle resize
    window.addEventListener('resize', () => {
      this.renderer.resize();
    });

    // Populate some demo data
    this._addDemoData();
  }

  get sheet() {
    return this.workbook.activeSheet;
  }

  // --- Selection ---

  onSelectionChange() {
    const cell = this.sheet.getCell(this.selection.activeCell.col, this.selection.activeCell.row);
    this.formulaBar.update(this.selection.activeCell.col, this.selection.activeCell.row, cell);
    this.toolbar.updateState(cell);
    this.renderer.requestRender();
  }

  // --- Editing ---

  startEdit() {
    const col = this.selection.activeCell.col;
    const row = this.selection.activeCell.row;
    const cell = this.sheet.getCell(col, row);
    const x = this.viewport.getColX(col);
    const y = this.viewport.getRowY(row);
    const w = this.sheet.getColWidth(col);
    const h = this.sheet.getRowHeight(row);
    const fmt = cell?.format || createDefaultFormat();
    const rawValue = cell?.rawValue ?? '';

    this.cellEditor.startEdit(col, row, x, y, w, h, String(rawValue), fmt,
      (value, direction) => this._commitEdit(col, row, value, direction),
      () => this.onSelectionChange(),
      (value) => this.formulaBar.input && (this.formulaBar.input.value = value)
    );
  }

  startEditWithChar(char) {
    const col = this.selection.activeCell.col;
    const row = this.selection.activeCell.row;
    const x = this.viewport.getColX(col);
    const y = this.viewport.getRowY(row);
    const w = this.sheet.getColWidth(col);
    const h = this.sheet.getRowHeight(row);
    const cell = this.sheet.getCell(col, row);
    const fmt = cell?.format || createDefaultFormat();

    this.cellEditor.startEditWithChar(col, row, x, y, w, h, char, fmt,
      (value, direction) => this._commitEdit(col, row, value, direction),
      () => this.onSelectionChange(),
      (value) => this.formulaBar.input && (this.formulaBar.input.value = value)
    );
  }

  _commitEdit(col, row, value, direction) {
    const cell = this.sheet.getCell(col, row);
    const oldRawValue = cell?.rawValue ?? null;
    const newRawValue = value === '' ? null : value;

    if (oldRawValue !== newRawValue) {
      const cmd = new SetCellValueCommand(this.sheet, [{
        col, row,
        oldRawValue,
        newRawValue,
        oldFormat: cell ? cloneFormat(cell.format) : createDefaultFormat(),
      }], this.formulaEngine);
      this.commandHistory.execute(cmd);
    }

    // Move to next cell
    if (direction === 'down') {
      this.selection.move(0, 1, this.sheet.maxCols, this.sheet.maxRows);
    } else if (direction === 'right') {
      this.selection.move(1, 0, this.sheet.maxCols, this.sheet.maxRows);
    } else if (direction === 'left') {
      this.selection.move(-1, 0, this.sheet.maxCols, this.sheet.maxRows);
    }

    this.viewport.scrollToCell(this.selection.activeCell.col, this.selection.activeCell.row);
    this.onSelectionChange();
    this.renderer.requestRender();
  }

  _commitFormulaBar(value) {
    const col = this.selection.activeCell.col;
    const row = this.selection.activeCell.row;
    this._commitEdit(col, row, value, null);
  }

  // --- Formatting ---

  toggleFormat(prop) {
    const range = this.selection.getPrimaryRange();
    const changes = [];

    // Check current state of active cell
    const activeCell = this.sheet.getCell(this.selection.activeCell.col, this.selection.activeCell.row);
    const currentValue = activeCell?.format?.[prop] || false;
    const newValue = !currentValue;

    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        const cell = this.sheet.getCell(col, row);
        const oldFormat = cell ? cloneFormat(cell.format) : createDefaultFormat();
        const newFormat = { ...oldFormat, [prop]: newValue };
        changes.push({ col, row, oldFormat, newFormat });
      }
    }

    if (changes.length > 0) {
      this.commandHistory.execute(new FormatCellsCommand(this.sheet, changes));
      this.onSelectionChange();
      this.renderer.requestRender();
    }
  }

  formatSelection(formatProps) {
    const range = this.selection.getPrimaryRange();
    const changes = [];

    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        const cell = this.sheet.getCell(col, row);
        const oldFormat = cell ? cloneFormat(cell.format) : createDefaultFormat();
        const newFormat = { ...oldFormat, ...formatProps };
        changes.push({ col, row, oldFormat, newFormat });
      }
    }

    if (changes.length > 0) {
      this.commandHistory.execute(new FormatCellsCommand(this.sheet, changes));
      this.onSelectionChange();
      this.renderer.requestRender();
    }
  }

  clearFormat() {
    const range = this.selection.getPrimaryRange();
    const changes = [];

    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        const cell = this.sheet.getCell(col, row);
        if (cell) {
          const oldFormat = cloneFormat(cell.format);
          changes.push({ col, row, oldFormat, newFormat: createDefaultFormat() });
        }
      }
    }

    if (changes.length > 0) {
      this.commandHistory.execute(new FormatCellsCommand(this.sheet, changes));
      this.onSelectionChange();
      this.renderer.requestRender();
    }
  }

  // --- Clipboard ---

  copy() {
    const range = this.selection.getPrimaryRange();
    const copyRange = this.clipboard.copy(this.sheet, range);
    this.renderer.copyRange = copyRange;
    this.renderer.startMarchingAnts();
    this.renderer.requestRender();
  }

  cut() {
    const range = this.selection.getPrimaryRange();
    const copyRange = this.clipboard.cut(this.sheet, range);
    this.renderer.copyRange = copyRange;
    this.renderer.startMarchingAnts();
    this.renderer.requestRender();
  }

  async paste() {
    const changes = await this.clipboard.paste(
      this.sheet,
      this.selection.activeCell.col,
      this.selection.activeCell.row,
      this.formulaEngine
    );

    if (changes.length > 0) {
      this.commandHistory.execute(new SetCellValueCommand(this.sheet, changes, this.formulaEngine));
      this.renderer.stopMarchingAnts();
      this.renderer.requestRender();
      this.onSelectionChange();
    }
  }

  // --- Clear ---

  clearSelection() {
    const range = this.selection.getPrimaryRange();
    const changes = [];

    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        const cell = this.sheet.getCell(col, row);
        if (cell && cell.rawValue != null) {
          changes.push({
            col, row,
            oldRawValue: cell.rawValue,
            newRawValue: null,
            oldFormat: cloneFormat(cell.format),
          });
        }
      }
    }

    if (changes.length > 0) {
      this.commandHistory.execute(new SetCellValueCommand(this.sheet, changes, this.formulaEngine));
      this.renderer.requestRender();
      this.onSelectionChange();
    }
  }

  // --- Undo/Redo ---

  undo() {
    if (this.commandHistory.undo()) {
      this.renderer.requestRender();
      this.onSelectionChange();
    }
  }

  redo() {
    if (this.commandHistory.redo()) {
      this.renderer.requestRender();
      this.onSelectionChange();
    }
  }

  // --- Resize ---

  commitResize(type, index, oldSize, newSize) {
    const cmd = new ResizeCommand(this.sheet, type, index, oldSize, newSize);
    this.commandHistory.execute(cmd);
    this.renderer.requestRender();
  }

  autoFitColumn(col) {
    const ctx = this.renderer.ctx;
    let maxWidth = 50;

    for (const [addr, cell] of this.sheet.cells) {
      if (cell.computedValue == null || cell.computedValue === '') continue;
      const parsed = addr.match(/^([A-Z]+)(\d+)$/);
      if (!parsed) continue;
      // Quick check if this cell is in the target column
      const fmt = cell.format || {};
      const fontSize = fmt.fontSize || 13;
      const fontStyle = (fmt.bold ? 'bold ' : '') + (fmt.italic ? 'italic ' : '');
      ctx.font = `${fontStyle}${fontSize}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif`;
      const text = String(cell.computedValue);
      const width = ctx.measureText(text).width + 16;
      if (width > maxWidth) maxWidth = width;
    }

    const oldSize = this.sheet.getColWidth(col);
    this.commitResize('col', col, oldSize, Math.ceil(maxWidth));
  }

  // --- Freeze Panes ---

  toggleFreeze() {
    const sel = this.selection.activeCell;
    if (this.sheet.freezeRows > 0 || this.sheet.freezeCols > 0) {
      this.sheet.freezeRows = 0;
      this.sheet.freezeCols = 0;
    } else {
      this.sheet.freezeRows = sel.row;
      this.sheet.freezeCols = sel.col;
    }
    this.sheet.invalidatePositionCache();
    this.renderer.requestRender();
  }

  // --- Insert/Delete rows/columns ---

  insertRow(below) {
    // Simple implementation: shift data down
    const insertAt = below ? this.selection.activeCell.row + 1 : this.selection.activeCell.row;
    const newCells = new Map();
    for (const [addr, cell] of this.sheet.cells) {
      const match = addr.match(/^([A-Z]+)(\d+)$/);
      if (!match) continue;
      const row = parseInt(match[2]) - 1;
      if (row >= insertAt) {
        newCells.set(match[1] + (row + 2), cell);
      } else {
        newCells.set(addr, cell);
      }
    }
    this.sheet.cells = newCells;
    this.formulaEngine.recalculate(this.sheet);
    this.renderer.requestRender();
  }

  insertCol(right) {
    // Simplified: just re-render
    this.renderer.requestRender();
  }

  deleteRow() {
    const deleteAt = this.selection.activeCell.row;
    const newCells = new Map();
    for (const [addr, cell] of this.sheet.cells) {
      const match = addr.match(/^([A-Z]+)(\d+)$/);
      if (!match) continue;
      const row = parseInt(match[2]) - 1;
      if (row === deleteAt) continue;
      if (row > deleteAt) {
        newCells.set(match[1] + row, cell);
      } else {
        newCells.set(addr, cell);
      }
    }
    this.sheet.cells = newCells;
    this.formulaEngine.recalculate(this.sheet);
    this.renderer.requestRender();
  }

  deleteCol() {
    // Simplified
    this.renderer.requestRender();
  }

  // --- Sheet management ---

  addSheet() {
    this.workbook.addSheet();
    this.workbook.switchSheet(this.workbook.sheets.length - 1);
    this._switchToActiveSheet();
  }

  switchSheet(index) {
    this.workbook.switchSheet(index);
    this._switchToActiveSheet();
  }

  renameSheet(index, name) {
    this.workbook.renameSheet(index, name);
    this.sheetTabs.update(this.workbook.sheets, this.workbook.activeSheetIndex);
  }

  duplicateSheet(index) {
    this.workbook.duplicateSheet(index);
    this.workbook.switchSheet(index + 1);
    this._switchToActiveSheet();
  }

  deleteSheet(index) {
    if (this.workbook.removeSheet(index)) {
      this._switchToActiveSheet();
    }
  }

  _switchToActiveSheet() {
    this.viewport.sheet = this.sheet;
    this.viewport.scrollX = 0;
    this.viewport.scrollY = 0;
    this.renderer.updateSheet(this.sheet);
    this.selection.setActiveCell(0, 0);
    this.sheetTabs.update(this.workbook.sheets, this.workbook.activeSheetIndex);
    this.onSelectionChange();
  }

  // --- Demo Data ---

  _addDemoData() {
    const sheet = this.sheet;

    // Headers
    const headers = ['Product', 'Q1', 'Q2', 'Q3', 'Q4', 'Total', 'Average'];
    headers.forEach((h, i) => {
      const cell = sheet.getOrCreateCell(i, 0);
      cell.rawValue = h;
      cell.computedValue = h;
      cell.format.bold = true;
      cell.format.bgColor = '#4285f4';
      cell.format.textColor = '#ffffff';
      cell.format.align = 'center';
    });

    // Data
    const products = [
      ['Widget A', 1200, 1350, 1100, 1500],
      ['Widget B', 800, 950, 1020, 880],
      ['Gadget X', 2100, 1900, 2300, 2150],
      ['Gadget Y', 450, 520, 490, 610],
      ['Service Z', 3200, 3400, 3100, 3600],
    ];

    products.forEach((row, r) => {
      row.forEach((val, c) => {
        const cell = sheet.getOrCreateCell(c, r + 1);
        cell.rawValue = val;
        cell.computedValue = val;
        if (c === 0) {
          cell.format.bold = true;
        }
      });

      // Total formula
      const totalCell = sheet.getOrCreateCell(5, r + 1);
      totalCell.rawValue = `=SUM(B${r + 2}:E${r + 2})`;

      // Average formula
      const avgCell = sheet.getOrCreateCell(6, r + 1);
      avgCell.rawValue = `=AVERAGE(B${r + 2}:E${r + 2})`;
    });

    // Summary row
    const summaryRow = products.length + 1;
    const sumLabel = sheet.getOrCreateCell(0, summaryRow);
    sumLabel.rawValue = 'Total';
    sumLabel.computedValue = 'Total';
    sumLabel.format.bold = true;
    sumLabel.format.bgColor = '#e8f0fe';

    for (let c = 1; c <= 6; c++) {
      const cell = sheet.getOrCreateCell(c, summaryRow);
      const colLetter = String.fromCharCode(65 + c);
      cell.rawValue = `=SUM(${colLetter}2:${colLetter}${summaryRow})`;
      cell.format.bold = true;
      cell.format.bgColor = '#e8f0fe';
    }

    // Set column widths
    sheet.setColWidth(0, 120);
    sheet.setColWidth(5, 100);
    sheet.setColWidth(6, 100);

    // Process all formulas
    for (const [addr, cell] of sheet.cells) {
      if (cell.rawValue && typeof cell.rawValue === 'string' && cell.rawValue.startsWith('=')) {
        const match = addr.match(/^([A-Z]+)(\d+)$/);
        if (match) {
          const col = match[1].charCodeAt(0) - 65;
          const row = parseInt(match[2]) - 1;
          this.formulaEngine.processCell(sheet, col, row);
        }
      }
    }

    this.formulaEngine.recalculate(sheet);
    this.renderer.requestRender();
  }
}

// Initialize
document.addEventListener('DOMContentLoaded', () => {
  window.app = new SpreadsheetApp();
});
