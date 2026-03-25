// Workbook - collection of sheets

import { Sheet } from './Sheet.js';

export class Workbook {
  constructor() {
    this.sheets = [new Sheet('Sheet1')];
    this.activeSheetIndex = 0;
  }

  get activeSheet() {
    return this.sheets[this.activeSheetIndex];
  }

  addSheet(name) {
    if (!name) {
      let i = this.sheets.length + 1;
      while (this.sheets.some(s => s.name === `Sheet${i}`)) i++;
      name = `Sheet${i}`;
    }
    const sheet = new Sheet(name);
    this.sheets.push(sheet);
    return sheet;
  }

  removeSheet(index) {
    if (this.sheets.length <= 1) return false;
    this.sheets.splice(index, 1);
    if (this.activeSheetIndex >= this.sheets.length) {
      this.activeSheetIndex = this.sheets.length - 1;
    }
    return true;
  }

  switchSheet(index) {
    if (index >= 0 && index < this.sheets.length) {
      this.activeSheetIndex = index;
      return true;
    }
    return false;
  }

  renameSheet(index, name) {
    if (index >= 0 && index < this.sheets.length) {
      this.sheets[index].name = name;
      return true;
    }
    return false;
  }

  duplicateSheet(index) {
    const src = this.sheets[index];
    if (!src) return null;
    let name = src.name + ' (copy)';
    let i = 2;
    while (this.sheets.some(s => s.name === name)) {
      name = src.name + ` (copy ${i})`;
      i++;
    }
    const sheet = new Sheet(name);
    // Copy cells
    for (const [addr, cell] of src.cells) {
      sheet.cells.set(addr, {
        rawValue: cell.rawValue,
        computedValue: cell.computedValue,
        formula: cell.formula,
        format: { ...cell.format },
        dependsOn: new Set(cell.dependsOn),
        dependedBy: new Set(cell.dependedBy),
      });
    }
    // Copy dimensions
    for (const [col, w] of src.colWidths) sheet.colWidths.set(col, w);
    for (const [row, h] of src.rowHeights) sheet.rowHeights.set(row, h);
    sheet.freezeRows = src.freezeRows;
    sheet.freezeCols = src.freezeCols;
    this.sheets.splice(index + 1, 0, sheet);
    return sheet;
  }
}
