// Clipboard - copy, cut, paste operations

import { toAddress, parseAddress, adjustReference } from './CellAddress.js';
import { cloneFormat, createCell, createDefaultFormat } from './Cell.js';

export class Clipboard {
  constructor() {
    this.data = null;    // Internal clipboard data
    this.isCut = false;
    this.sourceRange = null;
  }

  /**
   * Copy selected cells to clipboard
   */
  copy(sheet, range) {
    this.isCut = false;
    this.sourceRange = { ...range };
    this.data = this._extractData(sheet, range);
    this._writeToSystemClipboard(sheet, range);
    return this.sourceRange;
  }

  /**
   * Cut selected cells
   */
  cut(sheet, range) {
    this.isCut = true;
    this.sourceRange = { ...range };
    this.data = this._extractData(sheet, range);
    this._writeToSystemClipboard(sheet, range);
    return this.sourceRange;
  }

  /**
   * Paste at target position
   */
  async paste(sheet, targetCol, targetRow, formulaEngine) {
    const changes = [];

    if (this.data) {
      // Paste from internal clipboard
      const colOffset = targetCol - this.sourceRange.startCol;
      const rowOffset = targetRow - this.sourceRange.startRow;

      for (const item of this.data) {
        const newCol = targetCol + item.relCol;
        const newRow = targetRow + item.relRow;
        const addr = toAddress(newCol, newRow);
        const existingCell = sheet.getCell(newCol, newRow);

        let newRawValue = item.rawValue;
        // Adjust formula references
        if (typeof newRawValue === 'string' && newRawValue.startsWith('=')) {
          newRawValue = this._adjustFormula(newRawValue, colOffset, rowOffset);
        }

        changes.push({
          col: newCol,
          row: newRow,
          oldRawValue: existingCell?.rawValue ?? null,
          newRawValue,
          oldFormat: existingCell ? cloneFormat(existingCell.format) : createDefaultFormat(),
          newFormat: item.format ? cloneFormat(item.format) : createDefaultFormat(),
        });
      }

      // If cut, clear source cells
      if (this.isCut) {
        for (const item of this.data) {
          const srcCol = this.sourceRange.startCol + item.relCol;
          const srcRow = this.sourceRange.startRow + item.relRow;
          const srcCell = sheet.getCell(srcCol, srcRow);
          if (srcCell) {
            changes.push({
              col: srcCol,
              row: srcRow,
              oldRawValue: srcCell.rawValue,
              newRawValue: null,
              oldFormat: cloneFormat(srcCell.format),
              newFormat: createDefaultFormat(),
            });
          }
        }
        this.data = null;
        this.isCut = false;
      }
    } else {
      // Try system clipboard
      try {
        const text = await navigator.clipboard.readText();
        if (text) {
          const rows = text.split('\n');
          for (let r = 0; r < rows.length; r++) {
            const cols = rows[r].split('\t');
            for (let c = 0; c < cols.length; c++) {
              const newCol = targetCol + c;
              const newRow = targetRow + r;
              const existingCell = sheet.getCell(newCol, newRow);
              let value = cols[c];
              // Try to parse as number
              if (value !== '' && !isNaN(Number(value))) {
                value = value; // keep as string, processCell will convert
              }

              changes.push({
                col: newCol,
                row: newRow,
                oldRawValue: existingCell?.rawValue ?? null,
                newRawValue: value || null,
                oldFormat: existingCell ? cloneFormat(existingCell.format) : createDefaultFormat(),
                newFormat: existingCell ? cloneFormat(existingCell.format) : createDefaultFormat(),
              });
            }
          }
        }
      } catch (e) {
        // clipboard access denied
      }
    }

    return changes;
  }

  _extractData(sheet, range) {
    const data = [];
    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        const cell = sheet.getCell(col, row);
        data.push({
          relCol: col - range.startCol,
          relRow: row - range.startRow,
          rawValue: cell?.rawValue ?? null,
          format: cell ? cloneFormat(cell.format) : null,
        });
      }
    }
    return data;
  }

  _writeToSystemClipboard(sheet, range) {
    const rows = [];
    for (let row = range.startRow; row <= range.endRow; row++) {
      const cols = [];
      for (let col = range.startCol; col <= range.endCol; col++) {
        const cell = sheet.getCell(col, row);
        cols.push(cell ? String(cell.computedValue ?? '') : '');
      }
      rows.push(cols.join('\t'));
    }
    try {
      navigator.clipboard.writeText(rows.join('\n'));
    } catch (e) {
      // Fallback: do nothing
    }
  }

  _adjustFormula(formula, colOffset, rowOffset) {
    // Adjust cell references in the formula
    return '=' + formula.substring(1).replace(
      /(\$?)([A-Z]+)(\$?)(\d+)/gi,
      (match, absc, col, absr, row) => {
        return adjustReference(match, colOffset, rowOffset);
      }
    );
  }
}
