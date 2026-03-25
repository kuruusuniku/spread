// Undo/Redo command history

export class CommandHistory {
  constructor(maxSize = 200) {
    this.undoStack = [];
    this.redoStack = [];
    this.maxSize = maxSize;
  }

  execute(command) {
    command.execute();
    this.undoStack.push(command);
    this.redoStack.length = 0; // clear redo stack
    if (this.undoStack.length > this.maxSize) {
      this.undoStack.shift();
    }
  }

  undo() {
    const cmd = this.undoStack.pop();
    if (!cmd) return false;
    cmd.undo();
    this.redoStack.push(cmd);
    return true;
  }

  redo() {
    const cmd = this.redoStack.pop();
    if (!cmd) return false;
    cmd.execute();
    this.undoStack.push(cmd);
    return true;
  }

  get canUndo() {
    return this.undoStack.length > 0;
  }

  get canRedo() {
    return this.redoStack.length > 0;
  }

  clear() {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
  }
}

/**
 * Command to set cell values (batch)
 */
export class SetCellValueCommand {
  constructor(sheet, changes, formulaEngine) {
    // changes: [{ address, oldRawValue, newRawValue, oldFormat?, newFormat? }]
    this.sheet = sheet;
    this.changes = changes;
    this.formulaEngine = formulaEngine;
  }

  execute() {
    for (const change of this.changes) {
      const cell = this.sheet.getOrCreateCell(change.col, change.row);
      cell.rawValue = change.newRawValue;
      if (change.newFormat) {
        cell.format = { ...change.newFormat };
      }
      if (this.formulaEngine) {
        this.formulaEngine.processCell(this.sheet, change.col, change.row);
      }
    }
    if (this.formulaEngine) {
      this.formulaEngine.recalculate(this.sheet);
    }
  }

  undo() {
    for (const change of this.changes) {
      const cell = this.sheet.getOrCreateCell(change.col, change.row);
      cell.rawValue = change.oldRawValue;
      if (change.oldFormat) {
        cell.format = { ...change.oldFormat };
      }
      if (this.formulaEngine) {
        this.formulaEngine.processCell(this.sheet, change.col, change.row);
      }
    }
    if (this.formulaEngine) {
      this.formulaEngine.recalculate(this.sheet);
    }
  }
}

/**
 * Command to format cells
 */
export class FormatCellsCommand {
  constructor(sheet, changes) {
    // changes: [{ col, row, oldFormat, newFormat }]
    this.sheet = sheet;
    this.changes = changes;
  }

  execute() {
    for (const change of this.changes) {
      const cell = this.sheet.getOrCreateCell(change.col, change.row);
      cell.format = { ...change.newFormat };
    }
  }

  undo() {
    for (const change of this.changes) {
      const cell = this.sheet.getOrCreateCell(change.col, change.row);
      cell.format = { ...change.oldFormat };
    }
  }
}

/**
 * Command to resize columns/rows
 */
export class ResizeCommand {
  constructor(sheet, type, index, oldSize, newSize) {
    this.sheet = sheet;
    this.type = type; // 'col' or 'row'
    this.index = index;
    this.oldSize = oldSize;
    this.newSize = newSize;
  }

  execute() {
    if (this.type === 'col') {
      this.sheet.setColWidth(this.index, this.newSize);
    } else {
      this.sheet.setRowHeight(this.index, this.newSize);
    }
  }

  undo() {
    if (this.type === 'col') {
      this.sheet.setColWidth(this.index, this.oldSize);
    } else {
      this.sheet.setRowHeight(this.index, this.oldSize);
    }
  }
}
