// Event router - translates DOM events to spreadsheet actions

export class EventRouter {
  constructor(canvas, app) {
    this.canvas = canvas;
    this.app = app;
    this._resizing = null; // { type: 'col'|'row', index, startPos, startSize }
    this._fillDrag = null;

    this._bindEvents();
  }

  _bindEvents() {
    this.canvas.addEventListener('mousedown', (e) => this._onMouseDown(e));
    this.canvas.addEventListener('mousemove', (e) => this._onMouseMove(e));
    this.canvas.addEventListener('mouseup', (e) => this._onMouseUp(e));
    this.canvas.addEventListener('dblclick', (e) => this._onDoubleClick(e));
    this.canvas.addEventListener('contextmenu', (e) => this._onContextMenu(e));
    this.canvas.addEventListener('wheel', (e) => this._onWheel(e), { passive: false });

    document.addEventListener('keydown', (e) => this._onKeyDown(e));

    // Global mouse up for drag operations
    document.addEventListener('mousemove', (e) => {
      if (this._resizing) this._onResizeMove(e);
      if (this.app.selection.isSelecting) this._onSelectionMove(e);
    });
    document.addEventListener('mouseup', (e) => {
      if (this._resizing) this._onResizeEnd(e);
      if (this.app.selection.isSelecting) this.app.selection.endSelection();
    });
  }

  _getCanvasPos(e) {
    const rect = this.canvas.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  }

  _onMouseDown(e) {
    if (e.button !== 0) return; // Only left click

    const pos = this._getCanvasPos(e);
    const hit = this.app.viewport.hitTest(pos.x, pos.y);

    if (this.app.cellEditor.isEditing) {
      this.app.cellEditor.commit();
    }

    // Column resize
    if (hit.area === 'colHeader' && hit.resize) {
      this._resizing = {
        type: 'col',
        index: hit.col,
        startPos: e.clientX,
        startSize: this.app.sheet.getColWidth(hit.col),
      };
      e.preventDefault();
      return;
    }

    // Row resize
    if (hit.area === 'rowHeader' && hit.resize) {
      this._resizing = {
        type: 'row',
        index: hit.row,
        startPos: e.clientY,
        startSize: this.app.sheet.getRowHeight(hit.row),
      };
      e.preventDefault();
      return;
    }

    // Column header select (select entire column)
    if (hit.area === 'colHeader' && !hit.resize) {
      this.app.selection.setActiveCell(hit.col, 0);
      this.app.selection.ranges = [{
        startCol: hit.col, startRow: 0,
        endCol: hit.col, endRow: this.app.sheet.maxRows - 1
      }];
      this.app.onSelectionChange();
      return;
    }

    // Row header select (select entire row)
    if (hit.area === 'rowHeader' && !hit.resize) {
      this.app.selection.setActiveCell(0, hit.row);
      this.app.selection.ranges = [{
        startCol: 0, startRow: hit.row,
        endCol: this.app.sheet.maxCols - 1, endRow: hit.row
      }];
      this.app.onSelectionChange();
      return;
    }

    // Cell area click
    if (hit.area === 'cell') {
      // Check for fill handle
      const sel = this.app.selection;
      const acX = this.app.viewport.getColX(sel.activeCell.col);
      const acY = this.app.viewport.getRowY(sel.activeCell.row);
      const acW = this.app.sheet.getColWidth(sel.activeCell.col);
      const acH = this.app.sheet.getRowHeight(sel.activeCell.row);

      if (Math.abs(pos.x - (acX + acW)) < 6 && Math.abs(pos.y - (acY + acH)) < 6) {
        // Fill handle drag
        this._fillDrag = { startCol: sel.activeCell.col, startRow: sel.activeCell.row };
        return;
      }

      if (e.shiftKey) {
        this.app.selection.startSelection(hit.col, hit.row, true);
      } else {
        this.app.selection.startSelection(hit.col, hit.row);
      }
      this.app.onSelectionChange();
    }
  }

  _onSelectionMove(e) {
    const pos = this._getCanvasPos(e);
    const hit = this.app.viewport.hitTest(pos.x, pos.y);
    if (hit.area === 'cell' || hit.area === 'colHeader' || hit.area === 'rowHeader') {
      const col = Math.max(0, hit.col);
      const row = Math.max(0, hit.row);
      this.app.selection.updateSelection(col, row);
      this.app.renderer.requestRender();
    }
  }

  _onMouseUp(e) {
    if (this._fillDrag) {
      // TODO: implement fill logic
      this._fillDrag = null;
    }
  }

  _onMouseMove(e) {
    const pos = this._getCanvasPos(e);
    const hit = this.app.viewport.hitTest(pos.x, pos.y);

    // Update cursor
    if (hit.area === 'colHeader' && hit.resize) {
      this.canvas.style.cursor = 'col-resize';
    } else if (hit.area === 'rowHeader' && hit.resize) {
      this.canvas.style.cursor = 'row-resize';
    } else if (hit.area === 'cell') {
      // Check fill handle
      const sel = this.app.selection;
      const acX = this.app.viewport.getColX(sel.activeCell.col);
      const acY = this.app.viewport.getRowY(sel.activeCell.row);
      const acW = this.app.sheet.getColWidth(sel.activeCell.col);
      const acH = this.app.sheet.getRowHeight(sel.activeCell.row);
      if (Math.abs(pos.x - (acX + acW)) < 6 && Math.abs(pos.y - (acY + acH)) < 6) {
        this.canvas.style.cursor = 'crosshair';
      } else {
        this.canvas.style.cursor = 'cell';
      }
    } else {
      this.canvas.style.cursor = 'default';
    }
  }

  _onResizeMove(e) {
    if (!this._resizing) return;
    const delta = this._resizing.type === 'col'
      ? e.clientX - this._resizing.startPos
      : e.clientY - this._resizing.startPos;
    const newSize = Math.max(20, this._resizing.startSize + delta);
    if (this._resizing.type === 'col') {
      this.app.sheet.setColWidth(this._resizing.index, newSize);
    } else {
      this.app.sheet.setRowHeight(this._resizing.index, newSize);
    }
    this.app.renderer.requestRender();
  }

  _onResizeEnd(e) {
    if (!this._resizing) return;
    const delta = this._resizing.type === 'col'
      ? e.clientX - this._resizing.startPos
      : e.clientY - this._resizing.startPos;
    const newSize = Math.max(20, this._resizing.startSize + delta);
    this.app.commitResize(this._resizing.type, this._resizing.index, this._resizing.startSize, newSize);
    this._resizing = null;
  }

  _onDoubleClick(e) {
    const pos = this._getCanvasPos(e);
    const hit = this.app.viewport.hitTest(pos.x, pos.y);

    // Double-click column header border: auto-fit
    if (hit.area === 'colHeader' && hit.resize) {
      this.app.autoFitColumn(hit.col);
      return;
    }

    if (hit.area === 'cell') {
      this.app.startEdit();
    }
  }

  _onContextMenu(e) {
    e.preventDefault();
    this.app.contextMenu.show(e.clientX, e.clientY);
  }

  _onWheel(e) {
    e.preventDefault();
    const vp = this.app.viewport;

    if (e.shiftKey) {
      vp.scrollX += e.deltaY;
    } else {
      vp.scrollX += e.deltaX;
      vp.scrollY += e.deltaY;
    }
    vp.clampScroll();
    this.app.renderer.requestRender();
  }

  _onKeyDown(e) {
    // Don't handle if cell editor is active
    if (this.app.cellEditor.isEditing) return;
    // Don't handle if focus is in formula bar input
    if (document.activeElement?.classList.contains('formula-bar-input')) return;

    const sel = this.app.selection;
    const sheet = this.app.sheet;

    // Navigation
    if (e.key === 'ArrowUp' || e.key === 'ArrowDown' ||
        e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
      e.preventDefault();
      const dCol = e.key === 'ArrowLeft' ? -1 : e.key === 'ArrowRight' ? 1 : 0;
      const dRow = e.key === 'ArrowUp' ? -1 : e.key === 'ArrowDown' ? 1 : 0;

      if (e.ctrlKey || e.metaKey) {
        // Jump to edge of data
        this._jumpToEdge(dCol, dRow, e.shiftKey);
      } else {
        const pos = sel.move(dCol, dRow, sheet.maxCols, sheet.maxRows, e.shiftKey);
        this.app.viewport.scrollToCell(pos.col, pos.row);
      }
      this.app.onSelectionChange();
      return;
    }

    // Tab
    if (e.key === 'Tab') {
      e.preventDefault();
      const dCol = e.shiftKey ? -1 : 1;
      sel.move(dCol, 0, sheet.maxCols, sheet.maxRows);
      this.app.viewport.scrollToCell(sel.activeCell.col, sel.activeCell.row);
      this.app.onSelectionChange();
      return;
    }

    // Enter
    if (e.key === 'Enter') {
      e.preventDefault();
      const dRow = e.shiftKey ? -1 : 1;
      sel.move(0, dRow, sheet.maxCols, sheet.maxRows);
      this.app.viewport.scrollToCell(sel.activeCell.col, sel.activeCell.row);
      this.app.onSelectionChange();
      return;
    }

    // Delete/Backspace
    if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      this.app.clearSelection();
      return;
    }

    // Ctrl/Cmd shortcuts
    if (e.ctrlKey || e.metaKey) {
      switch (e.key.toLowerCase()) {
        case 'c': e.preventDefault(); this.app.copy(); return;
        case 'x': e.preventDefault(); this.app.cut(); return;
        case 'v': e.preventDefault(); this.app.paste(); return;
        case 'z':
          e.preventDefault();
          if (e.shiftKey) this.app.redo();
          else this.app.undo();
          return;
        case 'y': e.preventDefault(); this.app.redo(); return;
        case 'b': e.preventDefault(); this.app.toggleFormat('bold'); return;
        case 'i': e.preventDefault(); this.app.toggleFormat('italic'); return;
        case 'a':
          e.preventDefault();
          sel.ranges = [{ startCol: 0, startRow: 0, endCol: sheet.maxCols - 1, endRow: sheet.maxRows - 1 }];
          this.app.onSelectionChange();
          return;
      }
    }

    // F2 - edit mode
    if (e.key === 'F2') {
      e.preventDefault();
      this.app.startEdit();
      return;
    }

    // Escape
    if (e.key === 'Escape') {
      this.app.renderer.stopMarchingAnts();
      this.app.clipboard.data = null;
      this.app.renderer.requestRender();
      return;
    }

    // Home/End
    if (e.key === 'Home') {
      e.preventDefault();
      if (e.ctrlKey || e.metaKey) {
        sel.setActiveCell(0, 0);
      } else {
        sel.setActiveCell(0, sel.activeCell.row);
      }
      this.app.viewport.scrollToCell(sel.activeCell.col, sel.activeCell.row);
      this.app.onSelectionChange();
      return;
    }

    // Page Up/Down
    if (e.key === 'PageDown') {
      e.preventDefault();
      const visibleRows = Math.floor((this.app.viewport.canvasHeight - this.app.viewport.headerHeight) / 25);
      sel.move(0, visibleRows, sheet.maxCols, sheet.maxRows);
      this.app.viewport.scrollToCell(sel.activeCell.col, sel.activeCell.row);
      this.app.onSelectionChange();
      return;
    }
    if (e.key === 'PageUp') {
      e.preventDefault();
      const visibleRows = Math.floor((this.app.viewport.canvasHeight - this.app.viewport.headerHeight) / 25);
      sel.move(0, -visibleRows, sheet.maxCols, sheet.maxRows);
      this.app.viewport.scrollToCell(sel.activeCell.col, sel.activeCell.row);
      this.app.onSelectionChange();
      return;
    }

    // Type to start editing
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      this.app.startEditWithChar(e.key);
      return;
    }
  }

  _jumpToEdge(dCol, dRow, extend) {
    const sel = this.app.selection;
    const sheet = this.app.sheet;
    let col = extend ? (sel.ranges[0].endCol) : sel.activeCell.col;
    let row = extend ? (sel.ranges[0].endRow) : sel.activeCell.row;

    if (dCol !== 0) {
      const currentCell = sheet.getCell(col, row);
      const hasData = currentCell && currentCell.rawValue != null && currentCell.rawValue !== '';

      if (dCol > 0) {
        if (hasData) {
          // Jump to last non-empty cell in this direction
          while (col < sheet.maxCols - 1) {
            const next = sheet.getCell(col + 1, row);
            if (!next || next.rawValue == null || next.rawValue === '') break;
            col++;
          }
          if (col === sel.activeCell.col) col = sheet.maxCols - 1;
        } else {
          // Jump to next non-empty cell
          while (col < sheet.maxCols - 1) {
            col++;
            const next = sheet.getCell(col, row);
            if (next && next.rawValue != null && next.rawValue !== '') break;
          }
        }
      } else {
        if (hasData) {
          while (col > 0) {
            const next = sheet.getCell(col - 1, row);
            if (!next || next.rawValue == null || next.rawValue === '') break;
            col--;
          }
          if (col === sel.activeCell.col) col = 0;
        } else {
          while (col > 0) {
            col--;
            const next = sheet.getCell(col, row);
            if (next && next.rawValue != null && next.rawValue !== '') break;
          }
        }
      }
    }

    if (dRow !== 0) {
      const currentCell = sheet.getCell(col, row);
      const hasData = currentCell && currentCell.rawValue != null && currentCell.rawValue !== '';

      if (dRow > 0) {
        if (hasData) {
          while (row < sheet.maxRows - 1) {
            const next = sheet.getCell(col, row + 1);
            if (!next || next.rawValue == null || next.rawValue === '') break;
            row++;
          }
          if (row === sel.activeCell.row) row = sheet.maxRows - 1;
        } else {
          while (row < sheet.maxRows - 1) {
            row++;
            const next = sheet.getCell(col, row);
            if (next && next.rawValue != null && next.rawValue !== '') break;
          }
        }
      } else {
        if (hasData) {
          while (row > 0) {
            const next = sheet.getCell(col, row - 1);
            if (!next || next.rawValue == null || next.rawValue === '') break;
            row--;
          }
          if (row === sel.activeCell.row) row = 0;
        } else {
          while (row > 0) {
            row--;
            const next = sheet.getCell(col, row);
            if (next && next.rawValue != null && next.rawValue !== '') break;
          }
        }
      }
    }

    if (extend) {
      sel.ranges[0].endCol = col;
      sel.ranges[0].endRow = row;
    } else {
      sel.setActiveCell(col, row);
    }
    this.app.viewport.scrollToCell(col, row);
  }
}
