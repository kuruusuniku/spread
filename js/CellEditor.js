// Inline cell editor - textarea overlay on the canvas

import { DEFAULTS } from './constants.js';

export class CellEditor {
  constructor(container) {
    this.container = container;
    this.textarea = document.createElement('textarea');
    this.textarea.className = 'cell-editor';
    this.textarea.style.display = 'none';
    container.appendChild(this.textarea);

    this.isEditing = false;
    this.editCol = -1;
    this.editRow = -1;
    this._onCommit = null;
    this._onCancel = null;
    this._onInput = null;

    this.textarea.addEventListener('keydown', (e) => this._handleKeyDown(e));
    this.textarea.addEventListener('input', () => {
      this._autoSize();
      if (this._onInput) this._onInput(this.textarea.value);
    });
  }

  /**
   * Start editing a cell
   */
  startEdit(col, row, x, y, width, height, initialValue, format, onCommit, onCancel, onInput) {
    this.isEditing = true;
    this.editCol = col;
    this.editRow = row;
    this._onCommit = onCommit;
    this._onCancel = onCancel;
    this._onInput = onInput;

    const ta = this.textarea;
    ta.style.display = 'block';
    ta.style.left = x + 'px';
    ta.style.top = y + 'px';
    ta.style.minWidth = width + 'px';
    ta.style.minHeight = height + 'px';
    ta.style.width = width + 'px';
    ta.style.height = height + 'px';
    ta.style.font = `${format.bold ? 'bold ' : ''}${format.italic ? 'italic ' : ''}${format.fontSize || DEFAULTS.FONT_SIZE}px ${DEFAULTS.FONT_FAMILY}`;
    ta.style.textAlign = format.align || 'left';
    ta.style.color = format.textColor || '#000';
    ta.value = initialValue || '';
    ta.focus();

    // Place cursor at end
    ta.setSelectionRange(ta.value.length, ta.value.length);
  }

  /**
   * Start edit with a character (type-to-start)
   */
  startEditWithChar(col, row, x, y, width, height, char, format, onCommit, onCancel, onInput) {
    this.startEdit(col, row, x, y, width, height, '', format, onCommit, onCancel, onInput);
    this.textarea.value = char;
    if (this._onInput) this._onInput(char);
  }

  _handleKeyDown(e) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      this.commit();
    } else if (e.key === 'Tab') {
      e.preventDefault();
      this.commit(e.shiftKey ? 'left' : 'right');
    } else if (e.key === 'Escape') {
      e.preventDefault();
      this.cancel();
    }
  }

  commit(direction = 'down') {
    if (!this.isEditing) return;
    const value = this.textarea.value;
    this.isEditing = false;
    this.textarea.style.display = 'none';
    if (this._onCommit) this._onCommit(value, direction);
  }

  cancel() {
    if (!this.isEditing) return;
    this.isEditing = false;
    this.textarea.style.display = 'none';
    if (this._onCancel) this._onCancel();
  }

  _autoSize() {
    const ta = this.textarea;
    ta.style.height = 'auto';
    ta.style.height = Math.max(parseInt(ta.style.minHeight), ta.scrollHeight) + 'px';
    ta.style.width = 'auto';
    ta.style.width = Math.max(parseInt(ta.style.minWidth), ta.scrollWidth + 10) + 'px';
  }

  get value() {
    return this.textarea.value;
  }
}
