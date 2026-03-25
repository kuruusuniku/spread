// Formula bar - shows current cell address and content

import { toAddress } from './CellAddress.js';

export class FormulaBar {
  constructor(container, onCommit) {
    this.container = container;
    this.addressLabel = container.querySelector('.formula-bar-address');
    this.input = container.querySelector('.formula-bar-input');
    this._onCommit = onCommit;
    this._isEditing = false;

    this.input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        this._commit();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this._cancel();
      }
    });

    this.input.addEventListener('focus', () => {
      this._isEditing = true;
    });
  }

  update(col, row, cell) {
    this.addressLabel.textContent = toAddress(col, row);
    if (!this._isEditing) {
      this.input.value = cell ? (cell.rawValue ?? '') : '';
    }
  }

  _commit() {
    const value = this.input.value;
    this._isEditing = false;
    this.input.blur();
    if (this._onCommit) this._onCommit(value);
  }

  _cancel() {
    this._isEditing = false;
    this.input.blur();
  }
}
