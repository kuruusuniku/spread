// Right-click context menu

export class ContextMenu {
  constructor(app) {
    this.app = app;
    this.el = document.getElementById('context-menu');
    this._visible = false;

    this.el.addEventListener('click', (e) => {
      const item = e.target.closest('[data-action]');
      if (!item) return;
      this._handleAction(item.dataset.action);
      this.hide();
    });

    document.addEventListener('click', () => this.hide());
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') this.hide();
    });
  }

  show(x, y) {
    this.el.style.left = x + 'px';
    this.el.style.top = y + 'px';
    this.el.style.display = 'block';
    this._visible = true;

    // Keep menu on screen
    const rect = this.el.getBoundingClientRect();
    if (rect.right > window.innerWidth) {
      this.el.style.left = (x - rect.width) + 'px';
    }
    if (rect.bottom > window.innerHeight) {
      this.el.style.top = (y - rect.height) + 'px';
    }
  }

  hide() {
    this.el.style.display = 'none';
    this._visible = false;
  }

  _handleAction(action) {
    switch (action) {
      case 'cut': this.app.cut(); break;
      case 'copy': this.app.copy(); break;
      case 'paste': this.app.paste(); break;
      case 'insert-row-above': this.app.insertRow(false); break;
      case 'insert-row-below': this.app.insertRow(true); break;
      case 'insert-col-left': this.app.insertCol(false); break;
      case 'insert-col-right': this.app.insertCol(true); break;
      case 'delete-row': this.app.deleteRow(); break;
      case 'delete-col': this.app.deleteCol(); break;
      case 'clear': this.app.clearSelection(); break;
      case 'freeze': this.app.toggleFreeze(); break;
    }
  }
}
