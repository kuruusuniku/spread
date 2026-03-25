// Toolbar - formatting buttons and actions

export class Toolbar {
  constructor(container, app) {
    this.container = container;
    this.app = app;

    this._setupButtons();
  }

  _setupButtons() {
    this.container.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-action]');
      if (!btn) return;

      const action = btn.dataset.action;
      this._handleAction(action, btn);
    });

    // Color pickers
    const textColorInput = this.container.querySelector('#text-color');
    if (textColorInput) {
      textColorInput.addEventListener('input', (e) => {
        this.app.formatSelection({ textColor: e.target.value });
      });
    }

    const bgColorInput = this.container.querySelector('#bg-color');
    if (bgColorInput) {
      bgColorInput.addEventListener('input', (e) => {
        this.app.formatSelection({ bgColor: e.target.value });
      });
    }
  }

  _handleAction(action, btn) {
    switch (action) {
      case 'bold':
        this.app.toggleFormat('bold');
        break;
      case 'italic':
        this.app.toggleFormat('italic');
        break;
      case 'align-left':
        this.app.formatSelection({ align: 'left' });
        break;
      case 'align-center':
        this.app.formatSelection({ align: 'center' });
        break;
      case 'align-right':
        this.app.formatSelection({ align: 'right' });
        break;
      case 'undo':
        this.app.undo();
        break;
      case 'redo':
        this.app.redo();
        break;
      case 'freeze':
        this.app.toggleFreeze();
        break;
      case 'clear-format':
        this.app.clearFormat();
        break;
    }
  }

  updateState(cell) {
    const fmt = cell?.format || {};
    this._setActive('[data-action="bold"]', fmt.bold);
    this._setActive('[data-action="italic"]', fmt.italic);
    this._setActive('[data-action="align-left"]', fmt.align === 'left');
    this._setActive('[data-action="align-center"]', fmt.align === 'center');
    this._setActive('[data-action="align-right"]', fmt.align === 'right');
  }

  _setActive(selector, active) {
    const el = this.container.querySelector(selector);
    if (el) {
      el.classList.toggle('active', !!active);
    }
  }
}
