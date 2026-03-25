// Sheet tabs - bottom tab bar for multiple sheets

export class SheetTabs {
  constructor(container, app) {
    this.container = container;
    this.app = app;

    this.container.addEventListener('click', (e) => {
      const tab = e.target.closest('.sheet-tab');
      const addBtn = e.target.closest('.sheet-add');

      if (addBtn) {
        this.app.addSheet();
        return;
      }

      if (tab) {
        const index = parseInt(tab.dataset.index);
        this.app.switchSheet(index);
      }
    });

    this.container.addEventListener('dblclick', (e) => {
      const tab = e.target.closest('.sheet-tab');
      if (!tab) return;
      const index = parseInt(tab.dataset.index);
      this._startRename(tab, index);
    });

    this.container.addEventListener('contextmenu', (e) => {
      const tab = e.target.closest('.sheet-tab');
      if (!tab) return;
      e.preventDefault();
      const index = parseInt(tab.dataset.index);
      this._showTabMenu(e.clientX, e.clientY, index);
    });
  }

  update(sheets, activeIndex) {
    const tabsHtml = sheets.map((sheet, i) => {
      const active = i === activeIndex ? ' active' : '';
      return `<div class="sheet-tab${active}" data-index="${i}">${this._escapeHtml(sheet.name)}</div>`;
    }).join('');

    this.container.innerHTML = tabsHtml + '<div class="sheet-add" title="Add sheet">+</div>';
  }

  _escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  _startRename(tab, index) {
    const input = document.createElement('input');
    input.type = 'text';
    input.value = this.app.workbook.sheets[index].name;
    input.className = 'sheet-rename-input';
    tab.textContent = '';
    tab.appendChild(input);
    input.focus();
    input.select();

    const finish = () => {
      const newName = input.value.trim();
      if (newName) {
        this.app.renameSheet(index, newName);
      }
      this.update(this.app.workbook.sheets, this.app.workbook.activeSheetIndex);
    };

    input.addEventListener('blur', finish);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); finish(); }
      if (e.key === 'Escape') {
        this.update(this.app.workbook.sheets, this.app.workbook.activeSheetIndex);
      }
    });
  }

  _showTabMenu(x, y, index) {
    // Use a simple prompt-style menu
    const menu = document.createElement('div');
    menu.className = 'tab-context-menu';
    menu.style.left = x + 'px';
    menu.style.top = y + 'px';
    menu.innerHTML = `
      <div class="tab-menu-item" data-action="rename">Rename</div>
      <div class="tab-menu-item" data-action="duplicate">Duplicate</div>
      <div class="tab-menu-item" data-action="delete">Delete</div>
    `;
    document.body.appendChild(menu);

    const cleanup = () => menu.remove();

    menu.addEventListener('click', (e) => {
      const item = e.target.closest('[data-action]');
      if (!item) return;
      switch (item.dataset.action) {
        case 'rename':
          cleanup();
          const tab = this.container.querySelector(`[data-index="${index}"]`);
          if (tab) this._startRename(tab, index);
          break;
        case 'duplicate':
          this.app.duplicateSheet(index);
          cleanup();
          break;
        case 'delete':
          this.app.deleteSheet(index);
          cleanup();
          break;
      }
    });

    setTimeout(() => {
      document.addEventListener('click', cleanup, { once: true });
    }, 0);
  }
}
