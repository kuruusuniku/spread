# Spread

A lightweight spreadsheet application built with vanilla JavaScript and HTML5 Canvas.

## Features

- **Canvas-based rendering** - High-performance grid rendering using HTML5 Canvas
- **Formula engine** - Supports formulas with functions like `SUM`, `AVERAGE`, and more, with automatic dependency tracking and recalculation
- **Cell formatting** - Bold, italic, text color, background color, text alignment
- **Multiple sheets** - Create, rename, duplicate, and delete sheets with tab navigation
- **Undo/Redo** - Full command history with Ctrl+Z / Ctrl+Y
- **Copy/Cut/Paste** - Clipboard support with marching ants animation
- **Column/Row resizing** - Drag to resize, double-click to auto-fit
- **Freeze panes** - Lock rows and columns for easier navigation
- **Context menu** - Right-click for insert/delete rows and columns
- **Formula bar** - View and edit cell values and formulas

## Getting Started

No build step required. Simply serve the files with any HTTP server:

```bash
# Using Python
python3 -m http.server 8000

# Using Node.js
npx serve .
```

Then open `http://localhost:8000` in your browser.

## Project Structure

```
├── index.html          # Main HTML structure
├── style.css           # Styles for toolbar, formula bar, tabs, and menus
└── js/
    ├── main.js         # Application entry point (SpreadsheetApp)
    ├── constants.js    # Shared constants
    ├── Cell.js         # Cell data model and format utilities
    ├── CellAddress.js  # A1-style address parsing (e.g. "B3" ↔ col/row)
    ├── CellRange.js    # Range parsing and iteration (e.g. "A1:C3")
    ├── Sheet.js        # Sheet data: cells, row/col sizes, freeze state
    ├── Workbook.js     # Multi-sheet management
    ├── SelectionModel.js   # Active cell and selection ranges
    ├── CommandHistory.js   # Undo/Redo command pattern
    ├── FormulaEngine.js    # Tokenizer, parser, evaluator with dependency graph
    ├── Viewport.js     # Scroll and coordinate mapping
    ├── Renderer.js     # Canvas rendering (grid, cells, headers, selections)
    ├── CellEditor.js   # In-place cell editing overlay
    ├── FormulaBar.js   # Formula bar input handling
    ├── Toolbar.js      # Toolbar button actions
    ├── ContextMenu.js  # Right-click context menu
    ├── SheetTabs.js    # Sheet tab bar UI
    ├── Clipboard.js    # Copy/Cut/Paste logic
    └── EventRouter.js  # Keyboard and mouse event dispatch
```

## License

[MIT](LICENSE)
