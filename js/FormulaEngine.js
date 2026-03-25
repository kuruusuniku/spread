// Formula engine - tokenizer, parser, evaluator with dependency tracking

import { parseAddress, toAddress, labelToColIndex } from './CellAddress.js';
import { parseRange, iterateRange } from './CellRange.js';

// Token types
const T = {
  NUMBER: 'NUMBER',
  STRING: 'STRING',
  CELL_REF: 'CELL_REF',
  RANGE_REF: 'RANGE_REF',
  FUNCTION: 'FUNCTION',
  OPERATOR: 'OPERATOR',
  LPAREN: 'LPAREN',
  RPAREN: 'RPAREN',
  COMMA: 'COMMA',
  BOOLEAN: 'BOOLEAN',
};

export class FormulaEngine {
  constructor() {
    this._recalcOrder = null;
    this._functions = this._buildFunctions();
  }

  _buildFunctions() {
    return {
      SUM: (args) => {
        let total = 0;
        for (const v of this._flattenArgs(args)) {
          const n = Number(v);
          if (!isNaN(n)) total += n;
        }
        return total;
      },
      AVERAGE: (args) => {
        let total = 0, count = 0;
        for (const v of this._flattenArgs(args)) {
          const n = Number(v);
          if (!isNaN(n)) { total += n; count++; }
        }
        return count > 0 ? total / count : '#DIV/0!';
      },
      MAX: (args) => {
        let max = -Infinity;
        for (const v of this._flattenArgs(args)) {
          const n = Number(v);
          if (!isNaN(n) && n > max) max = n;
        }
        return max === -Infinity ? 0 : max;
      },
      MIN: (args) => {
        let min = Infinity;
        for (const v of this._flattenArgs(args)) {
          const n = Number(v);
          if (!isNaN(n) && n < min) min = n;
        }
        return min === Infinity ? 0 : min;
      },
      COUNT: (args) => {
        let count = 0;
        for (const v of this._flattenArgs(args)) {
          if (v !== null && v !== '' && !isNaN(Number(v))) count++;
        }
        return count;
      },
      COUNTA: (args) => {
        let count = 0;
        for (const v of this._flattenArgs(args)) {
          if (v !== null && v !== '') count++;
        }
        return count;
      },
      IF: (args) => {
        if (args.length < 2) return '#VALUE!';
        const condition = args[0];
        return condition ? (args[1] !== undefined ? args[1] : true) : (args[2] !== undefined ? args[2] : false);
      },
      CONCATENATE: (args) => {
        return args.map(a => Array.isArray(a) ? a.flat().join('') : String(a ?? '')).join('');
      },
      LEN: (args) => {
        return String(args[0] ?? '').length;
      },
      UPPER: (args) => String(args[0] ?? '').toUpperCase(),
      LOWER: (args) => String(args[0] ?? '').toLowerCase(),
      TRIM: (args) => String(args[0] ?? '').trim(),
      ABS: (args) => Math.abs(Number(args[0]) || 0),
      ROUND: (args) => {
        const val = Number(args[0]) || 0;
        const digits = Number(args[1]) || 0;
        const factor = Math.pow(10, digits);
        return Math.round(val * factor) / factor;
      },
      INT: (args) => Math.floor(Number(args[0]) || 0),
      MOD: (args) => {
        const n = Number(args[0]) || 0;
        const d = Number(args[1]) || 1;
        return d === 0 ? '#DIV/0!' : n % d;
      },
      POWER: (args) => Math.pow(Number(args[0]) || 0, Number(args[1]) || 0),
      SQRT: (args) => {
        const n = Number(args[0]) || 0;
        return n < 0 ? '#NUM!' : Math.sqrt(n);
      },
      NOW: () => new Date().toLocaleString(),
      TODAY: () => new Date().toLocaleDateString(),
      PI: () => Math.PI,
      VLOOKUP: (args) => {
        if (args.length < 3) return '#VALUE!';
        const lookupVal = args[0];
        const rangeData = args[1]; // Should be a 2D array from range ref
        const colIndex = Number(args[2]) - 1;
        const exactMatch = args.length > 3 ? !args[3] : false;
        if (!Array.isArray(rangeData) || !Array.isArray(rangeData[0])) return '#VALUE!';
        if (colIndex < 0 || colIndex >= (rangeData[0] ? rangeData[0].length : 0)) return '#REF!';
        for (const row of rangeData) {
          if (exactMatch) {
            if (row[0] == lookupVal) return row[colIndex];
          } else {
            if (row[0] <= lookupVal) {
              // Approximate - find largest value <= lookup
              continue;
            }
          }
        }
        if (!exactMatch && rangeData.length > 0) {
          return rangeData[rangeData.length - 1][colIndex];
        }
        return '#N/A';
      },
      SUMIF: (args) => {
        if (args.length < 2) return '#VALUE!';
        const rangeVals = this._flattenArgs([args[0]]);
        const criteria = args[1];
        const sumRange = args.length > 2 ? this._flattenArgs([args[2]]) : rangeVals;
        let total = 0;
        for (let i = 0; i < rangeVals.length; i++) {
          if (this._matchCriteria(rangeVals[i], criteria)) {
            const v = Number(sumRange[i]);
            if (!isNaN(v)) total += v;
          }
        }
        return total;
      },
      COUNTIF: (args) => {
        if (args.length < 2) return '#VALUE!';
        const rangeVals = this._flattenArgs([args[0]]);
        const criteria = args[1];
        let count = 0;
        for (const v of rangeVals) {
          if (this._matchCriteria(v, criteria)) count++;
        }
        return count;
      },
    };
  }

  _matchCriteria(value, criteria) {
    const cs = String(criteria);
    if (cs.startsWith('>=')) return Number(value) >= Number(cs.slice(2));
    if (cs.startsWith('<=')) return Number(value) <= Number(cs.slice(2));
    if (cs.startsWith('<>')) return String(value) !== cs.slice(2);
    if (cs.startsWith('>')) return Number(value) > Number(cs.slice(1));
    if (cs.startsWith('<')) return Number(value) < Number(cs.slice(1));
    if (cs.startsWith('=')) return String(value) === cs.slice(1);
    return String(value) === cs;
  }

  *_flattenArgs(args) {
    for (const a of args) {
      if (Array.isArray(a)) {
        for (const row of a) {
          if (Array.isArray(row)) {
            yield* row;
          } else {
            yield row;
          }
        }
      } else {
        yield a;
      }
    }
  }

  /**
   * Tokenize a formula string (without the leading '=')
   */
  tokenize(formula) {
    const tokens = [];
    let i = 0;
    while (i < formula.length) {
      const ch = formula[i];

      // Skip whitespace
      if (ch === ' ' || ch === '\t') { i++; continue; }

      // String literal
      if (ch === '"') {
        let str = '';
        i++;
        while (i < formula.length && formula[i] !== '"') {
          str += formula[i++];
        }
        i++; // skip closing quote
        tokens.push({ type: T.STRING, value: str });
        continue;
      }

      // Number
      if ((ch >= '0' && ch <= '9') || (ch === '.' && i + 1 < formula.length && formula[i + 1] >= '0' && formula[i + 1] <= '9')) {
        let num = '';
        while (i < formula.length && ((formula[i] >= '0' && formula[i] <= '9') || formula[i] === '.')) {
          num += formula[i++];
        }
        // Check for percentage
        if (i < formula.length && formula[i] === '%') {
          i++;
          tokens.push({ type: T.NUMBER, value: parseFloat(num) / 100 });
        } else {
          tokens.push({ type: T.NUMBER, value: parseFloat(num) });
        }
        continue;
      }

      // Cell reference, range reference, function name, or boolean
      if ((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || ch === '$') {
        let ident = '';
        while (i < formula.length && ((formula[i] >= 'A' && formula[i] <= 'Z') ||
               (formula[i] >= 'a' && formula[i] <= 'z') ||
               (formula[i] >= '0' && formula[i] <= '9') || formula[i] === '$')) {
          ident += formula[i++];
        }

        // Check for boolean
        if (ident.toUpperCase() === 'TRUE') {
          tokens.push({ type: T.BOOLEAN, value: true });
          continue;
        }
        if (ident.toUpperCase() === 'FALSE') {
          tokens.push({ type: T.BOOLEAN, value: false });
          continue;
        }

        // Check if it's a function (followed by open paren)
        if (i < formula.length && formula[i] === '(') {
          tokens.push({ type: T.FUNCTION, value: ident.toUpperCase() });
          continue;
        }

        // Check for range reference (A1:B5)
        if (i < formula.length && formula[i] === ':') {
          let range = ident + ':';
          i++; // skip ':'
          while (i < formula.length && ((formula[i] >= 'A' && formula[i] <= 'Z') ||
                 (formula[i] >= 'a' && formula[i] <= 'z') ||
                 (formula[i] >= '0' && formula[i] <= '9') || formula[i] === '$')) {
            range += formula[i++];
          }
          tokens.push({ type: T.RANGE_REF, value: range.toUpperCase() });
          continue;
        }

        // Check if it's a cell reference
        if (parseAddress(ident.replace(/\$/g, ''))) {
          tokens.push({ type: T.CELL_REF, value: ident.toUpperCase() });
        } else {
          // Unknown identifier - treat as function name that may not have parens
          tokens.push({ type: T.CELL_REF, value: ident.toUpperCase() });
        }
        continue;
      }

      // Operators
      if (ch === '+' || ch === '-' || ch === '*' || ch === '/' || ch === '^') {
        tokens.push({ type: T.OPERATOR, value: ch });
        i++;
        continue;
      }

      // Comparison operators
      if (ch === '>' || ch === '<' || ch === '=') {
        let op = ch;
        i++;
        if (i < formula.length && (formula[i] === '=' || (ch === '<' && formula[i] === '>'))) {
          op += formula[i++];
        }
        tokens.push({ type: T.OPERATOR, value: op });
        continue;
      }

      if (ch === '&') {
        tokens.push({ type: T.OPERATOR, value: '&' });
        i++;
        continue;
      }

      if (ch === '(') { tokens.push({ type: T.LPAREN }); i++; continue; }
      if (ch === ')') { tokens.push({ type: T.RPAREN }); i++; continue; }
      if (ch === ',') { tokens.push({ type: T.COMMA }); i++; continue; }

      // Unknown character, skip
      i++;
    }
    return tokens;
  }

  /**
   * Parse tokens into an AST
   */
  parse(tokens) {
    let pos = 0;

    const peek = () => tokens[pos] || null;
    const consume = () => tokens[pos++];

    const parseExpression = () => parseComparison();

    const parseComparison = () => {
      let left = parseConcatenation();
      while (peek() && peek().type === T.OPERATOR &&
             ['=', '<>', '<', '>', '<=', '>='].includes(peek().value)) {
        const op = consume().value;
        const right = parseConcatenation();
        left = { type: 'binary', op, left, right };
      }
      return left;
    };

    const parseConcatenation = () => {
      let left = parseAddSub();
      while (peek() && peek().type === T.OPERATOR && peek().value === '&') {
        consume();
        const right = parseAddSub();
        left = { type: 'binary', op: '&', left, right };
      }
      return left;
    };

    const parseAddSub = () => {
      let left = parseMulDiv();
      while (peek() && peek().type === T.OPERATOR && (peek().value === '+' || peek().value === '-')) {
        const op = consume().value;
        const right = parseMulDiv();
        left = { type: 'binary', op, left, right };
      }
      return left;
    };

    const parseMulDiv = () => {
      let left = parsePower();
      while (peek() && peek().type === T.OPERATOR && (peek().value === '*' || peek().value === '/')) {
        const op = consume().value;
        const right = parsePower();
        left = { type: 'binary', op, left, right };
      }
      return left;
    };

    const parsePower = () => {
      let left = parseUnary();
      while (peek() && peek().type === T.OPERATOR && peek().value === '^') {
        consume();
        const right = parseUnary();
        left = { type: 'binary', op: '^', left, right };
      }
      return left;
    };

    const parseUnary = () => {
      if (peek() && peek().type === T.OPERATOR && peek().value === '-') {
        consume();
        const operand = parseUnary();
        return { type: 'unary', op: '-', operand };
      }
      if (peek() && peek().type === T.OPERATOR && peek().value === '+') {
        consume();
        return parseUnary();
      }
      return parsePrimary();
    };

    const parsePrimary = () => {
      const tok = peek();
      if (!tok) return { type: 'literal', value: 0 };

      if (tok.type === T.NUMBER) {
        consume();
        return { type: 'literal', value: tok.value };
      }

      if (tok.type === T.STRING) {
        consume();
        return { type: 'literal', value: tok.value };
      }

      if (tok.type === T.BOOLEAN) {
        consume();
        return { type: 'literal', value: tok.value };
      }

      if (tok.type === T.FUNCTION) {
        const name = consume().value;
        consume(); // LPAREN
        const args = [];
        while (peek() && peek().type !== T.RPAREN) {
          args.push(parseExpression());
          if (peek() && peek().type === T.COMMA) consume();
        }
        if (peek() && peek().type === T.RPAREN) consume();
        return { type: 'function', name, args };
      }

      if (tok.type === T.RANGE_REF) {
        consume();
        return { type: 'range', value: tok.value };
      }

      if (tok.type === T.CELL_REF) {
        consume();
        return { type: 'cell', value: tok.value };
      }

      if (tok.type === T.LPAREN) {
        consume();
        const expr = parseExpression();
        if (peek() && peek().type === T.RPAREN) consume();
        return expr;
      }

      consume(); // skip unknown
      return { type: 'literal', value: '#ERROR!' };
    };

    const result = parseExpression();
    return result;
  }

  /**
   * Evaluate an AST node against a sheet
   */
  evaluate(node, sheet, visitedCells = new Set()) {
    if (!node) return 0;

    switch (node.type) {
      case 'literal':
        return node.value;

      case 'cell': {
        const addr = node.value.replace(/\$/g, '');
        if (visitedCells.has(addr)) return '#CIRCULAR!';
        const cell = sheet.getCellByAddress(addr);
        if (!cell) return 0;
        if (cell.formula) {
          // Already computed
          return cell.computedValue;
        }
        return cell.computedValue ?? cell.rawValue ?? 0;
      }

      case 'range': {
        const range = parseRange(node.value.replace(/\$/g, ''));
        if (!range) return '#REF!';
        const result = [];
        for (let r = range.startRow; r <= range.endRow; r++) {
          const rowData = [];
          for (let c = range.startCol; c <= range.endCol; c++) {
            const cell = sheet.getCell(c, r);
            rowData.push(cell ? (cell.computedValue ?? cell.rawValue ?? 0) : 0);
          }
          result.push(rowData);
        }
        return result;
      }

      case 'unary':
        if (node.op === '-') {
          const val = this.evaluate(node.operand, sheet, visitedCells);
          return typeof val === 'number' ? -val : '#VALUE!';
        }
        return this.evaluate(node.operand, sheet, visitedCells);

      case 'binary': {
        const left = this.evaluate(node.left, sheet, visitedCells);
        const right = this.evaluate(node.right, sheet, visitedCells);

        // Error propagation
        if (typeof left === 'string' && left.startsWith('#')) return left;
        if (typeof right === 'string' && right.startsWith('#')) return right;

        switch (node.op) {
          case '+': return Number(left) + Number(right);
          case '-': return Number(left) - Number(right);
          case '*': return Number(left) * Number(right);
          case '/': return Number(right) === 0 ? '#DIV/0!' : Number(left) / Number(right);
          case '^': return Math.pow(Number(left), Number(right));
          case '&': return String(left ?? '') + String(right ?? '');
          case '=': return left == right;
          case '<>': return left != right;
          case '<': return Number(left) < Number(right);
          case '>': return Number(left) > Number(right);
          case '<=': return Number(left) <= Number(right);
          case '>=': return Number(left) >= Number(right);
          default: return '#ERROR!';
        }
      }

      case 'function': {
        const fn = this._functions[node.name];
        if (!fn) return '#NAME?';
        const args = node.args.map(a => this.evaluate(a, sheet, visitedCells));
        try {
          return fn(args);
        } catch (e) {
          return '#ERROR!';
        }
      }

      default:
        return '#ERROR!';
    }
  }

  /**
   * Process a cell - detect if it has a formula and extract dependencies
   */
  processCell(sheet, col, row) {
    const cell = sheet.getCell(col, row);
    if (!cell) return;

    const addr = toAddress(col, row);

    // Remove old dependencies
    for (const dep of cell.dependsOn) {
      const depCell = sheet.getCellByAddress(dep);
      if (depCell) depCell.dependedBy.delete(addr);
    }
    cell.dependsOn.clear();
    cell.formula = null;

    const raw = cell.rawValue;
    if (typeof raw === 'string' && raw.startsWith('=')) {
      cell.formula = raw.substring(1);
      // Extract dependencies
      const deps = this._extractDependencies(cell.formula);
      for (const dep of deps) {
        cell.dependsOn.add(dep);
        const depCell = sheet.getOrCreateCell(
          ...(() => { const p = parseAddress(dep); return [p.col, p.row]; })()
        );
        depCell.dependedBy.add(addr);
      }
      // Evaluate
      const tokens = this.tokenize(cell.formula);
      const ast = this.parse(tokens);
      const visited = new Set([addr]);
      cell.computedValue = this.evaluate(ast, sheet, visited);
    } else {
      // Parse as number if possible
      if (raw !== null && raw !== '' && !isNaN(Number(raw))) {
        cell.computedValue = Number(raw);
      } else {
        cell.computedValue = raw;
      }
    }
  }

  /**
   * Extract all cell addresses referenced in a formula
   */
  _extractDependencies(formula) {
    const deps = new Set();
    const tokens = this.tokenize(formula);
    for (const tok of tokens) {
      if (tok.type === T.CELL_REF) {
        const addr = tok.value.replace(/\$/g, '');
        deps.add(addr);
      } else if (tok.type === T.RANGE_REF) {
        const range = parseRange(tok.value.replace(/\$/g, ''));
        if (range) {
          for (const cell of iterateRange(range)) {
            deps.add(cell.address);
          }
        }
      }
    }
    return deps;
  }

  /**
   * Recalculate all formula cells in dependency order
   */
  recalculate(sheet) {
    // Find all formula cells
    const formulaCells = [];
    for (const [addr, cell] of sheet.cells) {
      if (cell.formula) {
        formulaCells.push(addr);
      }
    }

    // Topological sort
    const visited = new Set();
    const order = [];
    const inStack = new Set();

    const visit = (addr) => {
      if (visited.has(addr)) return;
      if (inStack.has(addr)) {
        // Circular reference
        const cell = sheet.getCellByAddress(addr);
        if (cell) cell.computedValue = '#CIRCULAR!';
        return;
      }
      inStack.add(addr);
      const cell = sheet.getCellByAddress(addr);
      if (cell) {
        for (const dep of cell.dependsOn) {
          visit(dep);
        }
      }
      inStack.delete(addr);
      visited.add(addr);
      order.push(addr);
    };

    for (const addr of formulaCells) {
      visit(addr);
    }

    // Evaluate in order
    for (const addr of order) {
      const cell = sheet.getCellByAddress(addr);
      if (cell && cell.formula) {
        const tokens = this.tokenize(cell.formula);
        const ast = this.parse(tokens);
        const visitSet = new Set([addr]);
        cell.computedValue = this.evaluate(ast, sheet, visitSet);
      }
    }
  }
}
