// Cell data model

export function createDefaultFormat() {
  return {
    bold: false,
    italic: false,
    textColor: null,
    bgColor: null,
    align: null,     // null = auto (numbers right, text left)
    fontSize: null,  // null = default
  };
}

export function createCell(rawValue = null) {
  return {
    rawValue,
    computedValue: rawValue,
    formula: null,
    format: createDefaultFormat(),
    dependsOn: new Set(),
    dependedBy: new Set(),
  };
}

export function cloneFormat(fmt) {
  return { ...fmt };
}

export function cloneCell(cell) {
  return {
    rawValue: cell.rawValue,
    computedValue: cell.computedValue,
    formula: cell.formula,
    format: cloneFormat(cell.format),
    dependsOn: new Set(cell.dependsOn),
    dependedBy: new Set(cell.dependedBy),
  };
}
