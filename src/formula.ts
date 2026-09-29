// Formula Engine for CSV & Excel spreadsheet

export interface CellCoord {
  row: number; // 0-indexed
  col: number; // 0-indexed
}

export interface CellRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface FormulaColor {
  text: string;
  bg: string;
  border: string;
}

export const FORMULA_COLORS: FormulaColor[] = [
  { text: "#2563eb", bg: "rgba(37, 99, 235, 0.14)", border: "#2563eb" },   // Blue
  { text: "#ea580c", bg: "rgba(234, 88, 12, 0.14)", border: "#ea580c" },   // Orange
  { text: "#9333ea", bg: "rgba(147, 51, 234, 0.14)", border: "#9333ea" },   // Purple
  { text: "#16a34a", bg: "rgba(22, 163, 74, 0.14)", border: "#16a34a" },   // Green
  { text: "#db2777", bg: "rgba(219, 39, 119, 0.14)", border: "#db2777" }, // Pink
  { text: "#0891b2", bg: "rgba(8, 145, 178, 0.14)", border: "#0891b2" },   // Cyan
  { text: "#d97706", bg: "rgba(217, 119, 6, 0.14)", border: "#d97706" },   // Amber
  { text: "#4f46e5", bg: "rgba(79, 70, 229, 0.14)", border: "#4f46e5" },   // Indigo
];

export function getFormulaColor(index: number): FormulaColor {
  return FORMULA_COLORS[index % FORMULA_COLORS.length];
}

export function colIndexToLetter(col: number): string {
  let temp = col;
  let letter = "";
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

export function letterToColIndex(str: string): number {
  let col = 0;
  const upper = str.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    col = col * 26 + (upper.charCodeAt(i) - 64);
  }
  return col - 1;
}

export function parseCellCoord(ref: string): CellCoord | null {
  const match = ref.trim().toUpperCase().match(/^([A-Z]+)([1-9][0-9]*)$/);
  if (!match) return null;
  const col = letterToColIndex(match[1]);
  const row = parseInt(match[2], 10) - 1;
  if (row < 0 || col < 0) return null;
  return { row, col };
}

export function parseCellRange(ref: string): CellRange | null {
  const match = ref.trim().toUpperCase().match(/^([A-Z]+[1-9][0-9]*):([A-Z]+[1-9][0-9]*)$/);
  if (!match) return null;
  const start = parseCellCoord(match[1]);
  const end = parseCellCoord(match[2]);
  if (!start || !end) return null;
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

export function formatCellCoord(row: number, col: number): string {
  return `${colIndexToLetter(col)}${row + 1}`;
}

export function formatCellRange(
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): string {
  if (startRow === endRow && startCol === endCol) {
    return formatCellCoord(startRow, startCol);
  }
  const minR = Math.min(startRow, endRow);
  const maxR = Math.max(startRow, endRow);
  const minC = Math.min(startCol, endCol);
  const maxC = Math.max(startCol, endCol);
  return `${formatCellCoord(minR, minC)}:${formatCellCoord(maxR, maxC)}`;
}

// ---------------------------------------------------------------------------
// Formula Tokens & Syntax Highlighting
// ---------------------------------------------------------------------------

export type TokenType =
  | "equal"
  | "cell"
  | "range"
  | "function"
  | "operator"
  | "number"
  | "string"
  | "paren"
  | "comma"
  | "whitespace"
  | "unknown";

export interface FormulaToken {
  type: TokenType;
  value: string;
  colorIndex?: number;
  coord?: CellCoord;
  range?: CellRange;
}

export interface FormulaRefHighlight {
  label: string;
  color: FormulaColor;
  range: CellRange;
}

export function tokenizeFormula(formula: string): FormulaToken[] {
  if (!formula.startsWith("=")) {
    return [{ type: "unknown", value: formula }];
  }

  const tokens: FormulaToken[] = [];
  tokens.push({ type: "equal", value: "=" });

  let i = 1;
  const n = formula.length;
  const refColorMap = new Map<string, number>();
  let nextColorIndex = 0;

  const getRefColor = (refKey: string): number => {
    const norm = refKey.toUpperCase();
    if (!refColorMap.has(norm)) {
      refColorMap.set(norm, nextColorIndex++);
    }
    return refColorMap.get(norm)!;
  };

  while (i < n) {
    const ch = formula[i];

    // Whitespace
    if (/\s/.test(ch)) {
      let ws = "";
      while (i < n && /\s/.test(formula[i])) {
        ws += formula[i++];
      }
      tokens.push({ type: "whitespace", value: ws });
      continue;
    }

    // String literal "..."
    if (ch === '"') {
      let str = '"';
      i++;
      while (i < n) {
        if (formula[i] === '"') {
          str += '"';
          i++;
          if (i < n && formula[i] === '"') {
            str += '"';
            i++;
          } else {
            break;
          }
        } else {
          str += formula[i++];
        }
      }
      tokens.push({ type: "string", value: str });
      continue;
    }

    // Parentheses & commas
    if (ch === "(" || ch === ")") {
      tokens.push({ type: "paren", value: ch });
      i++;
      continue;
    }
    if (ch === ",") {
      tokens.push({ type: "comma", value: ch });
      i++;
      continue;
    }

    // Multi-char / Single-char Operators
    const twoChar = formula.slice(i, i + 2);
    if (["<=", ">=", "<>", "!=", "=="].includes(twoChar)) {
      tokens.push({ type: "operator", value: twoChar });
      i += 2;
      continue;
    }
    if (["+", "-", "*", "/", "^", "%", "&", "=", "<", ">"].includes(ch)) {
      tokens.push({ type: "operator", value: ch });
      i++;
      continue;
    }

    // Numbers: digits with optional decimal
    if (/\d/.test(ch) || (ch === "." && i + 1 < n && /\d/.test(formula[i + 1]))) {
      let num = "";
      while (i < n && /[\d.]/.test(formula[i])) {
        num += formula[i++];
      }
      tokens.push({ type: "number", value: num });
      continue;
    }

    // Identifiers (Function names, Cell references, or Range references)
    if (/[A-Za-z]/.test(ch)) {
      let id = "";
      while (i < n && /[A-Za-z0-9_]/.test(formula[i])) {
        id += formula[i++];
      }

      // Check if followed by colon ':' and another cell identifier => Range reference A1:B10
      if (i < n && formula[i] === ":") {
        let peek = i + 1;
        while (peek < n && /\s/.test(formula[peek])) peek++;
        let secondId = "";
        while (peek < n && /[A-Za-z0-9]/.test(formula[peek])) {
          secondId += formula[peek++];
        }
        const rangeText = `${id}:${secondId}`;
        const parsedRange = parseCellRange(rangeText);
        if (parsedRange) {
          const colorIdx = getRefColor(rangeText);
          tokens.push({
            type: "range",
            value: rangeText,
            colorIndex: colorIdx,
            range: parsedRange,
          });
          i = peek;
          continue;
        }
      }

      // Check if it's a single cell reference
      const parsedCoord = parseCellCoord(id);
      if (parsedCoord) {
        const colorIdx = getRefColor(id);
        tokens.push({
          type: "cell",
          value: id,
          colorIndex: colorIdx,
          coord: parsedCoord,
          range: {
            startRow: parsedCoord.row,
            startCol: parsedCoord.col,
            endRow: parsedCoord.row,
            endCol: parsedCoord.col,
          },
        });
        continue;
      }

      // Otherwise it's a function name or variable
      tokens.push({ type: "function", value: id });
      continue;
    }

    // Fallback unknown character
    tokens.push({ type: "unknown", value: ch });
    i++;
  }

  return tokens;
}

export function extractFormulaHighlights(formula: string): FormulaRefHighlight[] {
  if (!formula.startsWith("=")) return [];
  const tokens = tokenizeFormula(formula);
  const highlights: FormulaRefHighlight[] = [];
  const seenRanges = new Set<string>();

  for (const token of tokens) {
    if ((token.type === "cell" || token.type === "range") && token.range) {
      const key = `${token.range.startRow},${token.range.startCol}-${token.range.endRow},${token.range.endCol}`;
      if (!seenRanges.has(key)) {
        seenRanges.add(key);
        highlights.push({
          label: token.value.toUpperCase(),
          color: getFormulaColor(token.colorIndex ?? 0),
          range: token.range,
        });
      }
    }
  }

  return highlights;
}

// ---------------------------------------------------------------------------
// Formula Point & Click insertion helper
// ---------------------------------------------------------------------------

export function isFormulaPointingAllowed(formula: string): boolean {
  if (!formula.startsWith("=")) return false;
  const trimmed = formula.trimEnd();
  if (trimmed === "=") return true;
  const lastChar = trimmed[trimmed.length - 1];
  return ["=", "+", "-", "*", "/", "(", ",", "&", ":", "<", ">"].includes(lastChar);
}

export function applyCellReferenceToFormula(
  currentFormula: string,
  newRef: string,
  isReplacingLastRef: boolean,
): { formula: string; isReplacing: boolean } {
  if (!currentFormula.startsWith("=")) {
    return { formula: `=${newRef}`, isReplacing: true };
  }

  if (isReplacingLastRef) {
    // Replace the trailing cell or range reference
    const replaced = currentFormula.replace(
      /([A-Za-z]+[1-9][0-9]*(?::[A-Za-z]+[1-9][0-9]*)?)\s*$/,
      newRef,
    );
    if (replaced !== currentFormula) {
      return { formula: replaced, isReplacing: true };
    }
  }

  // If ends with an operator/delimiter or '=', append newRef
  const trimmed = currentFormula.trimEnd();
  const lastChar = trimmed[trimmed.length - 1];
  if (["=", "+", "-", "*", "/", "(", ",", "&", ":", "<", ">"].includes(lastChar)) {
    return { formula: `${currentFormula}${newRef}`, isReplacing: true };
  }

  // Fallback: replace or append with '+'
  return { formula: `${currentFormula}+${newRef}`, isReplacing: true };
}

// ---------------------------------------------------------------------------
// Formula Evaluation Engine
// ---------------------------------------------------------------------------

export type FormulaValue = string | number | boolean | null;

export function evaluateFormula(
  formula: string,
  getCellValue: (row: number, col: number) => FormulaValue,
  visited: Set<string> = new Set(),
): FormulaValue {
  if (!formula.startsWith("=")) {
    return formula;
  }

  const expr = formula.slice(1).trim();
  if (!expr) return "";

  try {
    const evaluator = new FormulaEvaluator(expr, getCellValue, visited);
    const result = evaluator.evaluate();
    if (result === null || result === undefined) return "";
    if (typeof result === "number") {
      if (Number.isNaN(result)) return "#VALUE!";
      if (!Number.isFinite(result)) return "#DIV/0!";
      // Round floating point inaccuracies (e.g. 0.1 + 0.2 = 0.3)
      const rounded = Math.round(result * 1e10) / 1e10;
      return rounded;
    }
    if (Array.isArray(result)) {
      const first = Array.isArray(result[0]) ? result[0][0] : result[0];
      return first ?? "";
    }
    return result;
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    if (msg.startsWith("#")) return msg;
    return "#ERROR!";
  }
}

class FormulaEvaluator {
  private pos = 0;
  private len: number;

  constructor(
    private expr: string,
    private getCellValue: (row: number, col: number) => FormulaValue,
    private visited: Set<string>,
  ) {
    this.len = expr.length;
  }

  private skipWhitespace(): void {
    while (this.pos < this.len && /\s/.test(this.expr[this.pos])) {
      this.pos++;
    }
  }

  public evaluate(): FormulaValue | FormulaValue[] | FormulaValue[][] {
    this.skipWhitespace();
    const val = this.parseExpression();
    this.skipWhitespace();
    if (this.pos < this.len) {
      throw new Error("#ERROR!");
    }
    return val;
  }

  // Recursive descent precedence:
  // parseExpression: Comparison (=, <>, <, <=, >, >=)
  // parseConcat: String concatenation (&)
  // parseAddSub: Addition (+, -)
  // parseMulDiv: Multiplication (*, /, %)
  // parsePower: Power (^)
  // parseUnary: Unary (+, -)
  // parsePrimary: Numbers, Strings, Functions, Cell Refs, Ranges, Parentheses

  private parseExpression(): any {
    let left = this.parseConcat();
    this.skipWhitespace();

    while (this.pos < this.len) {
      const two = this.expr.slice(this.pos, this.pos + 2);
      if (["<=", ">=", "<>", "!=", "=="].includes(two)) {
        this.pos += 2;
        const right = this.parseConcat();
        left = this.compare(left, right, two);
        this.skipWhitespace();
        continue;
      }

      const one = this.expr[this.pos];
      if (["=", "<", ">"].includes(one)) {
        this.pos++;
        const right = this.parseConcat();
        left = this.compare(left, right, one);
        this.skipWhitespace();
        continue;
      }
      break;
    }

    return left;
  }

  private compare(left: any, right: any, op: string): boolean {
    const l = this.unwrapScalar(left);
    const r = this.unwrapScalar(right);

    const lNum = Number(l);
    const rNum = Number(r);
    const bothNum = !Number.isNaN(lNum) && !Number.isNaN(rNum) && l !== "" && r !== "";

    const a = bothNum ? lNum : String(l ?? "").toLowerCase();
    const b = bothNum ? rNum : String(r ?? "").toLowerCase();

    switch (op) {
      case "=":
      case "==":
        return a === b;
      case "<>":
      case "!=":
        return a !== b;
      case "<":
        return a < b;
      case "<=":
        return a <= b;
      case ">":
        return a > b;
      case ">=":
        return a >= b;
      default:
        return false;
    }
  }

  private parseConcat(): any {
    let left = this.parseAddSub();
    this.skipWhitespace();

    while (this.pos < this.len && this.expr[this.pos] === "&") {
      this.pos++;
      const right = this.parseAddSub();
      const l = this.unwrapScalar(left);
      const r = this.unwrapScalar(right);
      left = `${l ?? ""}${r ?? ""}`;
      this.skipWhitespace();
    }

    return left;
  }

  private parseAddSub(): any {
    let left = this.parseMulDiv();
    this.skipWhitespace();

    while (this.pos < this.len) {
      const op = this.expr[this.pos];
      if (op !== "+" && op !== "-") break;
      this.pos++;
      const right = this.parseMulDiv();
      const l = this.toNumber(this.unwrapScalar(left));
      const r = this.toNumber(this.unwrapScalar(right));
      left = op === "+" ? l + r : l - r;
      this.skipWhitespace();
    }

    return left;
  }

  private parseMulDiv(): any {
    let left = this.parsePower();
    this.skipWhitespace();

    while (this.pos < this.len) {
      const op = this.expr[this.pos];
      if (op !== "*" && op !== "/" && op !== "%") break;
      this.pos++;
      const right = this.parsePower();
      const l = this.toNumber(this.unwrapScalar(left));
      const r = this.toNumber(this.unwrapScalar(right));
      if (op === "*") {
        left = l * r;
      } else if (op === "/") {
        if (r === 0) throw new Error("#DIV/0!");
        left = l / r;
      } else if (op === "%") {
        if (r === 0) throw new Error("#DIV/0!");
        left = l % r;
      }
      this.skipWhitespace();
    }

    return left;
  }

  private parsePower(): any {
    let left = this.parseUnary();
    this.skipWhitespace();

    while (this.pos < this.len && this.expr[this.pos] === "^") {
      this.pos++;
      const right = this.parseUnary();
      const l = this.toNumber(this.unwrapScalar(left));
      const r = this.toNumber(this.unwrapScalar(right));
      left = Math.pow(l, r);
      this.skipWhitespace();
    }

    return left;
  }

  private parseUnary(): any {
    this.skipWhitespace();
    if (this.pos < this.len) {
      const ch = this.expr[this.pos];
      if (ch === "+") {
        this.pos++;
        return +this.toNumber(this.unwrapScalar(this.parseUnary()));
      }
      if (ch === "-") {
        this.pos++;
        return -this.toNumber(this.unwrapScalar(this.parseUnary()));
      }
    }
    return this.parsePrimary();
  }

  private parsePrimary(): any {
    this.skipWhitespace();
    if (this.pos >= this.len) {
      throw new Error("#ERROR!");
    }

    const ch = this.expr[this.pos];

    // Parentheses
    if (ch === "(") {
      this.pos++;
      const val = this.parseExpression();
      this.skipWhitespace();
      if (this.pos >= this.len || this.expr[this.pos] !== ")") {
        throw new Error("#ERROR!");
      }
      this.pos++;
      return val;
    }

    // String literal
    if (ch === '"') {
      return this.parseString();
    }

    // Number literal
    if (/\d/.test(ch) || (ch === "." && this.pos + 1 < this.len && /\d/.test(this.expr[this.pos + 1]))) {
      return this.parseNumber();
    }

    // Identifiers: Functions, Cell references, Range references, Boolean (TRUE/FALSE)
    if (/[A-Za-z]/.test(ch)) {
      return this.parseIdentifierOrReference();
    }

    throw new Error("#ERROR!");
  }

  private parseString(): string {
    let str = "";
    this.pos++; // skip opening quote
    while (this.pos < this.len) {
      if (this.expr[this.pos] === '"') {
        this.pos++;
        if (this.pos < this.len && this.expr[this.pos] === '"') {
          str += '"';
          this.pos++;
        } else {
          return str;
        }
      } else {
        str += this.expr[this.pos++];
      }
    }
    return str;
  }

  private parseNumber(): number {
    let numStr = "";
    while (this.pos < this.len && /[\d.]/.test(this.expr[this.pos])) {
      numStr += this.expr[this.pos++];
    }
    const val = parseFloat(numStr);
    if (Number.isNaN(val)) throw new Error("#VALUE!");
    return val;
  }

  private parseIdentifierOrReference(): any {
    let id = "";
    while (this.pos < this.len && /[A-Za-z0-9_]/.test(this.expr[this.pos])) {
      id += this.expr[this.pos++];
    }

    this.skipWhitespace();

    // 1. Function Call: ID followed by '('
    if (this.pos < this.len && this.expr[this.pos] === "(") {
      this.pos++; // skip '('
      const args: any[] = [];
      this.skipWhitespace();
      if (this.pos < this.len && this.expr[this.pos] !== ")") {
        while (this.pos < this.len) {
          args.push(this.parseExpression());
          this.skipWhitespace();
          if (this.pos < this.len && this.expr[this.pos] === ",") {
            this.pos++;
            continue;
          }
          break;
        }
      }
      if (this.pos >= this.len || this.expr[this.pos] !== ")") {
        throw new Error("#ERROR!");
      }
      this.pos++; // skip ')'
      return this.callFunction(id.toUpperCase(), args);
    }

    // 2. Range Reference: ID followed by ':' and second ID
    if (this.pos < this.len && this.expr[this.pos] === ":") {
      this.pos++; // skip ':'
      this.skipWhitespace();
      let secondId = "";
      while (this.pos < this.len && /[A-Za-z0-9]/.test(this.expr[this.pos])) {
        secondId += this.expr[this.pos++];
      }
      const rangeText = `${id}:${secondId}`;
      const range = parseCellRange(rangeText);
      if (range) {
        return this.getRangeValues(range);
      }
      throw new Error("#REF!");
    }

    // 3. Single Cell Reference (e.g. B1, A10)
    const coord = parseCellCoord(id);
    if (coord) {
      return this.getSingleCellValue(coord.row, coord.col);
    }

    // 4. Boolean constants
    const upperId = id.toUpperCase();
    if (upperId === "TRUE") return true;
    if (upperId === "FALSE") return false;

    throw new Error("#NAME?");
  }

  private getSingleCellValue(row: number, col: number): FormulaValue {
    const key = `${row},${col}`;
    if (this.visited.has(key)) {
      throw new Error("#CYCLE!");
    }

    const raw = this.getCellValue(row, col);
    if (typeof raw === "string" && raw.startsWith("=")) {
      this.visited.add(key);
      try {
        const evalResult = evaluateFormula(raw, this.getCellValue, this.visited);
        return evalResult;
      } finally {
        this.visited.delete(key);
      }
    }
    return raw;
  }

  private getRangeValues(range: CellRange): FormulaValue[][] {
    const matrix: FormulaValue[][] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const rowVals: FormulaValue[] = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        rowVals.push(this.getSingleCellValue(r, c));
      }
      matrix.push(rowVals);
    }
    return matrix;
  }

  private unwrapScalar(val: any): any {
    if (Array.isArray(val)) {
      if (val.length === 0) return 0;
      if (Array.isArray(val[0])) {
        return val[0][0] ?? 0;
      }
      return val[0] ?? 0;
    }
    return val;
  }

  private toNumber(val: any): number {
    if (typeof val === "number") return val;
    if (val === null || val === undefined || val === "") return 0;
    if (typeof val === "boolean") return val ? 1 : 0;
    const num = Number(val);
    if (Number.isNaN(num)) throw new Error("#VALUE!");
    return num;
  }

  private flattenValues(args: any[]): FormulaValue[] {
    const result: FormulaValue[] = [];
    const walk = (item: any) => {
      if (Array.isArray(item)) {
        for (const sub of item) walk(sub);
      } else {
        result.push(item);
      }
    };
    walk(args);
    return result;
  }

  // ---------------------------------------------------------------------------
  // Function Implementations
  // ---------------------------------------------------------------------------

  private callFunction(name: string, args: any[]): any {
    switch (name) {
      // 1. SUM
      case "SUM": {
        const flat = this.flattenValues(args);
        let sum = 0;
        for (const v of flat) {
          if (typeof v === "number") {
            sum += v;
          } else if (typeof v === "string" && v.trim() !== "") {
            const n = Number(v);
            if (!Number.isNaN(n)) sum += n;
          }
        }
        return sum;
      }

      // 2. AVERAGE / AVG
      case "AVERAGE":
      case "AVG": {
        const flat = this.flattenValues(args);
        let sum = 0;
        let count = 0;
        for (const v of flat) {
          if (typeof v === "number") {
            sum += v;
            count++;
          } else if (typeof v === "string" && v.trim() !== "") {
            const n = Number(v);
            if (!Number.isNaN(n)) {
              sum += n;
              count++;
            }
          }
        }
        if (count === 0) throw new Error("#DIV/0!");
        return sum / count;
      }

      // 3. COUNT (counts numbers only)
      case "COUNT": {
        const flat = this.flattenValues(args);
        let count = 0;
        for (const v of flat) {
          if (typeof v === "number") count++;
          else if (typeof v === "string" && v.trim() !== "" && !Number.isNaN(Number(v))) count++;
        }
        return count;
      }

      // 4. COUNTA (counts non-empty cells)
      case "COUNTA": {
        const flat = this.flattenValues(args);
        let count = 0;
        for (const v of flat) {
          if (v !== null && v !== undefined && v !== "") count++;
        }
        return count;
      }

      // 5. COUNTBLANK
      case "COUNTBLANK": {
        const flat = this.flattenValues(args);
        let count = 0;
        for (const v of flat) {
          if (v === null || v === undefined || v === "") count++;
        }
        return count;
      }

      // 6. COUNTIF(range, criteria)
      case "COUNTIF": {
        if (args.length < 2) throw new Error("#VALUE!");
        const rangeVals = this.flattenValues([args[0]]);
        const criteria = this.unwrapScalar(args[1]);
        let count = 0;
        for (const val of rangeVals) {
          if (this.matchesCriteria(val, criteria)) {
            count++;
          }
        }
        return count;
      }

      // 7. SUMIF(range, criteria, [sum_range])
      case "SUMIF": {
        if (args.length < 2) throw new Error("#VALUE!");
        const rangeVals = this.flattenValues([args[0]]);
        const criteria = this.unwrapScalar(args[1]);
        const sumVals = args.length >= 3 ? this.flattenValues([args[2]]) : rangeVals;

        let sum = 0;
        for (let i = 0; i < rangeVals.length; i++) {
          if (this.matchesCriteria(rangeVals[i], criteria)) {
            const sItem = sumVals[i];
            const num = typeof sItem === "number" ? sItem : Number(sItem);
            if (!Number.isNaN(num)) {
              sum += num;
            }
          }
        }
        return sum;
      }

      // 8. VLOOKUP(lookup_value, table_array, col_index, [range_lookup])
      case "VLOOKUP": {
        if (args.length < 3) throw new Error("#VALUE!");
        const lookupVal = this.unwrapScalar(args[0]);
        const table = args[1];
        const colIndex = Math.floor(this.toNumber(this.unwrapScalar(args[2])));
        const rangeLookup = args.length >= 4 ? Boolean(this.unwrapScalar(args[3])) : false;

        if (!Array.isArray(table) || table.length === 0) {
          throw new Error("#VALUE!");
        }

        // Standardize table to 2D array of rows
        const rows: FormulaValue[][] = Array.isArray(table[0])
          ? (table as FormulaValue[][])
          : [table as FormulaValue[]];

        if (colIndex < 1 || colIndex > (rows[0]?.length || 0)) {
          throw new Error("#REF!");
        }

        const normLookup = typeof lookupVal === "string" ? lookupVal.toLowerCase() : lookupVal;

        if (!rangeLookup) {
          // Exact match
          for (let r = 0; r < rows.length; r++) {
            const cellVal = rows[r][0];
            const normCell = typeof cellVal === "string" ? cellVal.toLowerCase() : cellVal;
            if (
              normCell === normLookup ||
              (typeof normCell === "number" &&
                typeof normLookup === "number" &&
                normCell === normLookup)
            ) {
              return rows[r][colIndex - 1] ?? "";
            }
          }
          throw new Error("#N/A");
        } else {
          // Approximate match (assumes table is sorted ascending)
          let bestRow = -1;
          for (let r = 0; r < rows.length; r++) {
            const cellVal = rows[r][0];
            const cNum = Number(cellVal);
            const lNum = Number(normLookup);
            if (!Number.isNaN(cNum) && !Number.isNaN(lNum)) {
              if (cNum <= lNum) {
                bestRow = r;
              } else {
                break;
              }
            } else {
              const cStr = String(cellVal ?? "").toLowerCase();
              const lStr = String(normLookup ?? "").toLowerCase();
              if (cStr <= lStr) {
                bestRow = r;
              } else {
                break;
              }
            }
          }
          if (bestRow !== -1) {
            return rows[bestRow][colIndex - 1] ?? "";
          }
          throw new Error("#N/A");
        }
      }

      // 9. IF(condition, true_val, [false_val])
      case "IF": {
        if (args.length < 2) throw new Error("#VALUE!");
        const cond = Boolean(this.unwrapScalar(args[0]));
        if (cond) {
          return this.unwrapScalar(args[1]);
        }
        return args.length >= 3 ? this.unwrapScalar(args[2]) : false;
      }

      // 10. IFERROR(val, fallback)
      case "IFERROR": {
        if (args.length < 2) throw new Error("#VALUE!");
        try {
          const val = this.unwrapScalar(args[0]);
          if (typeof val === "string" && val.startsWith("#")) {
            return this.unwrapScalar(args[1]);
          }
          return val;
        } catch {
          return this.unwrapScalar(args[1]);
        }
      }

      // 11. MIN / MAX / MEDIAN
      case "MIN": {
        const flat = this.flattenValues(args);
        const nums = flat.map((v) => Number(v)).filter((n) => !Number.isNaN(n));
        if (nums.length === 0) return 0;
        return Math.min(...nums);
      }
      case "MAX": {
        const flat = this.flattenValues(args);
        const nums = flat.map((v) => Number(v)).filter((n) => !Number.isNaN(n));
        if (nums.length === 0) return 0;
        return Math.max(...nums);
      }
      case "PRODUCT": {
        const flat = this.flattenValues(args);
        let prod = 1;
        let hasNum = false;
        for (const v of flat) {
          if (typeof v === "number") {
            prod *= v;
            hasNum = true;
          } else if (typeof v === "string" && v.trim() !== "") {
            const n = Number(v);
            if (!Number.isNaN(n)) {
              prod *= n;
              hasNum = true;
            }
          }
        }
        return hasNum ? prod : 0;
      }

      // 12. Math Functions
      case "ABS":
        return Math.abs(this.toNumber(this.unwrapScalar(args[0])));
      case "ROUND": {
        const num = this.toNumber(this.unwrapScalar(args[0]));
        const digits = args.length >= 2 ? Math.floor(this.toNumber(this.unwrapScalar(args[1]))) : 0;
        const factor = Math.pow(10, digits);
        return Math.round(num * factor) / factor;
      }
      case "CEIL":
      case "CEILING":
        return Math.ceil(this.toNumber(this.unwrapScalar(args[0])));
      case "FLOOR":
        return Math.floor(this.toNumber(this.unwrapScalar(args[0])));
      case "SQRT": {
        const val = this.toNumber(this.unwrapScalar(args[0]));
        if (val < 0) throw new Error("#NUM!");
        return Math.sqrt(val);
      }
      case "POWER":
        return Math.pow(
          this.toNumber(this.unwrapScalar(args[0])),
          this.toNumber(this.unwrapScalar(args[1])),
        );
      case "MOD": {
        const n = this.toNumber(this.unwrapScalar(args[0]));
        const d = this.toNumber(this.unwrapScalar(args[1]));
        if (d === 0) throw new Error("#DIV/0!");
        return n % d;
      }

      // 13. Text Functions
      case "CONCAT":
      case "CONCATENATE": {
        const flat = this.flattenValues(args);
        return flat.map((v) => (v === null || v === undefined ? "" : String(v))).join("");
      }
      case "UPPER":
        return String(this.unwrapScalar(args[0]) ?? "").toUpperCase();
      case "LOWER":
        return String(this.unwrapScalar(args[0]) ?? "").toLowerCase();
      case "TRIM":
        return String(this.unwrapScalar(args[0]) ?? "").trim();
      case "LEN":
        return String(this.unwrapScalar(args[0]) ?? "").length;
      case "LEFT": {
        const str = String(this.unwrapScalar(args[0]) ?? "");
        const len = args.length >= 2 ? Math.floor(this.toNumber(this.unwrapScalar(args[1]))) : 1;
        return str.slice(0, Math.max(0, len));
      }
      case "RIGHT": {
        const str = String(this.unwrapScalar(args[0]) ?? "");
        const len = args.length >= 2 ? Math.floor(this.toNumber(this.unwrapScalar(args[1]))) : 1;
        return str.slice(Math.max(0, str.length - len));
      }
      case "MID": {
        const str = String(this.unwrapScalar(args[0]) ?? "");
        const start = Math.floor(this.toNumber(this.unwrapScalar(args[1]))) - 1;
        const len = Math.floor(this.toNumber(this.unwrapScalar(args[2])));
        return str.slice(Math.max(0, start), Math.max(0, start) + Math.max(0, len));
      }

      // 14. Dates
      case "TODAY":
        return new Date().toISOString().split("T")[0];
      case "NOW":
        return new Date().toLocaleString();

      default:
        throw new Error("#NAME?");
    }
  }

  private matchesCriteria(cellVal: FormulaValue, criteria: any): boolean {
    if (criteria === null || criteria === undefined) return cellVal === "" || cellVal === null;

    const critStr = String(criteria).trim();

    // Check operator prefix: >=, <=, <>, !=, >, <, =
    const matchOp = critStr.match(/^([><]=?|<>|!=|=)(.*)$/);
    if (matchOp) {
      const op = matchOp[1];
      const targetStr = matchOp[2].trim();
      const targetNum = Number(targetStr);
      const cellNum = Number(cellVal);

      if (!Number.isNaN(targetNum) && !Number.isNaN(cellNum) && cellVal !== "" && cellVal !== null) {
        if (op === ">") return cellNum > targetNum;
        if (op === ">=") return cellNum >= targetNum;
        if (op === "<") return cellNum < targetNum;
        if (op === "<=") return cellNum <= targetNum;
        if (op === "<>" || op === "!=") return cellNum !== targetNum;
        if (op === "=") return cellNum === targetNum;
      }

      const cStr = String(cellVal ?? "").toLowerCase();
      const tStr = targetStr.toLowerCase();
      if (op === ">") return cStr > tStr;
      if (op === ">=") return cStr >= tStr;
      if (op === "<") return cStr < tStr;
      if (op === "<=") return cStr <= tStr;
      if (op === "<>" || op === "!=") return cStr !== tStr;
      if (op === "=") return cStr === tStr;
    }

    // Direct equality (case-insensitive for string, wildcard * supported)
    if (critStr.includes("*")) {
      const regex = new RegExp(`^${critStr.replace(/\*/g, ".*")}$`, "i");
      return regex.test(String(cellVal ?? ""));
    }

    const cNum = Number(cellVal);
    const crNum = Number(critStr);
    if (!Number.isNaN(cNum) && !Number.isNaN(crNum) && cellVal !== "" && cellVal !== null) {
      return cNum === crNum;
    }

    return String(cellVal ?? "").toLowerCase() === critStr.toLowerCase();
  }
}
