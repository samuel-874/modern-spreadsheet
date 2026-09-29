import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";
import {
  colIndexToLetter,
  formatCellCoord,
  formatCellRange,
  tokenizeFormula,
  extractFormulaHighlights,
  applyCellReferenceToFormula,
  evaluateFormula,
  getFormulaColor,
  FormulaValue,
} from "./formula";
import { GridbookIcon } from "./GridbookLogo";

const renderFormulaSyntax = (formula: string) => {
  const tokens = tokenizeFormula(formula);
  return tokens.map((token, idx) => {
    if (token.type === "cell" || token.type === "range") {
      const color = getFormulaColor(token.colorIndex ?? 0);
      return (
        <span
          key={idx}
          className="formula-token-ref"
          style={{
            color: color.text,
            backgroundColor: color.bg,
            border: `1px solid ${color.border}40`,
          }}
        >
          {token.value}
        </span>
      );
    }
    if (token.type === "function") {
      return (
        <span key={idx} className="formula-token-fn">
          {token.value}
        </span>
      );
    }
    if (token.type === "operator" || token.type === "equal") {
      return (
        <span key={idx} className="formula-token-op">
          {token.value}
        </span>
      );
    }
    if (token.type === "paren") {
      return (
        <span key={idx} className="formula-token-paren">
          {token.value}
        </span>
      );
    }
    if (token.type === "number") {
      return (
        <span key={idx} className="formula-token-num">
          {token.value}
        </span>
      );
    }
    if (token.type === "string") {
      return (
        <span key={idx} className="formula-token-str">
          {token.value}
        </span>
      );
    }
    return <span key={idx}>{token.value}</span>;
  });
};

type CellStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: string;
  fontSize?: number;
  align?: "left" | "center" | "right";
};

type CellData = {
  value: string;
  style?: CellStyle;
};

type Cell = string | CellData;

type MergeRange = {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

type Sheet = {
  name: string;
  rows: Cell[][];
  merges?: MergeRange[];
};

type Document = {
  id: string;
  name: string;
  path?: string;
  sheets: Sheet[];
  activeSheet: number;
  dirty: boolean;
};

type StoredCell =
  | string
  | number
  | boolean
  | null
  | { value?: unknown; style?: CellStyle };

type StoredSheet = {
  name?: string;
  rows?: StoredCell[][];
  merges?: MergeRange[];
};

type StoredDoc = {
  id?: string;
  name?: string;
  path?: string;
  sheets?: StoredSheet[];
  activeSheet?: number;
  dirty?: boolean;
};

type ContextMenuTarget =
  | { type: "row"; rowIndex: number }
  | { type: "column"; columnIndex: number }
  | { type: "cell"; rowIndex: number; columnIndex: number }
  | { type: "sheet"; sheetIndex: number };

type ContextMenuState = {
  visible: boolean;
  x: number;
  y: number;
  target: ContextMenuTarget;
} | null;

const getCellValue = (cell: Cell | undefined): string => {
  if (cell === undefined || cell === null) return "";
  if (typeof cell === "string") return cell;
  return cell.value ?? "";
};

const getCellStyle = (cell: Cell | undefined): CellStyle => {
  if (!cell || typeof cell === "string") return {};
  return cell.style || {};
};

const FONT_FAMILIES: Record<string, string> = {
  Default: '"DM Mono", monospace',
  Inter: '"Inter", sans-serif',
  Roboto: '"Roboto", sans-serif',
  Arial: "Arial, sans-serif",
  Georgia: "Georgia, serif",
  Monaco: "Monaco, monospace",
  "Courier New": '"Courier New", Courier, monospace',
  "Times New Roman": '"Times New Roman", Times, serif',
  "Trebuchet MS": '"Trebuchet MS", sans-serif',
  Verdana: "Verdana, sans-serif",
};

const FONT_OPTIONS = [
  "Default",
  "Inter",
  "Roboto",
  "Arial",
  "Georgia",
  "Monaco",
  "Courier New",
  "Times New Roman",
  "Trebuchet MS",
  "Verdana",
];

const FONT_SIZES = [9, 10, 11, 12, 14, 16, 18, 20, 24, 32];

const getCellInlineStyle = (style: CellStyle): React.CSSProperties => {
  const result: React.CSSProperties = {};
  if (style.bold) result.fontWeight = "bold";
  if (style.italic) result.fontStyle = "italic";
  if (style.underline) result.textDecoration = "underline";
  if (style.fontFamily && style.fontFamily !== "Default") {
    result.fontFamily = FONT_FAMILIES[style.fontFamily] || style.fontFamily;
  }
  if (style.fontSize) {
    result.fontSize = `${style.fontSize}px`;
  }
  if (style.align) {
    result.textAlign = style.align;
  }
  return result;
};

const toCsv = (sheet: Sheet) =>
  sheet.rows
    .map((row) =>
      row
        .map((cell) => {
          const val = getCellValue(cell);
          return /[",\n]/.test(val) ? `"${val.replace(/"/g, '""')}"` : val;
        })
        .join(","),
    )
    .join("\n");

const letters = Array.from({ length: 26 }, (_, index) =>
  String.fromCharCode(65 + index),
);

const starterRows: Cell[][] = Array.from({ length: 18 }, (_, row) =>
  Array.from({ length: 8 }, (_, column) => {
    if (row === 0) {
      const headers = [
        "Item",
        "Category",
        "Owner",
        "Status",
        "Amount",
        "Due date",
        "Priority",
        "Notes",
      ];
      return {
        value: headers[column] || "",
        style: { bold: true },
      };
    }
    return { value: "", style: {} };
  }),
);

const newDocument = (): Document => ({
  id: crypto.randomUUID(),
  name: "Untitled Spreadsheet",
  sheets: [
    {
      name: "Sheet 1",
      rows: starterRows.map((row) =>
        row.map((cell) =>
          typeof cell === "object"
            ? { value: cell.value, style: { ...cell.style } }
            : { value: cell, style: {} },
        ),
      ),
      merges: [],
    },
  ],
  activeSheet: 0,
  dirty: false,
});

const adjustMergesOnRowInsert = (
  merges: MergeRange[],
  atRow: number,
  count: number,
): MergeRange[] => {
  return merges.map((m) => {
    if (m.startRow >= atRow) {
      return { ...m, startRow: m.startRow + count, endRow: m.endRow + count };
    }
    if (m.endRow >= atRow) {
      return { ...m, endRow: m.endRow + count };
    }
    return m;
  });
};

const adjustMergesOnRowDelete = (
  merges: MergeRange[],
  fromRow: number,
  toRow: number,
): MergeRange[] => {
  const count = toRow - fromRow + 1;
  return merges
    .filter((m) => !(m.startRow >= fromRow && m.endRow <= toRow))
    .map((m) => {
      let startR = m.startRow;
      let endR = m.endRow;
      if (startR > toRow) startR -= count;
      else if (startR >= fromRow) startR = fromRow;
      if (endR > toRow) endR -= count;
      else if (endR >= fromRow) endR = Math.max(startR, fromRow - 1);
      return { ...m, startRow: startR, endRow: Math.max(startR, endR) };
    })
    .filter((m) => m.endRow >= m.startRow);
};

const adjustMergesOnColInsert = (
  merges: MergeRange[],
  atCol: number,
  count: number,
): MergeRange[] => {
  return merges.map((m) => {
    if (m.startCol >= atCol) {
      return { ...m, startCol: m.startCol + count, endCol: m.endCol + count };
    }
    if (m.endCol >= atCol) {
      return { ...m, endCol: m.endCol + count };
    }
    return m;
  });
};

const adjustMergesOnColDelete = (
  merges: MergeRange[],
  fromCol: number,
  toCol: number,
): MergeRange[] => {
  const count = toCol - fromCol + 1;
  return merges
    .filter((m) => !(m.startCol >= fromCol && m.endCol <= toCol))
    .map((m) => {
      let startC = m.startCol;
      let endC = m.endCol;
      if (startC > toCol) startC -= count;
      else if (startC >= fromCol) startC = fromCol;
      if (endC > toCol) endC -= count;
      else if (endC >= fromCol) endC = Math.max(startC, fromCol - 1);
      return { ...m, startCol: startC, endCol: Math.max(startC, endC) };
    })
    .filter((m) => m.endCol >= m.startCol);
};

const expandRangeForMerges = (
  r1: number,
  c1: number,
  r2: number,
  c2: number,
  merges: MergeRange[],
): { startRow: number; startCol: number; endRow: number; endCol: number } => {
  let minR = Math.min(r1, r2);
  let maxR = Math.max(r1, r2);
  let minC = Math.min(c1, c2);
  let maxC = Math.max(c1, c2);

  let changed = true;
  let iterations = 0;
  while (changed && iterations < 30) {
    changed = false;
    iterations++;
    for (const m of merges) {
      const mMinR = Math.min(m.startRow, m.endRow);
      const mMaxR = Math.max(m.startRow, m.endRow);
      const mMinC = Math.min(m.startCol, m.endCol);
      const mMaxC = Math.max(m.startCol, m.endCol);

      const intersects =
        maxR >= mMinR &&
        minR <= mMaxR &&
        maxC >= mMinC &&
        minC <= mMaxC;

      if (intersects) {
        if (mMinR < minR) {
          minR = mMinR;
          changed = true;
        }
        if (mMaxR > maxR) {
          maxR = mMaxR;
          changed = true;
        }
        if (mMinC < minC) {
          minC = mMinC;
          changed = true;
        }
        if (mMaxC > maxC) {
          maxC = mMaxC;
          changed = true;
        }
      }
    }
  }

  return {
    startRow: r1 <= r2 ? minR : maxR,
    startCol: c1 <= c2 ? minC : maxC,
    endRow: r1 <= r2 ? maxR : minR,
    endCol: c1 <= c2 ? maxC : minC,
  };
};

function App() {
  const [documents, setDocuments] = useState<Document[]>(() => {
    try {
      const raw = JSON.parse(
        localStorage.getItem("paperclip-documents") || "[]",
      ) as unknown;
      if (!Array.isArray(raw)) return [];
      return (raw as StoredDoc[]).map((doc) => ({
        id: doc.id || crypto.randomUUID(),
        name: doc.name || "Untitled Spreadsheet",
        path: doc.path,
        activeSheet: doc.activeSheet || 0,
        dirty: Boolean(doc.dirty),
        sheets: (doc.sheets || []).map((sheet) => ({
          name: sheet.name || "Sheet 1",
          rows: (sheet.rows || []).map((row) =>
            (row || []).map((cell) => {
              if (
                typeof cell === "object" &&
                cell !== null &&
                "value" in cell
              ) {
                return {
                  value: String(cell.value ?? ""),
                  style: cell.style || {},
                };
              }
              return { value: String(cell ?? ""), style: {} };
            }),
          ),
          merges: sheet.merges || [],
        })),
      }));
    } catch {
      return [];
    }
  });

  const [activeId, setActiveId] = useState<string>(
    () => documents[0]?.id || "",
  );

  const [platform, setPlatform] = useState<string>("");
  const [isMaximized, setIsMaximized] = useState<boolean>(false);

  useEffect(() => {
    if (window.electronAPI) {
      window.electronAPI.getPlatform?.().then((p) => {
        if (p) setPlatform(p);
      });
      window.electronAPI.isMaximized?.().then((max) => {
        setIsMaximized(max);
      });
      const cleanup = window.electronAPI.onMaximizedChange?.((max) => {
        setIsMaximized(max);
      });
      return () => cleanup?.();
    }
  }, []);

  useEffect(() => {
    const activeDocName = documents.find((d) => d.id === activeId)?.name;
    if (activeDocName) {
      document.title = `${activeDocName} — Paperclip`;
    } else {
      document.title = "Paperclip";
    }
  }, [activeId, documents]);

  const [selection, setSelection] = useState({
    startRow: 1,
    startCol: 0,
    endRow: 1,
    endCol: 0,
  });
  const [isDragging, setIsDragging] = useState(false);
  const [dragType, setDragType] = useState<"cell" | "row" | "column" | null>(
    null,
  );
  const [editing, setEditing] = useState(false);
  const [isFormulaBarFocused, setIsFormulaBarFocused] = useState(false);
  const [formulaPointReplacing, setFormulaPointReplacing] = useState(false);
  const [formulaPointDragStart, setFormulaPointDragStart] = useState<{
    row: number;
    col: number;
  } | null>(null);
  const [isFormulaDragging, setIsFormulaDragging] = useState(false);
  const [search, setSearch] = useState("");
  const [darkMode, setDarkMode] = useState<boolean>(() => {
    try {
      return localStorage.getItem("paperclip-theme") === "dark";
    } catch {
      return false;
    }
  });
  const [autosave, setAutosave] = useState(true);
  const [undoStack, setUndoStack] = useState<string[]>([]);
  const [redoStack, setRedoStack] = useState<string[]>([]);
  const [saveStatus, setSaveStatus] = useState<"saved" | "saving" | "unsaved">(
    "saved",
  );
  const [saveFolder, setSaveFolder] = useState("");
  const [contextMenu, setContextMenu] = useState<ContextMenuState>(null);
  const [showMergeDropdown, setShowMergeDropdown] = useState(false);

  const fileInput = useRef<HTMLInputElement>(null);
  const activeCellRef = useRef<HTMLInputElement>(null);
  const internalClipboard = useRef<{ cells: Cell[][]; text: string } | null>(
    null,
  );

  const activeDocument = documents.find((document) => document.id === activeId);
  const activeSheet = activeDocument?.sheets[activeDocument.activeSheet];
  const rows = useMemo(() => activeSheet?.rows || [], [activeSheet]);
  const merges = useMemo(() => activeSheet?.merges || [], [activeSheet]);
  const columnCount = Math.max(8, ...rows.map((row) => row.length));
  const query = search.trim().toLowerCase();

  const activeRow = selection.startRow;
  const activeCol = selection.startCol;
  const minRow = Math.min(selection.startRow, selection.endRow);
  const maxRow = Math.max(selection.startRow, selection.endRow);
  const minCol = Math.min(selection.startCol, selection.endCol);
  const maxCol = Math.max(selection.startCol, selection.endCol);

  const activeCell = rows[activeRow]?.[activeCol];
  const activeStyle = getCellStyle(activeCell);
  const activeRawValue = getCellValue(activeCell);

  const isFormulaMode =
    activeRawValue.startsWith("=") &&
    (editing || isFormulaBarFocused || activeRawValue === "=");

  const formulaHighlights = useMemo(() => {
    if (activeRawValue.startsWith("=")) {
      return extractFormulaHighlights(activeRawValue);
    }
    return [];
  }, [activeRawValue]);

  const evaluatedSheet = useMemo(() => {
    const getVal = (r: number, c: number): FormulaValue => {
      const cell = rows[r]?.[c];
      return getCellValue(cell);
    };

    const matrix: string[][] = [];
    for (let r = 0; r < rows.length; r++) {
      const rowVals: string[] = [];
      const rowCells = rows[r] || [];
      for (let c = 0; c < columnCount; c++) {
        const raw = getCellValue(rowCells[c]);
        if (raw.startsWith("=")) {
          const evalRes = evaluateFormula(raw, getVal);
          rowVals.push(
            evalRes === null || evalRes === undefined ? "" : String(evalRes),
          );
        } else {
          rowVals.push(raw);
        }
      }
      matrix.push(rowVals);
    }
    return matrix;
  }, [rows, columnCount]);

  const getMergeAt = useCallback(
    (r: number, c: number): MergeRange | null => {
      for (const m of merges) {
        const minR = Math.min(m.startRow, m.endRow);
        const maxR = Math.max(m.startRow, m.endRow);
        const minC = Math.min(m.startCol, m.endCol);
        const maxC = Math.max(m.startCol, m.endCol);
        if (r >= minR && r <= maxR && c >= minC && c <= maxC) {
          return { startRow: minR, startCol: minC, endRow: maxR, endCol: maxC };
        }
      }
      return null;
    },
    [merges],
  );

  const isSelectionMerged = useMemo(() => {
    return merges.some((m) => {
      const minR = Math.min(m.startRow, m.endRow);
      const maxR = Math.max(m.startRow, m.endRow);
      const minC = Math.min(m.startCol, m.endCol);
      const maxC = Math.max(m.startCol, m.endCol);
      return (
        Math.max(minRow, minR) <= Math.min(maxRow, maxR) &&
        Math.max(minCol, minC) <= Math.min(maxCol, maxC)
      );
    });
  }, [merges, minRow, maxRow, minCol, maxCol]);

  useEffect(() => {
    window.electronAPI?.getSaveFolder().then(setSaveFolder);
  }, []);

  useEffect(() => {
    document.documentElement.dataset.theme = darkMode ? "dark" : "light";
    try {
      localStorage.setItem("paperclip-theme", darkMode ? "dark" : "light");
    } catch {}
  }, [darkMode]);

  useEffect(() => {
    localStorage.setItem("paperclip-documents", JSON.stringify(documents));
  }, [documents]);

  const updateSheetData = useCallback(
    (nextRows: Cell[][], nextMerges?: MergeRange[]) => {
      if (!activeDocument) return;
      setUndoStack((stack) => [
        ...stack,
        JSON.stringify(activeDocument.sheets),
      ]);
      setRedoStack([]);
      setDocuments((current) =>
        current.map((document) =>
          document.id === activeId
            ? {
                ...document,
                dirty: true,
                sheets: document.sheets.map((sheet, index) =>
                  index === document.activeSheet
                    ? {
                        ...sheet,
                        rows: nextRows,
                        merges:
                          nextMerges !== undefined
                            ? nextMerges
                            : sheet.merges || [],
                      }
                    : sheet,
                ),
              }
            : document,
        ),
      );
      setSaveStatus("unsaved");
    },
    [activeDocument, activeId],
  );

  const updateRows = useCallback(
    (nextRows: Cell[][]) => {
      updateSheetData(nextRows, activeSheet?.merges);
    },
    [updateSheetData, activeSheet?.merges],
  );

  const updateCell = (value: string) => {
    const nextRows = rows.map((row) => [...row]);
    while (nextRows.length <= activeRow) nextRows.push([]);
    while (nextRows[activeRow].length <= activeCol)
      nextRows[activeRow].push({ value: "", style: {} });

    const currentCell = nextRows[activeRow][activeCol];
    const existingStyle = getCellStyle(currentCell);
    nextRows[activeRow][activeCol] = {
      value,
      style: existingStyle,
    };
    updateRows(nextRows);
  };

  const applyStyleToSelection = useCallback(
    (updater: (prev: CellStyle) => CellStyle) => {
      if (!activeDocument) return;
      const nextRows = rows.map((row) => [...row]);
      while (nextRows.length <= maxRow) nextRows.push([]);
      for (let r = minRow; r <= maxRow; r++) {
        while (nextRows[r].length <= maxCol)
          nextRows[r].push({ value: "", style: {} });
        for (let c = minCol; c <= maxCol; c++) {
          const current = nextRows[r][c];
          const curVal = getCellValue(current);
          const curStyle = getCellStyle(current);
          nextRows[r][c] = {
            value: curVal,
            style: updater(curStyle),
          };
        }
      }
      updateRows(nextRows);
    },
    [activeDocument, rows, minRow, maxRow, minCol, maxCol, updateRows],
  );

  const toggleBold = useCallback(() => {
    const nextBold = !activeStyle.bold;
    applyStyleToSelection((s) => ({ ...s, bold: nextBold }));
  }, [activeStyle.bold, applyStyleToSelection]);

  const toggleItalic = useCallback(() => {
    const nextItalic = !activeStyle.italic;
    applyStyleToSelection((s) => ({ ...s, italic: nextItalic }));
  }, [activeStyle.italic, applyStyleToSelection]);

  const toggleUnderline = useCallback(() => {
    const nextUnderline = !activeStyle.underline;
    applyStyleToSelection((s) => ({ ...s, underline: nextUnderline }));
  }, [activeStyle.underline, applyStyleToSelection]);

  const setFontFamily = (fontFamily: string) => {
    applyStyleToSelection((s) => ({ ...s, fontFamily }));
  };

  const setFontSize = (fontSize: number) => {
    applyStyleToSelection((s) => ({ ...s, fontSize }));
  };

  const setAlignment = (align: "left" | "center" | "right") => {
    applyStyleToSelection((s) => ({ ...s, align }));
  };

  // Merge functions
  const mergeCells = useCallback(
    (r1 = minRow, c1 = minCol, r2 = maxRow, c2 = maxCol) => {
      const sR = Math.min(r1, r2);
      const eR = Math.max(r1, r2);
      const sC = Math.min(c1, c2);
      const eC = Math.max(c1, c2);
      if (sR === eR && sC === eC) return;

      const nextRows = rows.map((row) => [...row]);
      while (nextRows.length <= eR) nextRows.push([]);
      for (let r = sR; r <= eR; r++) {
        while (nextRows[r].length <= eC) {
          nextRows[r].push({ value: "", style: {} });
        }
      }

      // If top-left cell is empty, preserve first non-empty value in the range
      const masterVal = getCellValue(nextRows[sR][sC]);
      if (!masterVal.trim()) {
        findVal: for (let r = sR; r <= eR; r++) {
          for (let c = sC; c <= eC; c++) {
            const v = getCellValue(nextRows[r][c]);
            if (v.trim()) {
              nextRows[sR][sC] = {
                value: v,
                style: getCellStyle(nextRows[r][c]),
              };
              break findVal;
            }
          }
        }
      }

      const currentMerges = activeSheet?.merges || [];
      const filtered = currentMerges.filter(
        (m) =>
          Math.max(sR, m.startRow) > Math.min(eR, m.endRow) ||
          Math.max(sC, m.startCol) > Math.min(eC, m.endCol),
      );
      const newMerges = [
        ...filtered,
        { startRow: sR, startCol: sC, endRow: eR, endCol: eC },
      ];
      updateSheetData(nextRows, newMerges);
      setSelection({
        startRow: sR,
        startCol: sC,
        endRow: eR,
        endCol: eC,
      });
    },
    [activeSheet?.merges, minRow, minCol, maxRow, maxCol, rows, updateSheetData],
  );

  const mergeHorizontally = useCallback(
    (r1 = minRow, c1 = minCol, r2 = maxRow, c2 = maxCol) => {
      const sR = Math.min(r1, r2);
      const eR = Math.max(r1, r2);
      const sC = Math.min(c1, c2);
      const eC = Math.max(c1, c2);
      if (sC === eC) return;

      const nextRows = rows.map((row) => [...row]);
      while (nextRows.length <= eR) nextRows.push([]);
      for (let r = sR; r <= eR; r++) {
        while (nextRows[r].length <= eC) {
          nextRows[r].push({ value: "", style: {} });
        }
        const masterVal = getCellValue(nextRows[r][sC]);
        if (!masterVal.trim()) {
          for (let c = sC + 1; c <= eC; c++) {
            const v = getCellValue(nextRows[r][c]);
            if (v.trim()) {
              nextRows[r][sC] = {
                value: v,
                style: getCellStyle(nextRows[r][c]),
              };
              break;
            }
          }
        }
      }

      const currentMerges = (activeSheet?.merges || []).filter(
        (m) =>
          Math.max(sR, m.startRow) > Math.min(eR, m.endRow) ||
          Math.max(sC, m.startCol) > Math.min(eC, m.endCol),
      );
      for (let r = sR; r <= eR; r++) {
        currentMerges.push({
          startRow: r,
          startCol: sC,
          endRow: r,
          endCol: eC,
        });
      }
      updateSheetData(nextRows, currentMerges);
      setSelection({
        startRow: sR,
        startCol: sC,
        endRow: eR,
        endCol: eC,
      });
    },
    [activeSheet?.merges, minRow, minCol, maxRow, maxCol, rows, updateSheetData],
  );

  const mergeVertically = useCallback(
    (r1 = minRow, c1 = minCol, r2 = maxRow, c2 = maxCol) => {
      const sR = Math.min(r1, r2);
      const eR = Math.max(r1, r2);
      const sC = Math.min(c1, c2);
      const eC = Math.max(c1, c2);
      if (sR === eR) return;

      const nextRows = rows.map((row) => [...row]);
      while (nextRows.length <= eR) nextRows.push([]);
      for (let r = sR; r <= eR; r++) {
        while (nextRows[r].length <= eC) {
          nextRows[r].push({ value: "", style: {} });
        }
      }

      for (let c = sC; c <= eC; c++) {
        const masterVal = getCellValue(nextRows[sR][c]);
        if (!masterVal.trim()) {
          for (let r = sR + 1; r <= eR; r++) {
            const v = getCellValue(nextRows[r][c]);
            if (v.trim()) {
              nextRows[sR][c] = {
                value: v,
                style: getCellStyle(nextRows[r][c]),
              };
              break;
            }
          }
        }
      }

      const currentMerges = (activeSheet?.merges || []).filter(
        (m) =>
          Math.max(sR, m.startRow) > Math.min(eR, m.endRow) ||
          Math.max(sC, m.startCol) > Math.min(eC, m.endCol),
      );
      for (let c = sC; c <= eC; c++) {
        currentMerges.push({
          startRow: sR,
          startCol: c,
          endRow: eR,
          endCol: c,
        });
      }
      updateSheetData(nextRows, currentMerges);
      setSelection({
        startRow: sR,
        startCol: sC,
        endRow: eR,
        endCol: eC,
      });
    },
    [activeSheet?.merges, minRow, minCol, maxRow, maxCol, rows, updateSheetData],
  );

  const unmergeCells = useCallback(
    (r1 = minRow, c1 = minCol, r2 = maxRow, c2 = maxCol) => {
      const sR = Math.min(r1, r2);
      const eR = Math.max(r1, r2);
      const sC = Math.min(c1, c2);
      const eC = Math.max(c1, c2);

      const currentMerges = activeSheet?.merges || [];
      const filtered = currentMerges.filter(
        (m) =>
          Math.max(sR, m.startRow) > Math.min(eR, m.endRow) ||
          Math.max(sC, m.startCol) > Math.min(eC, m.endCol),
      );
      updateSheetData(rows, filtered);
    },
    [activeSheet?.merges, minRow, minCol, maxRow, maxCol, rows, updateSheetData],
  );

  // Clipboard operations
  const copySelection = useCallback(async () => {
    const selectedCells: Cell[][] = [];
    const textRows: string[] = [];
    for (let r = minRow; r <= maxRow; r++) {
      const rowCells: Cell[] = [];
      const rowTexts: string[] = [];
      for (let c = minCol; c <= maxCol; c++) {
        const cell = rows[r]?.[c];
        const val = getCellValue(cell);
        const st = getCellStyle(cell);
        rowCells.push({ value: val, style: { ...st } });
        rowTexts.push(val);
      }
      selectedCells.push(rowCells);
      textRows.push(rowTexts.join("\t"));
    }
    const text = textRows.join("\n");
    internalClipboard.current = { cells: selectedCells, text };
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      // Ignored
    }
  }, [rows, minRow, maxRow, minCol, maxCol]);

  const clearSelectionContents = useCallback(() => {
    const nextRows = rows.map((row) => [...row]);
    for (let r = minRow; r <= maxRow; r++) {
      if (nextRows[r]) {
        for (let c = minCol; c <= maxCol; c++) {
          if (nextRows[r][c] !== undefined) {
            const st = getCellStyle(nextRows[r][c]);
            nextRows[r][c] = { value: "", style: st };
          }
        }
      }
    }
    updateRows(nextRows);
  }, [rows, minRow, maxRow, minCol, maxCol, updateRows]);

  const cutSelection = useCallback(async () => {
    await copySelection();
    clearSelectionContents();
  }, [copySelection, clearSelectionContents]);

  const pasteSelection = useCallback(
    async (valuesOnly = false) => {
      let cellsToPaste: Cell[][] = [];

      try {
        const text = await navigator.clipboard.readText();
        if (text) {
          if (
            internalClipboard.current &&
            internalClipboard.current.text === text &&
            !valuesOnly
          ) {
            cellsToPaste = internalClipboard.current.cells;
          } else {
            const lines = text.split(/\r?\n/).filter((l) => l.length > 0);
            cellsToPaste = lines.map((line) =>
              line.split("\t").map((val) => ({ value: val, style: {} })),
            );
          }
        }
      } catch {
        if (internalClipboard.current) {
          cellsToPaste = valuesOnly
            ? internalClipboard.current.cells.map((row) =>
                row.map((c) => ({ value: getCellValue(c), style: {} })),
              )
            : internalClipboard.current.cells;
        }
      }

      if (!cellsToPaste.length) return;

      const nextRows = rows.map((row) => [...row]);
      const pasteRowCount = cellsToPaste.length;
      const pasteColCount = Math.max(...cellsToPaste.map((r) => r.length));

      while (nextRows.length < activeRow + pasteRowCount) {
        nextRows.push(
          Array.from({ length: columnCount }, () => ({
            value: "",
            style: {},
          })),
        );
      }

      for (let r = 0; r < pasteRowCount; r++) {
        const targetR = activeRow + r;
        while (nextRows[targetR].length < activeCol + pasteColCount) {
          nextRows[targetR].push({ value: "", style: {} });
        }
        for (let c = 0; c < cellsToPaste[r].length; c++) {
          const targetC = activeCol + c;
          const sourceCell = cellsToPaste[r][c];
          const sourceVal = getCellValue(sourceCell);
          const sourceStyle = getCellStyle(sourceCell);

          if (valuesOnly) {
            const existingStyle = getCellStyle(nextRows[targetR][targetC]);
            nextRows[targetR][targetC] = {
              value: sourceVal,
              style: existingStyle,
            };
          } else {
            nextRows[targetR][targetC] = {
              value: sourceVal,
              style: sourceStyle,
            };
          }
        }
      }

      updateRows(nextRows);
      setSelection({
        startRow: activeRow,
        startCol: activeCol,
        endRow: activeRow + pasteRowCount - 1,
        endCol: activeCol + pasteColCount - 1,
      });
    },
    [activeRow, activeCol, columnCount, rows, updateRows],
  );

  // Group row operations
  const insertRowsAbove = useCallback(
    (targetRow: number, count = 1) => {
      const nextRows = rows.map((row) => [...row]);
      const newRows: Cell[][] = Array.from({ length: count }, () =>
        Array.from({ length: columnCount }, () => ({ value: "", style: {} })),
      );
      nextRows.splice(targetRow, 0, ...newRows);
      const nextMerges = adjustMergesOnRowInsert(
        activeSheet?.merges || [],
        targetRow,
        count,
      );
      updateSheetData(nextRows, nextMerges);
      setSelection({
        startRow: targetRow,
        startCol: 0,
        endRow: targetRow + count - 1,
        endCol: columnCount - 1,
      });
    },
    [rows, columnCount, activeSheet?.merges, updateSheetData],
  );

  const insertRowsBelow = useCallback(
    (targetRow: number, count = 1) => {
      const insertAt = targetRow + 1;
      const nextRows = rows.map((row) => [...row]);
      const newRows: Cell[][] = Array.from({ length: count }, () =>
        Array.from({ length: columnCount }, () => ({ value: "", style: {} })),
      );
      nextRows.splice(insertAt, 0, ...newRows);
      const nextMerges = adjustMergesOnRowInsert(
        activeSheet?.merges || [],
        insertAt,
        count,
      );
      updateSheetData(nextRows, nextMerges);
      setSelection({
        startRow: insertAt,
        startCol: 0,
        endRow: insertAt + count - 1,
        endCol: columnCount - 1,
      });
    },
    [rows, columnCount, activeSheet?.merges, updateSheetData],
  );

  const deleteRows = useCallback(
    (fromRow: number, count = 1) => {
      let nextRows = rows.map((row) => [...row]);
      if (nextRows.length <= count) {
        nextRows = [
          Array.from({ length: columnCount }, () => ({
            value: "",
            style: {},
          })),
        ];
      } else {
        nextRows.splice(fromRow, count);
      }
      const nextMerges = adjustMergesOnRowDelete(
        activeSheet?.merges || [],
        fromRow,
        fromRow + count - 1,
      );
      updateSheetData(nextRows, nextMerges);
      const newActiveRow = Math.min(fromRow, nextRows.length - 1);
      setSelection({
        startRow: newActiveRow,
        startCol: 0,
        endRow: newActiveRow,
        endCol: columnCount - 1,
      });
    },
    [rows, columnCount, activeSheet?.merges, updateSheetData],
  );

  const clearRows = useCallback(
    (fromRow: number, count = 1) => {
      const nextRows = rows.map((row) => [...row]);
      for (let r = fromRow; r < fromRow + count && r < nextRows.length; r++) {
        nextRows[r] = nextRows[r].map((cell) => ({
          value: "",
          style: getCellStyle(cell),
        }));
      }
      updateRows(nextRows);
    },
    [rows, updateRows],
  );

  // Group column operations
  const insertColumnsLeft = useCallback(
    (targetCol: number, count = 1) => {
      const nextRows = rows.map((row) => {
        const nextRow = [...row];
        while (nextRow.length < targetCol) {
          nextRow.push({ value: "", style: {} });
        }
        const newCells = Array.from({ length: count }, () => ({
          value: "",
          style: {},
        }));
        nextRow.splice(targetCol, 0, ...newCells);
        return nextRow;
      });
      const nextMerges = adjustMergesOnColInsert(
        activeSheet?.merges || [],
        targetCol,
        count,
      );
      updateSheetData(nextRows, nextMerges);
      setSelection({
        startRow: 0,
        startCol: targetCol,
        endRow: Math.max(0, nextRows.length - 1),
        endCol: targetCol + count - 1,
      });
    },
    [rows, activeSheet?.merges, updateSheetData],
  );

  const insertColumnsRight = useCallback(
    (targetCol: number, count = 1) => {
      const insertAt = targetCol + 1;
      const nextRows = rows.map((row) => {
        const nextRow = [...row];
        while (nextRow.length < insertAt) {
          nextRow.push({ value: "", style: {} });
        }
        const newCells = Array.from({ length: count }, () => ({
          value: "",
          style: {},
        }));
        nextRow.splice(insertAt, 0, ...newCells);
        return nextRow;
      });
      const nextMerges = adjustMergesOnColInsert(
        activeSheet?.merges || [],
        insertAt,
        count,
      );
      updateSheetData(nextRows, nextMerges);
      setSelection({
        startRow: 0,
        startCol: insertAt,
        endRow: Math.max(0, nextRows.length - 1),
        endCol: insertAt + count - 1,
      });
    },
    [rows, activeSheet?.merges, updateSheetData],
  );

  const deleteColumns = useCallback(
    (fromCol: number, count = 1) => {
      const nextRows = rows.map((row) => {
        const nextRow = [...row];
        if (nextRow.length > fromCol) {
          nextRow.splice(fromCol, count);
        }
        return nextRow;
      });
      const nextMerges = adjustMergesOnColDelete(
        activeSheet?.merges || [],
        fromCol,
        fromCol + count - 1,
      );
      updateSheetData(nextRows, nextMerges);
      const newColCount = Math.max(1, columnCount - count);
      const newActiveCol = Math.min(fromCol, newColCount - 1);
      setSelection({
        startRow: 0,
        startCol: newActiveCol,
        endRow: Math.max(0, nextRows.length - 1),
        endCol: newActiveCol,
      });
    },
    [rows, columnCount, activeSheet?.merges, updateSheetData],
  );

  const clearColumns = useCallback(
    (fromCol: number, count = 1) => {
      const nextRows = rows.map((row) => {
        const nextRow = [...row];
        for (let c = fromCol; c < fromCol + count && c < nextRow.length; c++) {
          nextRow[c] = {
            value: "",
            style: getCellStyle(nextRow[c]),
          };
        }
        return nextRow;
      });
      updateRows(nextRows);
    },
    [rows, updateRows],
  );

  const sortColumn = useCallback(
    (targetCol: number, descending = false) => {
      if (rows.length > 1) {
        updateRows([
          rows[0],
          ...rows.slice(1).sort((a, b) => {
            const valA = getCellValue(a[targetCol]);
            const valB = getCellValue(b[targetCol]);
            const numA = Number(valA);
            const numB = Number(valB);
            if (!isNaN(numA) && !isNaN(numB) && valA !== "" && valB !== "") {
              return (numA - numB) * (descending ? -1 : 1);
            }
            return (
              valA.localeCompare(valB, undefined, { numeric: true }) *
              (descending ? -1 : 1)
            );
          }),
        ]);
      }
    },
    [rows, updateRows],
  );

  // Sheet operations
  const renameSheet = useCallback(
    (sheetIndex: number) => {
      const currentName =
        activeDocument?.sheets[sheetIndex]?.name || `Sheet ${sheetIndex + 1}`;
      const newName = window.prompt("Enter new sheet name:", currentName);
      if (!newName || !newName.trim() || newName === currentName) return;
      setDocuments((current) =>
        current.map((doc) =>
          doc.id === activeId
            ? {
                ...doc,
                dirty: true,
                sheets: doc.sheets.map((s, idx) =>
                  idx === sheetIndex ? { ...s, name: newName.trim() } : s,
                ),
              }
            : doc,
        ),
      );
    },
    [activeDocument, activeId],
  );

  const duplicateSheet = useCallback(
    (sheetIndex: number) => {
      if (!activeDocument) return;
      const sourceSheet = activeDocument.sheets[sheetIndex];
      if (!sourceSheet) return;
      const clonedRows = sourceSheet.rows.map((row) =>
        row.map((cell) => ({
          value: getCellValue(cell),
          style: { ...getCellStyle(cell) },
        })),
      );
      const clonedMerges = (sourceSheet.merges || []).map((m) => ({ ...m }));
      const newSheetName = `Copy of ${sourceSheet.name}`;
      setDocuments((current) =>
        current.map((doc) =>
          doc.id === activeId
            ? {
                ...doc,
                dirty: true,
                activeSheet: doc.sheets.length,
                sheets: [
                  ...doc.sheets,
                  {
                    name: newSheetName,
                    rows: clonedRows,
                    merges: clonedMerges,
                  },
                ],
              }
            : doc,
        ),
      );
    },
    [activeDocument, activeId],
  );

  const deleteSheet = useCallback(
    (sheetIndex: number) => {
      if (!activeDocument || activeDocument.sheets.length <= 1) return;
      const targetName = activeDocument.sheets[sheetIndex]?.name || "Sheet";
      if (!window.confirm(`Delete sheet "${targetName}"?`)) return;
      setDocuments((current) =>
        current.map((doc) => {
          if (doc.id !== activeId) return doc;
          const newSheets = doc.sheets.filter((_, idx) => idx !== sheetIndex);
          const nextActive = Math.min(doc.activeSheet, newSheets.length - 1);
          return {
            ...doc,
            dirty: true,
            activeSheet: nextActive,
            sheets: newSheets,
          };
        }),
      );
    },
    [activeDocument, activeId],
  );

  // Context menu open & close
  const openContextMenu = (
    event: React.MouseEvent,
    target: ContextMenuTarget,
  ) => {
    event.preventDefault();
    event.stopPropagation();
    if (showMergeDropdown) setShowMergeDropdown(false);

    const isFullRowsSelected =
      minCol === 0 &&
      maxCol >= columnCount - 1;
    const isFullColsSelected =
      minRow === 0 &&
      maxRow >= Math.max(0, rows.length - 1);

    if (target.type === "row") {
      const isInsideSelectedRows =
        isFullRowsSelected &&
        target.rowIndex >= minRow &&
        target.rowIndex <= maxRow;
      if (!isInsideSelectedRows) {
        setSelection({
          startRow: target.rowIndex,
          startCol: 0,
          endRow: target.rowIndex,
          endCol: Math.max(0, columnCount - 1),
        });
      }
    } else if (target.type === "column") {
      const isInsideSelectedCols =
        isFullColsSelected &&
        target.columnIndex >= minCol &&
        target.columnIndex <= maxCol;
      if (!isInsideSelectedCols) {
        setSelection({
          startRow: 0,
          startCol: target.columnIndex,
          endRow: Math.max(0, rows.length - 1),
          endCol: target.columnIndex,
        });
      }
    } else if (target.type === "cell") {
      const isInsideCurrent =
        target.rowIndex >= minRow &&
        target.rowIndex <= maxRow &&
        target.columnIndex >= minCol &&
        target.columnIndex <= maxCol;
      if (!isInsideCurrent) {
        const merge = getMergeAt(target.rowIndex, target.columnIndex);
        if (merge) {
          setSelection({
            startRow: merge.startRow,
            startCol: merge.startCol,
            endRow: merge.endRow,
            endCol: merge.endCol,
          });
        } else {
          setSelection({
            startRow: target.rowIndex,
            startCol: target.columnIndex,
            endRow: target.rowIndex,
            endCol: target.columnIndex,
          });
        }
      }
    }

    const menuWidth = 260;
    const menuHeight = 460;
    const x = Math.min(event.clientX, window.innerWidth - menuWidth - 10);
    const y = Math.min(event.clientY, window.innerHeight - menuHeight - 10);

    setContextMenu({
      visible: true,
      x: Math.max(10, x),
      y: Math.max(10, y),
      target,
    });
  };

  const closeContextMenu = useCallback(() => {
    setContextMenu(null);
  }, []);

  useEffect(() => {
    if (!contextMenu) return;
    const onPointerDown = (e: MouseEvent) => {
      const el = document.getElementById("spreadsheet-context-menu");
      if (el && !el.contains(e.target as Node)) {
        closeContextMenu();
      }
    };
    const onWheel = () => closeContextMenu();
    window.addEventListener("pointerdown", onPointerDown);
    window.addEventListener("wheel", onWheel, { passive: true });
    return () => {
      window.removeEventListener("pointerdown", onPointerDown);
      window.removeEventListener("wheel", onWheel);
    };
  }, [contextMenu, closeContextMenu]);

  useEffect(() => {
    if (!showMergeDropdown) return;
    const onPointerDown = (e: MouseEvent) => {
      const el = document.querySelector(".merge-group");
      if (el && !el.contains(e.target as Node)) {
        setShowMergeDropdown(false);
      }
    };
    window.addEventListener("pointerdown", onPointerDown);
    return () => window.removeEventListener("pointerdown", onPointerDown);
  }, [showMergeDropdown]);

  const handleCellMouseDown = (
    r: number,
    c: number,
    event: React.MouseEvent,
  ) => {
    if (event.button !== 0 || event.ctrlKey) return;
    if (contextMenu) closeContextMenu();
    if (showMergeDropdown) setShowMergeDropdown(false);
    const merge = getMergeAt(r, c);
    if (merge && !event.shiftKey) {
      setSelection({
        startRow: merge.startRow,
        startCol: merge.startCol,
        endRow: merge.endRow,
        endCol: merge.endCol,
      });
      setDragType("cell");
      setIsDragging(true);
      return;
    }
    if (event.shiftKey) {
      setSelection((prev) =>
        expandRangeForMerges(prev.startRow, prev.startCol, r, c, merges),
      );
    } else {
      setSelection({ startRow: r, startCol: c, endRow: r, endCol: c });
      setDragType("cell");
      setIsDragging(true);
    }
  };

  const handleCellMouseEnter = (r: number, c: number) => {
    if (!isDragging) return;
    window.getSelection()?.removeAllRanges();
    if (dragType === "cell") {
      setSelection((prev) =>
        expandRangeForMerges(prev.startRow, prev.startCol, r, c, merges),
      );
    } else if (dragType === "row") {
      setSelection((prev) => ({
        ...prev,
        endRow: r,
        startCol: 0,
        endCol: Math.max(0, columnCount - 1),
      }));
    } else if (dragType === "column") {
      setSelection((prev) => ({
        ...prev,
        startRow: 0,
        endRow: Math.max(0, rows.length - 1),
        endCol: c,
      }));
    }
  };

  const handleColumnHeaderMouseDown = (
    c: number,
    event: React.MouseEvent,
  ) => {
    if (event.button !== 0 || event.ctrlKey) return;
    if (contextMenu) closeContextMenu();
    if (showMergeDropdown) setShowMergeDropdown(false);
    const maxR = Math.max(0, rows.length - 1);
    if (event.shiftKey) {
      setSelection((prev) => ({
        ...prev,
        startRow: 0,
        endRow: maxR,
        endCol: c,
      }));
    } else {
      setSelection({
        startRow: 0,
        startCol: c,
        endRow: maxR,
        endCol: c,
      });
      setDragType("column");
      setIsDragging(true);
    }
  };

  const handleColumnHeaderMouseEnter = (c: number) => {
    if (!isDragging) return;
    window.getSelection()?.removeAllRanges();
    if (dragType === "column" || dragType === "cell") {
      setSelection((prev) => ({
        ...prev,
        endCol: c,
        startRow: 0,
        endRow: Math.max(0, rows.length - 1),
      }));
    }
  };

  const handleRowHeaderMouseDown = (r: number, event: React.MouseEvent) => {
    if (event.button !== 0 || event.ctrlKey) return;
    if (contextMenu) closeContextMenu();
    if (showMergeDropdown) setShowMergeDropdown(false);
    const maxC = Math.max(0, columnCount - 1);
    if (event.shiftKey) {
      setSelection((prev) => ({
        ...prev,
        startCol: 0,
        endCol: maxC,
        endRow: r,
      }));
    } else {
      setSelection({
        startRow: r,
        startCol: 0,
        endRow: r,
        endCol: maxC,
      });
      setDragType("row");
      setIsDragging(true);
    }
  };

  const handleRowHeaderMouseEnter = (r: number) => {
    if (!isDragging) return;
    window.getSelection()?.removeAllRanges();
    if (dragType === "row" || dragType === "cell") {
      setSelection((prev) => ({
        ...prev,
        endRow: r,
        startCol: 0,
        endCol: Math.max(0, columnCount - 1),
      }));
    }
  };

  useEffect(() => {
    const onMouseUp = () => {
      setIsDragging(false);
      setDragType(null);
      setIsFormulaDragging(false);
      setFormulaPointDragStart(null);
    };
    window.addEventListener("mouseup", onMouseUp);
    return () => window.removeEventListener("mouseup", onMouseUp);
  }, []);

  const handleCornerClick = useCallback(() => {
    if (contextMenu) closeContextMenu();
    setSelection({
      startRow: 0,
      startCol: 0,
      endRow: Math.max(0, rows.length - 1),
      endCol: Math.max(0, columnCount - 1),
    });
  }, [contextMenu, closeContextMenu, rows.length, columnCount]);

  const moveSelection = (row: number, column: number) => {
    const nextRow = Math.max(0, row);
    const nextCol = Math.max(0, column);
    setSelection({
      startRow: nextRow,
      startCol: nextCol,
      endRow: nextRow,
      endCol: nextCol,
    });
    setTimeout(() => {
      const el = document.querySelector<HTMLInputElement>(
        `input.cell[data-row="${nextRow}"][data-col="${nextCol}"]`,
      );
      el?.focus();
    }, 0);
  };

  const createDocument = () => {
    const document = newDocument();
    setDocuments((current) => [...current, document]);
    setActiveId(document.id);
    setSelection({ startRow: 1, startCol: 0, endRow: 1, endCol: 0 });
  };

  const removeDocument = (id: string) => {
    const remaining = documents.filter((document) => document.id !== id);
    setDocuments(remaining);
    if (id === activeId) setActiveId(remaining[0]?.id || "");
  };

  const download = (content: BlobPart, filename: string, type: string) => {
    const url = URL.createObjectURL(new Blob([content], { type }));
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  };

  const saveDocument = useCallback(async () => {
    if (!activeDocument || !window.electronAPI) return;
    setSaveStatus("saving");
    try {
      const result = await window.electronAPI.saveFile({
        filePath: activeDocument.path,
        name: `${activeDocument.name}.csv`,
        content: toCsv(activeDocument.sheets[activeDocument.activeSheet]),
      });
      if (!result) {
        setSaveStatus("unsaved");
        return;
      }
      setDocuments((current) =>
        current.map((document) =>
          document.id === activeId
            ? {
                ...document,
                path: result.filePath,
                name: result.name.replace(/\.[^.]+$/i, ""),
                dirty: false,
              }
            : document,
        ),
      );
      setSaveStatus("saved");
    } catch {
      setSaveStatus("unsaved");
    }
  }, [activeDocument, activeId]);

  useEffect(() => {
    if (!activeDocument?.dirty || !autosave) return;
    setSaveStatus("saving");
    const timer = window.setTimeout(() => {
      void saveDocument();
    }, 700);
    return () => window.clearTimeout(timer);
  }, [activeDocument?.dirty, autosave, saveDocument]);

  const exportCsv = () => {
    if (!activeSheet) return;
    download(
      toCsv(activeSheet),
      `${activeDocument?.name || "spreadsheet"}.csv`,
      "text/csv",
    );
  };

  const exportExcel = () => {
    if (!activeDocument) return;
    const wb = XLSX.utils.book_new();
    activeDocument.sheets.forEach((sheet) => {
      const data = sheet.rows.map((row) =>
        row.map((cell) => getCellValue(cell)),
      );
      const ws = XLSX.utils.aoa_to_sheet(data);
      if (sheet.merges && sheet.merges.length > 0) {
        ws["!merges"] = sheet.merges.map((m) => ({
          s: { r: m.startRow, c: m.startCol },
          e: { r: m.endRow, c: m.endCol },
        }));
      }
      XLSX.utils.book_append_sheet(wb, ws, sheet.name);
    });
    XLSX.writeFile(wb, `${activeDocument.name || "spreadsheet"}.xlsx`);
  };

  const parseCsv = (text: string, filePath?: string, fileName?: string) => {
    const parsed = text
      .split(/\r?\n/)
      .filter(Boolean)
      .map((line) =>
        line
          .split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/)
          .map((cell) => ({
            value: cell.replace(/^"|"$/g, "").replace(/""/g, '"'),
            style: {},
          })),
      );
    const document: Document = {
      ...newDocument(),
      name: (fileName || "Untitled Spreadsheet").replace(/\.[^.]+$/, ""),
      path: filePath,
      sheets: [
        {
          name: "Sheet 1",
          rows:
            parsed.length > 0
              ? parsed
              : starterRows.map((row) =>
                  row.map((cell) =>
                    typeof cell === "object"
                      ? { value: cell.value, style: { ...cell.style } }
                      : { value: cell, style: {} },
                  ),
                ),
          merges: [],
        },
      ],
      dirty: false,
    };
    setDocuments((current) => [...current, document]);
    setActiveId(document.id);
    setSelection({ startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
  };

  const parseExcel = (
    data: ArrayBuffer | string,
    filePath?: string,
    fileName?: string,
    isBase64 = false,
  ) => {
    try {
      const workbook = isBase64
        ? XLSX.read(data, { type: "base64" })
        : XLSX.read(data, { type: "array" });

      const sheets: Sheet[] = workbook.SheetNames.map((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const rawRows = XLSX.utils.sheet_to_json<
          (string | number | boolean | null | undefined)[]
        >(worksheet, {
          header: 1,
          defval: "",
        });
        const parsedRows: Cell[][] = rawRows.map((row) =>
          row.map((val) => ({
            value: val !== null && val !== undefined ? String(val) : "",
            style: {},
          })),
        );
        const parsedMerges: MergeRange[] = (worksheet["!merges"] || []).map(
          (m) => ({
            startRow: m.s.r,
            startCol: m.s.c,
            endRow: m.e.r,
            endCol: m.e.c,
          }),
        );
        return {
          name: sheetName,
          rows:
            parsedRows.length > 0
              ? parsedRows
              : starterRows.map((row) =>
                  row.map((cell) =>
                    typeof cell === "object"
                      ? { value: cell.value, style: { ...cell.style } }
                      : { value: cell, style: {} },
                  ),
                ),
          merges: parsedMerges,
        };
      });

      const document: Document = {
        ...newDocument(),
        name: (fileName || "Untitled Spreadsheet").replace(/\.[^.]+$/, ""),
        path: filePath,
        sheets:
          sheets.length > 0
            ? sheets
            : [
                {
                  name: "Sheet 1",
                  rows: starterRows.map((row) =>
                    row.map((cell) =>
                      typeof cell === "object"
                        ? { value: cell.value, style: { ...cell.style } }
                        : { value: cell, style: {} },
                    ),
                  ),
                  merges: [],
                },
              ],
        activeSheet: 0,
        dirty: false,
      };
      setDocuments((current) => [...current, document]);
      setActiveId(document.id);
      setSelection({ startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
    } catch (err) {
      console.error("Failed to parse Excel file", err);
    }
  };

  const handleFileInput = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const isExcel = /\.(xlsx|xls)$/i.test(file.name);
    if (isExcel) {
      const reader = new FileReader();
      reader.onload = () => {
        if (reader.result instanceof ArrayBuffer) {
          parseExcel(reader.result, undefined, file.name, false);
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      const reader = new FileReader();
      reader.onload = () => {
        parseCsv(String(reader.result || ""), undefined, file.name);
      };
      reader.readAsText(file);
    }
    event.target.value = "";
  };

  const openFile = async () => {
    const result = await window.electronAPI?.openFile();
    if (result) {
      if (result.isExcel && result.base64) {
        parseExcel(result.base64, result.filePath, result.name, true);
      } else if (result.content) {
        parseCsv(result.content, result.filePath, result.name);
      }
    } else if (!window.electronAPI) {
      fileInput.current?.click();
    }
  };

  const handleUndo = useCallback(() => {
    const previous = undoStack.at(-1);
    if (previous && activeDocument) {
      setRedoStack((stack) => [
        ...stack,
        JSON.stringify(activeDocument.sheets),
      ]);
      setUndoStack((stack) => stack.slice(0, -1));
      const sheets = JSON.parse(previous);
      setDocuments((current) =>
        current.map((item) =>
          item.id === activeId ? { ...item, sheets, dirty: true } : item,
        ),
      );
    }
  }, [undoStack, activeDocument, activeId]);

  const handleRedo = useCallback(() => {
    const next = redoStack.at(-1);
    if (next && activeDocument) {
      setUndoStack((stack) => [
        ...stack,
        JSON.stringify(activeDocument.sheets),
      ]);
      setRedoStack((stack) => stack.slice(0, -1));
      const sheets = JSON.parse(next);
      setDocuments((current) =>
        current.map((item) =>
          item.id === activeId ? { ...item, sheets, dirty: true } : item,
        ),
      );
    }
  }, [redoStack, activeDocument, activeId]);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      const command = event.metaKey || event.ctrlKey;

      if (event.key === "Escape") {
        if (showMergeDropdown) setShowMergeDropdown(false);
        if (contextMenu) {
          event.preventDefault();
          closeContextMenu();
          return;
        }
      }

      if (command && event.key.toLowerCase() === "s") {
        event.preventDefault();
        if (activeDocument) void saveDocument();
        return;
      }
      if (command && event.key.toLowerCase() === "f") {
        event.preventDefault();
        document.getElementById("search-input")?.focus();
        return;
      }
      if (command && event.key.toLowerCase() === "z") {
        event.preventDefault();
        if (event.shiftKey) {
          handleRedo();
        } else {
          handleUndo();
        }
        return;
      }
      if (command && event.key.toLowerCase() === "y") {
        event.preventDefault();
        handleRedo();
        return;
      }
      if (command && event.key.toLowerCase() === "b") {
        event.preventDefault();
        toggleBold();
        return;
      }
      if (command && event.key.toLowerCase() === "i") {
        event.preventDefault();
        toggleItalic();
        return;
      }
      if (command && event.key.toLowerCase() === "u") {
        event.preventDefault();
        toggleUnderline();
        return;
      }
      if (command && event.key.toLowerCase() === "c" && !editing) {
        event.preventDefault();
        void copySelection();
        return;
      }
      if (command && event.key.toLowerCase() === "x" && !editing) {
        event.preventDefault();
        void cutSelection();
        return;
      }
      if (command && event.key.toLowerCase() === "v") {
        if (event.shiftKey) {
          event.preventDefault();
          void pasteSelection(true);
        } else if (!editing) {
          event.preventDefault();
          void pasteSelection(false);
        }
        return;
      }
      if (command && event.key.toLowerCase() === "a" && !editing) {
        event.preventDefault();
        handleCornerClick();
        return;
      }

      if (
        (event.key === "Delete" || event.key === "Backspace") &&
        !editing &&
        document.activeElement?.id !== "search-input" &&
        document.activeElement?.getAttribute("aria-label") !== "Formula bar"
      ) {
        event.preventDefault();
        clearSelectionContents();
        return;
      }

      if (
        !editing &&
        ["ArrowDown", "ArrowUp", "ArrowLeft", "ArrowRight"].includes(event.key)
      ) {
        event.preventDefault();
        const delta = {
          ArrowDown: [1, 0],
          ArrowUp: [-1, 0],
          ArrowLeft: [0, -1],
          ArrowRight: [0, 1],
        }[event.key] || [0, 0];
        if (event.shiftKey) {
          setSelection((prev) => ({
            ...prev,
            endRow: Math.max(0, prev.endRow + delta[0]),
            endCol: Math.max(0, prev.endCol + delta[1]),
          }));
        } else {
          moveSelection(activeRow + delta[0], activeCol + delta[1]);
        }
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [
    activeDocument,
    activeId,
    editing,
    activeRow,
    activeCol,
    undoStack,
    redoStack,
    activeStyle,
    contextMenu,
    showMergeDropdown,
    handleUndo,
    handleRedo,
    toggleBold,
    toggleItalic,
    toggleUnderline,
    copySelection,
    cutSelection,
    pasteSelection,
    clearSelectionContents,
    handleCornerClick,
    closeContextMenu,
    saveDocument,
  ]);

  const selectionLabel = useMemo(() => {
    const colStart = colIndexToLetter(minCol);
    const rowStart = minRow + 1;
    if (minRow === maxRow && minCol === maxCol) {
      return `${colStart}${rowStart}`;
    }
    const colEnd = colIndexToLetter(maxCol);
    const rowEnd = maxRow + 1;
    return `${colStart}${rowStart}:${colEnd}${rowEnd}`;
  }, [minRow, maxRow, minCol, maxCol]);

  const visibleRows = useMemo(
    () =>
      rows
        .map((row, rowIndex) => ({ row, rowIndex }))
        .filter(
          ({ row }) =>
            !query ||
            row.some((cell) =>
              getCellValue(cell).toLowerCase().includes(query),
            ),
        ),
    [rows, query],
  );

  const selectedRowCount = maxRow - minRow + 1;
  const selectedColCount = maxCol - minCol + 1;

  return (
    <main className={`app-shell ${!activeDocument ? "is-home-view" : ""}`}>
      <header
        className={`topbar ${platform === "darwin" ? "platform-mac" : platform ? "platform-win" : ""} ${!activeDocument ? "is-home" : ""}`}
      >
        <div className="brand">
          <span className="brand-mark">
            <GridbookIcon size={13} />
          </span>
          <span>Paperclip</span>
          <button
            className="save-location"
            title={`Save new files to ${saveFolder || "Desktop"}`}
            onClick={async () => {
              const folder = await window.electronAPI?.chooseSaveFolder();
              if (folder) setSaveFolder(folder);
            }}
          >
            {saveFolder
              ? saveFolder.split("/").pop()?.toUpperCase()
              : "DESKTOP"}
          </button>
        </div>
        <div className="document-tabs">
          {documents.map((document) => (
            <button
              className={`document-tab ${document.id === activeId ? "active" : ""}`}
              key={document.id}
              onClick={() => setActiveId(document.id)}
            >
              <span>{document.name}</span>
              {document.dirty && <i>•</i>}
              <b
                onClick={(event) => {
                  event.stopPropagation();
                  removeDocument(document.id);
                }}
              >
                ×
              </b>
            </button>
          ))}
          <button className="add-tab" onClick={createDocument} title="New spreadsheet (⌘N)">
            ＋
          </button>
        </div>
        <div className="top-actions">
          <button
            className="icon-button"
            title="Toggle theme"
            onClick={() => setDarkMode(!darkMode)}
            aria-label="Toggle theme"
          >
            ☼
          </button>

          {/* Windows / Linux Custom Title Bar Controls */}
          {platform && platform !== "darwin" && (
            <div className="window-controls">
              <button
                type="button"
                className="win-control-button win-minimize"
                title="Minimize"
                onClick={() => window.electronAPI?.minimize()}
              >
                <svg width="10" height="1" viewBox="0 0 10 1">
                  <line x1="0" y1="0.5" x2="10" y2="0.5" stroke="currentColor" strokeWidth="1" />
                </svg>
              </button>
              <button
                type="button"
                className="win-control-button win-maximize"
                title={isMaximized ? "Restore" : "Maximize"}
                onClick={() => window.electronAPI?.maximize()}
              >
                {isMaximized ? (
                  <svg width="10" height="10" viewBox="0 0 10 10">
                    <path
                      d="M2.5 1.5H8.5V7.5M1.5 2.5H7.5V8.5H1.5V2.5Z"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="1"
                    />
                  </svg>
                ) : (
                  <svg width="10" height="10" viewBox="0 0 10 10">
                    <rect
                      x="1"
                      y="1"
                      width="8"
                      height="8"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="1"
                    />
                  </svg>
                )}
              </button>
              <button
                type="button"
                className="win-control-button win-close"
                title="Close"
                onClick={() => window.electronAPI?.close()}
              >
                <svg width="10" height="10" viewBox="0 0 10 10">
                  <line x1="1" y1="1" x2="9" y2="9" stroke="currentColor" strokeWidth="1.2" />
                  <line x1="9" y1="1" x2="1" y2="9" stroke="currentColor" strokeWidth="1.2" />
                </svg>
              </button>
            </div>
          )}
        </div>
      </header>

      {!activeDocument ? (
        <section className="empty-state">
          <div className="empty-mark">
            <GridbookIcon size={28} />
          </div>
          <p className="eyebrow">PAPERCLIP / WORKSPACE</p>
          <h1>
            Your spreadsheets,
            <br />
            <em>beautifully simple.</em>
          </h1>
          <p className="empty-copy">
            A calm, capable workspace for the data that moves your day.
          </p>
          <div className="empty-actions">
            <button className="primary-button" onClick={createDocument}>
              ＋ New spreadsheet
            </button>
            <button className="secondary-button" onClick={openFile}>
              Open a file
            </button>
          </div>
          <p className="empty-hint">
            CSV and Excel files are supported with full rich formatting...
          </p>
          <input
            ref={fileInput}
            type="file"
            accept=".csv,.xlsx,.xls"
            hidden
            onChange={handleFileInput}
          />
        </section>
      ) : (
        <>
      <section className="toolbar">
        <div className="toolbar-group">
          <button className="tool-button emphasized" onClick={createDocument}>
            ＋ <span>New</span>
          </button>
          <button className="tool-button" onClick={openFile}>
            ↥ <span>Open</span>
          </button>
          <button className="tool-button" onClick={() => void saveDocument()}>
            ↓ <span>Save</span>
          </button>
        </div>
        <div className="toolbar-divider" />
        <div className="toolbar-group">
          <button
            className="tool-button"
            disabled={!undoStack.length}
            onClick={handleUndo}
            title="Undo (⌘Z)"
          >
            ↶
          </button>
          <button
            className="tool-button"
            disabled={!redoStack.length}
            onClick={handleRedo}
            title="Redo (⌘⇧Z or ⌘Y)"
          >
            ↷
          </button>
        </div>
        <div className="toolbar-divider" />
        <div className="toolbar-group formatting">
          <select
            aria-label="Font"
            value={activeStyle.fontFamily || "Default"}
            onChange={(event) => setFontFamily(event.target.value)}
          >
            {FONT_OPTIONS.map((font) => (
              <option key={font} value={font}>
                {font}
              </option>
            ))}
          </select>
          <select
            aria-label="Font size"
            value={activeStyle.fontSize || 11}
            onChange={(event) => setFontSize(Number(event.target.value))}
          >
            {FONT_SIZES.map((size) => (
              <option key={size} value={size}>
                {size}
              </option>
            ))}
          </select>
          <button
            type="button"
            className={`format-button ${activeStyle.bold ? "active" : ""}`}
            onClick={toggleBold}
            onMouseDown={(e) => e.preventDefault()}
            title="Bold (⌘B)"
            aria-label="Bold"
          >
            <strong>B</strong>
          </button>
          <button
            type="button"
            className={`format-button ${activeStyle.italic ? "active" : ""}`}
            onClick={toggleItalic}
            onMouseDown={(e) => e.preventDefault()}
            title="Italic (⌘I)"
            aria-label="Italic"
          >
            <i>I</i>
          </button>
          <button
            type="button"
            className={`format-button ${activeStyle.underline ? "active" : ""}`}
            onClick={toggleUnderline}
            onMouseDown={(e) => e.preventDefault()}
            title="Underline (⌘U)"
            aria-label="Underline"
          >
            <u>U</u>
          </button>
          <button
            type="button"
            className={`format-button ${!activeStyle.align || activeStyle.align === "left" ? "active" : ""}`}
            onClick={() => setAlignment("left")}
            onMouseDown={(e) => e.preventDefault()}
            title="Align left"
            aria-label="Align left"
          >
            <svg
              width="14"
              height="14"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
              strokeLinecap="round"
            >
              <line x1="3" y1="6" x2="21" y2="6" />
              <line x1="3" y1="12" x2="15" y2="12" />
              <line x1="3" y1="18" x2="18" y2="18" />
            </svg>
          </button>
          <button
            type="button"
            className={`format-button ${activeStyle.align === "center" ? "active" : ""}`}
            onClick={() => setAlignment("center")}
            onMouseDown={(e) => e.preventDefault()}
            title="Align center"
            aria-label="Align center"
          >
            <svg
              width="14"
              height="14"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
              strokeLinecap="round"
            >
              <line x1="3" y1="6" x2="21" y2="6" />
              <line x1="6" y1="12" x2="18" y2="12" />
              <line x1="4" y1="18" x2="20" y2="18" />
            </svg>
          </button>
          <button
            type="button"
            className={`format-button ${activeStyle.align === "right" ? "active" : ""}`}
            onClick={() => setAlignment("right")}
            onMouseDown={(e) => e.preventDefault()}
            title="Align right"
            aria-label="Align right"
          >
            <svg
              width="14"
              height="14"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
              strokeLinecap="round"
            >
              <line x1="3" y1="6" x2="21" y2="6" />
              <line x1="9" y1="12" x2="21" y2="12" />
              <line x1="6" y1="18" x2="21" y2="18" />
            </svg>
          </button>

          {/* Merge Toolbar Button */}
          <div className="merge-group">
            <button
              type="button"
              className={`format-button ${isSelectionMerged ? "active" : ""}`}
              onClick={() => {
                if (isSelectionMerged) {
                  unmergeCells();
                } else {
                  mergeCells();
                }
              }}
              onMouseDown={(e) => e.preventDefault()}
              title={
                isSelectionMerged
                  ? "Unmerge cells"
                  : "Merge selected cells"
              }
              aria-label="Merge cells"
            >
              <svg
                width="14"
                height="14"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <rect x="3" y="4" width="18" height="16" rx="2" />
                <path d="M12 4v4M12 16v4" />
                <path d="M8 12h8" strokeWidth="2.5" />
                <path d="M11 9l-3 3 3 3M13 9l3 3-3 3" />
              </svg>
            </button>
            <button
              type="button"
              className="format-button dropdown-trigger"
              onClick={() => setShowMergeDropdown(!showMergeDropdown)}
              onMouseDown={(e) => e.preventDefault()}
              title="Merge options"
              aria-label="Merge options"
            >
              <svg width="8" height="8" viewBox="0 0 24 24" fill="currentColor">
                <path d="M7 10l5 5 5-5z" />
              </svg>
            </button>
            {showMergeDropdown && (
              <div
                className="toolbar-dropdown"
                onMouseDown={(e) => e.preventDefault()}
              >
                <button
                  onClick={() => {
                    mergeCells();
                    setShowMergeDropdown(false);
                  }}
                >
                  <span className="dropdown-item-content">
                    <svg
                      width="13"
                      height="13"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      strokeLinecap="round"
                    >
                      <rect x="3" y="3" width="18" height="18" rx="2" />
                      <path d="M9 12h6M12 9v6" />
                    </svg>
                    Merge all
                  </span>
                </button>
                <button
                  onClick={() => {
                    mergeHorizontally();
                    setShowMergeDropdown(false);
                  }}
                >
                  <span className="dropdown-item-content">
                    <svg
                      width="13"
                      height="13"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      strokeLinecap="round"
                    >
                      <rect x="2" y="4" width="20" height="16" rx="2" />
                      <path d="M12 4v16" />
                      <path d="M7 12h10" />
                    </svg>
                    Merge horizontally (columns)
                  </span>
                </button>
                <button
                  onClick={() => {
                    mergeVertically();
                    setShowMergeDropdown(false);
                  }}
                >
                  <span className="dropdown-item-content">
                    <svg
                      width="13"
                      height="13"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      strokeLinecap="round"
                    >
                      <rect x="4" y="2" width="16" height="20" rx="2" />
                      <path d="M4 12h16" />
                      <path d="M12 7v10" />
                    </svg>
                    Merge vertically (rows)
                  </span>
                </button>
                <button
                  onClick={() => {
                    unmergeCells();
                    setShowMergeDropdown(false);
                  }}
                >
                  <span className="dropdown-item-content">
                    <svg
                      width="13"
                      height="13"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      strokeLinecap="round"
                    >
                      <rect x="3" y="3" width="18" height="18" rx="2" />
                      <path d="M3 12h18M12 3v18" strokeDasharray="2 2" />
                    </svg>
                    Unmerge cells
                  </span>
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="toolbar-spacer" />
        <div className="toolbar-group">
          <button
            className="tool-button"
            onClick={() => sortColumn(activeCol, false)}
            title="Sort sheet by current column A to Z"
          >
            ↕ <span>Sort</span>
          </button>
          <button
            className="tool-button"
            onClick={() => document.getElementById("search-input")?.focus()}
          >
            ⌕ <span>Find</span>
          </button>
          <button
            className="tool-button"
            onClick={exportCsv}
            title="Export spreadsheet to CSV file"
          >
            ⇩ <span>CSV</span>
          </button>
          <button
            className="tool-button"
            onClick={exportExcel}
            title="Export spreadsheet to Excel (.xlsx) file"
          >
            ⇩ <span>Excel</span>
          </button>
        </div>
        <label className="autosave">
          <span>Autosave</span>
          <input
            type="checkbox"
            checked={autosave}
            onChange={(event) => setAutosave(event.target.checked)}
          />
          <i />
        </label>
      </section>
      <section className="workspace">
        <div className="formula-bar">
          <span className="name-box">{selectionLabel}</span>
          <div className="fx-group">
            <span className="fx">fx</span>
            {isFormulaMode && (
              <div className="formula-actions">
                <button
                  type="button"
                  className="formula-btn cancel"
                  title="Cancel (Esc)"
                  onMouseDown={(e) => {
                    e.preventDefault();
                    setEditing(false);
                    setIsFormulaBarFocused(false);
                    setFormulaPointReplacing(false);
                  }}
                >
                  ✕
                </button>
                <button
                  type="button"
                  className="formula-btn commit"
                  title="Commit (Enter)"
                  onMouseDown={(e) => {
                    e.preventDefault();
                    setEditing(false);
                    setIsFormulaBarFocused(false);
                    setFormulaPointReplacing(false);
                  }}
                >
                  ✓
                </button>
              </div>
            )}
          </div>
          <div
            className={`formula-input-container ${activeRawValue.startsWith("=") ? "is-formula" : ""}`}
          >
            {activeRawValue.startsWith("=") && (
              <div className="formula-syntax-backdrop" aria-hidden="true">
                {renderFormulaSyntax(activeRawValue)}
              </div>
            )}
            <input
              id="formula-bar-input"
              className="formula-real-input"
              value={activeRawValue}
              onChange={(event) => {
                const val = event.target.value;
                updateCell(val);
                if (val.startsWith("=")) {
                  const lastChar = val[val.length - 1];
                  if (
                    ["+", "-", "*", "/", "(", ",", "&", ":", "<", ">", "="].includes(
                      lastChar,
                    )
                  ) {
                    setFormulaPointReplacing(false);
                  }
                }
              }}
              onFocus={() => {
                setIsFormulaBarFocused(true);
              }}
              onBlur={() => {
                if (!isFormulaDragging) {
                  setIsFormulaBarFocused(false);
                  setFormulaPointReplacing(false);
                }
              }}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  setEditing(false);
                  setIsFormulaBarFocused(false);
                  setFormulaPointReplacing(false);
                  moveSelection(activeRow + 1, activeCol);
                } else if (e.key === "Escape") {
                  setEditing(false);
                  setIsFormulaBarFocused(false);
                  setFormulaPointReplacing(false);
                }
              }}
              aria-label="Formula bar"
              placeholder="Enter value or formula (=SUM(A1:A5), =B1+B2...)"
            />
          </div>
        </div>
        <div className="sheet-area">
          <div
            className="grid"
            style={{
              gridTemplateColumns: `44px repeat(${columnCount}, minmax(130px, 1fr))`,
            }}
          >
            <div
              className={`corner ${minRow === 0 && maxRow >= rows.length - 1 && minCol === 0 && maxCol >= columnCount - 1 ? "highlighted" : ""}`}
              style={{ gridRow: "1", gridColumn: "1" }}
              onClick={handleCornerClick}
              title="Select all cells"
            />
            {Array.from({ length: columnCount }, (_, column) => (
              <button
                className={`column-header ${column >= minCol && column <= maxCol ? "highlighted" : ""}`}
                key={column}
                style={{
                  gridRow: "1",
                  gridColumn: `${column + 2}`,
                }}
                onMouseDown={(e) => {
                  const isContextMenuTrigger = e.button === 2 || e.ctrlKey;
                  const isFullCols =
                    minRow === 0 && maxRow >= Math.max(0, rows.length - 1);
                  if (isContextMenuTrigger) {
                    if (
                      isFullCols &&
                      column >= minCol &&
                      column <= maxCol
                    ) {
                      return;
                    }
                    setSelection({
                      startRow: 0,
                      startCol: column,
                      endRow: Math.max(0, rows.length - 1),
                      endCol: column,
                    });
                    return;
                  }
                  handleColumnHeaderMouseDown(column, e);
                }}
                onMouseEnter={() => handleColumnHeaderMouseEnter(column)}
                onContextMenu={(e) =>
                  openContextMenu(e, { type: "column", columnIndex: column })
                }
              >
                {colIndexToLetter(column)}
              </button>
            ))}
            {visibleRows.map(({ row, rowIndex }) => (
              <React.Fragment key={rowIndex}>
                <button
                  className={`row-header ${rowIndex >= minRow && rowIndex <= maxRow ? "highlighted" : ""}`}
                  style={{
                    gridRow: `${rowIndex + 2}`,
                    gridColumn: "1",
                  }}
                  onMouseDown={(e) => {
                    const isContextMenuTrigger = e.button === 2 || e.ctrlKey;
                    const isFullRows =
                      minCol === 0 && maxCol >= columnCount - 1;
                    if (isContextMenuTrigger) {
                      if (
                        isFullRows &&
                        rowIndex >= minRow &&
                        rowIndex <= maxRow
                      ) {
                        return;
                      }
                      setSelection({
                        startRow: rowIndex,
                        startCol: 0,
                        endRow: rowIndex,
                        endCol: Math.max(0, columnCount - 1),
                      });
                      return;
                    }
                    handleRowHeaderMouseDown(rowIndex, e);
                  }}
                  onMouseEnter={() => handleRowHeaderMouseEnter(rowIndex)}
                  onContextMenu={(e) =>
                    openContextMenu(e, { type: "row", rowIndex })
                  }
                >
                  {rowIndex + 1}
                </button>
                {Array.from({ length: columnCount }, (_, column) => {
                  const merge = getMergeAt(rowIndex, column);
                  const isMerged = Boolean(merge);
                  const isMaster = merge
                    ? merge.startRow === rowIndex && merge.startCol === column
                    : false;
                  const isCovered = merge ? !isMaster : false;

                  if (isCovered) {
                    return null;
                  }

                  const rowSpan = merge
                    ? merge.endRow - merge.startRow + 1
                    : 1;
                  const colSpan = merge
                    ? merge.endCol - merge.startCol + 1
                    : 1;

                  const cell = row[column];
                  const style = getCellStyle(cell);
                  const isSelected =
                    activeRow === rowIndex && activeCol === column;
                  const inRange =
                    rowIndex >= minRow &&
                    rowIndex <= maxRow &&
                    column >= minCol &&
                    column <= maxCol;

                  const rawVal = getCellValue(cell);
                  const isFormula = rawVal.startsWith("=");
                  const evalVal = isFormula
                    ? evaluatedSheet[rowIndex]?.[column] ?? ""
                    : rawVal;
                  const isThisCellActive =
                    activeRow === rowIndex && activeCol === column;
                  const isThisCellEditing = editing && isThisCellActive;
                  const displayVal = isThisCellEditing ? rawVal : evalVal;
                  const isError =
                    isFormula &&
                    !isThisCellEditing &&
                    String(evalVal).startsWith("#");

                  const isRangeMulti =
                    selectedRowCount > 1 || selectedColCount > 1;
                  const isRangeTop =
                    isRangeMulti &&
                    inRange &&
                    (merge
                      ? merge.startRow === minRow
                      : rowIndex === minRow);
                  const isRangeBottom =
                    isRangeMulti &&
                    inRange &&
                    (merge
                      ? merge.endRow === maxRow
                      : rowIndex === maxRow);
                  const isRangeLeft =
                    isRangeMulti &&
                    inRange &&
                    (merge
                      ? merge.startCol === minCol
                      : column === minCol);
                  const isRangeRight =
                    isRangeMulti &&
                    inRange &&
                    (merge
                      ? merge.endCol === maxCol
                      : column === maxCol);

                  const cellElement = (
                    <input
                      key={`${rowIndex}-${column}`}
                      ref={isThisCellActive ? activeCellRef : undefined}
                      data-row={rowIndex}
                      data-col={column}
                      className={`cell ${isSelected ? "active-cell" : ""} ${inRange ? "in-range" : ""} ${isRangeTop ? "range-top" : ""} ${isRangeBottom ? "range-bottom" : ""} ${isRangeLeft ? "range-left" : ""} ${isRangeRight ? "range-right" : ""} ${isMerged ? "merged-cell" : ""} ${isThisCellEditing && isFormula ? "editing-formula" : ""} ${isError ? "cell-error" : ""} ${query && String(displayVal).toLowerCase().includes(query) ? "match" : ""}`}
                      style={{
                        ...getCellInlineStyle(style),
                        gridRow: `${rowIndex + 2} / span ${rowSpan}`,
                        gridColumn: `${column + 2} / span ${colSpan}`,
                        zIndex: isSelected ? 4 : isMerged ? 2 : 1,
                      }}
                      value={displayVal}
                      onMouseDown={(e) => {
                        const isContextMenuTrigger =
                          e.button === 2 || e.ctrlKey;
                        const isInside =
                          rowIndex >= minRow &&
                          rowIndex <= maxRow &&
                          column >= minCol &&
                          column <= maxCol;

                        if (isContextMenuTrigger) {
                          if (isInside) return;
                          const merge = getMergeAt(rowIndex, column);
                          if (merge) {
                            setSelection({
                              startRow: merge.startRow,
                              startCol: merge.startCol,
                              endRow: merge.endRow,
                              endCol: merge.endCol,
                            });
                          } else {
                            setSelection({
                              startRow: rowIndex,
                              startCol: column,
                              endRow: rowIndex,
                              endCol: column,
                            });
                          }
                          return;
                        }

                        // Formula pointing mode: user clicked another cell while in formula
                        if (
                          isFormulaMode &&
                          (rowIndex !== activeRow || column !== activeCol)
                        ) {
                          e.preventDefault();
                          const targetRef = formatCellCoord(rowIndex, column);
                          const { formula: nextFormula } =
                            applyCellReferenceToFormula(
                              activeRawValue,
                              targetRef,
                              formulaPointReplacing,
                            );
                          updateCell(nextFormula);
                          setFormulaPointReplacing(true);
                          setFormulaPointDragStart({
                            row: rowIndex,
                            col: column,
                          });
                          setIsFormulaDragging(true);
                          // Move cursor to end after inserting reference
                          requestAnimationFrame(() => {
                            const el = activeCellRef.current;
                            if (el) {
                              const len = nextFormula.length;
                              el.setSelectionRange(len, len);
                            }
                          });
                          return;
                        }

                        handleCellMouseDown(rowIndex, column, e);
                      }}
                      onMouseEnter={() => {
                        if (isFormulaDragging && formulaPointDragStart) {
                          const rangeRef = formatCellRange(
                            formulaPointDragStart.row,
                            formulaPointDragStart.col,
                            rowIndex,
                            column,
                          );
                          const { formula: nextFormula } =
                            applyCellReferenceToFormula(
                              activeRawValue,
                              rangeRef,
                              true,
                            );
                          updateCell(nextFormula);
                          // Move cursor to end after inserting range reference
                          requestAnimationFrame(() => {
                            const el = activeCellRef.current;
                            if (el) {
                              const len = nextFormula.length;
                              el.setSelectionRange(len, len);
                            }
                          });
                          return;
                        }
                        handleCellMouseEnter(rowIndex, column);
                      }}
                      onContextMenu={(e) =>
                        openContextMenu(e, {
                          type: "cell",
                          rowIndex,
                          columnIndex: column,
                        })
                      }
                      onFocus={() => {
                        const isInside =
                          rowIndex >= minRow &&
                          rowIndex <= maxRow &&
                          column >= minCol &&
                          column <= maxCol;
                        if (!isInside && !isFormulaMode) {
                          setSelection({
                            startRow: rowIndex,
                            startCol: column,
                            endRow: rowIndex,
                            endCol: column,
                          });
                        }
                      }}
                      onChange={(event) => {
                        const val = event.target.value;
                        updateCell(val);
                        if (!editing) setEditing(true);
                        if (val.startsWith("=")) {
                          const lastChar = val[val.length - 1];
                          if (
                            [
                              "+",
                              "-",
                              "*",
                              "/",
                              "(",
                              ",",
                              "&",
                              ":",
                              "<",
                              ">",
                              "=",
                            ].includes(lastChar)
                          ) {
                            setFormulaPointReplacing(false);
                          }
                        }
                      }}
                      onDoubleClick={() => setEditing(true)}
                      onBlur={() => {
                        if (!isFormulaDragging) {
                          setEditing(false);
                          setFormulaPointReplacing(false);
                        }
                      }}
                      onKeyDown={(event) => {
                        const command = event.metaKey || event.ctrlKey;
                        if (command && event.key.toLowerCase() === "b") {
                          event.preventDefault();
                          toggleBold();
                          return;
                        }
                        if (command && event.key.toLowerCase() === "i") {
                          event.preventDefault();
                          toggleItalic();
                          return;
                        }
                        if (command && event.key.toLowerCase() === "u") {
                          event.preventDefault();
                          toggleUnderline();
                          return;
                        }
                        if (event.key === "Enter") {
                          event.preventDefault();
                          setEditing(false);
                          setFormulaPointReplacing(false);
                          moveSelection(rowIndex + 1, column);
                        }
                        if (event.key === "Tab") {
                          event.preventDefault();
                          setEditing(false);
                          setFormulaPointReplacing(false);
                          moveSelection(
                            rowIndex,
                            column + (event.shiftKey ? -1 : 1),
                          );
                        }
                        if (event.key === "Escape") {
                          setEditing(false);
                          setFormulaPointReplacing(false);
                          event.currentTarget.blur();
                        }
                      }}
                      aria-label={`${colIndexToLetter(column)}${rowIndex + 1}`}
                    />
                  );

                  // Always wrap the active cell so the DOM structure stays
                  // stable when formula mode toggles (prevents focus loss
                  // when user types "=" and the wrapper would otherwise
                  // appear/disappear, causing React to remount the input).
                  if (isThisCellActive) {
                    return (
                      <div
                        key={`${rowIndex}-${column}`}
                        className="cell-wrapper"
                        style={{
                          gridRow: `${rowIndex + 2} / span ${rowSpan}`,
                          gridColumn: `${column + 2} / span ${colSpan}`,
                          zIndex: isSelected ? 5 : isMerged ? 2 : 1,
                        }}
                      >
                        {isThisCellEditing && isFormula && (
                          <div
                            className="cell-syntax-backdrop"
                            aria-hidden="true"
                          >
                            {renderFormulaSyntax(rawVal)}
                          </div>
                        )}
                        {cellElement}
                      </div>
                    );
                  }

                  return cellElement;
                })}
              </React.Fragment>
            ))}

            {/* Formula Reference Highlight Overlays on Grid */}
            {formulaHighlights.map((hl, idx) => (
              <div
                key={`formula-hl-${idx}`}
                className="formula-grid-overlay"
                style={{
                  gridRow: `${hl.range.startRow + 2} / ${hl.range.endRow + 3}`,
                  gridColumn: `${hl.range.startCol + 2} / ${hl.range.endCol + 3}`,
                  borderColor: hl.color.border,
                  backgroundColor: hl.color.bg,
                  zIndex: 8,
                  pointerEvents: "none",
                }}
              >
                <span
                  className="formula-grid-badge"
                  style={{ backgroundColor: hl.color.border }}
                >
                  {hl.label}
                </span>
              </div>
            ))}
          </div>
        </div>
      </section>
      <footer className="statusbar">
        <div className="sheet-tabs">
          {activeDocument.sheets.map((sheet, index) => (
            <button
              className={index === activeDocument.activeSheet ? "active" : ""}
              key={sheet.name}
              onClick={() => {
                setDocuments((current) =>
                  current.map((document) =>
                    document.id === activeId
                      ? { ...document, activeSheet: index }
                      : document,
                  ),
                );
                setSelection({
                  startRow: 0,
                  startCol: 0,
                  endRow: 0,
                  endCol: 0,
                });
              }}
              onContextMenu={(e) =>
                openContextMenu(e, { type: "sheet", sheetIndex: index })
              }
            >
              {sheet.name}
            </button>
          ))}
          <button
            onClick={() =>
              setDocuments((current) =>
                current.map((document) =>
                  document.id === activeId
                    ? {
                        ...document,
                        sheets: [
                          ...document.sheets,
                          {
                            name: `Sheet ${document.sheets.length + 1}`,
                            rows: starterRows.map((row) =>
                              row.map((cell) =>
                                typeof cell === "object"
                                  ? {
                                      value: cell.value,
                                      style: { ...cell.style },
                                    }
                                  : { value: cell, style: {} },
                              ),
                            ),
                            merges: [],
                          },
                        ],
                      }
                    : document,
                ),
              )
            }
          >
            ＋
          </button>
        </div>
        <div className="status-details">
          <div className={`save-status ${saveStatus}`}>
            {saveStatus === "saved"
              ? "✓ Saved"
              : saveStatus === "saving"
                ? "◌ Saving…"
                : "● Unsaved changes"}
          </div>
          <span>
            {rows.length} rows · {columnCount} columns
          </span>
          <span>100%</span>
        </div>
      </footer>
      <input
        ref={fileInput}
        type="file"
        accept=".csv,.xlsx,.xls"
        hidden
        onChange={handleFileInput}
      />
      <div className="search-popover">
        <span>⌕</span>
        <input
          id="search-input"
          value={search}
          onChange={(event) => setSearch(event.target.value)}
          placeholder="Find in spreadsheet"
        />
        <kbd>⌘ F</kbd>
      </div>

      {/* Context Menu */}
      {contextMenu?.visible && (
        <div
          id="spreadsheet-context-menu"
          className="context-menu"
          style={{ left: contextMenu.x, top: contextMenu.y }}
        >
          {contextMenu.target.type === "row" && (
            <>
              <button
                className="context-menu-item"
                onClick={() => {
                  insertRowsAbove(minRow, selectedRowCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M12 5v14M5 12h14" />
                  </svg>
                  Insert {selectedRowCount} row
                  {selectedRowCount > 1 ? "s" : ""} above
                </span>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  insertRowsBelow(maxRow, selectedRowCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M12 5v14M5 12h14" />
                  </svg>
                  Insert {selectedRowCount} row
                  {selectedRowCount > 1 ? "s" : ""} below
                </span>
              </button>
              <div className="context-menu-divider" />
              <button
                className="context-menu-item danger"
                onClick={() => {
                  deleteRows(minRow, selectedRowCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
                  </svg>
                  Delete {selectedRowCount} row
                  {selectedRowCount > 1 ? "s" : ""} ({minRow + 1}
                  {selectedRowCount > 1 ? ` - ${maxRow + 1}` : ""})
                </span>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  clearRows(minRow, selectedRowCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M18 6L6 18M6 6l12 12" />
                  </svg>
                  Clear {selectedRowCount} row
                  {selectedRowCount > 1 ? "s" : ""}
                </span>
              </button>
              <div className="context-menu-divider" />
              {selectedRowCount > 1 && (
                <>
                  <button
                    className="context-menu-item"
                    onClick={() => {
                      mergeVertically();
                      closeContextMenu();
                    }}
                  >
                    <span className="menu-label">
                      <svg
                        width="14"
                        height="14"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                      >
                        <rect x="4" y="2" width="16" height="20" rx="2" />
                        <path d="M4 12h16" />
                        <path d="M12 7v10" />
                      </svg>
                      Merge rows vertically
                    </span>
                  </button>
                  <button
                    className="context-menu-item"
                    onClick={() => {
                      mergeCells();
                      closeContextMenu();
                    }}
                  >
                    <span className="menu-label">
                      <svg
                        width="14"
                        height="14"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                      >
                        <rect x="3" y="3" width="18" height="18" rx="2" />
                        <path d="M9 12h6M12 9v6" />
                      </svg>
                      Merge all cells in rows
                    </span>
                  </button>
                </>
              )}
              {isSelectionMerged && (
                <button
                  className="context-menu-item"
                  onClick={() => {
                    unmergeCells();
                    closeContextMenu();
                  }}
                >
                  <span className="menu-label">
                    <svg
                      width="14"
                      height="14"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      strokeLinecap="round"
                    >
                      <rect x="3" y="3" width="18" height="18" rx="2" />
                      <path d="M3 12h18M12 3v18" strokeDasharray="2 2" />
                    </svg>
                    Unmerge cells
                  </span>
                </button>
              )}
              <div className="context-menu-divider" />
              <button
                className="context-menu-item"
                onClick={() => {
                  void cutSelection();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">Cut</span>
                <kbd>⌘X</kbd>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  void copySelection();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">Copy</span>
                <kbd>⌘C</kbd>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  void pasteSelection(false);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">Paste</span>
                <kbd>⌘V</kbd>
              </button>
              <div className="context-menu-divider" />
              <button
                className="context-menu-item"
                onClick={() => {
                  toggleBold();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <strong>B</strong> Bold
                </span>
                <kbd>⌘B</kbd>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  toggleItalic();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <i>I</i> Italic
                </span>
                <kbd>⌘I</kbd>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  toggleUnderline();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <u>U</u> Underline
                </span>
                <kbd>⌘U</kbd>
              </button>
            </>
          )}

          {contextMenu.target.type === "column" && (
            <>
              <button
                className="context-menu-item"
                onClick={() => {
                  insertColumnsLeft(minCol, selectedColCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M12 5v14M5 12h14" />
                  </svg>
                  Insert {selectedColCount} column
                  {selectedColCount > 1 ? "s" : ""} left
                </span>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  insertColumnsRight(maxCol, selectedColCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M12 5v14M5 12h14" />
                  </svg>
                  Insert {selectedColCount} column
                  {selectedColCount > 1 ? "s" : ""} right
                </span>
              </button>
              <div className="context-menu-divider" />
              <button
                className="context-menu-item danger"
                onClick={() => {
                  deleteColumns(minCol, selectedColCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
                  </svg>
                  Delete {selectedColCount} column
                  {selectedColCount > 1 ? "s" : ""} ({letters[minCol]}
                  {selectedColCount > 1 ? ` - ${letters[maxCol]}` : ""})
                </span>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  clearColumns(minCol, selectedColCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M18 6L6 18M6 6l12 12" />
                  </svg>
                  Clear {selectedColCount} column
                  {selectedColCount > 1 ? "s" : ""}
                </span>
              </button>
              <div className="context-menu-divider" />
              <button
                className="context-menu-item"
                onClick={() => {
                  sortColumn(minCol, false);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M3 4h18M3 10h12M3 16h6" />
                  </svg>
                  Sort sheet A → Z
                </span>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  sortColumn(minCol, true);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M3 4h6M3 10h12M3 16h18" />
                  </svg>
                  Sort sheet Z → A
                </span>
              </button>
              <div className="context-menu-divider" />
              {selectedColCount > 1 && (
                <>
                  <button
                    className="context-menu-item"
                    onClick={() => {
                      mergeHorizontally();
                      closeContextMenu();
                    }}
                  >
                    <span className="menu-label">
                      <svg
                        width="14"
                        height="14"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                      >
                        <rect x="2" y="4" width="20" height="16" rx="2" />
                        <path d="M12 4v16" />
                        <path d="M7 12h10" />
                      </svg>
                      Merge columns horizontally
                    </span>
                  </button>
                  <button
                    className="context-menu-item"
                    onClick={() => {
                      mergeCells();
                      closeContextMenu();
                    }}
                  >
                    <span className="menu-label">
                      <svg
                        width="14"
                        height="14"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                      >
                        <rect x="3" y="3" width="18" height="18" rx="2" />
                        <path d="M9 12h6M12 9v6" />
                      </svg>
                      Merge all cells in columns
                    </span>
                  </button>
                </>
              )}
              {isSelectionMerged && (
                <button
                  className="context-menu-item"
                  onClick={() => {
                    unmergeCells();
                    closeContextMenu();
                  }}
                >
                  <span className="menu-label">
                    <svg
                      width="14"
                      height="14"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      strokeLinecap="round"
                    >
                      <rect x="3" y="3" width="18" height="18" rx="2" />
                      <path d="M3 12h18M12 3v18" strokeDasharray="2 2" />
                    </svg>
                    Unmerge cells
                  </span>
                </button>
              )}
              <div className="context-menu-divider" />
              <button
                className="context-menu-item"
                onClick={() => {
                  void cutSelection();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">Cut</span>
                <kbd>⌘X</kbd>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  void copySelection();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">Copy</span>
                <kbd>⌘C</kbd>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  void pasteSelection(false);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">Paste</span>
                <kbd>⌘V</kbd>
              </button>
              <div className="context-menu-divider" />
              <button
                className="context-menu-item"
                onClick={() => {
                  toggleBold();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <strong>B</strong> Bold
                </span>
                <kbd>⌘B</kbd>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  toggleItalic();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <i>I</i> Italic
                </span>
                <kbd>⌘I</kbd>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  toggleUnderline();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <u>U</u> Underline
                </span>
                <kbd>⌘U</kbd>
              </button>
            </>
          )}

          {contextMenu.target.type === "cell" && (
            <>
              <button
                className="context-menu-item"
                onClick={() => {
                  void cutSelection();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">Cut</span>
                <kbd>⌘X</kbd>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  void copySelection();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">Copy</span>
                <kbd>⌘C</kbd>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  void pasteSelection(false);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">Paste</span>
                <kbd>⌘V</kbd>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  void pasteSelection(true);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">Paste values only</span>
                <kbd>⌘⇧V</kbd>
              </button>
              <div className="context-menu-divider" />
              {(selectedRowCount > 1 || selectedColCount > 1) && (
                <>
                  <button
                    className="context-menu-item"
                    onClick={() => {
                      mergeCells();
                      closeContextMenu();
                    }}
                  >
                    <span className="menu-label">
                      <svg
                        width="14"
                        height="14"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                      >
                        <rect x="3" y="3" width="18" height="18" rx="2" />
                        <path d="M9 12h6M12 9v6" />
                      </svg>
                      Merge all cells
                    </span>
                  </button>
                  {selectedColCount > 1 && (
                    <button
                      className="context-menu-item"
                      onClick={() => {
                        mergeHorizontally();
                        closeContextMenu();
                      }}
                    >
                      <span className="menu-label">
                        <svg
                          width="14"
                          height="14"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2"
                          strokeLinecap="round"
                        >
                          <rect x="2" y="4" width="20" height="16" rx="2" />
                          <path d="M12 4v16" />
                          <path d="M7 12h10" />
                        </svg>
                        Merge columns horizontally
                      </span>
                    </button>
                  )}
                  {selectedRowCount > 1 && (
                    <button
                      className="context-menu-item"
                      onClick={() => {
                        mergeVertically();
                        closeContextMenu();
                      }}
                    >
                      <span className="menu-label">
                        <svg
                          width="14"
                          height="14"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2"
                          strokeLinecap="round"
                        >
                          <rect x="4" y="2" width="16" height="20" rx="2" />
                          <path d="M4 12h16" />
                          <path d="M12 7v10" />
                        </svg>
                        Merge rows vertically
                      </span>
                    </button>
                  )}
                </>
              )}
              {isSelectionMerged && (
                <button
                  className="context-menu-item"
                  onClick={() => {
                    unmergeCells();
                    closeContextMenu();
                  }}
                >
                  <span className="menu-label">
                    <svg
                      width="14"
                      height="14"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      strokeLinecap="round"
                    >
                      <rect x="3" y="3" width="18" height="18" rx="2" />
                      <path d="M3 12h18M12 3v18" strokeDasharray="2 2" />
                    </svg>
                    Unmerge cells
                  </span>
                </button>
              )}
              <div className="context-menu-divider" />
              <button
                className="context-menu-item"
                onClick={() => {
                  insertRowsAbove(minRow, selectedRowCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M12 5v14M5 12h14" />
                  </svg>
                  Insert {selectedRowCount} row
                  {selectedRowCount > 1 ? "s" : ""} above
                </span>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  insertRowsBelow(maxRow, selectedRowCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M12 5v14M5 12h14" />
                  </svg>
                  Insert {selectedRowCount} row
                  {selectedRowCount > 1 ? "s" : ""} below
                </span>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  insertColumnsLeft(minCol, selectedColCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M12 5v14M5 12h14" />
                  </svg>
                  Insert {selectedColCount} column
                  {selectedColCount > 1 ? "s" : ""} left
                </span>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  insertColumnsRight(maxCol, selectedColCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M12 5v14M5 12h14" />
                  </svg>
                  Insert {selectedColCount} column
                  {selectedColCount > 1 ? "s" : ""} right
                </span>
              </button>
              <div className="context-menu-divider" />
              <button
                className="context-menu-item danger"
                onClick={() => {
                  deleteRows(minRow, selectedRowCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
                  </svg>
                  Delete {selectedRowCount} row
                  {selectedRowCount > 1 ? "s" : ""}{" "}
                  {minCol === 0 && maxCol >= columnCount - 1
                    ? `(${minRow + 1}${selectedRowCount > 1 ? ` - ${maxRow + 1}` : ""})`
                    : ""}
                </span>
              </button>
              <button
                className="context-menu-item danger"
                onClick={() => {
                  deleteColumns(minCol, selectedColCount);
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
                  </svg>
                  Delete {selectedColCount} column
                  {selectedColCount > 1 ? "s" : ""}{" "}
                  {minRow === 0 && maxRow >= Math.max(0, rows.length - 1)
                    ? `(${letters[minCol]}${selectedColCount > 1 ? ` - ${letters[maxCol]}` : ""})`
                    : ""}
                </span>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  clearSelectionContents();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M18 6L6 18M6 6l12 12" />
                  </svg>
                  Clear contents ({selectedRowCount * selectedColCount} cells)
                </span>
                <kbd>Del</kbd>
              </button>
              <div className="context-menu-divider" />
              <button
                className="context-menu-item"
                onClick={() => {
                  toggleBold();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <strong>B</strong> Bold
                </span>
                <kbd>⌘B</kbd>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  toggleItalic();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <i>I</i> Italic
                </span>
                <kbd>⌘I</kbd>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  toggleUnderline();
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <u>U</u> Underline
                </span>
                <kbd>⌘U</kbd>
              </button>
            </>
          )}

          {contextMenu.target.type === "sheet" && (
            <>
              <button
                className="context-menu-item"
                onClick={() => {
                  renameSheet(
                    (
                      contextMenu.target as {
                        type: "sheet";
                        sheetIndex: number;
                      }
                    ).sheetIndex,
                  );
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <path d="M12 20h9M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z" />
                  </svg>
                  Rename sheet
                </span>
              </button>
              <button
                className="context-menu-item"
                onClick={() => {
                  duplicateSheet(
                    (
                      contextMenu.target as {
                        type: "sheet";
                        sheetIndex: number;
                      }
                    ).sheetIndex,
                  );
                  closeContextMenu();
                }}
              >
                <span className="menu-label">
                  <svg
                    width="14"
                    height="14"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                  >
                    <rect
                      x="9"
                      y="9"
                      width="13"
                      height="13"
                      rx="2"
                      ry="2"
                    />
                    <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
                  </svg>
                  Duplicate sheet
                </span>
              </button>
              {activeDocument.sheets.length > 1 && (
                <button
                  className="context-menu-item danger"
                  onClick={() => {
                    deleteSheet(
                      (
                        contextMenu.target as {
                          type: "sheet";
                          sheetIndex: number;
                        }
                      ).sheetIndex,
                    );
                    closeContextMenu();
                  }}
                >
                  <span className="menu-label">
                    <svg
                      width="14"
                      height="14"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      strokeLinecap="round"
                    >
                      <path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
                    </svg>
                    Delete sheet
                  </span>
                </button>
              )}
            </>
          )}
        </div>
      )}
        </>
      )}
    </main>
  );
}

export default App;
