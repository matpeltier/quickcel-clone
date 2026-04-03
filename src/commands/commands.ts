import { getCachedSettings, refreshSettingsCache } from "../shared/settings";
import {
  buildNumberFormat,
  buildPercentFormat,
  buildMultipleFormat,
  DATE_FORMAT,
  getBorderStyleForExcel,
  getBorderColorForEntry,
  hexToRgb,
  detectCellType,
} from "../shared/excel-utils";
import { BorderStyleEntry, ColorEntry } from "../shared/types";

let specialCopyClipboard: { formula: string; address: string } | null = null;
const cycleState: Record<string, number> = {};

function getCycleIndex(key: string, listLength: number): number {
  if (cycleState[key] === undefined) {
    cycleState[key] = 0;
  } else {
    cycleState[key] = (cycleState[key] + 1) % listLength;
  }
  return cycleState[key];
}

async function getSelectedRange(): Promise<Excel.Range> {
  await Excel.run(async (context) => {
    return context;
  });
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["rowCount", "columnCount", "address", "formulas", "values", "numberFormat", "format"]);
    range.format.load(["font", "fill", "borders", "horizontalAlignment", "verticalAlignment", "indentLevel"]);
    range.format.font.load(["color"]);
    range.format.fill.load(["color"]);
    range.format.borders.load(["getItem"]);
    await context.sync();
    return range;
  }) as any;
}

async function runWithRange(callback: (context: Excel.RequestContext, range: Excel.Range) => Promise<void>): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["rowCount", "columnCount", "address"]);
    await context.sync();
    await callback(context, range);
    await context.sync();
  });
}

function hexToExcelColor(hex: string): string {
  const rgb = hexToRgb(hex);
  return `${rgb.r},${rgb.g},${rgb.b}`;
}

async function fontColorCycle(): Promise<void> {
  const settings = getCachedSettings();
  const colors = settings.fontColorCycle;
  const idx = getCycleIndex("fontColor", colors.length);
  const color = colors[idx];

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.font.color = color.fontColor || "#000000";
    await context.sync();
  });
}

async function autoColor(): Promise<void> {
  const settings = getCachedSettings();
  const inputColor = settings.fontColorCycle.find((c) => c.name.includes("Input"))?.fontColor || "#0000FF";
  const linkColor = settings.fontColorCycle.find((c) => c.name.includes("Link"))?.fontColor || "#008000";
  const formulaColor = "#000000";

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["rowCount", "columnCount", "formulas", "values"]);
    await context.sync();

    for (let r = 0; r < range.rowCount; r++) {
      for (let c = 0; c < range.columnCount; c++) {
        const formula = range.formulas[r][c] as string;
        const value = range.values[r][c] as string;
        const cellType = detectCellType(formula, value);

        let color = formulaColor;
        if (cellType === "input") color = inputColor;
        else if (cellType === "link") color = linkColor;

        const cell = range.getCell(r, c);
        cell.format.font.color = color;
      }
    }
    await context.sync();
  });
}

async function cellColorCycle(): Promise<void> {
  const settings = getCachedSettings();
  const colors = settings.cellColorCycle;
  const idx = getCycleIndex("cellColor", colors.length);
  const color = colors[idx];

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    if (color.fillColor === "none" || color.name === "None") {
      range.format.fill.clear();
    } else {
      range.format.fill.color = color.fillColor!;
    }
    await context.sync();
  });
}

async function brandCycle(): Promise<void> {
  const settings = getCachedSettings();
  const brands = settings.brandStyles;
  const idx = getCycleIndex("brand", brands.length);
  const brand = brands[idx];

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.fill.color = brand.fillColor;
    range.format.font.color = brand.fontColor;

    const borderStyle = getBorderStyleForExcel(brand.borderStyle);
    const borderColor = brand.borderColor;
    const outlineBorders = ["EdgeTop" as const, "EdgeBottom" as const, "EdgeLeft" as const, "EdgeRight" as const];
    for (const bi of outlineBorders) {
      const border = range.format.borders.getItem(bi);
      border.style = borderStyle;
      border.color = borderColor;
    }
    await context.sync();
  });
}

async function yearCycle(): Promise<void> {
  const settings = getCachedSettings();
  const suffixes = settings.yearSuffixes;
  const idx = getCycleIndex("year", suffixes.length);
  const suffix = suffixes[idx].suffix;

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["values", "formulas", "numberFormat"]);
    await context.sync();

    for (let r = 0; r < range.rowCount; r++) {
      for (let c = 0; c < range.columnCount; c++) {
        const val = range.values[r][c] as string;
        const formula = range.formulas[r][c] as string;
        const cell = range.getCell(r, c);

        let baseVal: string;
        if (formula && formula.startsWith("=")) {
          baseVal = formula;
        } else {
          baseVal = String(val || "");
        }

        const currentSuffixes = suffixes.map((s) => s.suffix).filter((s) => s !== "");
        let cleaned = baseVal;
        for (const s of currentSuffixes) {
          if (cleaned.endsWith(s)) {
            cleaned = cleaned.slice(0, -s.length);
            break;
          }
        }

        if (suffix) {
          if (cleaned.startsWith("=")) {
            cell.formulas = [[cleaned + "&\"" + suffix + "\""]];
          } else {
            cell.values = [[cleaned + suffix]];
          }
        } else {
          if (cleaned.startsWith("=")) {
            cell.formulas = [[cleaned]];
          } else {
            cell.values = [[cleaned]];
          }
        }
      }
    }
    await context.sync();
  });
}

async function numberFormat(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.numberFormat = [[buildNumberFormat(0)]];
    await context.sync();
  });
}

async function dateFormat(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.numberFormat = [[DATE_FORMAT]];
    await context.sync();
  });
}

async function percentageFormat(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.numberFormat = [[buildPercentFormat(0)]];
    await context.sync();
  });
}

async function multipleFormat(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.numberFormat = [[buildMultipleFormat(1)]];
    await context.sync();
  });
}

async function increaseDecimal(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["numberFormat"]);
    await context.sync();

    const currentFmt = (range.numberFormat as any[][])[0][0] as string;
    const decimals = countDecimalsInFormat(currentFmt);
    const newFmt = incrementDecimals(currentFmt, decimals + 1);
    range.numberFormat = [[newFmt]];
    await context.sync();
  });
}

async function decreaseDecimal(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["numberFormat"]);
    await context.sync();

    const currentFmt = (range.numberFormat as any[][])[0][0] as string;
    const decimals = countDecimalsInFormat(currentFmt);
    const newFmt = incrementDecimals(currentFmt, Math.max(0, decimals - 1));
    range.numberFormat = [[newFmt]];
    await context.sync();
  });
}

function countDecimalsInFormat(fmt: string): number {
  const match = fmt.match(/0+\.(0+)/);
  return match ? match[1].length : 0;
}

function incrementDecimals(fmt: string, targetDecimals: number): string {
  if (fmt === "General") return buildNumberFormat(targetDecimals);
  return fmt.replace(/0+\.(0+)/g, (_, zeros) => "0." + "0".repeat(targetDecimals));
}

async function outlineCycle(): Promise<void> {
  const settings = getCachedSettings();
  const borders = settings.outlineBorderCycle;
  const idx = getCycleIndex("outline", borders.length);
  const borderEntry = borders[idx];

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const borderIndices = ["EdgeTop" as const, "EdgeBottom" as const, "EdgeLeft" as const, "EdgeRight" as const];

    if (borderEntry.style === "None") {
      for (const bi of borderIndices) {
        range.format.borders.getItem(bi).style = "None";
      }
    } else {
      const style = getBorderStyleForExcel(borderEntry.style);
      const color = getBorderColorForEntry(borderEntry);
      for (const bi of borderIndices) {
        const border = range.format.borders.getItem(bi);
        border.style = style;
        border.color = color;
      }
    }
    await context.sync();
  });
}

async function borderBottom(): Promise<void> {
  const settings = getCachedSettings();
  const borders = settings.borderBottomCycle;
  const idx = getCycleIndex("borderBottom", borders.length);
  const borderEntry = borders[idx];

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const border = range.format.borders.getItem("EdgeBottom");

    if (borderEntry.style === "None") {
      border.style = "None";
    } else {
      border.style = getBorderStyleForExcel(borderEntry.style);
      border.color = "#000000";
    }
    await context.sync();
  });
}

async function centerAcross(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.horizontalAlignment = "CenterAcrossSelection";
    await context.sync();
  });
}

async function decreaseIndent(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.load(["indentLevel"]);
    await context.sync();
    const current = range.format.indentLevel || 0;
    range.format.indentLevel = Math.max(0, current - 1);
    await context.sync();
  });
}

async function increaseIndent(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.load(["indentLevel"]);
    await context.sync();
    const current = range.format.indentLevel || 0;
    range.format.indentLevel = Math.min(15, current + 1);
    await context.sync();
  });
}

async function heightCycle(): Promise<void> {
  const settings = getCachedSettings();
  const heights = settings.heightCycle;
  const idx = getCycleIndex("height", heights.length);
  const height = heights[idx];

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.rowHeight = height.value;
    await context.sync();
  });
}

async function widthCycle(): Promise<void> {
  const settings = getCachedSettings();
  const widths = settings.widthCycle;
  const idx = getCycleIndex("width", widths.length);
  const width = widths[idx];

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.columnWidth = width.value;
    await context.sync();
  });
}

async function divideBy1000(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["values", "formulas"]);
    await context.sync();

    for (let r = 0; r < range.rowCount; r++) {
      for (let c = 0; c < range.columnCount; c++) {
        const formula = range.formulas[r][c] as string;
        const cell = range.getCell(r, c);

        if (formula && formula.startsWith("=")) {
          cell.formulas = [[`=(${formula})/1000`]];
        } else {
          const val = range.values[r][c] as number;
          if (typeof val === "number") {
            cell.values = [[val / 1000]];
          }
        }
      }
    }
    await context.sync();
  });
}

async function multiplyBy1000(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["values", "formulas"]);
    await context.sync();

    for (let r = 0; r < range.rowCount; r++) {
      for (let c = 0; c < range.columnCount; c++) {
        const formula = range.formulas[r][c] as string;
        const cell = range.getCell(r, c);

        if (formula && formula.startsWith("=")) {
          cell.formulas = [[`=(${formula})*1000`]];
        } else {
          const val = range.values[r][c] as number;
          if (typeof val === "number") {
            cell.values = [[val * 1000]];
          }
        }
      }
    }
    await context.sync();
  });
}

async function negative(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["values", "formulas"]);
    await context.sync();

    for (let r = 0; r < range.rowCount; r++) {
      for (let c = 0; c < range.columnCount; c++) {
        const formula = range.formulas[r][c] as string;
        const cell = range.getCell(r, c);

        if (formula && formula.startsWith("=")) {
          cell.formulas = [[`=-(${formula})`]];
        } else {
          const val = range.values[r][c] as number;
          if (typeof val === "number") {
            cell.values = [[-val]];
          }
        }
      }
    }
    await context.sync();
  });
}

async function specialCopy(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["formulas", "address"]);
    await context.sync();

    const formula = (range.formulas as any[][])[0][0] as string;
    specialCopyClipboard = {
      formula: formula,
      address: range.address,
    };
  });
}

async function pasteExact(): Promise<void> {
  if (!specialCopyClipboard) return;

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const formula = specialCopyClipboard!.formula;

    if (formula && formula.startsWith("=")) {
      const rows = range.rowCount;
      const cols = range.columnCount;
      const formulasGrid: string[][] = [];
      for (let r = 0; r < rows; r++) {
        const row: string[] = [];
        for (let c = 0; c < cols; c++) {
          row.push(formula);
        }
        formulasGrid.push(row);
      }
      range.formulas = formulasGrid;
    } else {
      const rows = range.rowCount;
      const cols = range.columnCount;
      const valuesGrid: any[][] = [];
      for (let r = 0; r < rows; r++) {
        const row: any[] = [];
        for (let c = 0; c < cols; c++) {
          row.push(formula);
        }
        valuesGrid.push(row);
      }
      range.values = valuesGrid;
    }
    await context.sync();
  });
}

async function pasteTranspose(): Promise<void> {
  if (!specialCopyClipboard) return;

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address"]);
    await context.sync();

    const sourceFormula = specialCopyClipboard!.formula;

    const sourceRange = context.workbook
      .worksheets
      .getActiveWorksheet()
      .getRange(specialCopyClipboard!.address);

    sourceRange.load(["rowCount", "columnCount", "formulas"]);
    await context.sync();

    const sourceRows = sourceRange.rowCount;
    const sourceCols = sourceRange.columnCount;
    const sourceFormulas = sourceRange.formulas as any[][];

    const transposed: string[][] = [];
    for (let c = 0; c < sourceCols; c++) {
      const row: string[] = [];
      for (let r = 0; r < sourceRows; r++) {
        row.push(sourceFormulas[r][c] as string);
      }
      transposed.push(row);
    }

    const targetRange = range.getCell(0, 0).getResizedRange(transposed.length - 1, transposed[0].length - 1);

    for (let r = 0; r < transposed.length; r++) {
      for (let c = 0; c < transposed[0].length; c++) {
        const cell = targetRange.getCell(r, c);
        const val = transposed[r][c];
        if (val && (val as string).startsWith("=")) {
          cell.formulas = [[val]];
        } else {
          cell.values = [[val]];
        }
      }
    }
    await context.sync();
  });
}

const actionMap: Record<string, () => Promise<void>> = {
  fontColorCycle,
  autoColor,
  cellColorCycle,
  brandCycle,
  yearCycle,
  numberFormat,
  dateFormat,
  percentageFormat,
  multipleFormat,
  increaseDecimal,
  decreaseDecimal,
  outlineCycle,
  borderBottom,
  centerAcross,
  decreaseIndent,
  increaseIndent,
  heightCycle,
  widthCycle,
  divideBy1000,
  multiplyBy1000,
  negative,
  specialCopy,
  pasteExact,
  pasteTranspose,
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    Office.actions.associate("fontColorCycle", fontColorCycle);
    Office.actions.associate("autoColor", autoColor);
    Office.actions.associate("cellColorCycle", cellColorCycle);
    Office.actions.associate("brandCycle", brandCycle);
    Office.actions.associate("yearCycle", yearCycle);
    Office.actions.associate("numberFormat", numberFormat);
    Office.actions.associate("dateFormat", dateFormat);
    Office.actions.associate("percentageFormat", percentageFormat);
    Office.actions.associate("multipleFormat", multipleFormat);
    Office.actions.associate("increaseDecimal", increaseDecimal);
    Office.actions.associate("decreaseDecimal", decreaseDecimal);
    Office.actions.associate("outlineCycle", outlineCycle);
    Office.actions.associate("borderBottom", borderBottom);
    Office.actions.associate("centerAcross", centerAcross);
    Office.actions.associate("decreaseIndent", decreaseIndent);
    Office.actions.associate("increaseIndent", increaseIndent);
    Office.actions.associate("heightCycle", heightCycle);
    Office.actions.associate("widthCycle", widthCycle);
    Office.actions.associate("divideBy1000", divideBy1000);
    Office.actions.associate("multiplyBy1000", multiplyBy1000);
    Office.actions.associate("negative", negative);
    Office.actions.associate("specialCopy", specialCopy);
    Office.actions.associate("pasteExact", pasteExact);
    Office.actions.associate("pasteTranspose", pasteTranspose);
  }
});
