import { BorderStyleEntry } from "./types";

function rgbToHex(r: number, g: number, b: number): string {
  return (
    "#" +
    [r, g, b]
      .map((c) => {
        const hex = Math.round(c * 255).toString(16);
        return hex.length === 1 ? "0" + hex : hex;
      })
      .join("")
  );
}

function hexToRgb(hex: string): { r: number; g: number; b: number } {
  const cleaned = hex.replace("#", "");
  return {
    r: parseInt(cleaned.substring(0, 2), 16) / 255,
    g: parseInt(cleaned.substring(2, 4), 16) / 255,
    b: parseInt(cleaned.substring(4, 6), 16) / 255,
  };
}

export function buildNumberFormat(numDecimals: number): string {
  const z = "0".repeat(numDecimals);
  const pos = `#,##0${z ? "." + z : ""}`;
  const neg = `(#,##0${z ? "." + z : ""})`;
  const zero = `"-"${z ? `0.${"0".repeat(numDecimals - 1)}` : ""}`;
  return `${pos};${neg};${zero}`;
}

export function buildPercentFormat(numDecimals: number): string {
  const z = "0".repeat(numDecimals);
  const pos = `0${z ? "." + z : ""}%`;
  const neg = `(0${z ? "." + z : ""}%)`;
  const zero = `"-"`;
  return `${pos};${neg};${zero}`;
}

export function buildMultipleFormat(numDecimals: number): string {
  const z = "0".repeat(numDecimals);
  const pos = `0${z ? "." + z : ""}"x"`;
  const neg = `(0${z ? "." + z : ""}"x")`;
  const zero = `"-"`;
  return `${pos};${neg};${zero}`;
}

export const DATE_FORMAT = "MMM-YY";

export function getBorderStyleForExcel(style: string): Excel.BorderLineStyle {
  const map: Record<string, Excel.BorderLineStyle> = {
    Continuous: "Continuous" as Excel.BorderLineStyle,
    Dash: "Dash" as Excel.BorderLineStyle,
    Dot: "Dot" as Excel.BorderLineStyle,
    None: "None" as Excel.BorderLineStyle,
  };
  return map[style] || ("None" as Excel.BorderLineStyle);
}

export function getBorderColorForEntry(entry: BorderStyleEntry): string {
  if (entry.name === "White") return "#FFFFFF";
  return "#000000";
}

export function detectCellType(formula: string | undefined, value: string | undefined): "input" | "link" | "formula" | "constant" {
  if (!formula || formula === "") {
    return value && !isNaN(Number(value)) ? "constant" : "constant";
  }
  if (formula.startsWith("='") || formula.startsWith("=[") || formula.match(/^='?[^!]+!/) !== null) {
    return "link";
  }
  if (formula.startsWith("=")) {
    return "formula";
  }
  return "input";
}

export { hexToRgb, rgbToHex };
