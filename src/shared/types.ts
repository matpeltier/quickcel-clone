export interface ColorEntry {
  name: string;
  fontColor?: string;
  fillColor?: string;
}

export interface BorderStyleEntry {
  name: string;
  style: string;
}

export interface BrandStyleEntry {
  name: string;
  fillColor: string;
  fontColor: string;
  borderStyle: string;
  borderColor: string;
}

export interface YearSuffixEntry {
  name: string;
  suffix: string;
}

export interface HeightEntry {
  name: string;
  value: number;
}

export interface WidthEntry {
  name: string;
  value: number;
}

export interface UserSettings {
  fontColorCycle: ColorEntry[];
  cellColorCycle: ColorEntry[];
  brandStyles: BrandStyleEntry[];
  yearSuffixes: YearSuffixEntry[];
  outlineBorderCycle: BorderStyleEntry[];
  borderBottomCycle: BorderStyleEntry[];
  heightCycle: HeightEntry[];
  widthCycle: WidthEntry[];
  shortcuts: Record<string, string>;
}

export const DEFAULT_SETTINGS: UserSettings = {
  fontColorCycle: [
    { name: "Black", fontColor: "#000000" },
    { name: "Blue (Input)", fontColor: "#0000FF" },
    { name: "Green (Link)", fontColor: "#008000" },
    { name: "Gray (Constant)", fontColor: "#808080" },
  ],
  cellColorCycle: [
    { name: "Yellow", fillColor: "#FFFF00" },
    { name: "Light Gray", fillColor: "#D9D9D9" },
    { name: "Light Blue", fillColor: "#BDD7EE" },
    { name: "Cracked Gray", fillColor: "#D6DCE4" },
  ],
  brandStyles: [
    {
      name: "Brand 1",
      fillColor: "#4472C4",
      fontColor: "#FFFFFF",
      borderStyle: "Solid",
      borderColor: "#4472C4",
    },
    {
      name: "Brand 2",
      fillColor: "#ED7D31",
      fontColor: "#FFFFFF",
      borderStyle: "Solid",
      borderColor: "#ED7D31",
    },
    {
      name: "Brand 3",
      fillColor: "#70AD47",
      fontColor: "#FFFFFF",
      borderStyle: "Solid",
      borderColor: "#70AD47",
    },
  ],
  yearSuffixes: [
    { name: "None", suffix: "" },
    { name: "Actual", suffix: "A" },
    { name: "Budget", suffix: "B" },
    { name: "Estimated", suffix: "E" },
  ],
  outlineBorderCycle: [
    { name: "Solid", style: "Continuous" },
    { name: "Dashed", style: "Dash" },
    { name: "Dotted", style: "Dot" },
    { name: "White", style: "Continuous" },
    { name: "None", style: "None" },
  ],
  borderBottomCycle: [
    { name: "Solid", style: "Continuous" },
    { name: "Dashed", style: "Dash" },
    { name: "Dotted", style: "Dot" },
    { name: "None", style: "None" },
  ],
  heightCycle: [
    { name: "Normal", value: 15 },
    { name: "Thin", value: 5 },
  ],
  widthCycle: [
    { name: "Normal", value: 8.43 },
    { name: "Thin", value: 1 },
  ],
  shortcuts: {
    fontColorCycle: "Ctrl+'",
    autoColor: "Ctrl+Alt+A",
    cellColorCycle: "Ctrl+Shift+K",
    brandCycle: "Ctrl+Shift+J",
    yearCycle: "Ctrl+Shift+Y",
    numberFormat: "Ctrl+Shift+1",
    dateFormat: "Ctrl+Shift+3",
    percentageFormat: "Ctrl+Shift+5",
    multipleFormat: "Ctrl+Shift+8",
    increaseDecimal: "Ctrl+Alt+.",
    decreaseDecimal: "Ctrl+Alt+,",
    outlineCycle: "Ctrl+Shift+7",
    borderBottom: "Ctrl+Shift+B",
    centerAcross: "Ctrl+Alt+Down",
    decreaseIndent: "Ctrl+Alt+Left",
    increaseIndent: "Ctrl+Alt+Right",
    heightCycle: "Ctrl+Shift+H",
    widthCycle: "Ctrl+Shift+W",
    divideBy1000: "Ctrl+Shift+D",
    multiplyBy1000: "Ctrl+Shift+M",
    negative: "Ctrl+Shift+N",
    specialCopy: "Ctrl+Alt+X",
    pasteExact: "Ctrl+Alt+E",
    pasteTranspose: "Ctrl+Alt+T",
  },
};
