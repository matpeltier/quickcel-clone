using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace QuickCelVSTO
{
    public class ExcelOperations
    {
        private readonly Dictionary<string, int> _cycleIndices = new Dictionary<string, int>();
        private string _specialCopyFormula;
        private string _specialCopyAddress;

        private Excel.Range Sel() => Globals.ThisAddIn.Application.Selection as Excel.Range;

        private int Cycle(string key, int count)
        {
            if (!_cycleIndices.ContainsKey(key)) _cycleIndices[key] = 0;
            else _cycleIndices[key] = (_cycleIndices[key] + 1) % count;
            return _cycleIndices[key];
        }

        public void FontColorCycle()
        {
            var s = SettingsManager.Settings;
            var r = Sel(); if (r == null) return;
            int i = Cycle("fontColor", s.FontColorCycle.Count);
            r.Font.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(s.FontColorCycle[i].Hex));
        }

        public void AutoColor()
        {
            var s = SettingsManager.Settings;
            string inputHex = s.FontColorCycle.Find(c => c.Name.Contains("Input"))?.Hex ?? "#0000FF";
            string linkHex = s.FontColorCycle.Find(c => c.Name.Contains("Link"))?.Hex ?? "#008000";
            int inputClr = ColorTranslator.ToOle(ColorTranslator.FromHtml(inputHex));
            int linkClr = ColorTranslator.ToOle(ColorTranslator.FromHtml(linkHex));
            int formulaClr = ColorTranslator.ToOle(Color.Black);

            var r = Sel(); if (r == null) return;
            for (int row = 1; row <= r.Rows.Count; row++)
            {
                for (int col = 1; col <= r.Columns.Count; col++)
                {
                    var cell = r.Cells[row, col] as Excel.Range;
                    string f = cell.Formula as string ?? "";
                    int clr;
                    if (string.IsNullOrEmpty(f) || !f.StartsWith("="))
                        clr = inputClr;
                    else if (f.StartsWith("='") || f.Contains("!"))
                        clr = linkClr;
                    else
                        clr = formulaClr;
                    cell.Font.Color = clr;
                    Marshal.ReleaseComObject(cell);
                }
            }
        }

        public void CellColorCycle()
        {
            var s = SettingsManager.Settings;
            var r = Sel(); if (r == null) return;
            int i = Cycle("cellColor", s.CellColorCycle.Count);
            var entry = s.CellColorCycle[i];
            if (entry.Name == "None")
                r.Interior.Pattern = XlPattern.xlPatternNone;
            else
            {
                r.Interior.Pattern = XlPattern.xlPatternSolid;
                r.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(entry.Hex));
            }
        }

        public void BrandCycle()
        {
            var s = SettingsManager.Settings;
            var r = Sel(); if (r == null) return;
            int i = Cycle("brand", s.BrandStyles.Count);
            var b = s.BrandStyles[i];
            r.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(b.FillColor));
            r.Font.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(b.FontColor));

            XlLineStyle style = b.Name.Contains("Dash") ? XlLineStyle.xlDash : b.Name.Contains("Dot") ? XlLineStyle.xlDot : XlLineStyle.xlContinuous;
            int borderClr = ColorTranslator.ToOle(ColorTranslator.FromHtml(b.BorderColor));

            r.Borders[XlBordersIndex.xlEdgeTop].LineStyle = style;
            r.Borders[XlBordersIndex.xlEdgeTop].Color = borderClr;
            r.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = style;
            r.Borders[XlBordersIndex.xlEdgeBottom].Color = borderClr;
            r.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = style;
            r.Borders[XlBordersIndex.xlEdgeLeft].Color = borderClr;
            r.Borders[XlBordersIndex.xlEdgeRight].LineStyle = style;
            r.Borders[XlBordersIndex.xlEdgeRight].Color = borderClr;
        }

        public void YearCycle()
        {
            var s = SettingsManager.Settings;
            var r = Sel(); if (r == null) return;
            int i = Cycle("year", s.YearSuffixes.Count);
            string suffix = s.YearSuffixes[i].Suffix;

            for (int row = 1; row <= r.Rows.Count; row++)
            {
                for (int col = 1; col <= r.Columns.Count; col++)
                {
                    var cell = r.Cells[row, col] as Excel.Range;
                    string f = cell.Formula as string ?? "";
                    string clean = f;
                    foreach (var ys in s.YearSuffixes)
                    {
                        if (!string.IsNullOrEmpty(ys.Suffix) && clean.EndsWith(ys.Suffix))
                        {
                            clean = clean.Substring(0, clean.Length - ys.Suffix.Length);
                            break;
                        }
                    }
                    if (!string.IsNullOrEmpty(suffix))
                    {
                        if (clean.StartsWith("="))
                            cell.Formula = clean + "&\"" + suffix + "\"";
                        else
                            cell.Value = clean + suffix;
                    }
                    else
                    {
                        if (clean.StartsWith("=")) cell.Formula = clean;
                        else cell.Value = clean;
                    }
                    Marshal.ReleaseComObject(cell);
                }
            }
        }

        public void NumberFormat()
        {
            var r = Sel(); if (r == null) return;
            r.NumberFormat = "#,##0;(#,##0);\"-\"";
        }

        public void DateFormat()
        {
            var r = Sel(); if (r == null) return;
            r.NumberFormat = "MMM-YY";
        }

        public void PercentageFormat()
        {
            var r = Sel(); if (r == null) return;
            r.NumberFormat = "0%;(0%);\"-\"";
        }

        public void MultipleFormat()
        {
            var r = Sel(); if (r == null) return;
            r.NumberFormat = "0.0\"x\";(0.0\"x\");\"-\"";
        }

        public void IncreaseDecimal()
        {
            var r = Sel(); if (r == null) return;
            string fmt = r.NumberFormat as string ?? "General";
            int d = CountDecimals(fmt);
            r.NumberFormat = BuildNumberFmt(d + 1);
        }

        public void DecreaseDecimal()
        {
            var r = Sel(); if (r == null) return;
            string fmt = r.NumberFormat as string ?? "General";
            int d = CountDecimals(fmt);
            r.NumberFormat = BuildNumberFmt(Math.Max(0, d - 1));
        }

        public void OutlineCycle()
        {
            var s = SettingsManager.Settings;
            var r = Sel(); if (r == null) return;
            int i = Cycle("outline", s.OutlineBorderCycle.Count);
            var entry = s.OutlineBorderCycle[i];

            if (entry.Style == "None")
            {
                r.Borders.LineStyle = XlLineStyle.xlLineStyleNone;
            }
            else
            {
                var style = ToXlLineStyle(entry.Style);
                var clr = entry.Name == "White" ? ColorTranslator.ToOle(Color.White) : ColorTranslator.ToOle(Color.Black);
                r.Borders[XlBordersIndex.xlEdgeTop].LineStyle = style;
                r.Borders[XlBordersIndex.xlEdgeTop].Color = clr;
                r.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = style;
                r.Borders[XlBordersIndex.xlEdgeBottom].Color = clr;
                r.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = style;
                r.Borders[XlBordersIndex.xlEdgeLeft].Color = clr;
                r.Borders[XlBordersIndex.xlEdgeRight].LineStyle = style;
                r.Borders[XlBordersIndex.xlEdgeRight].Color = clr;
            }
        }

        public void BorderBottom()
        {
            var s = SettingsManager.Settings;
            var r = Sel(); if (r == null) return;
            int i = Cycle("borderBottom", s.BorderBottomCycle.Count);
            var entry = s.BorderBottomCycle[i];
            if (entry.Style == "None")
                r.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlLineStyleNone;
            else
            {
                r.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = ToXlLineStyle(entry.Style);
                r.Borders[XlBordersIndex.xlEdgeBottom].Color = ColorTranslator.ToOle(Color.Black);
            }
        }

        public void CenterAcross()
        {
            var r = Sel(); if (r == null) return;
            r.HorizontalAlignment = XlHAlign.xlHAlignCenterAcrossSelection;
        }

        public void DecreaseIndent()
        {
            var r = Sel(); if (r == null) return;
            int cur = r.IndentLevel;
            r.IndentLevel = Math.Max(0, cur - 1);
        }

        public void IncreaseIndent()
        {
            var r = Sel(); if (r == null) return;
            int cur = r.IndentLevel;
            r.IndentLevel = Math.Min(15, cur + 1);
        }

        public void HeightCycle()
        {
            var s = SettingsManager.Settings;
            var r = Sel(); if (r == null) return;
            int i = Cycle("height", s.HeightCycle.Count);
            r.EntireRow.RowHeight = s.HeightCycle[i].Value;
        }

        public void WidthCycle()
        {
            var s = SettingsManager.Settings;
            var r = Sel(); if (r == null) return;
            int i = Cycle("width", s.WidthCycle.Count);
            r.EntireColumn.ColumnWidth = s.WidthCycle[i].Value;
        }

        public void DivideBy1000()
        {
            var r = Sel(); if (r == null) return;
            ApplyMath(r, v => v / 1000.0, "/1000");
        }

        public void MultiplyBy1000()
        {
            var r = Sel(); if (r == null) return;
            ApplyMath(r, v => v * 1000.0, "*1000");
        }

        public void Negative()
        {
            var r = Sel(); if (r == null) return;
            ApplyMath(r, v => -v, "*(-1)");
        }

        public void SpecialCopy()
        {
            var r = Sel(); if (r == null) return;
            _specialCopyFormula = r.Formula as string ?? (r.Value?.ToString() ?? "");
            _specialCopyAddress = r.Address;
        }

        public void PasteExact()
        {
            if (_specialCopyFormula == null) return;
            var r = Sel(); if (r == null) return;
            for (int row = 1; row <= r.Rows.Count; row++)
            {
                for (int col = 1; col <= r.Columns.Count; col++)
                {
                    var cell = r.Cells[row, col] as Excel.Range;
                    if (_specialCopyFormula.StartsWith("="))
                        cell.Formula = _specialCopyFormula;
                    else
                        cell.Value = _specialCopyFormula;
                    Marshal.ReleaseComObject(cell);
                }
            }
        }

        public void PasteTranspose()
        {
            if (_specialCopyFormula == null) return;
            var r = Sel(); if (r == null) return;
            var cell = r.Cells[1, 1] as Excel.Range;
            if (_specialCopyFormula.StartsWith("="))
                cell.Formula = _specialCopyFormula;
            else
                cell.Value = _specialCopyFormula;
            Marshal.ReleaseComObject(cell);
        }

        private void ApplyMath(Excel.Range r, Func<double, double> op, string formulaSuffix)
        {
            for (int row = 1; row <= r.Rows.Count; row++)
            {
                for (int col = 1; col <= r.Columns.Count; col++)
                {
                    var cell = r.Cells[row, col] as Excel.Range;
                    try
                    {
                        string f = cell.Formula as string ?? "";
                        if (f.StartsWith("="))
                            cell.Formula = "=(" + f + ")" + formulaSuffix;
                        else
                        {
                            object v = cell.Value;
                            if (v is double d) cell.Value = op(d);
                            else if (v is int iv) cell.Value = op(iv);
                            else if (v is float fv) cell.Value = op(fv);
                            else if (v != null && double.TryParse(v.ToString(), out double parsed))
                                cell.Value = op(parsed);
                        }
                    }
                    catch { }
                    Marshal.ReleaseComObject(cell);
                }
            }
        }

        private int CountDecimals(string fmt)
        {
            if (string.IsNullOrEmpty(fmt) || fmt == "General") return 2;
            int dot = fmt.IndexOf('.');
            if (dot < 0) return 0;
            int count = 0;
            for (int i = dot + 1; i < fmt.Length && fmt[i] == '0'; i++) count++;
            return count;
        }

        private string BuildNumberFmt(int decimals)
        {
            string z = decimals > 0 ? new string('0', decimals) : "";
            string dec = z.Length > 0 ? "." + z : "";
            return $"#,##0{dec};(#,##0{dec});\"-\"";
        }

        private XlLineStyle ToXlLineStyle(string style)
        {
            switch (style)
            {
                case "Solid": return XlLineStyle.xlContinuous;
                case "Dashed": return XlLineStyle.xlDash;
                case "Dotted": return XlLineStyle.xlDot;
                default: return XlLineStyle.xlLineStyleNone;
            }
        }
    }
}
