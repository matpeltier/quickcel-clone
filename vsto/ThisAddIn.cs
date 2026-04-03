using System;

namespace QuickCelVSTO
{
    public partial class ThisAddIn
    {
        private KeyboardShortcutManager _keyboardManager;
        private ExcelOperations _operations;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            SettingsManager.Reload();
            _operations = new ExcelOperations();
            _keyboardManager = new KeyboardShortcutManager();
            RegisterShortcuts();
            _keyboardManager.Start();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _keyboardManager?.Stop();
            _keyboardManager?.Dispose();
        }

        private void RegisterShortcuts()
        {
            var k = _keyboardManager;

            // Colors
            k.RegisterShortcut(true, false, false, Keys.Oemtilde, _operations.FontColorCycle);
            k.RegisterShortcut(true, false, true, Keys.A, _operations.AutoColor);
            k.RegisterShortcut(true, true, false, Keys.K, _operations.CellColorCycle);
            k.RegisterShortcut(true, true, false, Keys.J, _operations.BrandCycle);
            k.RegisterShortcut(true, true, false, Keys.Y, _operations.YearCycle);

            // Number formats
            k.RegisterShortcut(true, true, false, Keys.D1, _operations.NumberFormat);
            k.RegisterShortcut(true, true, false, Keys.D3, _operations.DateFormat);
            k.RegisterShortcut(true, true, false, Keys.D5, _operations.PercentageFormat);
            k.RegisterShortcut(true, true, false, Keys.D8, _operations.MultipleFormat);
            k.RegisterShortcut(true, false, true, Keys.OemPeriod, _operations.IncreaseDecimal);
            k.RegisterShortcut(true, false, true, Keys.Oemcomma, _operations.DecreaseDecimal);

            // Borders
            k.RegisterShortcut(true, true, false, Keys.D7, _operations.OutlineCycle);
            k.RegisterShortcut(true, true, false, Keys.B, _operations.BorderBottom);

            // Alignment
            k.RegisterShortcut(true, false, true, Keys.Down, _operations.CenterAcross);
            k.RegisterShortcut(true, false, true, Keys.Left, _operations.DecreaseIndent);
            k.RegisterShortcut(true, false, true, Keys.Right, _operations.IncreaseIndent);
            k.RegisterShortcut(true, true, false, Keys.H, _operations.HeightCycle);
            k.RegisterShortcut(true, true, false, Keys.W, _operations.WidthCycle);

            // Math
            k.RegisterShortcut(true, true, false, Keys.D, _operations.DivideBy1000);
            k.RegisterShortcut(true, true, false, Keys.M, _operations.MultiplyBy1000);
            k.RegisterShortcut(true, true, false, Keys.N, _operations.Negative);

            // Copy/Paste
            k.RegisterShortcut(true, false, true, Keys.X, _operations.SpecialCopy);
            k.RegisterShortcut(true, false, true, Keys.E, _operations.PasteExact);
            k.RegisterShortcut(true, false, true, Keys.T, _operations.PasteTranspose);
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility[] CreateRibbonExtensibilityObject()
        {
            return new Microsoft.Office.Core.IRibbonExtensibility[] { new Ribbon() };
        }

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}
