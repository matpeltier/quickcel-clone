using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace QuickCelVSTO
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        public string GetCustomUI(string ribbonID)
        {
            if (ribbonID == "Microsoft.Excel.Workbook")
            {
                string path = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "Ribbon.xml");
                if (File.Exists(path))
                    return File.ReadAllText(path);
            }
            return string.Empty;
        }

        public void OnSettingsButton(IRibbonControl control)
        {
            SettingsManager.Reload();
            System.Windows.Forms.MessageBox.Show(
                "Settings reloaded.\n\nFile: %APPDATA%\\QuickCel\\settings.json",
                "QuickCel",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }
    }
}
