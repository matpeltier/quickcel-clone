using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace QuickCelVSTO
{
    public class QuickCelSettings
    {
        public List<ColorEntry> FontColorCycle { get; set; }
        public List<ColorEntry> CellColorCycle { get; set; }
        public List<BrandStyle> BrandStyles { get; set; }
        public List<YearSuffix> YearSuffixes { get; set; }
        public List<BorderEntry> OutlineBorderCycle { get; set; }
        public List<BorderEntry> BorderBottomCycle { get; set; }
        public List<SizeEntry> HeightCycle { get; set; }
        public List<SizeEntry> WidthCycle { get; set; }
    }

    public class ColorEntry { public string Name { get; set; } public string Hex { get; set; } }
    public class BrandStyle { public string Name { get; set; } public string FillColor { get; set; } public string FontColor { get; set; } public string BorderColor { get; set; } }
    public class YearSuffix { public string Name { get; set; } public string Suffix { get; set; } }
    public class BorderEntry { public string Name { get; set; } public string Style { get; set; } }
    public class SizeEntry { public string Name { get; set; } public double Value { get; set; } }

    public static class SettingsManager
    {
        private static readonly string Dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "QuickCel");
        private static readonly string File = Path.Combine(Dir, "settings.json");

        public static QuickCelSettings Settings { get; private set; } = CreateDefaults();

        public static void Reload()
        {
            try
            {
                if (System.IO.File.Exists(File))
                {
                    string json = System.IO.File.ReadAllText(File);
                    Settings = JsonConvert.DeserializeObject<QuickCelSettings>(json) ?? CreateDefaults();
                }
            }
            catch { }
        }

        public static void Save(QuickCelSettings s)
        {
            try
            {
                System.IO.Directory.CreateDirectory(Dir);
                System.IO.File.WriteAllText(File, JsonConvert.SerializeObject(s, Formatting.Indented));
                Settings = s;
            }
            catch { }
        }

        public static QuickCelSettings CreateDefaults()
        {
            return new QuickCelSettings
            {
                FontColorCycle = new List<ColorEntry>
                {
                    new ColorEntry { Name = "Black", Hex = "#000000" },
                    new ColorEntry { Name = "Blue (Input)", Hex = "#0000FF" },
                    new ColorEntry { Name = "Green (Link)", Hex = "#008000" },
                    new ColorEntry { Name = "Gray (Constant)", Hex = "#808080" }
                },
                CellColorCycle = new List<ColorEntry>
                {
                    new ColorEntry { Name = "Yellow", Hex = "#FFFF00" },
                    new ColorEntry { Name = "Light Gray", Hex = "#D9D9D9" },
                    new ColorEntry { Name = "Light Blue", Hex = "#BDD7EE" },
                    new ColorEntry { Name = "Cracked Gray", Hex = "#D6DCE4" }
                },
                BrandStyles = new List<BrandStyle>
                {
                    new BrandStyle { Name = "Brand 1", FillColor = "#4472C4", FontColor = "#FFFFFF", BorderColor = "#4472C4" },
                    new BrandStyle { Name = "Brand 2", FillColor = "#ED7D31", FontColor = "#FFFFFF", BorderColor = "#ED7D31" },
                    new BrandStyle { Name = "Brand 3", FillColor = "#70AD47", FontColor = "#FFFFFF", BorderColor = "#70AD47" }
                },
                YearSuffixes = new List<YearSuffix>
                {
                    new YearSuffix { Name = "None", Suffix = "" },
                    new YearSuffix { Name = "Actual", Suffix = "A" },
                    new YearSuffix { Name = "Budget", Suffix = "B" },
                    new YearSuffix { Name = "Estimated", Suffix = "E" }
                },
                OutlineBorderCycle = new List<BorderEntry>
                {
                    new BorderEntry { Name = "Solid", Style = "Solid" },
                    new BorderEntry { Name = "Dashed", Style = "Dashed" },
                    new BorderEntry { Name = "Dotted", Style = "Dotted" },
                    new BorderEntry { Name = "White", Style = "Solid" },
                    new BorderEntry { Name = "None", Style = "None" }
                },
                BorderBottomCycle = new List<BorderEntry>
                {
                    new BorderEntry { Name = "Solid", Style = "Solid" },
                    new BorderEntry { Name = "Dashed", Style = "Dashed" },
                    new BorderEntry { Name = "Dotted", Style = "Dotted" },
                    new BorderEntry { Name = "None", Style = "None" }
                },
                HeightCycle = new List<SizeEntry>
                {
                    new SizeEntry { Name = "Normal", Value = 15 },
                    new SizeEntry { Name = "Thin", Value = 5 }
                },
                WidthCycle = new List<SizeEntry>
                {
                    new SizeEntry { Name = "Normal", Value = 8.43 },
                    new SizeEntry { Name = "Thin", Value = 1 }
                }
            };
        }
    }
}
