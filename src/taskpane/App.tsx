import React, { useState, useCallback } from "react";
import { loadSettings, saveSettings } from "../shared/settings";
import { UserSettings, DEFAULT_SETTINGS, ColorEntry, BrandStyleEntry, HeightEntry, WidthEntry, YearSuffixEntry } from "../shared/types";

type TabId = "colors" | "formats" | "brands" | "shortcuts";

export default function App() {
  const [settings, setSettings] = useState<UserSettings>(loadSettings);
  const [activeTab, setActiveTab] = useState<TabId>("colors");
  const [saved, setSaved] = useState(false);

  const handleSave = useCallback(() => {
    saveSettings(settings);
    setSaved(true);
    setTimeout(() => setSaved(false), 1500);
  }, [settings]);

  const handleReset = useCallback(() => {
    const defaults: UserSettings = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
    setSettings(defaults);
    saveSettings(defaults);
  }, []);

  const updateFontColor = useCallback((index: number, color: string) => {
    setSettings((prev) => {
      const updated = { ...prev };
      const cycle = [...updated.fontColorCycle];
      cycle[index] = { ...cycle[index], fontColor: color };
      updated.fontColorCycle = cycle;
      return updated;
    });
  }, []);

  const updateCellColor = useCallback((index: number, color: string) => {
    setSettings((prev) => {
      const updated = { ...prev };
      const cycle = [...updated.cellColorCycle];
      cycle[index] = { ...cycle[index], fillColor: color };
      updated.cellColorCycle = cycle;
      return updated;
    });
  }, []);

  const updateBrandStyle = useCallback(
    (index: number, field: keyof BrandStyleEntry, value: string) => {
      setSettings((prev) => {
        const updated = { ...prev };
        const brands = [...updated.brandStyles];
        brands[index] = { ...brands[index], [field]: value };
        updated.brandStyles = brands;
        return updated;
      });
    },
    []
  );

  const updateYearSuffix = useCallback((index: number, suffix: string) => {
    setSettings((prev) => {
      const updated = { ...prev };
      const suffixes = [...updated.yearSuffixes];
      suffixes[index] = { ...suffixes[index], suffix };
      updated.yearSuffixes = suffixes;
      return updated;
    });
  }, []);

  const updateHeight = useCallback((index: number, value: number) => {
    setSettings((prev) => {
      const updated = { ...prev };
      const heights = [...updated.heightCycle];
      heights[index] = { ...heights[index], value };
      updated.heightCycle = heights;
      return updated;
    });
  }, []);

  const updateWidth = useCallback((index: number, value: number) => {
    setSettings((prev) => {
      const updated = { ...prev };
      const widths = [...updated.widthCycle];
      widths[index] = { ...widths[index], value };
      updated.widthCycle = widths;
      return updated;
    });
  }, []);

  const tabs: { id: TabId; label: string }[] = [
    { id: "colors", label: "Colors" },
    { id: "formats", label: "Formats" },
    { id: "brands", label: "Brands" },
    { id: "shortcuts", label: "Shortcuts" },
  ];

  return (
    <div>
      <h1>QuickCel Clone</h1>
      <p className="subtitle">Customize your keyboard shortcuts</p>

      <div className="tabs">
        {tabs.map((t) => (
          <button
            key={t.id}
            className={`tab ${activeTab === t.id ? "active" : ""}`}
            onClick={() => setActiveTab(t.id)}
          >
            {t.label}
          </button>
        ))}
      </div>

      {activeTab === "colors" && (
        <ColorsTab
          settings={settings}
          onUpdateFontColor={updateFontColor}
          onUpdateCellColor={updateCellColor}
          onUpdateYearSuffix={updateYearSuffix}
        />
      )}
      {activeTab === "formats" && (
        <FormatsTab
          settings={settings}
          onUpdateHeight={updateHeight}
          onUpdateWidth={updateWidth}
        />
      )}
      {activeTab === "brands" && (
        <BrandsTab
          settings={settings}
          onUpdateBrand={updateBrandStyle}
        />
      )}
      {activeTab === "shortcuts" && (
        <ShortcutsTab settings={settings} />
      )}

      <div className="actions">
        <button className="btn btn-primary" onClick={handleSave}>
          Save Settings
        </button>
        <button className="btn btn-danger" onClick={handleReset}>
          Reset to Defaults
        </button>
      </div>

      {saved && <div className="notification show">Settings saved!</div>}
    </div>
  );
}

function ColorsTab({
  settings,
  onUpdateFontColor,
  onUpdateCellColor,
  onUpdateYearSuffix,
}: {
  settings: UserSettings;
  onUpdateFontColor: (i: number, color: string) => void;
  onUpdateCellColor: (i: number, color: string) => void;
  onUpdateYearSuffix: (i: number, suffix: string) => void;
}) {
  return (
    <div>
      <div className="section">
        <h2>Font Color Cycle (Ctrl+')</h2>
        <ul className="cycle-list">
          {settings.fontColorCycle.map((entry: ColorEntry, i: number) => (
            <li key={i}>
              <span className="step-number">{i + 1}</span>
              <span>{entry.name}</span>
              <input
                type="color"
                value={entry.fontColor || "#000000"}
                onChange={(e) => onUpdateFontColor(i, e.target.value)}
              />
              {i < settings.fontColorCycle.length - 1 && (
                <span className="arrow">→</span>
              )}
            </li>
          ))}
        </ul>
      </div>

      <div className="section">
        <h2>Cell Color Cycle (Ctrl+Shift+K)</h2>
        <ul className="cycle-list">
          {settings.cellColorCycle.map((entry: ColorEntry, i: number) => (
            <li key={i}>
              <span className="step-number">{i + 1}</span>
              <span>{entry.name}</span>
              <input
                type="color"
                value={entry.fillColor || "#FFFFFF"}
                onChange={(e) => onUpdateCellColor(i, e.target.value)}
              />
              {i < settings.cellColorCycle.length - 1 && (
                <span className="arrow">→</span>
              )}
            </li>
          ))}
        </ul>
      </div>

      <div className="section">
        <h2>Year Cycle (Ctrl+Shift+Y)</h2>
        <ul className="cycle-list">
          {settings.yearSuffixes.map((entry: YearSuffixEntry, i: number) => (
            <li key={i}>
              <span className="step-number">{i + 1}</span>
              <span>{entry.name}</span>
              <input
                type="text"
                value={entry.suffix}
                onChange={(e) => onUpdateYearSuffix(i, e.target.value)}
                placeholder="suffix"
              />
              {i < settings.yearSuffixes.length - 1 && (
                <span className="arrow">→</span>
              )}
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
}

function FormatsTab({
  settings,
  onUpdateHeight,
  onUpdateWidth,
}: {
  settings: UserSettings;
  onUpdateHeight: (i: number, v: number) => void;
  onUpdateWidth: (i: number, v: number) => void;
}) {
  return (
    <div>
      <div className="section">
        <h2>Height Cycle (Ctrl+Shift+H)</h2>
        <ul className="cycle-list">
          {settings.heightCycle.map((entry: HeightEntry, i: number) => (
            <li key={i}>
              <span className="step-number">{i + 1}</span>
              <span>{entry.name}</span>
              <input
                type="number"
                value={entry.value}
                onChange={(e) => onUpdateHeight(i, parseFloat(e.target.value) || 0)}
                style={{ width: 60, fontSize: 12, padding: "3px 6px", border: "1px solid #ccc", borderRadius: 3 }}
              />
              <span>pt</span>
            </li>
          ))}
        </ul>
      </div>

      <div className="section">
        <h2>Width Cycle (Ctrl+Shift+W)</h2>
        <ul className="cycle-list">
          {settings.widthCycle.map((entry: WidthEntry, i: number) => (
            <li key={i}>
              <span className="step-number">{i + 1}</span>
              <span>{entry.name}</span>
              <input
                type="number"
                value={entry.value}
                step="0.01"
                onChange={(e) => onUpdateWidth(i, parseFloat(e.target.value) || 0)}
                style={{ width: 60, fontSize: 12, padding: "3px 6px", border: "1px solid #ccc", borderRadius: 3 }}
              />
              <span>ch</span>
            </li>
          ))}
        </ul>
      </div>

      <div className="section">
        <h2>Number Formats Reference</h2>
        <div className="shortcut-row"><label>Number (Ctrl+Shift+1)</label> <kbd>#,##0;(#,##0);"−"</kbd></div>
        <div className="shortcut-row"><label>Date (Ctrl+Shift+3)</label> <kbd>MMM-YY</kbd></div>
        <div className="shortcut-row"><label>Percent (Ctrl+Shift+5)</label> <kbd>0%;(0%);"−"</kbd></div>
        <div className="shortcut-row"><label>Multiple (Ctrl+Shift+8)</label> <kbd>0.0"x";(0.0"x");"−"</kbd></div>
      </div>
    </div>
  );
}

function BrandsTab({
  settings,
  onUpdateBrand,
}: {
  settings: UserSettings;
  onUpdateBrand: (i: number, field: keyof BrandStyleEntry, value: string) => void;
}) {
  return (
    <div>
      <h2>Brand Styles (Ctrl+Shift+J)</h2>
      {settings.brandStyles.map((brand: BrandStyleEntry, i: number) => (
        <div className="brand-card" key={i}>
          <div className="brand-header">
            <div
              className="brand-swatch"
              style={{ background: brand.fillColor, color: brand.fontColor, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: "bold" }}
            >
              Aa
            </div>
            <strong>Brand {i + 1}</strong>
          </div>
          <div className="brand-fields">
            <div className="field-group">
              <label>Fill</label>
              <input
                type="color"
                value={brand.fillColor}
                onChange={(e) => onUpdateBrand(i, "fillColor", e.target.value)}
              />
            </div>
            <div className="field-group">
              <label>Font</label>
              <input
                type="color"
                value={brand.fontColor}
                onChange={(e) => onUpdateBrand(i, "fontColor", e.target.value)}
              />
            </div>
            <div className="field-group">
              <label>Border Color</label>
              <input
                type="color"
                value={brand.borderColor}
                onChange={(e) => onUpdateBrand(i, "borderColor", e.target.value)}
              />
            </div>
            <div className="field-group">
              <label>Border Style</label>
              <select
                value={brand.borderStyle}
                onChange={(e) => onUpdateBrand(i, "borderStyle", e.target.value)}
                style={{ fontSize: 12, padding: "3px 6px" }}
              >
                <option value="Solid">Solid</option>
                <option value="Dash">Dashed</option>
                <option value="Dot">Dotted</option>
              </select>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

function ShortcutsTab({ settings }: { settings: UserSettings }) {
  const shortcuts = [
    { action: "Font Color Cycle", key: settings.shortcuts.fontColorCycle },
    { action: "Auto Color", key: settings.shortcuts.autoColor },
    { action: "Cell Color Cycle", key: settings.shortcuts.cellColorCycle },
    { action: "Brand Cycle", key: settings.shortcuts.brandCycle },
    { action: "Year Cycle", key: settings.shortcuts.yearCycle },
    { action: "Number Format", key: settings.shortcuts.numberFormat },
    { action: "Date Format", key: settings.shortcuts.dateFormat },
    { action: "Percentage Format", key: settings.shortcuts.percentageFormat },
    { action: "Multiple Format", key: settings.shortcuts.multipleFormat },
    { action: "Increase Decimal", key: settings.shortcuts.increaseDecimal },
    { action: "Decrease Decimal", key: settings.shortcuts.decreaseDecimal },
    { action: "Outline Cycle", key: settings.shortcuts.outlineCycle },
    { action: "Border Bottom", key: settings.shortcuts.borderBottom },
    { action: "Center Across", key: settings.shortcuts.centerAcross },
    { action: "Decrease Indent", key: settings.shortcuts.decreaseIndent },
    { action: "Increase Indent", key: settings.shortcuts.increaseIndent },
    { action: "Height Cycle", key: settings.shortcuts.heightCycle },
    { action: "Width Cycle", key: settings.shortcuts.widthCycle },
    { action: "Divide by 1000", key: settings.shortcuts.divideBy1000 },
    { action: "Multiply by 1000", key: settings.shortcuts.multiplyBy1000 },
    { action: "Negative", key: settings.shortcuts.negative },
    { action: "Special Copy", key: settings.shortcuts.specialCopy },
    { action: "Paste Exact", key: settings.shortcuts.pasteExact },
    { action: "Paste Transpose", key: settings.shortcuts.pasteTranspose },
  ];

  return (
    <div>
      <h2>All Shortcuts</h2>
      {shortcuts.map((s, i) => (
        <div className="shortcut-row" key={i}>
          <label>{s.action}</label>
          <kbd>{s.key}</kbd>
        </div>
      ))}
    </div>
  );
}
