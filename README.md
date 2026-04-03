# QuickCel Clone - Excel Add-in

Clone de QuickCel avec raccourcis clavier personnalisables pour le financial modeling.

## Architecture hybride

- **VSTO (C#/.NET 4.8)** : Intercepte les raccourcis clavier au niveau système (peut surcharger les raccourcis natifs Excel)
- **Office.js (React/TypeScript)** : Panneau de settings web pour personnaliser les cycles de couleurs, formats, etc.

## Comment build le VSTO

### Prérequis (Windows uniquement)
- Visual Studio 2019/2022 avec le workload "Office/SharePoint development"
- .NET Framework 4.8
- Excel 2016+ (desktop)

### Build
1. Ouvrir le dossier `vsto/` dans Visual Studio (ou créer un nouveau projet VSTO Excel Add-in et copier les fichiers)
2. Installer le package NuGet `Newtonsoft.Json` via le Package Manager Console :
   ```
   Install-Package Newtonsoft.Json -Version 13.0.3
   ```
3. Build la solution

### Méthode recommandée : Créer le projet dans Visual Studio

La façon la plus simple est de laisser Visual Studio générer la structure du projet :

1. Ouvrir Visual Studio
2. **File > New > Project**
3. Chercher "Excel VSTO Add-in"
4. Nommer le projet `QuickCelVSTO`
5. Une fois le projet créé, **remplacer** les fichiers générés par nos fichiers :
   - Remplacer `ThisAddIn.cs` par notre version
   - Ajouter `KeyboardShortcutManager.cs`
   - Ajouter `ExcelOperations.cs`
   - Ajouter `SettingsManager.cs`
   - Ajouter `Ribbon.cs` + `Ribbon.xml` (propriété: Embedded Resource)
6. Installer `Newtonsoft.Json` via NuGet
7. Build + Run

## Comment build le Taskpane (Office.js)

```bash
npm install
npm start        # Dev server sur https://localhost:3000
npm run build    # Build production
```

## Paramètres partagés

Les settings sont stockés dans `%APPDATA%\QuickCel\settings.json`.
Le VSTO et le taskpane web lisent/écrivent le même fichier.

## 24 raccourcis implémentés

| Shortcut | Action |
|----------|--------|
| Ctrl+' | Font Color Cycle |
| Ctrl+Alt+A | Auto Color |
| Ctrl+Shift+K | Cell Color Cycle |
| Ctrl+Shift+J | Brand Cycle |
| Ctrl+Shift+Y | Year Cycle |
| Ctrl+Shift+1 | Number Format |
| Ctrl+Shift+3 | Date Format |
| Ctrl+Shift+5 | Percentage Format |
| Ctrl+Shift+8 | Multiple Format |
| Ctrl+Alt+. | Increase Decimal |
| Ctrl+Alt+, | Decrease Decimal |
| Ctrl+Shift+7 | Outline Cycle |
| Ctrl+Shift+B | Border Bottom |
| Ctrl+Alt+Down | Center Across |
| Ctrl+Alt+Left | Decrease Indent |
| Ctrl+Alt+Right | Increase Indent |
| Ctrl+Shift+H | Height Cycle |
| Ctrl+Shift+W | Width Cycle |
| Ctrl+Shift+D | Divide by 1000 |
| Ctrl+Shift+M | Multiply by 1000 |
| Ctrl+Shift+N | Negative |
| Ctrl+Alt+X | Special Copy |
| Ctrl+Alt+E | Paste Exact |
| Ctrl+Alt+T | Paste Transpose |
