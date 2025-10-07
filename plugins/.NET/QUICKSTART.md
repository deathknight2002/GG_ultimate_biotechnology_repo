# Quick Start Guide - Biotech Plugin System

## 🚀 Getting Started in 5 Minutes

### Step 1: Verify Installation

The plugin system is now integrated into your SeleniumVBA repository. All files are located in:
```
/plugins/.NET/
```

### Step 2: Understanding the Structure

```
plugins/.NET/
├── Interfaces/         ← Plugin interface definitions (IPlugin, IBiotechPlugin, IUIPlugin)
├── Core/              ← Core infrastructure (PluginManager, PluginBase)
├── Biotech/           ← Biotech plugins (DNA, Protein analysis)
├── UI/                ← Glass OLED theme and dashboard
├── Examples/          ← Starter template plugin
├── VBA/               ← VBA integration helper
└── README.md          ← Full documentation
```

### Step 3: View the Dashboard

1. Open your browser
2. Navigate to: `file:///[YOUR_PATH]/plugins/.NET/UI/templates/dashboard.html`
3. See the Glass OLED biotech dashboard in action!

### Step 4: Use from VBA (Option A - Simple)

```vba
Sub SimpleBiotechExample()
    ' Import the PluginHelper.bas module into your VBA project first
    
    ' Analyze a DNA sequence
    Dim result As Object
    Set result = AnalyzeDNA("ATCGATCGATCG")
    
    ' Display results
    Debug.Print "Sequence Length: " & result("Length")
    Debug.Print "GC Content: " & result("GCContent")
End Sub
```

### Step 5: Use from VBA (Option B - Advanced)

```vba
Sub AdvancedBiotechExample()
    ' Initialize plugin system
    InitializePlugins
    
    ' Process Excel data
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Get sequence from cell A2
    Dim sequence As String
    sequence = ws.Range("A2").Value
    
    ' Analyze it
    Dim result As Object
    Set result = ExecutePlugin("biotech.dna.sequencing", _
        "action", "analyze", _
        "sequence", sequence)
    
    ' Write results to Excel
    ws.Range("B2").Value = result("Length")
    ws.Range("C2").Value = result("GCContent")
    
    ' Cleanup
    CleanupPlugins
End Sub
```

## 🎨 Glass OLED UI Features

The included Glass OLED theme provides:

- **Modern Design**: Glass morphism with backdrop blur effects
- **OLED Optimized**: Pure black backgrounds save power on OLED displays
- **Biotech Colors**: Neon green (#00ff88) and cyan (#00d4ff) accent colors
- **Smooth Animations**: Shimmer effects and smooth transitions
- **Responsive**: Works on desktop, tablet, and mobile

### Using the Theme

Just link to the CSS in your HTML:

```html
<link rel="stylesheet" href="../styles/glass-oled-biotech.css">
```

Then use the provided classes:

```html
<div class="glass-container">
    <h2>My Biotech Data</h2>
    <p>Content with glass effect</p>
</div>

<button class="btn-glass">Click Me</button>

<div class="metric-card">
    <div class="metric-value">42</div>
    <div class="metric-label">Samples</div>
</div>
```

## 🧬 Available Plugins

### 1. DNA Sequencing Plugin
**ID**: `biotech.dna.sequencing`

Analyzes DNA sequences with:
- Quality control
- GC content calculation
- Sequence complexity
- Motif finding

### 2. Protein Analysis Plugin
**ID**: `biotech.protein.analysis`

Analyzes proteins with:
- Molecular weight calculation
- Isoelectric point prediction
- Structure prediction
- Domain identification

### 3. Glass OLED Theme Plugin
**ID**: `ui.theme.glass-oled`

Provides modern UI with:
- Glass morphism styling
- Multiple themes
- UI components
- Responsive design

## 📊 Example: Complete Biotech Workflow

```vba
Sub CompleteBiotechWorkflow()
    ' 1. Initialize
    InitializePlugins
    
    ' 2. Launch dashboard
    Dim driver As WebDriver
    Set driver = New WebDriver
    driver.StartChrome
    driver.OpenBrowser
    
    Dim dashboardPath As String
    dashboardPath = "file:///" & Replace(ThisWorkbook.Path, "\", "/") & "/plugins/.NET/UI/templates/dashboard.html"
    driver.NavigateTo dashboardPath
    driver.Wait 1000
    
    ' 3. Process data
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim i As Long
    For i = 2 To 10 ' Process rows 2-10
        Dim sequence As String
        sequence = ws.Cells(i, 1).Value
        
        If sequence <> "" Then
            ' Analyze DNA
            Dim result As Object
            Set result = AnalyzeDNA(sequence)
            
            ' Write to Excel
            ws.Cells(i, 2).Value = result("Length")
            ws.Cells(i, 3).Value = result("GCContent")
            
            ' Update dashboard (via JavaScript)
            driver.ExecuteScript "updateSampleCount(" & i - 1 & ")"
        End If
    Next i
    
    ' 4. Generate report
    Dim reportParams As Object
    Set reportParams = CreateObject("Scripting.Dictionary")
    reportParams.Add "action", "report"
    reportParams.Add "format", "HTML"
    
    Dim report As Object
    Set report = ExecutePlugin("biotech.dna.sequencing", reportParams)
    
    ' 5. Cleanup
    MsgBox "Analysis complete! Check dashboard for results.", vbInformation
    CleanupPlugins
End Sub
```

## 🔧 Creating Your Own Plugin

1. Create a new .NET class library project
2. Reference the Interfaces and Core assemblies
3. Inherit from `PluginBase`
4. Implement required interface methods
5. Build the DLL
6. Copy to plugins directory

Example:

```csharp
using SeleniumVBA.Plugins.Core;

public class MyPlugin : PluginBase
{
    public override string Id => "mycompany.myplugin";
    public override string Name => "My Custom Plugin";
    // ... implement other properties and methods
}
```

See `Examples/StarterPlugin.cs` for a complete template.

## 📚 Resources

- **Full Documentation**: See `README.md` in this directory
- **Plugin Manifest**: See `plugin-manifest.json` for configuration
- **VBA Helper**: See `VBA/PluginHelper.bas` for integration code
- **Example Dashboard**: `UI/templates/dashboard.html`
- **Theme CSS**: `UI/styles/glass-oled-biotech.css`

## 🆘 Troubleshooting

**Problem**: Plugins not loading
- **Solution**: Ensure .NET Framework 4.7.2+ is installed
- Check plugin manifest JSON is valid
- Verify DLL files exist in correct directories

**Problem**: Theme not displaying correctly
- **Solution**: Check CSS file path is correct
- Use Chrome or Edge (Safari/Firefox may not support backdrop-filter)
- Check browser console for errors

**Problem**: VBA can't create plugin manager
- **Solution**: DLLs may need COM registration: `regasm /codebase PluginManager.dll`
- Run as administrator
- Check Windows Event Log for errors

## 🎯 Next Steps

1. Review the full `README.md` for comprehensive documentation
2. Explore the example plugins in `Biotech/` and `Examples/`
3. Customize the Glass OLED theme in `UI/styles/`
4. Create your own plugins using `StarterPlugin.cs` as template
5. Build biotech workflows combining VBA, plugins, and WebDriver

## 💡 Tips

- Use the VBA `PluginHelper` module for easy integration
- Start with the `StarterPlugin` example to learn the architecture
- The Glass OLED theme works great on OLED monitors/TVs
- Combine plugins with SeleniumVBA's web automation for powerful workflows
- Check `plugin-manifest.json` to see all available plugins and capabilities

---

**Happy Coding! 🧬💻🎨**
