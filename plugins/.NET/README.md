# SeleniumVBA .NET Plugin System

## Overview

The SeleniumVBA .NET Plugin System provides a powerful, extensible architecture for adding custom functionality to SeleniumVBA. This system is specifically enhanced for biotechnology applications with modern Glass/OLED UI styling.

## Features

- **Plug-and-Play Architecture**: Simply drop DLLs into the plugin directory
- **Biotech-Specific Plugins**: DNA sequencing, protein analysis, data visualization
- **Modern UI Themes**: Glass morphism with OLED-optimized colors
- **Easy Integration**: Works seamlessly with VBA through COM interop
- **Extensible**: Create custom plugins using the provided interfaces

## Quick Start

### 1. Plugin Directory Structure

```
plugins/
├── .NET/
│   ├── Interfaces/          # Plugin interface definitions
│   │   ├── IPlugin.cs
│   │   ├── IBiotechPlugin.cs
│   │   └── IUIPlugin.cs
│   ├── Core/               # Core plugin infrastructure
│   │   ├── PluginManager.cs
│   │   └── PluginBase.cs
│   ├── Biotech/            # Biotech-specific plugins
│   │   ├── DNASequencingPlugin.cs
│   │   └── ProteinAnalysisPlugin.cs
│   ├── UI/                 # UI enhancement plugins
│   │   ├── GlassOLEDThemePlugin.cs
│   │   ├── styles/
│   │   │   └── glass-oled-biotech.css
│   │   ├── templates/
│   │   │   └── dashboard.html
│   │   └── scripts/
│   │       └── dashboard.js
│   └── Examples/           # Example plugins
│       └── StarterPlugin.cs
```

### 2. Building Plugins

#### Requirements
- .NET Framework 4.7.2 or higher
- Visual Studio 2019 or later (or .NET CLI)

#### Building a Plugin DLL

```bash
# Using .NET CLI
cd plugins/.NET/YourPlugin
dotnet build -c Release

# The compiled DLL will be in bin/Release/
```

#### Creating a Plugin Project

```bash
dotnet new classlib -n MyBiotechPlugin
cd MyBiotechPlugin
dotnet add reference ../Interfaces/
dotnet add reference ../Core/
```

### 3. Creating a Custom Plugin

```csharp
using System;
using System.Collections.Generic;
using SeleniumVBA.Plugins.Core;
using SeleniumVBA.Plugins.Interfaces;

namespace MyCompany.Plugins
{
    public class MyCustomPlugin : PluginBase, IBiotechPlugin
    {
        public override string Id => "mycompany.custom";
        public override string Name => "My Custom Plugin";
        public override Version Version => new Version(1, 0, 0);
        public override string Author => "My Company";
        public override string Description => "Custom biotech analysis";
        public override string Category => "Biotech";

        public override object Execute(Dictionary<string, object> parameters)
        {
            // Your plugin logic here
            return new Dictionary<string, object>
            {
                { "Success", true },
                { "Result", "Plugin executed successfully" }
            };
        }

        // Implement IBiotechPlugin interface methods
        public DataTable ProcessBiologicalData(DataTable data, string processingType)
        {
            // Process biotech data
            return data;
        }

        // ... other interface methods
    }
}
```

## Using Plugins from VBA

### Loading the Plugin System

```vba
' In VBA module
Sub InitializePluginSystem()
    Dim pluginManager As Object
    Set pluginManager = CreateObject("SeleniumVBA.Plugins.Core.PluginManager")
    
    ' Add plugin directory
    pluginManager.AddPluginDirectory("C:\path\to\plugins")
    
    ' Load all plugins
    pluginManager.LoadPlugins()
    
    ' Get available plugins
    Dim plugins As Object
    Set plugins = pluginManager.GetAllPlugins()
    
    Debug.Print "Loaded " & plugins.Count & " plugins"
End Sub
```

### Executing a Plugin

```vba
Sub ExecuteBiotechPlugin()
    Dim pluginManager As Object
    Dim parameters As Object
    Dim result As Object
    
    Set pluginManager = CreateObject("SeleniumVBA.Plugins.Core.PluginManager")
    pluginManager.AddPluginDirectory("C:\path\to\plugins")
    pluginManager.LoadPlugins()
    
    ' Create parameters dictionary
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "action", "analyze"
    parameters.Add "sequence", "ATCGATCGATCG"
    
    ' Execute plugin
    Set result = pluginManager.ExecutePlugin("biotech.dna.sequencing", parameters)
    
    ' Process results
    Debug.Print "Length: " & result("Length")
    Debug.Print "GC Content: " & result("GCContent")
End Sub
```

## Available Plugins

### 1. DNA Sequencing Plugin

**ID**: `biotech.dna.sequencing`

**Features**:
- Quality control analysis
- Sequence alignment
- Motif finding
- GC content calculation

**Usage**:
```vba
parameters.Add "action", "analyze"
parameters.Add "sequence", "ATCGATCG..."
Set result = pluginManager.ExecutePlugin("biotech.dna.sequencing", parameters)
```

### 2. Protein Analysis Plugin

**ID**: `biotech.protein.analysis`

**Features**:
- Structure prediction
- Domain identification
- Molecular weight calculation
- Isoelectric point calculation

**Usage**:
```vba
parameters.Add "action", "calculateproperties"
parameters.Add "sequence", "MKWVTFISLLFLFSSAYS..."
Set result = pluginManager.ExecutePlugin("biotech.protein.analysis", parameters)
```

### 3. Glass OLED Theme Plugin

**ID**: `ui.theme.glass-oled`

**Features**:
- Modern glass morphism styling
- OLED-optimized dark theme
- Dynamic components
- Responsive design

**Usage**:
```vba
parameters.Add "action", "applytheme"
parameters.Add "theme", "glass-biotech"
Set result = pluginManager.ExecutePlugin("ui.theme.glass-oled", parameters)
```

## UI Themes

### Glass Biotech Theme

A modern, reflective glass morphism design with biotech-specific color schemes:

- **Primary Color**: Neon green (#00ff88) - DNA/life theme
- **Secondary Color**: Cyan (#00d4ff) - Technology theme
- **Background**: Pure black (#000000) - OLED optimized
- **Effects**: Backdrop blur, shimmer animations, smooth transitions

### Using the Theme

1. **In HTML**:
```html
<link rel="stylesheet" href="plugins/.NET/UI/styles/glass-oled-biotech.css">
```

2. **In VBA (via WebDriver)**:
```vba
Sub ApplyBiotechTheme()
    Dim driver As New WebDriver
    driver.StartChrome
    driver.OpenBrowser
    
    ' Navigate to dashboard
    driver.NavigateTo "file:///C:/path/to/plugins/.NET/UI/templates/dashboard.html"
    
    ' The theme is automatically applied
    driver.Wait 2000
End Sub
```

## Component Library

### Glass Container
```html
<div class="glass-container">
    <h2>Your Content</h2>
    <p>Glass morphism effect applied</p>
</div>
```

### OLED Card
```html
<div class="oled-card">
    <h3>Data Analysis</h3>
    <p>Card with shimmer effect</p>
</div>
```

### Metric Card
```html
<div class="metric-card">
    <div class="metric-value">42</div>
    <div class="metric-label">Active Samples</div>
</div>
```

### Glass Button
```html
<button class="btn-glass" onclick="yourFunction()">
    Click Me
</button>
```

### Data Table
```html
<table class="data-table">
    <thead>
        <tr>
            <th>Sample ID</th>
            <th>Status</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>BT-001</td>
            <td><span class="status-indicator status-active"></span>Active</td>
        </tr>
    </tbody>
</table>
```

## Plugin Development Guide

### Interface Hierarchy

1. **IPlugin** - Base interface for all plugins
   - Initialize(), Execute(), Dispose()
   - GetCapabilities()

2. **IBiotechPlugin** - Extends IPlugin for biotech
   - ProcessBiologicalData()
   - GenerateReport()
   - PerformAnalysis()
   - Visualize()
   - ValidateData()

3. **IUIPlugin** - Extends IPlugin for UI
   - GetStyleSheet()
   - GetHTMLTemplate()
   - ApplyTheme()
   - RenderComponent()

### Best Practices

1. **Always inherit from PluginBase**
   - Provides common functionality
   - Handles initialization/disposal
   - Provides logging

2. **Implement proper error handling**
   ```csharp
   public override object Execute(Dictionary<string, object> parameters)
   {
       try {
           // Your code
       }
       catch (Exception ex) {
           Log($"Error: {ex.Message}", "ERROR");
           throw;
       }
   }
   ```

3. **Validate parameters**
   ```csharp
   if (!ValidateParameters(parameters, "required1", "required2"))
   {
       throw new ArgumentException("Missing required parameters");
   }
   ```

4. **Document capabilities**
   ```csharp
   protected override void InitializeCapabilities()
   {
       base.InitializeCapabilities();
       _capabilities["SupportedFormats"] = new List<string> { "FASTA", "FASTQ" };
       _capabilities["MaxDataSize"] = 10000000;
   }
   ```

## Integration with SeleniumVBA

### Complete Example

```vba
' Module: BiotechAutomation
Option Explicit

Private pluginManager As Object
Private driver As WebDriver

Sub InitializeBiotechSystem()
    ' Initialize plugin system
    Set pluginManager = CreateObject("SeleniumVBA.Plugins.Core.PluginManager")
    pluginManager.AddPluginDirectory(ThisWorkbook.Path & "\plugins")
    pluginManager.LoadPlugins()
    
    ' Initialize WebDriver
    Set driver = New WebDriver
    driver.StartChrome
    driver.OpenBrowser
    
    ' Load biotech dashboard
    driver.NavigateTo "file:///" & Replace(ThisWorkbook.Path, "\", "/") & "/plugins/.NET/UI/templates/dashboard.html"
End Sub

Sub AnalyzeDNASample()
    Dim params As Object
    Dim result As Object
    
    ' Create parameters
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "action", "analyze"
    params.Add "sequence", Range("A1").Value ' Get sequence from Excel
    
    ' Execute plugin
    Set result = pluginManager.ExecutePlugin("biotech.dna.sequencing", params)
    
    ' Write results back to Excel
    Range("B1").Value = result("Length")
    Range("C1").Value = result("GCContent")
    Range("D1").Value = result("ComplexityScore")
    
    ' Update dashboard via WebDriver
    driver.ExecuteScript "updateMetrics(" & result("Length") & ")"
End Sub

Sub CleanupBiotechSystem()
    pluginManager.UnloadAllPlugins
    driver.CloseBrowser
    driver.Shutdown
End Sub
```

## Troubleshooting

### Plugin Not Loading

1. Check that the DLL is in the correct directory
2. Ensure .NET Framework 4.7.2+ is installed
3. Verify the plugin implements required interfaces
4. Check Windows Event Log for errors

### COM Registration Issues

For VBA to access .NET plugins, they may need COM registration:

```bash
regasm /codebase YourPlugin.dll
```

### Theme Not Applying

1. Verify CSS file path is correct
2. Check browser console for errors
3. Ensure backdrop-filter is supported (Chrome/Edge required)

## Resources

- SeleniumVBA Wiki: https://github.com/GCuser99/SeleniumVBA/wiki
- .NET Plugin Development: https://docs.microsoft.com/dotnet/
- Glass Morphism Design: https://glassmorphism.com/
- VBA-COM Interop: https://docs.microsoft.com/office/vba/

## License

MIT License - See LICENSE.txt for details

## Support

For issues and questions:
- GitHub Issues: https://github.com/deathknight2002/GG_ultimate_biotechnology_repo/issues
- SeleniumVBA Community: https://github.com/GCuser99/SeleniumVBA
