<img src="https://github.com/GCuser99/SeleniumVBA/blob/main/dev/logo/logo.png" alt="SeleniumVBA" width=25%>

A comprehensive Selenium wrapper for browser automation developed for MS Office VBA running in Windows

## Features

### Browser Automation
- Edge, Chrome, and Firefox browser automation support
- MS Excel Add-in, MS Access DB, and [twinBASIC](https://twinbasic.com/preview.html) ActiveX DLL solutions available
- A superset of Selenium's [W3C WebDriver](https://w3c.github.io/webdriver/) commands - [over 400 public methods and properties](https://github.com/GCuser99/SeleniumVBA/wiki/Object-Model-Overview)
- Support for HTML DOM, Action Chains, SendKeys, Shadow Roots, Cookies, ExecuteScript, and Capabilities
- ExecuteCDP method exposing [Chrome DevTools Protocol](https://chromedevtools.github.io/devtools-protocol/) - a low-level interface for browser interaction
- Automated Browser/WebDriver version alignment - works out-of-the-box with no manual downloads necessary!

### 🧬 Biotechnology Plugin System (NEW)
- **Extensible .NET Plugin Architecture** - Create custom plugins using C#
- **DNA Sequencing Analysis** - Quality control, GC content, motif finding
- **Protein Analysis** - Structure prediction, molecular weight, domain identification
- **Data Processing Pipeline** - Batch process biological data from Excel/VBA
- **Plug-and-Play Integration** - Simply drop DLL files into plugin directory

### 🎨 Modern UI Framework (NEW)
- **Glass OLED Theme** - Glass morphism with OLED-optimized pure blacks
- **Biotech Color Scheme** - Neon green (#00ff88) and cyan accents
- **Interactive Dashboard** - Real-time metrics and data visualization
- **Responsive Components** - Pre-built cards, tables, buttons, and charts
- **Dynamic Animations** - Shimmer effects, smooth transitions, status indicators

### Resources
- Help documentation is available in the [SeleniumVBA Wiki](https://github.com/GCuser99/SeleniumVBA/wiki)
- Plugin documentation in [plugins/.NET/README.md](plugins/.NET/README.md)
- Quick start guide in [plugins/.NET/QUICKSTART.md](plugins/.NET/QUICKSTART.md)

**What's New?**

- 🧬 **NEW: Biotechnology .NET Plugin System** - Extensible plugin architecture with biotech-specific modules for DNA sequencing, protein analysis, and data visualization
- 🎨 **NEW: Glass OLED UI Theme** - Modern glass morphism design with OLED-optimized colors and dynamic biotech styling
- 🔌 **NEW: Plug-and-Play Architecture** - Easy integration of custom .NET plugins with VBA through COM interop
- 📊 **NEW: Interactive Dashboard** - Real-time biotech data visualization with responsive HTML/CSS/JS components
- Improved performance/reliability by executing expensive/complex code using in-browser JavaScript
- NavigateToString allowing direct navigation to HTML strings to help facilitate testing and issue debugging
- Featured on [Sancarn](https://github.com/sancarn)'s curated [Awesome VBA list](https://github.com/sancarn/awesome-vba?tab=readme-ov-file#web-tools) [![Awesome](https://awesome.re/badge.svg)](https://awesome.re)

## Setup

**SeleniumVBA will function right out-of-the-box**. Just download/install any one of the provided [SeleniumVBA solutions](https://github.com/GCuser99/SeleniumVBA/tree/main/dist) and then run one of the subs in the "test" Standard modules. If the Selenium WebDriver does not exist, or is out-of-date, SeleniumVBA will detect this automatically and download the appropriate driver to a [configurable location](https://github.com/GCuser99/SeleniumVBA/wiki#advanced-customization) on your system.

Driver updates can also be programmatically invoked via the [WebDriverManager class](https://github.com/GCuser99/SeleniumVBA/wiki/Object-Model-Overview#webdrivermanager).

The [twinBASIC](https://twinbasic.com/preview.html) ActiveX DLL solution requires no dependencies (such as .Net Framework). To try it, download and run the installer in the [dist folder](https://github.com/GCuser99/SeleniumVBA/tree/main/dist).

## SendKeys Example

```vba
Sub doSendKeys()
    Dim driver As New WebDriver
    Dim keys As New WebKeyboard
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    keySeq = "This is COOKL!" & keys.Repeat(keys.LeftKey, 3) & keys.DeleteKey & keys.ReturnKey
    
    driver.FindElement(By.Name, "q").SendKeys keySeq
    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
```

## File Download Example
```vba
Sub doFileDownload()
    Dim driver As New WebDriver
    Dim caps As WebCapabilities
   
    driver.StartChrome
    
    'set the directory path for saving download to
    Set caps = driver.CreateCapabilities
    caps.SetDownloadPrefs downloadFolderPath:=".\"
    driver.OpenBrowser caps
    
    'delete legacy copy if it exists
    driver.DeleteFiles ".\test.pdf"
    
    driver.NavigateTo "https://github.com/GCuser99/SeleniumVBA/raw/main/dev/test_files/test.pdf"

    'wait until the download is complete before closing browser
    driver.WaitForDownload ".\test.pdf"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
```

## Action Chain Example
```vba
Sub doActionChain()
    Dim driver As New WebDriver
    Dim keys As New WebKeyboard
    Dim actions As WebActionChain
    Dim elemSearch As WebElement
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    Set elemSearch = driver.FindElement(By.Name, "btnK")
    
    Set actions = driver.ActionChain
    
    'build the chain and then execute with Perform method
    actions.KeyDown(keys.ShiftKey).SendKeys("upper case").KeyUp(keys.ShiftKey)
    actions.MoveToElement(elemSearch).Click().Perform

    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
```

## Biotech Plugin Example (NEW)
```vba
Sub doBiotechAnalysis()
    ' Import PluginHelper.bas module first
    
    ' Analyze DNA sequence
    Dim result As Object
    Set result = AnalyzeDNA("ATCGATCGATCGATCG")
    
    ' Display results
    Debug.Print "Sequence Length: " & result("Length")
    Debug.Print "GC Content: " & result("GCContent")
    Debug.Print "Complexity Score: " & result("ComplexityScore")
    
    ' Launch interactive dashboard
    LaunchBiotechDashboard
End Sub
```

## Glass OLED Dashboard Example (NEW)
```vba
Sub LaunchDashboard()
    Dim driver As New WebDriver
    driver.StartChrome
    driver.OpenBrowser
    
    ' Navigate to the Glass OLED biotech dashboard
    Dim dashboardPath As String
    dashboardPath = "file:///" & Replace(ThisWorkbook.Path, "\", "/") & _
                    "/plugins/.NET/UI/templates/dashboard.html"
    driver.NavigateTo dashboardPath
    
    ' Dashboard features:
    ' - Real-time metrics with animated counters
    ' - Glass morphism effects with backdrop blur
    ' - OLED-optimized pure black background
    ' - Interactive data tables and charts
    ' - Biotech-themed color scheme (neon green/cyan)
    
    driver.Wait 5000
End Sub
```

## Plugin System Setup

### Quick Start

1. **View the Dashboard**: Open `plugins/.NET/UI/templates/dashboard.html` in your browser
2. **Import VBA Helper**: Add `plugins/.NET/VBA/PluginHelper.bas` to your VBA project
3. **Use Plugins**: Call biotech functions directly from VBA

```vba
' Initialize plugin system
InitializePlugins

' Analyze DNA from Excel
Dim sequence As String
sequence = Range("A2").Value
Dim result As Object
Set result = AnalyzeDNA(sequence)

' Write results back to Excel
Range("B2").Value = result("Length")
Range("C2").Value = result("GCContent")
```

### Available Plugins

- **DNA Sequencing** (`biotech.dna.sequencing`) - Sequence analysis and quality control
- **Protein Analysis** (`biotech.protein.analysis`) - Structure and property prediction
- **Glass OLED Theme** (`ui.theme.glass-oled`) - Modern UI components and styling

For full documentation, see [plugins/.NET/README.md](plugins/.NET/README.md)

## Creating Custom Plugins

Extend the system with your own .NET plugins:

```csharp
using SeleniumVBA.Plugins.Core;

public class MyPlugin : PluginBase, IBiotechPlugin
{
    public override string Id => "mycompany.customplugin";
    public override string Name => "My Custom Plugin";
    // Implement plugin logic...
}
```

See [plugins/.NET/Examples/StarterPlugin.cs](plugins/.NET/Examples/StarterPlugin.cs) for a complete template.

## Collaborators

[@6DiegoDiego9](https://github.com/6DiegoDiego9) and [@GCUser99](https://github.com/GCUser99)

## Credits

This project is an extensively modified/extended version of uezo's [TinySeleniumVBA](https://github.com/uezo/TinySeleniumVBA/)

[VBA-JSON](https://github.com/VBA-tools/VBA-JSON) by Tim Hall, JSON converter for VBA

[RubberDuck](https://rubberduckvba.com/) by Mathieu Guindon

[twinBASIC](https://twinbasic.com/preview.html) by Wayne Phillips

[Inno Setup](https://jrsoftware.org/isinfo.php) by Jordan Russell and [UninsIS](https://github.com/Bill-Stewart/UninsIS) by Bill Stewart

[vba-regex](https://github.com/sihlfall/vba-regex) by sihlfall


