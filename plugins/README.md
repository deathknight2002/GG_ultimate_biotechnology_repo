# SeleniumVBA Plugin System

This directory contains the extensible plugin architecture for SeleniumVBA, with special focus on biotechnology applications and modern UI theming.

## 🚀 Quick Links

- **[.NET Plugins](.NET/)** - Main plugin system implementation
  - **[Quick Start Guide](.NET/QUICKSTART.md)** - Get started in 5 minutes
  - **[Full Documentation](.NET/README.md)** - Complete plugin development guide
  - **[Implementation Summary](.NET/IMPLEMENTATION_SUMMARY.md)** - What was delivered

## 📁 Directory Overview

```
plugins/
└── .NET/                          # .NET Plugin System
    ├── Interfaces/                # Plugin contracts (IPlugin, IBiotechPlugin, IUIPlugin)
    ├── Core/                     # Core infrastructure (PluginManager, PluginBase)
    ├── Biotech/                  # Biotech plugins (DNA, Protein)
    ├── UI/                       # UI theme plugin (Glass OLED)
    │   ├── styles/              # CSS themes
    │   ├── templates/           # HTML templates
    │   └── scripts/             # JavaScript
    ├── Examples/                 # Starter plugin templates
    ├── VBA/                     # VBA integration helper
    └── resources/               # Sample data and templates
```

## 🎯 What's Inside

### Biotech Plugins
- **DNA Sequencing** - Analyze sequences, calculate GC content, find motifs
- **Protein Analysis** - Calculate properties, predict structure, identify domains

### UI Theme
- **Glass OLED** - Modern glass morphism with OLED-optimized styling
- **Dashboard** - Interactive biotech data visualization
- **Components** - Pre-built UI elements (cards, tables, buttons)

### Integration
- **VBA Helper** - Easy integration with Excel/Access VBA
- **Plugin Manager** - Automatic plugin discovery and loading
- **COM Interop** - Seamless VBA ↔ .NET communication

## 🔧 Getting Started

### View the Dashboard
Open in your browser:
```
plugins/.NET/UI/templates/dashboard.html
```

### Use from VBA
1. Import `plugins/.NET/VBA/PluginHelper.bas` into your VBA project
2. Call the helper functions:

```vba
' Analyze DNA
Dim result As Object
Set result = AnalyzeDNA("ATCGATCGATCG")
Debug.Print "GC Content: " & result("GCContent")

' Launch dashboard
LaunchBiotechDashboard
```

### Create Custom Plugins
See `plugins/.NET/Examples/StarterPlugin.cs` for a template.

## 📚 Documentation

- **[Full Plugin System Documentation](.NET/README.md)** - Comprehensive guide
- **[Quick Start Guide](.NET/QUICKSTART.md)** - 5-minute tutorial
- **[Implementation Summary](.NET/IMPLEMENTATION_SUMMARY.md)** - Technical details
- **[Plugin Manifest](.NET/plugin-manifest.json)** - Plugin registry

## 🎨 Features

- ✅ Plug-and-play architecture
- ✅ Biotech-specific modules
- ✅ Modern Glass OLED UI
- ✅ VBA integration
- ✅ Extensible design
- ✅ Well-documented
- ✅ Production-ready

## 💡 Use Cases

- **DNA/RNA Analysis** - Process sequencing data from Excel
- **Protein Studies** - Calculate properties and predict structure
- **Data Visualization** - Create professional biotech dashboards
- **Automation** - Combine with SeleniumVBA for web scraping + analysis
- **Custom Workflows** - Build your own plugins for specific needs

## 🆘 Support

For help:
1. Check the documentation in `.NET/README.md`
2. See examples in `.NET/Examples/`
3. Open an issue on GitHub

## 📄 License

MIT License - See LICENSE.txt in the root directory

---

**Part of SeleniumVBA** - A comprehensive Selenium wrapper for browser automation
