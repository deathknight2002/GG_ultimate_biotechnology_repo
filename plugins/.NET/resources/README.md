# .NET Plugin System Resources

This directory contains biotechnology-specific resources and datasets for use with the SeleniumVBA .NET plugin system.

## Contents

### Sample Data

Example datasets for testing and development:

- **DNA Sequences**: Sample FASTA/FASTQ files
- **Protein Data**: Sample protein sequences and structures
- **Analysis Results**: Example output formats

### Protocols

Standard operating procedures and protocols:

- **DNA Sequencing Protocol**: Step-by-step sequencing analysis
- **Protein Analysis Protocol**: Protein characterization workflow
- **Quality Control Standards**: QC thresholds and validation

### Templates

Reusable templates for reports and visualizations:

- **Report Templates**: HTML, PDF, Excel report formats
- **Dashboard Templates**: Custom dashboard layouts
- **Visualization Templates**: Chart and graph templates

## Usage

These resources can be accessed programmatically through the plugin system:

```vba
' Load sample DNA data
Dim sampleData As String
sampleData = LoadResource("samples/dna_sample_001.fasta")

' Use in analysis
Dim result As Object
Set result = AnalyzeDNA(sampleData)
```

## Adding Your Own Resources

1. Create appropriate subdirectories
2. Add your data files
3. Update resource manifest if needed
4. Reference in your plugins or VBA code

## File Formats Supported

- **Sequences**: FASTA, FASTQ, GenBank
- **Structures**: PDB, mmCIF
- **Data**: CSV, TSV, Excel, JSON
- **Images**: PNG, JPEG, SVG
- **Reports**: HTML, PDF, DOCX

## Directory Structure

```
resources/
├── samples/          ← Sample datasets
├── protocols/        ← Standard protocols
├── templates/        ← Report/viz templates
├── images/          ← Icons and graphics
└── documentation/   ← Additional docs
```

## Integration

Resources are automatically discovered by the plugin system and can be referenced by path or ID in your VBA code or plugin implementations.

For more information, see the main plugin documentation in `/plugins/.NET/README.md`.
