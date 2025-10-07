using System;
using System.Collections.Generic;
using System.Data;

namespace SeleniumVBA.Plugins.Interfaces
{
    /// <summary>
    /// Interface for biotechnology-specific plugins
    /// Extends IPlugin with biotech domain functionality
    /// </summary>
    public interface IBiotechPlugin : IPlugin
    {
        /// <summary>
        /// Process biological data
        /// </summary>
        /// <param name="data">Input data table</param>
        /// <param name="processingType">Type of processing (sequencing, analysis, etc.)</param>
        /// <returns>Processed data table</returns>
        DataTable ProcessBiologicalData(DataTable data, string processingType);

        /// <summary>
        /// Generate biotech report
        /// </summary>
        /// <param name="data">Data to report on</param>
        /// <param name="reportType">Report format (PDF, HTML, Excel, etc.)</param>
        /// <returns>Report file path or content</returns>
        string GenerateReport(DataTable data, string reportType);

        /// <summary>
        /// Perform statistical analysis on biological data
        /// </summary>
        /// <param name="data">Input data</param>
        /// <returns>Analysis results</returns>
        Dictionary<string, object> PerformAnalysis(DataTable data);

        /// <summary>
        /// Visualize data in biotech-specific formats
        /// </summary>
        /// <param name="data">Data to visualize</param>
        /// <param name="visualizationType">Type of visualization (chart, heatmap, etc.)</param>
        /// <returns>Visualization output path or base64 encoded image</returns>
        string Visualize(DataTable data, string visualizationType);

        /// <summary>
        /// Validate biological data against standards
        /// </summary>
        /// <param name="data">Data to validate</param>
        /// <returns>Validation results</returns>
        Dictionary<string, object> ValidateData(DataTable data);
    }
}
