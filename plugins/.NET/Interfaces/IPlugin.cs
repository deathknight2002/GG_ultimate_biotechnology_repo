using System;
using System.Collections.Generic;

namespace SeleniumVBA.Plugins.Interfaces
{
    /// <summary>
    /// Base interface for all SeleniumVBA plugins
    /// Provides core functionality for plugin lifecycle management
    /// </summary>
    public interface IPlugin
    {
        /// <summary>
        /// Unique identifier for the plugin
        /// </summary>
        string Id { get; }

        /// <summary>
        /// Display name of the plugin
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Plugin version
        /// </summary>
        Version Version { get; }

        /// <summary>
        /// Plugin author/vendor
        /// </summary>
        string Author { get; }

        /// <summary>
        /// Plugin description
        /// </summary>
        string Description { get; }

        /// <summary>
        /// Plugin category (Biotech, Analysis, Visualization, etc.)
        /// </summary>
        string Category { get; }

        /// <summary>
        /// Initialize the plugin
        /// </summary>
        /// <param name="config">Configuration parameters</param>
        /// <returns>True if initialization successful</returns>
        bool Initialize(Dictionary<string, object> config);

        /// <summary>
        /// Execute the plugin's main functionality
        /// </summary>
        /// <param name="parameters">Execution parameters</param>
        /// <returns>Execution result</returns>
        object Execute(Dictionary<string, object> parameters);

        /// <summary>
        /// Cleanup plugin resources
        /// </summary>
        void Dispose();

        /// <summary>
        /// Get plugin capabilities
        /// </summary>
        /// <returns>Dictionary of supported capabilities</returns>
        Dictionary<string, object> GetCapabilities();
    }
}
