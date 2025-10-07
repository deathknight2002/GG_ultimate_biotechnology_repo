using System;
using System.Collections.Generic;

namespace SeleniumVBA.Plugins.Interfaces
{
    /// <summary>
    /// Interface for UI enhancement plugins
    /// Supports modern glass/OLED styling and dynamic interfaces
    /// </summary>
    public interface IUIPlugin : IPlugin
    {
        /// <summary>
        /// Get CSS stylesheet for the plugin UI
        /// </summary>
        /// <param name="theme">Theme name (glass, oled-dark, oled-light, biotech, etc.)</param>
        /// <returns>CSS content</returns>
        string GetStyleSheet(string theme);

        /// <summary>
        /// Generate HTML template for the plugin
        /// </summary>
        /// <param name="templateName">Template identifier</param>
        /// <returns>HTML content</returns>
        string GetHTMLTemplate(string templateName);

        /// <summary>
        /// Get JavaScript for dynamic UI interactions
        /// </summary>
        /// <returns>JavaScript content</returns>
        string GetJavaScript();

        /// <summary>
        /// Apply theme to UI elements
        /// </summary>
        /// <param name="themeName">Theme to apply</param>
        /// <returns>True if theme applied successfully</returns>
        bool ApplyTheme(string themeName);

        /// <summary>
        /// Get list of supported themes
        /// </summary>
        /// <returns>List of theme names</returns>
        List<string> GetAvailableThemes();

        /// <summary>
        /// Render UI component
        /// </summary>
        /// <param name="componentName">Component identifier</param>
        /// <param name="data">Data to render</param>
        /// <returns>Rendered HTML</returns>
        string RenderComponent(string componentName, Dictionary<string, object> data);
    }
}
