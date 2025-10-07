using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SeleniumVBA.Plugins.Core
{
    /// <summary>
    /// Plugin manager for loading, managing, and executing plugins
    /// </summary>
    public class PluginManager
    {
        private Dictionary<string, IPlugin> _loadedPlugins;
        private List<string> _pluginDirectories;
        private Dictionary<string, object> _globalConfig;

        public PluginManager()
        {
            _loadedPlugins = new Dictionary<string, IPlugin>();
            _pluginDirectories = new List<string>();
            _globalConfig = new Dictionary<string, object>();
        }

        /// <summary>
        /// Add a directory to scan for plugins
        /// </summary>
        public void AddPluginDirectory(string path)
        {
            if (Directory.Exists(path))
            {
                _pluginDirectories.Add(path);
            }
        }

        /// <summary>
        /// Discover and load all plugins from registered directories
        /// </summary>
        public void LoadPlugins()
        {
            foreach (var directory in _pluginDirectories)
            {
                LoadPluginsFromDirectory(directory);
            }
        }

        /// <summary>
        /// Load plugins from a specific directory
        /// </summary>
        private void LoadPluginsFromDirectory(string directory)
        {
            var dllFiles = Directory.GetFiles(directory, "*.dll", SearchOption.AllDirectories);

            foreach (var dllFile in dllFiles)
            {
                try
                {
                    LoadPluginFromAssembly(dllFile);
                }
                catch (Exception ex)
                {
                    // Log error but continue loading other plugins
                    Console.WriteLine($"Error loading plugin from {dllFile}: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Load plugin from assembly file
        /// </summary>
        private void LoadPluginFromAssembly(string assemblyPath)
        {
            var assembly = Assembly.LoadFrom(assemblyPath);
            var pluginTypes = assembly.GetTypes()
                .Where(t => typeof(IPlugin).IsAssignableFrom(t) && !t.IsInterface && !t.IsAbstract);

            foreach (var type in pluginTypes)
            {
                var plugin = (IPlugin)Activator.CreateInstance(type);
                if (plugin.Initialize(_globalConfig))
                {
                    _loadedPlugins[plugin.Id] = plugin;
                }
            }
        }

        /// <summary>
        /// Get loaded plugin by ID
        /// </summary>
        public IPlugin GetPlugin(string pluginId)
        {
            return _loadedPlugins.ContainsKey(pluginId) ? _loadedPlugins[pluginId] : null;
        }

        /// <summary>
        /// Get all loaded plugins
        /// </summary>
        public List<IPlugin> GetAllPlugins()
        {
            return _loadedPlugins.Values.ToList();
        }

        /// <summary>
        /// Get plugins by category
        /// </summary>
        public List<IPlugin> GetPluginsByCategory(string category)
        {
            return _loadedPlugins.Values
                .Where(p => p.Category.Equals(category, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        /// <summary>
        /// Execute a plugin
        /// </summary>
        public object ExecutePlugin(string pluginId, Dictionary<string, object> parameters)
        {
            var plugin = GetPlugin(pluginId);
            if (plugin != null)
            {
                return plugin.Execute(parameters);
            }
            throw new InvalidOperationException($"Plugin '{pluginId}' not found");
        }

        /// <summary>
        /// Unload all plugins
        /// </summary>
        public void UnloadAllPlugins()
        {
            foreach (var plugin in _loadedPlugins.Values)
            {
                plugin.Dispose();
            }
            _loadedPlugins.Clear();
        }

        /// <summary>
        /// Set global configuration
        /// </summary>
        public void SetGlobalConfig(Dictionary<string, object> config)
        {
            _globalConfig = config;
        }
    }
}
