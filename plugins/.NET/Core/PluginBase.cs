using System;
using System.Collections.Generic;
using SeleniumVBA.Plugins.Interfaces;

namespace SeleniumVBA.Plugins.Core
{
    /// <summary>
    /// Abstract base class for plugin implementations
    /// Provides common functionality for all plugins
    /// </summary>
    public abstract class PluginBase : IPlugin
    {
        protected bool _isInitialized;
        protected Dictionary<string, object> _config;
        protected Dictionary<string, object> _capabilities;

        public abstract string Id { get; }
        public abstract string Name { get; }
        public abstract Version Version { get; }
        public abstract string Author { get; }
        public abstract string Description { get; }
        public abstract string Category { get; }

        protected PluginBase()
        {
            _isInitialized = false;
            _config = new Dictionary<string, object>();
            _capabilities = new Dictionary<string, object>();
            InitializeCapabilities();
        }

        public virtual bool Initialize(Dictionary<string, object> config)
        {
            if (_isInitialized)
                return true;

            _config = config ?? new Dictionary<string, object>();
            _isInitialized = OnInitialize();
            return _isInitialized;
        }

        public abstract object Execute(Dictionary<string, object> parameters);

        public virtual void Dispose()
        {
            OnDispose();
            _isInitialized = false;
        }

        public virtual Dictionary<string, object> GetCapabilities()
        {
            return new Dictionary<string, object>(_capabilities);
        }

        /// <summary>
        /// Override to implement plugin-specific initialization
        /// </summary>
        protected virtual bool OnInitialize()
        {
            return true;
        }

        /// <summary>
        /// Override to implement plugin-specific cleanup
        /// </summary>
        protected virtual void OnDispose()
        {
            // Base implementation does nothing
        }

        /// <summary>
        /// Override to define plugin capabilities
        /// </summary>
        protected virtual void InitializeCapabilities()
        {
            _capabilities["SupportedVersions"] = new List<string> { "1.0" };
        }

        /// <summary>
        /// Validate required parameters
        /// </summary>
        protected bool ValidateParameters(Dictionary<string, object> parameters, params string[] requiredKeys)
        {
            if (parameters == null)
                return false;

            foreach (var key in requiredKeys)
            {
                if (!parameters.ContainsKey(key))
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Log message (can be overridden for custom logging)
        /// </summary>
        protected virtual void Log(string message, string level = "INFO")
        {
            Console.WriteLine($"[{level}] {Name}: {message}");
        }
    }
}
