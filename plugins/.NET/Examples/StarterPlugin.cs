using System;
using System.Collections.Generic;
using SeleniumVBA.Plugins.Core;
using SeleniumVBA.Plugins.Interfaces;

namespace SeleniumVBA.Plugins.Examples
{
    /// <summary>
    /// Example starter plugin demonstrating the plugin architecture
    /// Use this as a template for creating custom plugins
    /// </summary>
    public class StarterPlugin : PluginBase
    {
        public override string Id => "example.starter";
        public override string Name => "Starter Plugin Template";
        public override Version Version => new Version(1, 0, 0);
        public override string Author => "Your Name";
        public override string Description => "A template plugin showing basic functionality";
        public override string Category => "Examples";

        protected override void InitializeCapabilities()
        {
            base.InitializeCapabilities();
            
            // Define what your plugin can do
            _capabilities["SupportedOperations"] = new List<string> 
            { 
                "HelloWorld", 
                "ProcessData", 
                "GenerateOutput" 
            };
        }

        protected override bool OnInitialize()
        {
            // Custom initialization logic
            Log("Starter plugin initializing...");
            
            // Check configuration
            if (_config.ContainsKey("debug"))
            {
                bool debug = Convert.ToBoolean(_config["debug"]);
                Log($"Debug mode: {debug}");
            }

            return true; // Return false if initialization fails
        }

        public override object Execute(Dictionary<string, object> parameters)
        {
            if (!ValidateParameters(parameters, "operation"))
            {
                throw new ArgumentException("Missing required parameter: operation");
            }

            string operation = parameters["operation"].ToString();
            
            Log($"Executing operation: {operation}");

            switch (operation.ToLower())
            {
                case "helloworld":
                    return HelloWorld(parameters);
                    
                case "processdata":
                    return ProcessData(parameters);
                    
                case "generateoutput":
                    return GenerateOutput(parameters);
                    
                default:
                    throw new ArgumentException($"Unknown operation: {operation}");
            }
        }

        protected override void OnDispose()
        {
            // Cleanup resources
            Log("Starter plugin disposing...");
        }

        // Example operation methods
        private object HelloWorld(Dictionary<string, object> parameters)
        {
            string name = parameters.ContainsKey("name") 
                ? parameters["name"].ToString() 
                : "World";

            string message = $"Hello, {name}!";
            Log(message);

            return new Dictionary<string, object>
            {
                { "Message", message },
                { "Timestamp", DateTime.Now },
                { "Success", true }
            };
        }

        private object ProcessData(Dictionary<string, object> parameters)
        {
            if (!ValidateParameters(parameters, "data"))
            {
                throw new ArgumentException("Missing required parameter: data");
            }

            var inputData = parameters["data"];
            Log($"Processing data: {inputData}");

            // Your data processing logic here
            var result = new Dictionary<string, object>
            {
                { "OriginalData", inputData },
                { "ProcessedData", $"Processed: {inputData}" },
                { "ProcessingTime", DateTime.Now },
                { "Success", true }
            };

            return result;
        }

        private object GenerateOutput(Dictionary<string, object> parameters)
        {
            string outputType = parameters.ContainsKey("type") 
                ? parameters["type"].ToString() 
                : "json";

            Log($"Generating {outputType} output");

            return new Dictionary<string, object>
            {
                { "OutputType", outputType },
                { "Content", "This is sample output content" },
                { "GeneratedAt", DateTime.Now },
                { "Success", true }
            };
        }
    }
}
