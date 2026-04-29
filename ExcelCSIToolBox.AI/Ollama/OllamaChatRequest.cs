using System.Collections.Generic;

namespace ExcelCSIToolBox.AI.Ollama
{
    /// <summary>
    /// Request payload sent to the Ollama /api/chat endpoint.
    /// </summary>
    public class OllamaChatRequest
    {
        /// <summary>Model name, e.g. "qwen2.5-coder:3b".</summary>
        public string model { get; set; }

        /// <summary>List of messages forming the conversation context.</summary>
        public List<OllamaMessage> messages { get; set; }

        /// <summary>
        /// When false (default), Ollama returns the full response as a single JSON object.
        /// When true, Ollama streams tokens one by one.
        /// </summary>
        public bool stream { get; set; } = false;
        
        /// <summary>
        /// How long the model should stay in memory. Default "30m".
        /// </summary>
        public string keep_alive { get; set; } = "30m";

        /// <summary>
        /// Model configuration options.
        /// </summary>
        public OllamaOptions options { get; set; } = new OllamaOptions();
    }

    public class OllamaOptions
    {
        public int num_ctx { get; set; } = 4096;
        public int num_predict { get; set; } = 400;
        public double temperature { get; set; } = 0.2;
    }
}
