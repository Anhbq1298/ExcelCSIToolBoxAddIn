using System.Collections.Generic;

namespace ExcelCSIToolBox.AI.Ollama
{
    /// <summary>
    /// Request payload sent to the Ollama /api/chat endpoint.
    /// </summary>
    public class OllamaChatRequest
    {
        /// <summary>Model name, e.g. "qwen2.5-coder:7b".</summary>
        public string model { get; set; }

        /// <summary>List of messages forming the conversation context.</summary>
        public List<OllamaMessage> messages { get; set; }

        /// <summary>
        /// When false (default), Ollama returns the full response as a single JSON object.
        /// We always use false for simplicity in this add-in.
        /// </summary>
        public bool stream { get; set; } = false;
    }
}
