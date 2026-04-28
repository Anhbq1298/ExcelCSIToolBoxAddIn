namespace ExcelCSIToolBox.AI.Ollama
{
    /// <summary>
    /// Parsed response from the Ollama /api/chat endpoint (stream=false).
    /// </summary>
    public class OllamaChatResponse
    {
        /// <summary>Model that generated the response.</summary>
        public string model { get; set; }

        /// <summary>The assistant message returned by the model.</summary>
        public OllamaMessage message { get; set; }

        /// <summary>Whether generation is finished.</summary>
        public bool done { get; set; }
    }
}
