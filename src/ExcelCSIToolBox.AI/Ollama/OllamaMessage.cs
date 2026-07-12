namespace ExcelCSIToolBox.AI.Ollama
{
    /// <summary>
    /// Represents a single chat message in an Ollama conversation.
    /// </summary>
    public class OllamaMessage
    {
        /// <summary>"system", "user", or "assistant".</summary>
        public string role { get; set; }

        /// <summary>Text content of the message.</summary>
        public string content { get; set; }
    }
}
