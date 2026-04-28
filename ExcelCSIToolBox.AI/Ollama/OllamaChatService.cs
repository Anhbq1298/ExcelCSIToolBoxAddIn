using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ExcelCSIToolBox.AI.Ollama
{
    /// <summary>
    /// HTTP client service for communicating with a local Ollama instance.
    ///
    /// Configuration:
    ///   Endpoint : http://localhost:11434/api/chat   (change OllamaEndpoint to override)
    ///   Model    : qwen2.5-coder:7b                  (change DefaultModel to override)
    ///   Stream   : false (always — we collect the full response at once)
    /// </summary>
    public class OllamaChatService
    {
        // ── Configuration ─────────────────────────────────────────────────────────

        /// <summary>Base URL of the Ollama /api/chat endpoint.</summary>
        public static string OllamaEndpoint = "http://localhost:11434/api/chat";

        /// <summary>Default model pulled from Ollama ("ollama pull qwen2.5-coder:7b").</summary>
        public static string DefaultModel = "qwen2.5-coder:7b";

        // ── Private state ─────────────────────────────────────────────────────────

        private static readonly HttpClient _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(120)
        };

        // ── Public API ────────────────────────────────────────────────────────────

        /// <summary>
        /// Send a list of messages to Ollama and return the assistant reply text.
        /// Throws HttpRequestException or TaskCanceledException on network failure.
        /// </summary>
        public async Task<string> ChatAsync(
            List<OllamaMessage> messages,
            CancellationToken   cancellationToken,
            string              model = null)
        {
            OllamaChatRequest request = new OllamaChatRequest
            {
                model    = string.IsNullOrWhiteSpace(model) ? DefaultModel : model,
                messages = messages,
                stream   = false
            };

            string requestJson = JsonConvert.SerializeObject(request);
            StringContent content = new StringContent(requestJson, Encoding.UTF8, "application/json");

            HttpResponseMessage httpResponse = await _httpClient.PostAsync(OllamaEndpoint, content, cancellationToken);
            httpResponse.EnsureSuccessStatusCode();

            string responseBody = await httpResponse.Content.ReadAsStringAsync();
            OllamaChatResponse ollamaResponse = JsonConvert.DeserializeObject<OllamaChatResponse>(responseBody);

            return ollamaResponse?.message?.content ?? string.Empty;
        }
    }
}
