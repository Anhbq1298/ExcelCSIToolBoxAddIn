using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelCSIToolBox.AI.Agent
{
    public sealed class AgentTaskExecutorService
    {
        private readonly Func<string, CancellationToken, Task<AiAgentResponse>> _executeTaskAsync;

        public AgentTaskExecutorService(Func<string, CancellationToken, Task<AiAgentResponse>> executeTaskAsync)
        {
            _executeTaskAsync = executeTaskAsync ?? throw new ArgumentNullException(nameof(executeTaskAsync));
        }

        public async Task<AgentTaskExecutionSummary> ExecuteAsync(
            IReadOnlyList<AgentTaskItem> tasks,
            CancellationToken cancellationToken)
        {
            var summary = new AgentTaskExecutionSummary
            {
                Tasks = new List<AgentTaskItem>()
            };

            if (tasks == null)
            {
                return summary;
            }

            for (int i = 0; i < tasks.Count; i++)
            {
                AgentTaskItem task = tasks[i];
                if (task == null)
                {
                    continue;
                }

                task.Status = "Running";

                if (IsInstructionOnly(task))
                {
                    task.Status = "Completed";
                    task.ResultMessage = "Applied as a response-format instruction.";
                    summary.Tasks.Add(task);
                    continue;
                }

                if (NeedsClarification(task))
                {
                    task.Status = "NeedsClarification";
                    task.NeedsClarification = true;
                    task.ResultMessage = "This task needs a clearer target or parameters before it can be executed safely.";
                    summary.Tasks.Add(task);
                    continue;
                }

                try
                {
                    AiAgentResponse response = await _executeTaskAsync(task.OriginalText, cancellationToken);
                    task.ResultMessage = response == null || string.IsNullOrWhiteSpace(response.AssistantText)
                        ? "No response was produced."
                        : response.AssistantText.Trim();
                    task.Status = IsFailureResponse(response) ? "Failed" : "Completed";
                }
                catch (Exception ex)
                {
                    task.Status = "Failed";
                    task.ResultMessage = ex.Message;
                }

                summary.Tasks.Add(task);
            }

            return summary;
        }

        public static string FormatDetectedTasks(IReadOnlyList<AgentTaskItem> tasks)
        {
            var builder = new StringBuilder();
            builder.AppendLine("Detected tasks:");
            if (tasks == null || tasks.Count == 0)
            {
                builder.AppendLine("1. No executable task detected.");
                return builder.ToString();
            }

            for (int i = 0; i < tasks.Count; i++)
            {
                builder.AppendLine((i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ". " + tasks[i].OriginalText);
            }

            return builder.ToString();
        }

        public static string FormatFinalResponse(AgentTaskExecutionSummary summary, bool includeDetectedTasks = true)
        {
            var builder = new StringBuilder();
            IReadOnlyList<AgentTaskItem> tasks = summary == null ? null : summary.Tasks;

            if (includeDetectedTasks)
            {
                builder.AppendLine("Tasks detected");
                AppendTaskList(builder, tasks, null);
                builder.AppendLine();
            }

            builder.AppendLine("Tasks completed");
            AppendTaskList(builder, tasks, "Completed");

            builder.AppendLine();
            builder.AppendLine("Tasks failed");
            AppendTaskList(builder, tasks, "Failed");

            builder.AppendLine();
            builder.AppendLine("Tasks requiring clarification");
            AppendTaskList(builder, tasks, "NeedsClarification");

            builder.AppendLine();
            builder.AppendLine("Model/API changes made successfully");
            AppendSuccessfulChanges(builder, tasks);

            return builder.ToString().Trim();
        }

        private static void AppendTaskList(StringBuilder builder, IReadOnlyList<AgentTaskItem> tasks, string status)
        {
            int count = 0;
            if (tasks != null)
            {
                for (int i = 0; i < tasks.Count; i++)
                {
                    AgentTaskItem task = tasks[i];
                    if (task == null || status != null && !string.Equals(task.Status, status, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    count++;
                    string suffix = string.IsNullOrWhiteSpace(task.ResultMessage) || status == null
                        ? string.Empty
                        : " - " + task.ResultMessage;
                    builder.AppendLine(count.ToString(System.Globalization.CultureInfo.InvariantCulture) + ". " + task.OriginalText + suffix);
                }
            }

            if (count == 0)
            {
                builder.AppendLine("- None.");
            }
        }

        private static void AppendSuccessfulChanges(StringBuilder builder, IReadOnlyList<AgentTaskItem> tasks)
        {
            int count = 0;
            if (tasks != null)
            {
                for (int i = 0; i < tasks.Count; i++)
                {
                    AgentTaskItem task = tasks[i];
                    if (task == null ||
                        !string.Equals(task.Status, "Completed", StringComparison.OrdinalIgnoreCase) ||
                        !IsModelOrApiChange(task))
                    {
                        continue;
                    }

                    count++;
                    builder.AppendLine(count.ToString(System.Globalization.CultureInfo.InvariantCulture) + ". " + task.ResultMessage);
                }
            }

            if (count == 0)
            {
                builder.AppendLine("- None.");
            }
        }

        private static bool IsInstructionOnly(AgentTaskItem task)
        {
            string text = task.NormalizedIntent ?? string.Empty;
            return task.ActionType == "Query" &&
                   (text.StartsWith("return ", StringComparison.OrdinalIgnoreCase) ||
                    text.StartsWith("summarize ", StringComparison.OrdinalIgnoreCase));
        }

        private static bool NeedsClarification(AgentTaskItem task)
        {
            if (task == null)
            {
                return false;
            }

            if (task.ActionType == "CodeChange")
            {
                return true;
            }

            return task.ActionType == "Unknown" && task.TargetObjectType == "General";
        }

        private static bool IsFailureResponse(AiAgentResponse response)
        {
            if (response == null)
            {
                return true;
            }

            if (response.ToolWasCalled && response.ToolResponse != null)
            {
                return !response.ToolResponse.Success;
            }

            string text = response.AssistantText ?? string.Empty;
            return text.IndexOf("failed", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   text.IndexOf("could not", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   text.IndexOf("not approved", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsModelOrApiChange(AgentTaskItem task)
        {
            return task.ActionType == "Create" ||
                   task.ActionType == "Assign" ||
                   task.ActionType == "Select" ||
                   task.ActionType == "CodeChange";
        }
    }

    public sealed class AgentTaskExecutionSummary
    {
        public List<AgentTaskItem> Tasks { get; set; }
    }
}
