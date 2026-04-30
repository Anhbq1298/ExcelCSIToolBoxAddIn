using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelCSIToolBox.Data.CSISapModel.Intent;

namespace ExcelCSIToolBox.AI.Agent
{
    public sealed class IntentClarificationFactory
    {
        public string CreateMessage(IReadOnlyList<CsiRequestTaskClassificationDto> tasks)
        {
            var messages = new List<string>();
            for (int i = 0; i < tasks.Count; i++)
            {
                CsiRequestTaskClassificationDto task = tasks[i];
                if (task == null || string.IsNullOrWhiteSpace(task.ClarificationMessage))
                {
                    continue;
                }

                if (tasks.Count == 1)
                {
                    messages.Add(task.ClarificationMessage);
                }
                else
                {
                    messages.Add((i + 1).ToString(CultureInfo.InvariantCulture) + ". " + task.ClarificationMessage);
                }
            }

            return messages.Count == 0
                ? "Please provide the missing action, target object, and required parameters."
                : string.Join(Environment.NewLine, messages);
        }
    }
}
