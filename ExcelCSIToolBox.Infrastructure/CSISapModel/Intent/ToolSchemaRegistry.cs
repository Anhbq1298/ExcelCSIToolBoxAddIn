using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using ExcelCSIToolBox.Data.CSISapModel.Intent;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel.Intent
{
    public sealed class ToolSchemaRegistry
    {
        private readonly List<ToolSchema> _schemas = new List<ToolSchema>
        {
            new ToolSchema(
                "PointObj_AddCartesian",
                "Add",
                "PointObj",
                new[] { "pointName", "x", "y", "z" },
                new string[0],
                "Please provide the point name and X, Y, Z coordinates."),

            new ToolSchema(
                "FrameObj_AddByPoint",
                "Add",
                "FrameObj",
                new[] { "frameName", "pointI", "pointJ" },
                new[] { "sectionName" },
                "Please provide the frame definition, either by point names or by start/end coordinates."),

            new ToolSchema(
                "FrameObj_AddByCoordinate",
                "Add",
                "FrameObj",
                new[] { "frameName", "xi", "yi", "zi", "xj", "yj", "zj" },
                new[] { "sectionName" },
                "Please provide the frame definition, either by point names or by start/end coordinates."),

            new ToolSchema(
                "FrameObj_SetSection",
                "Assign",
                "FrameObj",
                new[] { "frameNames", "sectionName" },
                new string[0],
                "Please provide the frame name(s) and the section property to assign.")
        };

        public void Validate(CsiRequestTaskClassificationDto task)
        {
            if (task == null)
            {
                return;
            }

            if (task.Parameters == null)
            {
                task.Parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            NormalizeTaskNames(task);
            task.MissingParameters = new List<string>();
            task.ToolName = null;

            if (IsUnknown(task.Action))
            {
                task.MissingParameters.Add("action");
            }

            if (IsUnknown(task.TargetObject))
            {
                task.MissingParameters.Add("targetObject");
                task.ClarificationMessage = GetTargetClarification(task.Action);
                return;
            }

            if (task.MissingParameters.Count > 0)
            {
                task.ClarificationMessage = "Please clarify the action you want to perform.";
                return;
            }

            if (IsActionTarget(task, "Add", "FrameObj"))
            {
                ValidateAddFrame(task);
                return;
            }

            ToolSchema schema = FindSchema(task.Action, task.TargetObject);
            if (schema == null)
            {
                return;
            }

            ApplySchemaValidation(task, schema);
        }

        private void ValidateAddFrame(CsiRequestTaskClassificationDto task)
        {
            ToolSchema byPoint = FindSchema("Add", "FrameObj", "FrameObj_AddByPoint");
            ToolSchema byCoordinate = FindSchema("Add", "FrameObj", "FrameObj_AddByCoordinate");

            bool hasPointGeometry = HasText(task, "pointI") && HasText(task, "pointJ");
            bool hasCoordinateGeometry =
                HasNumber(task, "xi") && HasNumber(task, "yi") && HasNumber(task, "zi") &&
                HasNumber(task, "xj") && HasNumber(task, "yj") && HasNumber(task, "zj");

            if (hasPointGeometry)
            {
                ApplySchemaValidation(task, byPoint);
                return;
            }

            if (hasCoordinateGeometry)
            {
                ApplySchemaValidation(task, byCoordinate);
                return;
            }

            task.MissingParameters.Add("geometry");
            task.ClarificationMessage = byPoint.ClarificationMessage;
        }

        private static void ApplySchemaValidation(CsiRequestTaskClassificationDto task, ToolSchema schema)
        {
            for (int i = 0; i < schema.RequiredParameters.Count; i++)
            {
                string parameterName = schema.RequiredParameters[i];
                if (IsMissing(task, parameterName))
                {
                    task.MissingParameters.Add(parameterName);
                }
            }

            if (task.MissingParameters.Count > 0)
            {
                task.ClarificationMessage = schema.ClarificationMessage;
                return;
            }

            task.ToolName = schema.ToolName;
            task.ClarificationMessage = null;
        }

        private ToolSchema FindSchema(string action, string targetObject)
        {
            for (int i = 0; i < _schemas.Count; i++)
            {
                ToolSchema schema = _schemas[i];
                if (EqualsText(schema.SupportedAction, action) &&
                    EqualsText(schema.TargetObject, targetObject))
                {
                    return schema;
                }
            }

            return null;
        }

        private ToolSchema FindSchema(string action, string targetObject, string toolName)
        {
            for (int i = 0; i < _schemas.Count; i++)
            {
                ToolSchema schema = _schemas[i];
                if (EqualsText(schema.SupportedAction, action) &&
                    EqualsText(schema.TargetObject, targetObject) &&
                    EqualsText(schema.ToolName, toolName))
                {
                    return schema;
                }
            }

            return null;
        }

        private static void NormalizeTaskNames(CsiRequestTaskClassificationDto task)
        {
            task.Action = NormalizeAction(task.Action);
            task.TargetObject = NormalizeTargetObject(task.TargetObject);

            if (task.Parameters == null)
            {
                return;
            }

            CopyParameter(task.Parameters, "name", "pointName");
            CopyParameter(task.Parameters, "userName", "frameName");
            CopyParameter(task.Parameters, "name", "frameName");
            CopyParameter(task.Parameters, "propName", "sectionName");
            CopyParameter(task.Parameters, "pointIName", "pointI");
            CopyParameter(task.Parameters, "pointJName", "pointJ");
        }

        private static string NormalizeAction(string action)
        {
            string value = (action ?? string.Empty).Trim();
            if (EqualsText(value, "Create") || EqualsText(value, "Draw") || EqualsText(value, "Insert"))
            {
                return "Add";
            }

            if (EqualsText(value, "Apply") || EqualsText(value, "Set") || EqualsText(value, "Modify"))
            {
                return "Assign";
            }

            return string.IsNullOrWhiteSpace(value) ? "Unknown" : value;
        }

        private static string NormalizeTargetObject(string targetObject)
        {
            string value = (targetObject ?? string.Empty).Trim();
            if (EqualsText(value, "Point") || EqualsText(value, "Joint"))
            {
                return "PointObj";
            }

            if (EqualsText(value, "Frame") || EqualsText(value, "Beam") || EqualsText(value, "Member") || EqualsText(value, "Column") || EqualsText(value, "Brace"))
            {
                return "FrameObj";
            }

            if (EqualsText(value, "Area"))
            {
                return "AreaObj";
            }

            if (EqualsText(value, "Shell"))
            {
                return "ShellObj";
            }

            return string.IsNullOrWhiteSpace(value) ? "Unknown" : value;
        }

        private static bool IsMissing(CsiRequestTaskClassificationDto task, string parameterName)
        {
            if (EqualsText(parameterName, "frameNames"))
            {
                return !HasText(task, "frameNames");
            }

            if (IsNumericParameter(parameterName))
            {
                return !HasNumber(task, parameterName);
            }

            return !HasText(task, parameterName);
        }

        private static bool HasText(CsiRequestTaskClassificationDto task, string key)
        {
            string value;
            return task.Parameters != null &&
                   task.Parameters.TryGetValue(key, out value) &&
                   !string.IsNullOrWhiteSpace(value);
        }

        private static bool HasNumber(CsiRequestTaskClassificationDto task, string key)
        {
            string value;
            return task.Parameters != null &&
                   task.Parameters.TryGetValue(key, out value) &&
                   double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _);
        }

        private static bool IsNumericParameter(string parameterName)
        {
            return Regex.IsMatch(parameterName ?? string.Empty, @"^(?:x|y|z|xi|yi|zi|xj|yj|zj)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static bool IsUnknown(string value)
        {
            return string.IsNullOrWhiteSpace(value) || EqualsText(value, "Unknown");
        }

        private static bool IsActionTarget(CsiRequestTaskClassificationDto task, string action, string targetObject)
        {
            return EqualsText(task.Action, action) && EqualsText(task.TargetObject, targetObject);
        }

        private static string GetTargetClarification(string action)
        {
            return EqualsText(action, "Add")
                ? "What would you like to add: point, frame, shell/area, load pattern, load case, or load combination?"
                : "Please specify what object the action should apply to.";
        }

        private static void CopyParameter(Dictionary<string, string> parameters, string source, string target)
        {
            string value;
            if (!parameters.ContainsKey(target) &&
                parameters.TryGetValue(source, out value) &&
                !string.IsNullOrWhiteSpace(value))
            {
                parameters[target] = value;
            }
        }

        private static bool EqualsText(string left, string right)
        {
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }

        private sealed class ToolSchema
        {
            public ToolSchema(
                string toolName,
                string supportedAction,
                string targetObject,
                IEnumerable<string> requiredParameters,
                IEnumerable<string> optionalParameters,
                string clarificationMessage)
            {
                ToolName = toolName;
                SupportedAction = supportedAction;
                TargetObject = targetObject;
                RequiredParameters = new List<string>(requiredParameters);
                OptionalParameters = new List<string>(optionalParameters);
                ClarificationMessage = clarificationMessage;
            }

            public string ToolName { get; private set; }
            public string SupportedAction { get; private set; }
            public string TargetObject { get; private set; }
            public List<string> RequiredParameters { get; private set; }
            public List<string> OptionalParameters { get; private set; }
            public string ClarificationMessage { get; private set; }
        }
    }
}
