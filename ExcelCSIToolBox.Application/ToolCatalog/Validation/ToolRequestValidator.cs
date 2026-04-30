using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using ExcelCSIToolBox.Application.ToolCatalog.Contracts;
using ExcelCSIToolBox.Application.ToolCatalog.Registry;

namespace ExcelCSIToolBox.Application.ToolCatalog.Validation
{
    public sealed class ToolRequestValidator
    {
        private readonly ToolSchemaRegistry _registry;

        public ToolRequestValidator(ToolSchemaRegistry registry)
        {
            _registry = registry ?? throw new ArgumentNullException(nameof(registry));
        }

        public Contracts.ToolValidationResult Validate(ToolRequest request)
        {
            var result = new Contracts.ToolValidationResult
            {
                IsValid = false,
                MissingParameters = new List<string>()
            };

            if (request == null)
            {
                result.MissingParameters.Add("request");
                result.ClarificationMessage = "Please provide a request.";
                return result;
            }

            NormalizeRequest(request);

            if (IsUnknown(request.Action))
            {
                result.MissingParameters.Add("action");
            }

            if (IsUnknown(request.TargetObject))
            {
                result.MissingParameters.Add("targetObject");
                result.ClarificationMessage = GetTargetClarification(request.Action);
                return result;
            }

            if (result.MissingParameters.Count > 0)
            {
                result.ClarificationMessage = "Please clarify the action you want to perform.";
                return result;
            }

            IReadOnlyList<ToolSchema> candidates = _registry.FindByActionAndTarget(request.Action, request.TargetObject);
            if (candidates.Count == 0)
            {
                return result;
            }

            ToolSchema bestSchema = ChooseSchema(request, candidates);
            if (bestSchema == null)
            {
                bestSchema = candidates[0];
            }

            ValidateRequiredParameters(request, bestSchema, result);
            result.Schema = bestSchema;
            if (result.MissingParameters.Count > 0)
            {
                result.ClarificationMessage = ResolveClarification(bestSchema, result.MissingParameters);
                return result;
            }

            result.IsValid = true;
            result.ToolName = bestSchema.ToolName;
            return result;
        }

        private static ToolSchema ChooseSchema(ToolRequest request, IReadOnlyList<ToolSchema> candidates)
        {
            if (EqualsText(request.Action, "Add") && EqualsText(request.TargetObject, "FrameObject"))
            {
                bool hasPointGeometry = HasText(request, "pointI") && HasText(request, "pointJ");
                bool hasCoordinateGeometry =
                    HasNumber(request, "xi") && HasNumber(request, "yi") && HasNumber(request, "zi") &&
                    HasNumber(request, "xj") && HasNumber(request, "yj") && HasNumber(request, "zj");

                if (!hasPointGeometry && !hasCoordinateGeometry)
                {
                    return FindCandidate(candidates, "FrameObject_AddByPoint");
                }

                return hasPointGeometry
                    ? FindCandidate(candidates, "FrameObject_AddByPoint")
                    : FindCandidate(candidates, "FrameObject_AddByCoordinate");
            }

            for (int i = 0; i < candidates.Count; i++)
            {
                if (HasRequiredParameters(request, candidates[i]))
                {
                    return candidates[i];
                }
            }

            return candidates.Count > 0 ? candidates[0] : null;
        }

        private static ToolSchema FindCandidate(IReadOnlyList<ToolSchema> candidates, string toolName)
        {
            for (int i = 0; i < candidates.Count; i++)
            {
                if (EqualsText(candidates[i].ToolName, toolName))
                {
                    return candidates[i];
                }
            }

            return null;
        }

        private static void ValidateRequiredParameters(ToolRequest request, ToolSchema schema, Contracts.ToolValidationResult result)
        {
            if (schema == null || schema.RequiredParameters == null)
            {
                return;
            }

            for (int i = 0; i < schema.RequiredParameters.Count; i++)
            {
                string parameterName = schema.RequiredParameters[i].Name;
                if (IsMissing(request, parameterName))
                {
                    if (EqualsText(request.Action, "Add") &&
                        EqualsText(request.TargetObject, "FrameObject") &&
                        (EqualsText(parameterName, "pointI") || EqualsText(parameterName, "pointJ")))
                    {
                        AddMissingOnce(result.MissingParameters, "geometry");
                    }
                    else
                    {
                        AddMissingOnce(result.MissingParameters, parameterName);
                    }
                }
            }
        }

        private static bool HasRequiredParameters(ToolRequest request, ToolSchema schema)
        {
            if (schema == null || schema.RequiredParameters == null)
            {
                return false;
            }

            for (int i = 0; i < schema.RequiredParameters.Count; i++)
            {
                if (IsMissing(request, schema.RequiredParameters[i].Name))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IsMissing(ToolRequest request, string parameterName)
        {
            if (IsNumericParameter(parameterName))
            {
                return !HasNumber(request, parameterName);
            }

            return !HasText(request, parameterName);
        }

        private static bool HasText(ToolRequest request, string key)
        {
            string value;
            return request.Parameters != null &&
                   request.Parameters.TryGetValue(key, out value) &&
                   !string.IsNullOrWhiteSpace(value);
        }

        private static bool HasNumber(ToolRequest request, string key)
        {
            string value;
            return request.Parameters != null &&
                   request.Parameters.TryGetValue(key, out value) &&
                   double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _);
        }

        private static void NormalizeRequest(ToolRequest request)
        {
            request.Action = NormalizeAction(request.Action);
            request.TargetObject = NormalizeTargetObject(request.TargetObject);

            if (request.Parameters == null)
            {
                request.Parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            CopyParameter(request.Parameters, "name", "pointName");
            CopyParameter(request.Parameters, "userName", "frameName");
            CopyParameter(request.Parameters, "name", "frameName");
            CopyParameter(request.Parameters, "propName", "sectionName");
            CopyParameter(request.Parameters, "pointIName", "pointI");
            CopyParameter(request.Parameters, "pointJName", "pointJ");
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
            if (EqualsText(value, "PointObj") || EqualsText(value, "Point") || EqualsText(value, "Joint"))
            {
                return "PointObject";
            }

            if (EqualsText(value, "FrameObj") || EqualsText(value, "Frame") || EqualsText(value, "Beam") || EqualsText(value, "Member") || EqualsText(value, "Column") || EqualsText(value, "Brace"))
            {
                return "FrameObject";
            }

            if (EqualsText(value, "AreaObj") || EqualsText(value, "ShellObj") || EqualsText(value, "Area") || EqualsText(value, "Shell"))
            {
                return "ShellObject";
            }

            return string.IsNullOrWhiteSpace(value) ? "Unknown" : value;
        }

        private static string ResolveClarification(ToolSchema schema, IReadOnlyList<string> missingParameters)
        {
            if (schema != null && !string.IsNullOrWhiteSpace(schema.ClarificationMessage))
            {
                return schema.ClarificationMessage;
            }

            return "Please provide the missing required parameter(s): " + string.Join(", ", missingParameters);
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

        private static void AddMissingOnce(List<string> missingParameters, string parameterName)
        {
            for (int i = 0; i < missingParameters.Count; i++)
            {
                if (EqualsText(missingParameters[i], parameterName))
                {
                    return;
                }
            }

            missingParameters.Add(parameterName);
        }

        private static bool IsNumericParameter(string parameterName)
        {
            return Regex.IsMatch(parameterName ?? string.Empty, @"^(?:x|y|z|xi|yi|zi|xj|yj|zj)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static bool IsUnknown(string value)
        {
            return string.IsNullOrWhiteSpace(value) || EqualsText(value, "Unknown");
        }

        private static bool EqualsText(string left, string right)
        {
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }
    }
}
