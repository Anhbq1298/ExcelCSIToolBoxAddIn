using System.Collections.Generic;
using ExcelCSIToolBox.Application.Tooling.Contracts;

namespace ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI
{
    internal static class SchemaModuleHelpers
    {
        public static ToolSchema Schema(
            string domain,
            string toolName,
            string action,
            string targetObject,
            IEnumerable<string> requiredParameters,
            IEnumerable<string> optionalParameters,
            IEnumerable<string> aliases,
            IEnumerable<string> intentHints,
            string clarificationMessage,
            bool isModelMutating)
        {
            return new ToolSchema
            {
                Domain = domain,
                ToolName = toolName,
                Action = action,
                TargetObject = targetObject,
                RequiredParameters = Parameters(requiredParameters, true),
                OptionalParameters = Parameters(optionalParameters, false),
                Aliases = aliases == null ? new List<string>() : new List<string>(aliases),
                IntentHints = intentHints == null ? new List<string>() : new List<string>(intentHints),
                ClarificationMessage = clarificationMessage,
                IsModelMutating = isModelMutating
            };
        }

        private static List<ToolParameterSchema> Parameters(IEnumerable<string> names, bool required)
        {
            var parameters = new List<ToolParameterSchema>();
            if (names == null)
            {
                return parameters;
            }

            foreach (string name in names)
            {
                parameters.Add(new ToolParameterSchema
                {
                    Name = name,
                    IsRequired = required,
                    ValueType = "string"
                });
            }

            return parameters;
        }
    }
}
