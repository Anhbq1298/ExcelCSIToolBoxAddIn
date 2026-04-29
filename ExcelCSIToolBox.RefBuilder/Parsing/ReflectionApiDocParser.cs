using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelCSIToolBox.RefBuilder.Parsing
{
    public sealed class ReflectionApiDocParser : IApiDocParser
    {
        public IReadOnlyList<ApiMethodDefinition> Parse(string productName, string dllPath)
        {
            if (!File.Exists(dllPath))
            {
                return new List<ApiMethodDefinition>();
            }

            Assembly assembly = Assembly.LoadFrom(dllPath);
            string rootNamespace = productName.Equals("ETABS", StringComparison.OrdinalIgnoreCase) ? "ETABSv1" : "SAP2000v1";
            List<ApiMethodDefinition> methods = new List<ApiMethodDefinition>();

            foreach (Type type in assembly.GetTypes()
                .Where(t => t.IsInterface && t.FullName != null && t.FullName.StartsWith(rootNamespace + ".c", StringComparison.Ordinal))
                .OrderBy(t => t.Name))
            {
                foreach (MethodInfo method in type.GetMethods()
                    .Where(m => !m.IsSpecialName)
                    .OrderBy(m => m.Name)
                    .ThenBy(m => BuildSignatureKey(m)))
                {
                    methods.Add(CreateMethod(productName, GetObjectName(type), type, method, dllPath));
                }
            }

            return methods;
        }

        private static ApiMethodDefinition CreateMethod(string productName, string objectName, Type type, MethodInfo method, string dllPath)
        {
            ApiMethodDefinition definition = new ApiMethodDefinition
            {
                ProductName = productName,
                ObjectName = objectName,
                InterfaceName = type.FullName,
                MethodName = method.Name,
                ReturnType = method.ReturnType.Name,
                Category = ClassifyCategory(objectName, method.Name),
                SafetyFlag = Classify(objectName, method.Name),
                SourceDocumentationFile = dllPath
            };

            foreach (ParameterInfo parameter in method.GetParameters())
            {
                definition.Parameters.Add(new ApiParameterDefinition
                {
                    Name = parameter.Name,
                    TypeName = CleanTypeName(parameter.ParameterType),
                    IsByRef = parameter.ParameterType.IsByRef,
                    IsOut = parameter.IsOut,
                    Description = string.Empty
                });
            }

            definition.FullSignature = method.ReturnType.Name + " " + method.Name + "(" +
                string.Join(", ", definition.Parameters.Select(p => p.TypeName + " " + p.Name)) + ")";
            return definition;
        }

        private static string GetObjectName(Type type)
        {
            return type.Name.StartsWith("c", StringComparison.Ordinal) && type.Name.Length > 1
                ? type.Name.Substring(1)
                : type.Name;
        }

        private static string BuildSignatureKey(MethodInfo method)
        {
            return method.Name + "(" + string.Join(",", method.GetParameters().Select(p => CleanTypeName(p.ParameterType))) + ")";
        }

        private static string CleanTypeName(Type type)
        {
            Type valueType = type.IsByRef ? type.GetElementType() : type;
            return valueType == null ? type.Name : valueType.Name;
        }

        private static string Classify(string objectName, string methodName)
        {
            if (methodName.StartsWith("Get", StringComparison.OrdinalIgnoreCase))
            {
                return "ReadOnly";
            }

            if (methodName.StartsWith("Count", StringComparison.OrdinalIgnoreCase))
            {
                return "ReadOnly";
            }

            if (Contains(objectName, "Result") &&
                !methodName.StartsWith("Set", StringComparison.OrdinalIgnoreCase) &&
                !methodName.StartsWith("Delete", StringComparison.OrdinalIgnoreCase))
            {
                return "ReadOnly";
            }

            if (methodName.IndexOf("Delete", StringComparison.OrdinalIgnoreCase) >= 0 ||
                methodName.IndexOf("Remove", StringComparison.OrdinalIgnoreCase) >= 0 ||
                methodName.IndexOf("Save", StringComparison.OrdinalIgnoreCase) >= 0 ||
                methodName.IndexOf("Open", StringComparison.OrdinalIgnoreCase) >= 0 ||
                methodName.IndexOf("Run", StringComparison.OrdinalIgnoreCase) >= 0 ||
                methodName.IndexOf("New", StringComparison.OrdinalIgnoreCase) >= 0 ||
                methodName.IndexOf("Lock", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "HighRiskWrite";
            }

            if (methodName.StartsWith("Set", StringComparison.OrdinalIgnoreCase) ||
                methodName.StartsWith("Add", StringComparison.OrdinalIgnoreCase) ||
                methodName.StartsWith("Change", StringComparison.OrdinalIgnoreCase))
            {
                return "Write";
            }

            return "Write";
        }

        private static string ClassifyCategory(string objectName, string methodName)
        {
            string combined = objectName + "." + methodName;

            if (Contains(combined, "PointObj"))
            {
                return "Points";
            }

            if (Contains(combined, "FrameObj"))
            {
                return "Frames";
            }

            if (Contains(combined, "AreaObj") || Contains(combined, "Shell"))
            {
                return "Shells / Areas";
            }

            if (Contains(combined, "LoadPattern"))
            {
                return "Load Patterns";
            }

            if (Contains(combined, "Result"))
            {
                return "Results";
            }

            if (Contains(combined, "LoadCase") || Contains(combined, "Case"))
            {
                return "Load Cases";
            }

            if (Contains(combined, "Combo"))
            {
                return "Load Combinations";
            }

            if (Contains(combined, "Load"))
            {
                return "Loads";
            }

            if (Contains(combined, "Prop") || Contains(combined, "Section"))
            {
                return "Sections / Properties";
            }

            if (Contains(combined, "Analyze") || Contains(combined, "Analysis"))
            {
                return "Analysis";
            }

            if (Contains(combined, "Design"))
            {
                return "Design";
            }

            if (Contains(combined, "File"))
            {
                return "Model / File / Units";
            }

            return "Other";
        }

        private static bool Contains(string value, string token)
        {
            return value.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
