using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelCSIToolBox.RefBuilder.Parsing;

namespace ExcelCSIToolBox.RefBuilder.Indexing
{
    public sealed class ApiIndexWriter
    {
        public void Write(string outputPath, IReadOnlyList<ApiMethodDefinition> methods)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));
            File.WriteAllText(outputPath, ToJson(methods), Encoding.UTF8);
        }

        private static string ToJson(IReadOnlyList<ApiMethodDefinition> methods)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("[");
            for (int i = 0; i < methods.Count; i++)
            {
                ApiMethodDefinition method = methods[i];
                builder.AppendLine("  {");
                Append(builder, "productName", method.ProductName, true);
                Append(builder, "objectName", method.ObjectName, true);
                Append(builder, "interfaceName", method.InterfaceName, true);
                Append(builder, "methodName", method.MethodName, true);
                Append(builder, "returnType", method.ReturnType, true);
                Append(builder, "fullSignature", method.FullSignature, true);
                Append(builder, "category", method.Category, true);
                Append(builder, "safetyFlag", method.SafetyFlag, true);
                Append(builder, "sourceDocumentationFile", method.SourceDocumentationFile, true);
                builder.AppendLine("    \"parameters\": [");
                for (int p = 0; p < method.Parameters.Count; p++)
                {
                    ApiParameterDefinition parameter = method.Parameters[p];
                    builder.AppendLine("      {");
                    Append(builder, "name", parameter.Name, true, 8);
                    Append(builder, "typeName", parameter.TypeName, true, 8);
                    builder.AppendLine("        \"isByRef\": " + parameter.IsByRef.ToString().ToLowerInvariant() + ",");
                    builder.AppendLine("        \"isOut\": " + parameter.IsOut.ToString().ToLowerInvariant() + ",");
                    Append(builder, "description", parameter.Description, false, 8);
                    builder.Append("      }");
                    builder.AppendLine(p == method.Parameters.Count - 1 ? string.Empty : ",");
                }
                builder.AppendLine("    ]");
                builder.Append("  }");
                builder.AppendLine(i == methods.Count - 1 ? string.Empty : ",");
            }
            builder.AppendLine("]");
            return builder.ToString();
        }

        private static void Append(StringBuilder builder, string name, string value, bool comma, int spaces = 4)
        {
            builder.Append(new string(' ', spaces));
            builder.Append('"').Append(name).Append("\": \"").Append(Escape(value)).Append('"');
            if (comma)
            {
                builder.Append(',');
            }
            builder.AppendLine();
        }

        private static string Escape(string value)
        {
            return (value ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"");
        }
    }
}
