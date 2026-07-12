using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelCSIToolBox.RefBuilder.Parsing;

namespace ExcelCSIToolBox.RefBuilder.Generation
{
    public sealed class ServiceScaffoldGenerator : IServiceScaffoldGenerator
    {
        public void Generate(string infrastructureRoot, IReadOnlyList<ApiMethodDefinition> methods)
        {
            IReadOnlyList<ApiMethodDefinition> pointMethods = methods.Where(m => m.ObjectName == "PointObj").ToList();
            IReadOnlyList<ApiMethodDefinition> frameMethods = methods.Where(m => m.ObjectName == "FrameObj").ToList();
            IReadOnlyList<ApiMethodDefinition> shellMethods = methods.Where(m => m.ObjectName == "AreaObj").ToList();
            IReadOnlyList<ApiMethodDefinition> loadPatternMethods = methods.Where(m => m.ObjectName == "LoadPatterns").ToList();
            IReadOnlyList<ApiMethodDefinition> loadCombinationMethods = methods.Where(m => m.ObjectName == "Combo").ToList();

            WriteSummary(infrastructureRoot, "PointObjectService", pointMethods);
            WriteSummary(infrastructureRoot, "FrameObjectService", frameMethods);
            WriteSummary(infrastructureRoot, "ShellObjectService", shellMethods);
            WriteSummary(infrastructureRoot, "LoadPatternService", loadPatternMethods);
            WriteSummary(infrastructureRoot, "LoadCombinationService", loadCombinationMethods);
            WriteFullReference(infrastructureRoot, methods);
            WriteMethodCatalog(infrastructureRoot, methods);

            WriteCompiledScaffold(infrastructureRoot, "PointObjectService", "CSISapModelPointObjectService.generated.cs", BuildPointObjectScaffold());
            WriteCompiledScaffold(infrastructureRoot, "FrameObjectService", "CSISapModelFrameObjectService.generated.cs", BuildFrameObjectScaffold());
            WriteCompiledScaffold(infrastructureRoot, "ShellObjectService", "CSISapModelShellObjectService.generated.cs", BuildShellObjectScaffold());
        }

        private static void WriteSummary(string infrastructureRoot, string serviceFolder, IReadOnlyList<ApiMethodDefinition> methods)
        {
            string path = Path.Combine(infrastructureRoot, "CSISapModel", serviceFolder, serviceFolder + ".reference.generated.md");
            Directory.CreateDirectory(Path.GetDirectoryName(path));

            StringBuilder builder = new StringBuilder();
            builder.AppendLine("# AUTO-GENERATED FILE. DO NOT EDIT MANUALLY.");
            builder.AppendLine();
            builder.AppendLine("This file is generated from CSI API reference metadata and is for review/scaffold planning.");
            builder.AppendLine("Compiled wrapper scaffolds live beside the manual service files as `*.generated.cs`.");
            builder.AppendLine();

            foreach (ApiMethodDefinition method in methods.OrderBy(m => m.MethodName))
            {
                builder.Append("- ");
                builder.Append(method.ProductName).Append(" ");
                builder.Append(method.InterfaceName).Append(".").Append(method.MethodName);
                builder.Append(" -> ").Append(method.SafetyFlag);
                builder.AppendLine();
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private static void WriteMethodCatalog(string infrastructureRoot, IReadOnlyList<ApiMethodDefinition> methods)
        {
            string path = Path.Combine(infrastructureRoot, "CSISapModel", "CsiMethodCatalog.generated.cs");
            Directory.CreateDirectory(Path.GetDirectoryName(path));

            StringBuilder builder = new StringBuilder();
            builder.AppendLine("using System.Collections.Generic;");
            builder.AppendLine("using ExcelCSIToolBox.Core.Models.CSI;");
            builder.AppendLine();
            builder.AppendLine("namespace ExcelCSIToolBox.Infrastructure.CSISapModel");
            builder.AppendLine("{");
            builder.AppendLine("    // AUTO-GENERATED FILE. DO NOT EDIT MANUALLY.");
            builder.AppendLine("    // Full CSI API reference catalog generated from ETABSv1.dll and SAP2000v1.dll.");
            builder.AppendLine("    // These descriptors are metadata only. Reviewed MCP tools must call explicit safe services.");
            builder.AppendLine("    public sealed partial class CsiMethodCatalog");
            builder.AppendLine("    {");
            builder.AppendLine("        private static IReadOnlyList<CsiMethodDescriptor> GetGeneratedReferenceDescriptors()");
            builder.AppendLine("        {");
            builder.AppendLine("            return new[]");
            builder.AppendLine("            {");

            foreach (ApiMethodDefinition method in methods
                .OrderBy(m => m.ProductName)
                .ThenBy(m => m.Category)
                .ThenBy(m => m.ObjectName)
                .ThenBy(m => m.MethodName)
                .ThenBy(m => m.FullSignature))
            {
                AppendDescriptor(builder, method);
            }

            builder.AppendLine("            };");
            builder.AppendLine("        }");
            builder.AppendLine("    }");
            builder.AppendLine("}");

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private static void AppendDescriptor(StringBuilder builder, ApiMethodDefinition method)
        {
            string riskLevel = GetRiskLevel(method);
            bool isReadOnly = method.SafetyFlag == "ReadOnly";
            bool isWrite = !isReadOnly;
            bool requiresConfirmation = riskLevel == "Medium" || riskLevel == "High" || riskLevel == "Dangerous";
            bool supportsDryRun = isWrite && riskLevel != "Dangerous";

            builder.AppendLine("                new CsiMethodDescriptor");
            builder.AppendLine("                {");
            AppendProperty(builder, "ProductType", method.ProductName, 20);
            AppendProperty(builder, "Category", method.Category, 20);
            AppendProperty(builder, "SubCategory", method.ObjectName, 20);
            AppendProperty(builder, "InterfaceName", method.InterfaceName, 20);
            AppendProperty(builder, "MethodName", method.MethodName, 20);
            builder.AppendLine("                    Parameters = new CsiParameterDescriptor[]");
            builder.AppendLine("                    {");

            foreach (ApiParameterDefinition parameter in method.Parameters)
            {
                builder.AppendLine("                        new CsiParameterDescriptor");
                builder.AppendLine("                        {");
                AppendProperty(builder, "Name", parameter.Name, 28);
                AppendProperty(builder, "TypeName", parameter.TypeName, 28);
                builder.AppendLine("                            IsOut = " + ToCSharpBool(parameter.IsOut) + ",");
                builder.AppendLine("                            IsOptional = false");
                builder.AppendLine("                        },");
            }

            builder.AppendLine("                    },");
            AppendProperty(builder, "ReturnType", method.ReturnType, 20);
            builder.AppendLine("                    IsReadOnly = " + ToCSharpBool(isReadOnly) + ",");
            builder.AppendLine("                    IsWrite = " + ToCSharpBool(isWrite) + ",");
            builder.AppendLine("                    RiskLevel = CsiMethodRiskLevel." + riskLevel + ",");
            builder.AppendLine("                    RequiresConfirmation = " + ToCSharpBool(requiresConfirmation) + ",");
            builder.AppendLine("                    SupportsDryRun = " + ToCSharpBool(supportsDryRun) + ",");
            AppendProperty(builder, "ToolName", string.Empty, 20);
            AppendProperty(builder, "Description", method.FullSignature, 20);
            AppendProperty(builder, "Notes", "Generated reference metadata only. Implement a reviewed Infrastructure wrapper and MCP tool before execution.", 20, false);
            builder.AppendLine("                },");
        }

        private static string GetRiskLevel(ApiMethodDefinition method)
        {
            if (method.SafetyFlag == "ReadOnly")
            {
                return "None";
            }

            if (method.SafetyFlag == "HighRiskWrite")
            {
                if (method.MethodName.IndexOf("Save", System.StringComparison.OrdinalIgnoreCase) >= 0 ||
                    method.MethodName.IndexOf("Open", System.StringComparison.OrdinalIgnoreCase) >= 0 ||
                    method.MethodName.IndexOf("New", System.StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return "Dangerous";
                }

                return "High";
            }

            if (method.MethodName.StartsWith("Add", System.StringComparison.OrdinalIgnoreCase))
            {
                return "Low";
            }

            return "Medium";
        }

        private static void AppendProperty(StringBuilder builder, string name, string value, int indent, bool comma = true)
        {
            builder.Append(new string(' ', indent));
            builder.Append(name).Append(" = \"").Append(EscapeCSharp(value)).Append("\"");
            if (comma)
            {
                builder.Append(",");
            }

            builder.AppendLine();
        }

        private static string EscapeCSharp(string value)
        {
            return (value ?? string.Empty)
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", string.Empty)
                .Replace("\n", "\\n");
        }

        private static string ToCSharpBool(bool value)
        {
            return value ? "true" : "false";
        }

        private static void WriteFullReference(string infrastructureRoot, IReadOnlyList<ApiMethodDefinition> methods)
        {
            string path = Path.Combine(infrastructureRoot, "CSISapModel", "CsiApiReference.generated.md");
            Directory.CreateDirectory(Path.GetDirectoryName(path));

            StringBuilder builder = new StringBuilder();
            builder.AppendLine("# AUTO-GENERATED FILE. DO NOT EDIT MANUALLY.");
            builder.AppendLine();
            builder.AppendLine("Full CSI API reflection catalog from ETABSv1.dll and SAP2000v1.dll.");
            builder.AppendLine("This is development-time reference metadata only. Do not expose methods directly to MCP without reviewed services.");
            builder.AppendLine();

            foreach (IGrouping<string, ApiMethodDefinition> category in methods
                .OrderBy(m => m.Category)
                .ThenBy(m => m.ObjectName)
                .ThenBy(m => m.MethodName)
                .GroupBy(m => m.Category))
            {
                builder.AppendLine("## " + category.Key);
                builder.AppendLine();

                foreach (ApiMethodDefinition method in category)
                {
                    builder.Append("- ");
                    builder.Append(method.ProductName).Append(" ");
                    builder.Append(method.InterfaceName).Append(".").Append(method.MethodName);
                    builder.Append(" -> ").Append(method.SafetyFlag);
                    builder.AppendLine();
                }

                builder.AppendLine();
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private static void WriteCompiledScaffold(string infrastructureRoot, string serviceFolder, string fileName, string content)
        {
            string path = Path.Combine(infrastructureRoot, "CSISapModel", serviceFolder, fileName);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, content, Encoding.UTF8);
        }

        private static string BuildPointObjectScaffold()
        {
            return @"using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.CSISapModel.PointObject;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    // AUTO-GENERATED FILE. DO NOT EDIT MANUALLY.
    // Generated scaffold from CSI API reference metadata. Keep safety/business logic in the manual companion file.
    internal static partial class CSISapModelPointObjectService
    {
        internal delegate int CSISapModelDeletePointGenerated<TSapModel>(TSapModel sapModel, string name);

        internal static OperationResult<IReadOnlyList<string>> GetNameListGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelGetNameList<TSapModel> getNameList)
        {
            return GetNameList(productName, sapModel, getNameList);
        }

        internal static OperationResult<PointObjectInfo> GetCoordCartesianGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string pointName,
            CSISapModelReadPointCoordinates<TSapModel> getPointCoordinates)
        {
            return GetByName(productName, sapModel, pointName, getPointCoordinates, null);
        }

        internal static OperationResult<CSISapModelAddPointsResultDTO> AddCartesianGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelPointCartesianInput pointInput,
            CSISapModelAddCartesianPoint<TSapModel> addPoint,
            Func<TSapModel, OperationResult> refreshView)
        {
            return AddPointsByCartesian(new[] { pointInput }, productName, sapModel, addPoint, refreshView);
        }

        internal static OperationResult DeleteGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string pointName,
            CSISapModelDeletePointGenerated<TSapModel> deletePoint)
        {
            if (string.IsNullOrWhiteSpace(pointName))
            {
                return OperationResult.Failure(""Point name is required."");
            }

            int result = deletePoint(sapModel, pointName.Trim());
            return result == 0
                ? OperationResult.Success($""Deleted {productName} point '{pointName}'."")
                : OperationResult.Failure($""{productName} PointObj.Delete failed for '{pointName}' (return code {result})."");
        }
    }
}
";
        }

        private static string BuildFrameObjectScaffold()
        {
            return @"using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.CSISapModel.FrameObject;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    // AUTO-GENERATED FILE. DO NOT EDIT MANUALLY.
    // Generated scaffold from CSI API reference metadata. Keep safety/business logic in the manual companion file.
    internal static partial class CSISapModelFrameObjectService
    {
        internal delegate int CSISapModelSetFrameSectionGenerated<TSapModel>(TSapModel sapModel, string name, string propertyName);
        internal delegate int CSISapModelDeleteFrameGenerated<TSapModel>(TSapModel sapModel, string name);

        internal static OperationResult<IReadOnlyList<string>> GetNameListGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelGetNameList<TSapModel> getNameList)
        {
            return GetNameList(productName, sapModel, getNameList);
        }

        internal static OperationResult<FrameEndPointInfo> GetPointsGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string frameName,
            CSISapModelReadFramePoints<TSapModel> getPoints)
        {
            return GetPoints(productName, sapModel, frameName, getPoints);
        }

        internal static OperationResult<FrameSectionInfo> GetSectionGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string frameName,
            CSISapModelReadFrameSection<TSapModel> getSection)
        {
            return GetSection(productName, sapModel, frameName, getSection);
        }

        internal static OperationResult<CSISapModelAddFramesResultDTO> AddByCoordGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelFrameByCoordInput frameInput,
            CSISapModelAddFrameByCoord<TSapModel> addFrame,
            Func<TSapModel, OperationResult> refreshView)
        {
            return AddFramesByCoordinates(new[] { frameInput }, productName, sapModel, addFrame, refreshView);
        }

        internal static OperationResult<CSISapModelAddFramesResultDTO> AddByPointGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelFrameByPointInput frameInput,
            CSISapModelAddFrameByPoint<TSapModel> addFrame,
            Func<TSapModel, OperationResult> refreshView)
        {
            return AddFramesByPoint(new[] { frameInput }, productName, sapModel, addFrame, refreshView);
        }

        internal static OperationResult SetSectionGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string frameName,
            string propertyName,
            CSISapModelSetFrameSectionGenerated<TSapModel> setSection)
        {
            if (string.IsNullOrWhiteSpace(frameName) || string.IsNullOrWhiteSpace(propertyName))
            {
                return OperationResult.Failure(""Frame name and section property are required."");
            }

            int result = setSection(sapModel, frameName.Trim(), propertyName.Trim());
            return result == 0
                ? OperationResult.Success($""Assigned section '{propertyName}' to {productName} frame '{frameName}'."")
                : OperationResult.Failure($""{productName} FrameObj.SetSection failed for '{frameName}' (return code {result})."");
        }

        internal static OperationResult DeleteGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string frameName,
            CSISapModelDeleteFrameGenerated<TSapModel> deleteFrame)
        {
            if (string.IsNullOrWhiteSpace(frameName))
            {
                return OperationResult.Failure(""Frame name is required."");
            }

            int result = deleteFrame(sapModel, frameName.Trim());
            return result == 0
                ? OperationResult.Success($""Deleted {productName} frame '{frameName}'."")
                : OperationResult.Failure($""{productName} FrameObj.Delete failed for '{frameName}' (return code {result})."");
        }
    }
}
";
        }

        private static string BuildShellObjectScaffold()
        {
            return @"using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    // AUTO-GENERATED FILE. DO NOT EDIT MANUALLY.
    // Generated scaffold from CSI API reference metadata. Keep safety/business logic in the manual companion file.
    internal static partial class CSISapModelShellObjectService
    {
        internal delegate int CSISapModelSetShellPropertyGenerated<TSapModel>(TSapModel sapModel, string name, string propertyName);
        internal delegate int CSISapModelDeleteShellGenerated<TSapModel>(TSapModel sapModel, string name);

        internal static OperationResult<IReadOnlyList<string>> GetNameListGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            CSISapModelAreaGetNameList<TSapModel> getNameList)
        {
            return GetNameList<TSapModel>(sapModel, productName, getNameList);
        }

        internal static OperationResult<IReadOnlyList<string>> GetPointsGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string areaName,
            CSISapModelAreaGetPoints<TSapModel> getPoints)
        {
            return GetPoints<TSapModel>(sapModel, productName, areaName, getPoints);
        }

        internal static OperationResult SetPropertyGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string areaName,
            string propertyName,
            CSISapModelSetShellPropertyGenerated<TSapModel> setProperty)
        {
            if (string.IsNullOrWhiteSpace(areaName) || string.IsNullOrWhiteSpace(propertyName))
            {
                return OperationResult.Failure(""Shell/area name and property name are required."");
            }

            int result = setProperty(sapModel, areaName.Trim(), propertyName.Trim());
            return result == 0
                ? OperationResult.Success($""Assigned property '{propertyName}' to {productName} shell/area '{areaName}'."")
                : OperationResult.Failure($""{productName} AreaObj.SetProperty failed for '{areaName}' (return code {result})."");
        }

        internal static OperationResult DeleteGenerated<TSapModel>(
            string productName,
            TSapModel sapModel,
            string areaName,
            CSISapModelDeleteShellGenerated<TSapModel> deleteShell)
        {
            if (string.IsNullOrWhiteSpace(areaName))
            {
                return OperationResult.Failure(""Shell/area name is required."");
            }

            int result = deleteShell(sapModel, areaName.Trim());
            return result == 0
                ? OperationResult.Success($""Deleted {productName} shell/area '{areaName}'."")
                : OperationResult.Failure($""{productName} AreaObj.Delete failed for '{areaName}' (return code {result})."");
        }
    }
}
";
        }
    }
}
