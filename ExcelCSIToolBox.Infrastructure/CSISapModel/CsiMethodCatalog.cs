using System.Collections.Generic;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    public sealed class CsiMethodCatalog
    {
        public IReadOnlyList<CsiMethodDescriptor> GetReviewedToolDescriptors()
        {
            return new[]
            {
                Tool("Both", "Points", "Creation", "cPointObj", "AddCartesian", "points.add_by_coordinates", CsiMethodRiskLevel.Low, false, true, "Adds one point by Cartesian coordinates."),
                Tool("Both", "Frames", "Creation", "cFrameObj", "AddByCoord", "frames.add_by_coordinates", CsiMethodRiskLevel.Low, false, true, "Adds one frame by end coordinates."),
                Tool("Both", "Frames", "Creation", "cFrameObj", "AddByPoint", "frames.add_by_points", CsiMethodRiskLevel.Low, false, true, "Adds one frame by existing point names."),
                Tool("Both", "Frames", "Assignments", "cFrameObj", "SetSection", "frames.assign_section", CsiMethodRiskLevel.Medium, true, true, "Assigns section property to frame objects."),
                Tool("Both", "Loads", "Frame", "cFrameObj", "SetLoadDistributed", "loads.frame.assign_distributed", CsiMethodRiskLevel.Medium, true, true, "Assigns distributed frame load."),
                Tool("Both", "Loads", "Frame", "cFrameObj", "SetLoadPoint", "loads.frame.assign_point_load", CsiMethodRiskLevel.Medium, true, true, "Assigns frame point load."),
                Tool("Both", "Selection", "General", "cSelect", "ClearSelection", "selection.clear", CsiMethodRiskLevel.Low, false, true, "Clears active selection."),
                Tool("Both", "Frames", "Deletion", "cFrameObj", "Delete", "frames.delete", CsiMethodRiskLevel.High, true, true, "Deletes frame objects."),
                Tool("Both", "Analysis", "Run", "cAnalyze", "RunAnalysis", "analysis.run", CsiMethodRiskLevel.High, true, false, "Runs model analysis."),
                Tool("Both", "Model / File / Units", "File", "cFile", "Save", "file.save_model", CsiMethodRiskLevel.Dangerous, true, false, "Saves model file. Blocked by default.")
            };
        }

        private static CsiMethodDescriptor Tool(
            string productType,
            string category,
            string subCategory,
            string interfaceName,
            string methodName,
            string toolName,
            CsiMethodRiskLevel riskLevel,
            bool requiresConfirmation,
            bool supportsDryRun,
            string description)
        {
            return new CsiMethodDescriptor
            {
                ProductType = productType,
                Category = category,
                SubCategory = subCategory,
                InterfaceName = interfaceName,
                MethodName = methodName,
                ToolName = toolName,
                Parameters = new CsiParameterDescriptor[0],
                ReturnType = "int",
                IsReadOnly = false,
                IsWrite = true,
                RiskLevel = riskLevel,
                RequiresConfirmation = requiresConfirmation,
                SupportsDryRun = supportsDryRun,
                Description = description,
                Notes = "Reviewed against ETABSv1.dll/SAP2000v1.dll signatures and CHM object grouping. CHM files are reference-only and not required at runtime."
            };
        }
    }
}
