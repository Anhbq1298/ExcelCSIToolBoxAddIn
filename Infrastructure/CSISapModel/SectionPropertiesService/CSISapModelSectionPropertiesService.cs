using System;
using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Data;
using ExcelCSIToolBoxAddIn.Data.DTOs;
using ExcelCSIToolBoxAddIn.Data.Models;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel
{
    internal delegate int CSISapModelSetSectionCreationUnits<TSapModel>(TSapModel sapModel);

    internal delegate int CSISapModelCreateSection<TSapModel, TInput>(
        TSapModel sapModel,
        TInput input);

    internal delegate bool CSISapModelSectionExists<TSapModel>(
        TSapModel sapModel,
        string sectionName);

    internal static class CSISapModelSectionPropertiesService
    {
        internal static OperationResult AddSteelISections(
            ETABSv1.cSapModel sapModel,
            IReadOnlyList<CSISapModelSteelISectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "ETABS",
                "Creating Steel I-Sections...",
                sapModel,
                SetEtabsSectionCreationUnits,
                EtabsFrameSectionExists,
                (model, input) => model.PropFrame.SetISection(input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw, input.B, input.Tf));
            return result;
        }

        internal static OperationResult AddSteelISections(
            SAP2000v1.cSapModel sapModel,
            IReadOnlyList<CSISapModelSteelISectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "SAP2000",
                "Creating Steel I-Sections...",
                sapModel,
                SetSap2000SectionCreationUnits,
                Sap2000FrameSectionExists,
                (model, input) => model.PropFrame.SetISection(input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw, input.B, input.Tf, -1, "", ""));
            return result;
        }

        internal static OperationResult AddSteelChannelSections(
            ETABSv1.cSapModel sapModel,
            IReadOnlyList<CSISapModelSteelChannelSectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "ETABS",
                "Creating Steel Channel Sections...",
                sapModel,
                SetEtabsSectionCreationUnits,
                EtabsFrameSectionExists,
                (model, input) => model.PropFrame.SetChannel(input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw));
            return result;
        }

        internal static OperationResult AddSteelChannelSections(
            SAP2000v1.cSapModel sapModel,
            IReadOnlyList<CSISapModelSteelChannelSectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "SAP2000",
                "Creating Steel Channel Sections...",
                sapModel,
                SetSap2000SectionCreationUnits,
                Sap2000FrameSectionExists,
                (model, input) => model.PropFrame.SetChannel(input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw, -1, "", ""));
            return result;
        }

        internal static OperationResult AddSteelAngleSections(
            ETABSv1.cSapModel sapModel,
            IReadOnlyList<CSISapModelSteelAngleSectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "ETABS",
                "Creating Steel Angle Sections...",
                sapModel,
                SetEtabsSectionCreationUnits,
                EtabsFrameSectionExists,
                (model, input) => model.PropFrame.SetAngle(input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw));
            return result;
        }

        internal static OperationResult AddSteelAngleSections(
            SAP2000v1.cSapModel sapModel,
            IReadOnlyList<CSISapModelSteelAngleSectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "SAP2000",
                "Creating Steel Angle Sections...",
                sapModel,
                SetSap2000SectionCreationUnits,
                Sap2000FrameSectionExists,
                (model, input) => model.PropFrame.SetAngle(input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw, -1, "", ""));
            return result;
        }

        internal static OperationResult AddSteelPipeSections(
            ETABSv1.cSapModel sapModel,
            IReadOnlyList<CSISapModelSteelPipeSectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "ETABS",
                "Creating Steel Pipe Sections...",
                sapModel,
                SetEtabsSectionCreationUnits,
                EtabsFrameSectionExists,
                (model, input) => model.PropFrame.SetPipe(input.SectionName, input.MaterialName, input.OutsideDiameter, input.WallThickness));
            return result;
        }

        internal static OperationResult AddSteelPipeSections(
            SAP2000v1.cSapModel sapModel,
            IReadOnlyList<CSISapModelSteelPipeSectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "SAP2000",
                "Creating Steel Pipe Sections...",
                sapModel,
                SetSap2000SectionCreationUnits,
                Sap2000FrameSectionExists,
                (model, input) => model.PropFrame.SetPipe(input.SectionName, input.MaterialName, input.OutsideDiameter, input.WallThickness, -1, "", ""));
            return result;
        }

        internal static OperationResult AddSteelTubeSections(
            ETABSv1.cSapModel sapModel,
            IReadOnlyList<CSISapModelSteelTubeSectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "ETABS",
                "Creating Steel Tube Sections...",
                sapModel,
                SetEtabsSectionCreationUnits,
                EtabsFrameSectionExists,
                (model, input) => model.PropFrame.SetTube_1(input.SectionName, input.MaterialName, input.H, input.B, input.T, input.T, 0.000000001, -1, "", "Default"));
            return result;
        }

        internal static OperationResult AddSteelTubeSections(
            SAP2000v1.cSapModel sapModel,
            IReadOnlyList<CSISapModelSteelTubeSectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "SAP2000",
                "Creating Steel Tube Sections...",
                sapModel,
                SetSap2000SectionCreationUnits,
                Sap2000FrameSectionExists,
                (model, input) => model.PropFrame.SetTube_1(input.SectionName, input.MaterialName, input.H, input.B, input.T, input.T, 0.000000001, -1, "", ""));
            return result;
        }

        internal static OperationResult AddConcreteRectangleSections(
            ETABSv1.cSapModel sapModel,
            IReadOnlyList<CSISapModelConcreteRectangleSectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "ETABS",
                "Creating Concrete Rectangle Sections...",
                sapModel,
                SetEtabsSectionCreationUnits,
                EtabsFrameSectionExists,
                (model, input) => model.PropFrame.SetRectangle(input.SectionName, input.MaterialName, input.H, input.B));
            return result;
        }

        internal static OperationResult AddConcreteRectangleSections(
            SAP2000v1.cSapModel sapModel,
            IReadOnlyList<CSISapModelConcreteRectangleSectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "SAP2000",
                "Creating Concrete Rectangle Sections...",
                sapModel,
                SetSap2000SectionCreationUnits,
                Sap2000FrameSectionExists,
                (model, input) => model.PropFrame.SetRectangle(input.SectionName, input.MaterialName, input.H, input.B, -1, "", ""));
            return result;
        }

        internal static OperationResult AddConcreteCircleSections(
            ETABSv1.cSapModel sapModel,
            IReadOnlyList<CSISapModelConcreteCircleSectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "ETABS",
                "Creating Concrete Circle Sections...",
                sapModel,
                SetEtabsSectionCreationUnits,
                EtabsFrameSectionExists,
                (model, input) => model.PropFrame.SetCircle(input.SectionName, input.MaterialName, input.D));
            return result;
        }

        internal static OperationResult AddConcreteCircleSections(
            SAP2000v1.cSapModel sapModel,
            IReadOnlyList<CSISapModelConcreteCircleSectionInput> inputs)
        {
            var result = CreateSections(
                inputs,
                "SAP2000",
                "Creating Concrete Circle Sections...",
                sapModel,
                SetSap2000SectionCreationUnits,
                Sap2000FrameSectionExists,
                (model, input) => model.PropFrame.SetCircle(input.SectionName, input.MaterialName, input.D, -1, "", ""));
            return result;
        }

        private static OperationResult CreateSections<TSapModel, TInput>(
            IReadOnlyList<TInput> inputs,
            string productName,
            string title,
            TSapModel sapModel,
            CSISapModelSetSectionCreationUnits<TSapModel> setSectionCreationUnits,
            CSISapModelSectionExists<TSapModel> sectionExists,
            CSISapModelCreateSection<TSapModel, TInput> createSection)
        {
            if (inputs == null || inputs.Count == 0)
            {
                return OperationResult.Failure("No valid rows were found in the selected range.");
            }

            int unitRet = setSectionCreationUnits(sapModel);
            if (unitRet != 0)
            {
                return OperationResult.Failure($"Failed to set {productName} present units to N-mm-C.");
            }

            int failCount = 0;
            var progress = BatchProgressWindow.RunWithProgress(inputs.Count, title, ctx =>
            {
                foreach (var input in inputs)
                {
                    if (ctx.IsCancellationRequested) break;

                    string sectionName = GetSectionName(input);
                    if (sectionExists(sapModel, sectionName))
                    {
                        ctx.IncrementSkipped();
                        continue;
                    }

                    int ret = createSection(sapModel, input);
                    if (ret == 0)
                    {
                        ctx.IncrementRan();
                    }
                    else
                    {
                        failCount++;
                        ctx.IncrementSkipped();
                    }
                }
            });

            string msg = $"Created: {progress.RanCount}, Skipped: {progress.SkippedCount}, Failed: {failCount}";
            if (progress.WasCancelled)
            {
                msg += " (Cancelled)";
            }

            return OperationResult.Success(msg);
        }

        private static int SetEtabsSectionCreationUnits(ETABSv1.cSapModel sapModel)
        {
            return sapModel.SetPresentUnits(ETABSv1.eUnits.N_mm_C);
        }

        private static int SetSap2000SectionCreationUnits(SAP2000v1.cSapModel sapModel)
        {
            return sapModel.SetPresentUnits(SAP2000v1.eUnits.N_mm_C);
        }

        private static bool EtabsFrameSectionExists(ETABSv1.cSapModel sapModel, string sectionName)
        {
            if (string.IsNullOrWhiteSpace(sectionName))
            {
                return false;
            }

            ETABSv1.eFramePropType propType = ETABSv1.eFramePropType.I;
            int ret = sapModel.PropFrame.GetTypeOAPI(sectionName, ref propType);
            return ret == 0;
        }

        private static bool Sap2000FrameSectionExists(SAP2000v1.cSapModel sapModel, string sectionName)
        {
            if (string.IsNullOrWhiteSpace(sectionName))
            {
                return false;
            }

            SAP2000v1.eFramePropType propType = SAP2000v1.eFramePropType.I;
            int ret = sapModel.PropFrame.GetTypeOAPI(sectionName, ref propType);
            return ret == 0;
        }

        private static string GetSectionName<TInput>(TInput input)
        {
            if (input == null)
            {
                return string.Empty;
            }

            var property = input.GetType().GetProperty("SectionName");
            return property == null ? string.Empty : (string)property.GetValue(input, null);
        }
    }
}
