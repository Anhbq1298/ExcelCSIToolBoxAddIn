using System;
using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
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

    internal delegate int CSISapModelSetSteelISection<TSapModel>(
        TSapModel sapModel,
        string sectionName,
        string materialName,
        double h,
        double b,
        double tf,
        double tw);

    internal delegate int CSISapModelSetSteelChannelSection<TSapModel>(
        TSapModel sapModel,
        string sectionName,
        string materialName,
        double h,
        double b,
        double tf,
        double tw);

    internal delegate int CSISapModelSetSteelAngleSection<TSapModel>(
        TSapModel sapModel,
        string sectionName,
        string materialName,
        double h,
        double b,
        double tf,
        double tw);

    internal delegate int CSISapModelSetSteelPipeSection<TSapModel>(
        TSapModel sapModel,
        string sectionName,
        string materialName,
        double outsideDiameter,
        double wallThickness);

    internal delegate int CSISapModelSetSteelTubeSection<TSapModel>(
        TSapModel sapModel,
        string sectionName,
        string materialName,
        double h,
        double b,
        double t);

    internal delegate int CSISapModelSetConcreteRectangleSection<TSapModel>(
        TSapModel sapModel,
        string sectionName,
        string materialName,
        double h,
        double b);

    internal delegate int CSISapModelSetConcreteCircleSection<TSapModel>(
        TSapModel sapModel,
        string sectionName,
        string materialName,
        double d);

    internal static class CSISapModelSectionPropertiesService
    {
        internal static OperationResult AddSteelISections<TSapModel>(
            IReadOnlyList<CSISapModelSteelISectionInput> inputs,
            string productName,
            TSapModel sapModel,
            CSISapModelSetSectionCreationUnits<TSapModel> setSectionCreationUnits,
            CSISapModelSectionExists<TSapModel> sectionExists,
            CSISapModelSetSteelISection<TSapModel> setSection)
        {
            return CreateSections(
                inputs,
                productName,
                "Creating Steel I-Sections...",
                sapModel,
                setSectionCreationUnits,
                sectionExists,
                (model, input) => setSection(model, input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw));
        }

        internal static OperationResult AddSteelChannelSections<TSapModel>(
            IReadOnlyList<CSISapModelSteelChannelSectionInput> inputs,
            string productName,
            TSapModel sapModel,
            CSISapModelSetSectionCreationUnits<TSapModel> setSectionCreationUnits,
            CSISapModelSectionExists<TSapModel> sectionExists,
            CSISapModelSetSteelChannelSection<TSapModel> setSection)
        {
            return CreateSections(
                inputs,
                productName,
                "Creating Steel Channel Sections...",
                sapModel,
                setSectionCreationUnits,
                sectionExists,
                (model, input) => setSection(model, input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw));
        }

        internal static OperationResult AddSteelAngleSections<TSapModel>(
            IReadOnlyList<CSISapModelSteelAngleSectionInput> inputs,
            string productName,
            TSapModel sapModel,
            CSISapModelSetSectionCreationUnits<TSapModel> setSectionCreationUnits,
            CSISapModelSectionExists<TSapModel> sectionExists,
            CSISapModelSetSteelAngleSection<TSapModel> setSection)
        {
            return CreateSections(
                inputs,
                productName,
                "Creating Steel Angle Sections...",
                sapModel,
                setSectionCreationUnits,
                sectionExists,
                (model, input) => setSection(model, input.SectionName, input.MaterialName, input.H, input.B, input.Tf, input.Tw));
        }

        internal static OperationResult AddSteelPipeSections<TSapModel>(
            IReadOnlyList<CSISapModelSteelPipeSectionInput> inputs,
            string productName,
            TSapModel sapModel,
            CSISapModelSetSectionCreationUnits<TSapModel> setSectionCreationUnits,
            CSISapModelSectionExists<TSapModel> sectionExists,
            CSISapModelSetSteelPipeSection<TSapModel> setSection)
        {
            return CreateSections(
                inputs,
                productName,
                "Creating Steel Pipe Sections...",
                sapModel,
                setSectionCreationUnits,
                sectionExists,
                (model, input) => setSection(model, input.SectionName, input.MaterialName, input.OutsideDiameter, input.WallThickness));
        }

        internal static OperationResult AddSteelTubeSections<TSapModel>(
            IReadOnlyList<CSISapModelSteelTubeSectionInput> inputs,
            string productName,
            TSapModel sapModel,
            CSISapModelSetSectionCreationUnits<TSapModel> setSectionCreationUnits,
            CSISapModelSectionExists<TSapModel> sectionExists,
            CSISapModelSetSteelTubeSection<TSapModel> setSection)
        {
            return CreateSections(
                inputs,
                productName,
                "Creating Steel Tube Sections...",
                sapModel,
                setSectionCreationUnits,
                sectionExists,
                (model, input) => setSection(model, input.SectionName, input.MaterialName, input.H, input.B, input.T));
        }

        internal static OperationResult AddConcreteRectangleSections<TSapModel>(
            IReadOnlyList<CSISapModelConcreteRectangleSectionInput> inputs,
            string productName,
            TSapModel sapModel,
            CSISapModelSetSectionCreationUnits<TSapModel> setSectionCreationUnits,
            CSISapModelSectionExists<TSapModel> sectionExists,
            CSISapModelSetConcreteRectangleSection<TSapModel> setSection)
        {
            return CreateSections(
                inputs,
                productName,
                "Creating Concrete Rectangle Sections...",
                sapModel,
                setSectionCreationUnits,
                sectionExists,
                (model, input) => setSection(model, input.SectionName, input.MaterialName, input.H, input.B));
        }

        internal static OperationResult AddConcreteCircleSections<TSapModel>(
            IReadOnlyList<CSISapModelConcreteCircleSectionInput> inputs,
            string productName,
            TSapModel sapModel,
            CSISapModelSetSectionCreationUnits<TSapModel> setSectionCreationUnits,
            CSISapModelSectionExists<TSapModel> sectionExists,
            CSISapModelSetConcreteCircleSection<TSapModel> setSection)
        {
            return CreateSections(
                inputs,
                productName,
                "Creating Concrete Circle Sections...",
                sapModel,
                setSectionCreationUnits,
                sectionExists,
                (model, input) => setSection(model, input.SectionName, input.MaterialName, input.D));
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
