using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.Core.Abstractions.CSI
{
    public interface ICsiModelCommandService
    {
        CsiWritePreview PreviewAddPoint(double x, double y, double z, string userName);
        OperationResult AddPoint(double x, double y, double z, string userName, bool confirmed);

        CsiWritePreview PreviewAddFrameByCoordinates(
            double xi, double yi, double zi,
            double xj, double yj, double zj,
            string sectionName,
            string userName);

        OperationResult AddFrameByCoordinates(
            double xi, double yi, double zi,
            double xj, double yj, double zj,
            string sectionName,
            string userName,
            bool confirmed);

        CsiWritePreview PreviewAddFrameByPoints(string point1Name, string point2Name, string sectionName, string userName);
        OperationResult AddFrameByPoints(string point1Name, string point2Name, string sectionName, string userName, bool confirmed);

        CsiWritePreview PreviewAssignFrameSection(IReadOnlyList<string> frameNames, string sectionName);
        OperationResult AssignFrameSection(IReadOnlyList<string> frameNames, string sectionName, bool confirmed);

        CsiWritePreview PreviewAssignFrameDistributedLoad(
            IReadOnlyList<string> frameNames,
            string loadPattern,
            int direction,
            double value1,
            double value2);

        OperationResult AssignFrameDistributedLoad(
            IReadOnlyList<string> frameNames,
            string loadPattern,
            int direction,
            double value1,
            double value2,
            bool confirmed);

        CsiWritePreview PreviewAssignFramePointLoad(
            IReadOnlyList<string> frameNames,
            string loadPattern,
            int direction,
            double distance,
            double value);

        OperationResult AssignFramePointLoad(
            IReadOnlyList<string> frameNames,
            string loadPattern,
            int direction,
            double distance,
            double value,
            bool confirmed);

        CsiWritePreview PreviewSetObjectSelection(IReadOnlyList<string> objectNames, string objectType);
        OperationResult SetObjectSelection(IReadOnlyList<string> objectNames, string objectType, bool confirmed);

        CsiWritePreview PreviewClearSelection();
        OperationResult ClearSelection(bool confirmed);

        CsiWritePreview PreviewDeleteObjects(IReadOnlyList<string> objectNames, string objectType);
        OperationResult DeleteObjects(IReadOnlyList<string> objectNames, string objectType, bool confirmed);

        CsiWritePreview PreviewRunAnalysis();
        OperationResult RunAnalysis(bool confirmed);

        CsiWritePreview PreviewSaveModel();
        OperationResult SaveModel(bool confirmed);
    }
}
