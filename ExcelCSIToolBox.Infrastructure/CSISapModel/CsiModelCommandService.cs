using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Data.CSISapModel.PointObject;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Data.Models;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    public sealed class CsiModelCommandService : ICsiModelCommandService
    {
        private readonly ICSISapModelConnectionService _etabsService;
        private readonly ICSISapModelConnectionService _sap2000Service;
        private readonly IMcpWriteGuard _writeGuard;
        private readonly CsiOperationLogger _logger;

        public CsiModelCommandService(
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service,
            IMcpWriteGuard writeGuard,
            CsiOperationLogger logger)
        {
            _etabsService = etabsService ?? throw new ArgumentNullException(nameof(etabsService));
            _sap2000Service = sap2000Service ?? throw new ArgumentNullException(nameof(sap2000Service));
            _writeGuard = writeGuard ?? throw new ArgumentNullException(nameof(writeGuard));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public CsiWritePreview PreviewAddPoint(double x, double y, double z, string userName)
        {
            return Preview("points.add_by_coordinates", CsiMethodRiskLevel.Low, false, true,
                $"This will add one point at ({x}, {y}, {z}) with name '{CleanName(userName)}'.",
                One(CleanName(userName)));
        }

        public OperationResult AddPoint(double x, double y, double z, string userName, bool confirmed)
        {
            return Execute("points.add_by_coordinates", "Points", "Creation", CsiMethodRiskLevel.Low, confirmed, One(CleanName(userName)),
                $"x={x}, y={y}, z={z}, userName={CleanName(userName)}",
                service =>
                {
                    string requestedName = CleanName(userName);
                    OperationResult<CSISapModelAddPointsResultDTO> addResult = service.AddPointsByCartesian(new[]
                    {
                        new CSISapModelPointCartesianInput
                        {
                            ExcelRowNumber = 1,
                            UniqueName = requestedName,
                            X = x,
                            Y = y,
                            Z = z
                        }
                    });

                    if (!addResult.IsSuccess)
                    {
                        return OperationResult.Failure(addResult.Message);
                    }

                    if (addResult.Data == null || addResult.Data.AddedCount <= 0)
                    {
                        return OperationResult.Failure("PointObj.AddCartesian returned without adding a point.");
                    }

                    string failedMessages = JoinMessages(addResult.Data.FailedRowMessages);
                    if (!string.IsNullOrWhiteSpace(requestedName))
                    {
                        OperationResult<PointObjectInfo> verifyResult = service.GetPointCoordinates(requestedName);
                        if (!verifyResult.IsSuccess)
                        {
                            string detail = string.IsNullOrWhiteSpace(failedMessages) ? string.Empty : " " + failedMessages;
                            return OperationResult.Failure("ETABS/SAP2000 reported success, but the requested point '" + requestedName + "' could not be verified in the active model." + detail);
                        }
                    }

                    string message = "Added point '" + (string.IsNullOrWhiteSpace(requestedName) ? "(auto name)" : requestedName) + "' at (" + x + ", " + y + ", " + z + ").";
                    if (!string.IsNullOrWhiteSpace(failedMessages))
                    {
                        message += " Note: " + failedMessages;
                    }

                    return OperationResult.Success(message);
                });
        }

        public CsiWritePreview PreviewAddFrameByCoordinates(double xi, double yi, double zi, double xj, double yj, double zj, string sectionName, string userName)
        {
            return Preview("frames.add_by_coordinates", CsiMethodRiskLevel.Low, false, true,
                $"This will add one frame from ({xi}, {yi}, {zi}) to ({xj}, {yj}, {zj}) using section '{sectionName}'.",
                One(CleanName(userName)));
        }

        public OperationResult AddFrameByCoordinates(double xi, double yi, double zi, double xj, double yj, double zj, string sectionName, string userName, bool confirmed)
        {
            if (string.IsNullOrWhiteSpace(sectionName))
            {
                return OperationResult.Failure("SectionName is required.");
            }

            return Execute("frames.add_by_coordinates", "Frames", "Creation", CsiMethodRiskLevel.Low, confirmed, One(CleanName(userName)),
                $"section={sectionName}, userName={CleanName(userName)}",
                service => service.AddFramesByCoordinates(new[]
                {
                    new CSISapModelFrameByCoordInput
                    {
                        ExcelRowNumber = 1,
                        UniqueName = CleanName(userName),
                        SectionName = sectionName,
                        Xi = xi,
                        Yi = yi,
                        Zi = zi,
                        Xj = xj,
                        Yj = yj,
                        Zj = zj
                    }
                }));
        }

        public CsiWritePreview PreviewAddFrameByPoints(string point1Name, string point2Name, string sectionName, string userName)
        {
            return Preview("frames.add_by_points", CsiMethodRiskLevel.Low, false, true,
                $"This will add one frame between points '{point1Name}' and '{point2Name}' using section '{sectionName}'.",
                One(CleanName(userName)));
        }

        public OperationResult AddFrameByPoints(string point1Name, string point2Name, string sectionName, string userName, bool confirmed)
        {
            if (string.IsNullOrWhiteSpace(point1Name) || string.IsNullOrWhiteSpace(point2Name))
            {
                return OperationResult.Failure("Point1Name and Point2Name are required.");
            }

            if (string.IsNullOrWhiteSpace(sectionName))
            {
                return OperationResult.Failure("SectionName is required.");
            }

            return Execute("frames.add_by_points", "Frames", "Creation", CsiMethodRiskLevel.Low, confirmed, One(CleanName(userName)),
                $"point1={point1Name}, point2={point2Name}, section={sectionName}, userName={CleanName(userName)}",
                service => service.AddFramesByPoint(new[]
                {
                    new CSISapModelFrameByPointInput
                    {
                        ExcelRowNumber = 1,
                        UniqueName = CleanName(userName),
                        SectionName = sectionName,
                        Point1Name = point1Name,
                        Point2Name = point2Name
                    }
                }));
        }

        public CsiWritePreview PreviewAssignFrameSection(IReadOnlyList<string> frameNames, string sectionName)
        {
            return Preview("frames.assign_section", CsiMethodRiskLevel.Medium, true, true,
                $"This will assign section '{sectionName}' to {Count(frameNames)} frame(s).",
                frameNames);
        }

        public OperationResult AssignFrameSection(IReadOnlyList<string> frameNames, string sectionName, bool confirmed)
        {
            if (frameNames == null || frameNames.Count == 0)
            {
                return OperationResult.Failure("At least one frame name is required.");
            }

            if (string.IsNullOrWhiteSpace(sectionName))
            {
                return OperationResult.Failure("SectionName is required.");
            }

            return Execute("frames.assign_section", "Frames", "Assignments", CsiMethodRiskLevel.Medium, confirmed, frameNames,
                $"section={sectionName}, frames={Count(frameNames)}",
                service => service.AssignFrameSection(frameNames, sectionName));
        }

        public CsiWritePreview PreviewAssignFrameDistributedLoad(IReadOnlyList<string> frameNames, string loadPattern, int direction, double value1, double value2)
        {
            return Preview("loads.frame.assign_distributed", CsiMethodRiskLevel.Medium, true, true,
                $"This will assign distributed load pattern '{loadPattern}' to {Count(frameNames)} frame(s).",
                frameNames);
        }

        public OperationResult AssignFrameDistributedLoad(IReadOnlyList<string> frameNames, string loadPattern, int direction, double value1, double value2, bool confirmed)
        {
            if (frameNames == null || frameNames.Count == 0)
            {
                return OperationResult.Failure("At least one frame name is required.");
            }

            if (string.IsNullOrWhiteSpace(loadPattern))
            {
                return OperationResult.Failure("LoadPattern is required.");
            }

            return Execute("loads.frame.assign_distributed", "Loads", "Frame", CsiMethodRiskLevel.Medium, confirmed, frameNames,
                $"loadPattern={loadPattern}, direction={direction}, value1={value1}, value2={value2}, frames={Count(frameNames)}",
                service => service.AssignFrameDistributedLoad(frameNames, loadPattern, direction, value1, value2));
        }

        public CsiWritePreview PreviewAssignFramePointLoad(IReadOnlyList<string> frameNames, string loadPattern, int direction, double distance, double value)
        {
            return Preview("loads.frame.assign_point_load", CsiMethodRiskLevel.Medium, true, true,
                $"This will assign point load pattern '{loadPattern}' to {Count(frameNames)} frame(s).",
                frameNames);
        }

        public OperationResult AssignFramePointLoad(IReadOnlyList<string> frameNames, string loadPattern, int direction, double distance, double value, bool confirmed)
        {
            if (frameNames == null || frameNames.Count == 0)
            {
                return OperationResult.Failure("At least one frame name is required.");
            }

            if (string.IsNullOrWhiteSpace(loadPattern))
            {
                return OperationResult.Failure("LoadPattern is required.");
            }

            return Execute("loads.frame.assign_point_load", "Loads", "Frame", CsiMethodRiskLevel.Medium, confirmed, frameNames,
                $"loadPattern={loadPattern}, direction={direction}, distance={distance}, value={value}, frames={Count(frameNames)}",
                service => service.AssignFramePointLoad(frameNames, loadPattern, direction, distance, value));
        }

        public CsiWritePreview PreviewSetObjectSelection(IReadOnlyList<string> objectNames, string objectType)
        {
            return Preview("selection.set_objects", CsiMethodRiskLevel.Low, true, true,
                $"This will select {Count(objectNames)} {objectType} object(s).",
                objectNames);
        }

        public OperationResult SetObjectSelection(IReadOnlyList<string> objectNames, string objectType, bool confirmed)
        {
            if (objectNames == null || objectNames.Count == 0)
            {
                return OperationResult.Failure("At least one object name is required.");
            }

            string normalized = (objectType ?? string.Empty).Trim().ToLowerInvariant();
            return Execute("selection.set_objects", "Selection", normalized, CsiMethodRiskLevel.Low, confirmed, objectNames,
                $"objectType={objectType}, count={Count(objectNames)}",
                service =>
                {
                    if (normalized == "point" || normalized == "points")
                    {
                        return service.SelectPointsByUniqueNames(objectNames);
                    }

                    if (normalized == "frame" || normalized == "frames")
                    {
                        return service.SelectFramesByUniqueNames(objectNames);
                    }

                    return OperationResult.Failure("Only point and frame selection is currently supported.");
                });
        }

        public CsiWritePreview PreviewClearSelection()
        {
            return Preview("selection.clear", CsiMethodRiskLevel.Low, true, true,
                "This will clear the active CSI object selection.",
                new string[0]);
        }

        public OperationResult ClearSelection(bool confirmed)
        {
            return Execute("selection.clear", "Selection", "General", CsiMethodRiskLevel.Low, confirmed, new string[0],
                "clear active selection",
                service => service.ClearSelection());
        }

        public CsiWritePreview PreviewDeleteObjects(IReadOnlyList<string> objectNames, string objectType)
        {
            return Preview("frames.delete", CsiMethodRiskLevel.High, true, true,
                $"This will delete {Count(objectNames)} {objectType} object(s). This is high risk.",
                objectNames);
        }

        public OperationResult DeleteObjects(IReadOnlyList<string> objectNames, string objectType, bool confirmed)
        {
            if (objectNames == null || objectNames.Count == 0)
            {
                return OperationResult.Failure("At least one object name is required.");
            }

            string normalized = (objectType ?? string.Empty).Trim().ToLowerInvariant();
            if (normalized != "frame" && normalized != "frames")
            {
                return OperationResult.Failure("Only frame deletion is currently implemented.");
            }

            return Execute("frames.delete", "Frames", "Deletion", CsiMethodRiskLevel.High, confirmed, objectNames,
                $"delete frame count={Count(objectNames)}",
                service => service.DeleteFrameObjects(objectNames));
        }

        public CsiWritePreview PreviewRunAnalysis()
        {
            return Preview("analysis.run", CsiMethodRiskLevel.High, true, false,
                "This will run CSI analysis on the attached model. This is high risk and may take time.",
                new[] { "Active model" });
        }

        public OperationResult RunAnalysis(bool confirmed)
        {
            return Execute("analysis.run", "Analysis", "Run", CsiMethodRiskLevel.High, confirmed, new[] { "Active model" },
                "run analysis",
                service => service.RunAnalysis());
        }

        public CsiWritePreview PreviewSaveModel()
        {
            return Preview("file.save_model", CsiMethodRiskLevel.Dangerous, true, false,
                "Saving the model is dangerous and blocked by default for AI usage.",
                new[] { "Active model file" });
        }

        public OperationResult SaveModel(bool confirmed)
        {
            return Execute("file.save_model", "Model / File / Units", "File", CsiMethodRiskLevel.Dangerous, confirmed, new[] { "Active model file" },
                "save model",
                service => service.SaveModel());
        }

        private OperationResult Execute(
            string operationName,
            string category,
            string subCategory,
            CsiMethodRiskLevel riskLevel,
            bool confirmed,
            IReadOnlyList<string> affectedObjects,
            string argumentsSummary,
            Func<ICSISapModelConnectionService, OperationResult> action)
        {
            OperationResult guardResult = _writeGuard.ValidateWrite(operationName, riskLevel, confirmed, affectedObjects);
            if (!guardResult.IsSuccess)
            {
                _logger.Log("Unknown", operationName, category, subCategory, riskLevel, argumentsSummary, affectedObjects, confirmed, false, guardResult.Message);
                return guardResult;
            }

            OperationResult<ICSISapModelConnectionService> serviceResult = GetActiveService();
            if (!serviceResult.IsSuccess)
            {
                _logger.Log("None", operationName, category, subCategory, riskLevel, argumentsSummary, affectedObjects, confirmed, false, serviceResult.Message);
                return OperationResult.Failure(serviceResult.Message);
            }

            try
            {
                OperationResult result = action(serviceResult.Data);
                _logger.Log(serviceResult.Data.ProductName, operationName, category, subCategory, riskLevel, argumentsSummary, affectedObjects, confirmed, result.IsSuccess, result.Message);
                return result;
            }
            catch (Exception ex)
            {
                _logger.Log(serviceResult.Data.ProductName, operationName, category, subCategory, riskLevel, argumentsSummary, affectedObjects, confirmed, false, ex.Message);
                return OperationResult.Failure("CSI write operation failed: " + ex.Message);
            }
        }

        private OperationResult<ICSISapModelConnectionService> GetActiveService()
        {
            OperationResult<CSISapModelConnectionInfoDTO> etabs = _etabsService.GetCurrentConnection();
            if (etabs.IsSuccess && etabs.Data != null && etabs.Data.IsConnected)
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_etabsService);
            }

            OperationResult<CSISapModelConnectionInfoDTO> sap2000 = _sap2000Service.GetCurrentConnection();
            if (sap2000.IsSuccess && sap2000.Data != null && sap2000.Data.IsConnected)
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_sap2000Service);
            }

            etabs = _etabsService.TryAttachToRunningInstance();
            if (etabs.IsSuccess && etabs.Data != null && etabs.Data.IsConnected)
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_etabsService);
            }

            sap2000 = _sap2000Service.TryAttachToRunningInstance();
            if (sap2000.IsSuccess && sap2000.Data != null && sap2000.Data.IsConnected)
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_sap2000Service);
            }

            return OperationResult<ICSISapModelConnectionService>.Failure("No ETABS or SAP2000 model is attached.");
        }

        private static CsiWritePreview Preview(
            string operationName,
            CsiMethodRiskLevel riskLevel,
            bool requiresConfirmation,
            bool supportsDryRun,
            string summary,
            IReadOnlyList<string> affectedObjects)
        {
            return new CsiWritePreview
            {
                OperationName = operationName,
                RiskLevel = riskLevel,
                RequiresConfirmation = requiresConfirmation,
                SupportsDryRun = supportsDryRun,
                Summary = summary,
                AffectedObjects = affectedObjects ?? new string[0]
            };
        }

        private static string CleanName(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private static IReadOnlyList<string> One(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? new string[0] : new[] { value };
        }

        private static int Count(IReadOnlyList<string> values)
        {
            return values == null ? 0 : values.Count;
        }

        private static string JoinMessages(IReadOnlyList<string> values)
        {
            if (values == null || values.Count == 0)
            {
                return string.Empty;
            }

            return string.Join(" ", values);
        }
    }
}
