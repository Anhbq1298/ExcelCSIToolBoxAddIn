using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelCSIToolBox.Application.Modelling.PileCaps;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;
using ExcelCSIToolBox.Core.Contracts.CSI.PileCap;

namespace ExcelCSIToolBox.Infrastructure.CSI.Etabs.Modelling
{
    internal static class EtabsPileCapCreationService
    {
        private const double CoordinateToleranceFallback = 1.0;
        private const double SectionTolerance = 0.001;

        public static OperationResult<PileCapAssignmentSummaryDto> QuickCreatePileCaps(
            ETABSv1.cSapModel sapModel,
            PileCapAssignmentRequestDto request)
        {
            if (sapModel == null)
            {
                return OperationResult<PileCapAssignmentSummaryDto>.Failure("Active ETABS model is not available.");
            }

            if (request == null)
            {
                return OperationResult<PileCapAssignmentSummaryDto>.Failure("Pile-cap request is required.");
            }

            var summary = new PileCapAssignmentSummaryDto();
            ETABSv1.eForce originalForce = ETABSv1.eForce.kN;
            ETABSv1.eLength originalLength = ETABSv1.eLength.m;
            ETABSv1.eTemperature originalTemperature = ETABSv1.eTemperature.C;
            bool capturedUnits = false;

            try
            {
                int unitReadRet = sapModel.GetPresentUnits_2(ref originalForce, ref originalLength, ref originalTemperature);
                capturedUnits = unitReadRet == 0;

                if (sapModel.GetModelIsLocked())
                {
                    return OperationResult<PileCapAssignmentSummaryDto>.Failure("The ETABS model is locked. Unlock the model before creating pile caps.");
                }

                int setUnitRet = sapModel.SetPresentUnits(ETABSv1.eUnits.N_mm_C);
                if (setUnitRet != 0)
                {
                    return OperationResult<PileCapAssignmentSummaryDto>.Failure("Failed to set ETABS units to N-mm-C for pile-cap creation (return code " + setUnitRet.ToString(CultureInfo.InvariantCulture) + ").");
                }

                PileCapInputParameters input = CreateInputParameters(request);
                IReadOnlyList<string> validationMessages = new PileCapInputValidator().Validate(input);
                if (validationMessages.Count > 0)
                {
                    return OperationResult<PileCapAssignmentSummaryDto>.Failure(string.Join(Environment.NewLine, validationMessages));
                }

                HashSet<string> concreteMaterials = GetConcreteMaterialNameSet(sapModel);
                if (!concreteMaterials.Contains(input.PileMaterial))
                {
                    return OperationResult<PileCapAssignmentSummaryDto>.Failure("Pile material '" + input.PileMaterial + "' does not exist as a concrete material in ETABS.");
                }

                if (!concreteMaterials.Contains(input.PileCapMaterial))
                {
                    return OperationResult<PileCapAssignmentSummaryDto>.Failure("Pile-cap material '" + input.PileCapMaterial + "' does not exist as a concrete material in ETABS.");
                }

                summary.PilePropertyName = PileCapPropertyNameBuilder.BuildPileFrameSectionName(input.PileDiameterMillimeters, input.PileMaterial);
                summary.PileCapPropertyName = PileCapPropertyNameBuilder.BuildPileCapAreaSectionName(input.PileCapThicknessMillimeters, input.PileCapMaterial);

                OperationResult pilePropertyResult = EnsurePileFrameProperty(
                    sapModel,
                    summary.PilePropertyName,
                    input.PileMaterial,
                    input.PileDiameterMillimeters);
                if (!pilePropertyResult.IsSuccess)
                {
                    return OperationResult<PileCapAssignmentSummaryDto>.Failure(pilePropertyResult.Message);
                }

                OperationResult pileCapPropertyResult = EnsurePileCapAreaProperty(
                    sapModel,
                    summary.PileCapPropertyName,
                    input.PileCapMaterial,
                    input.PileCapThicknessMillimeters);
                if (!pileCapPropertyResult.IsSuccess)
                {
                    return OperationResult<PileCapAssignmentSummaryDto>.Failure(pileCapPropertyResult.Message);
                }

                int ignoredNonPointCount;
                IReadOnlyList<string> selectedPointNames = GetSelectedPointNames(sapModel, out ignoredNonPointCount);
                summary.IgnoredNonPointObjectCount = ignoredNonPointCount;
                summary.SelectedPointCount = selectedPointNames.Count;
                foreach (string selectedPointName in selectedPointNames)
                {
                    summary.SelectedPointNames.Add(selectedPointName);
                }

                if (selectedPointNames.Count == 0)
                {
                    return OperationResult<PileCapAssignmentSummaryDto>.Failure("Please select at least one ETABS point object.");
                }

                double mergeTolerance = GetModelMergeTolerance(sapModel);
                PointIndex pointIndex = LoadPointIndex(sapModel);
                FrameIndex frameIndex = LoadFrameIndex(sapModel, pointIndex);
                AreaIndex areaIndex = LoadAreaIndex(sapModel, pointIndex);
                var geometryCalculator = new PileCapGeometryCalculator();
                PileCapGeometry geometry = geometryCalculator.Calculate(input);

                foreach (string pointName in selectedPointNames)
                {
                    ProcessSelectedPoint(
                        sapModel,
                        pointName,
                        input,
                        geometry,
                        mergeTolerance,
                        pointIndex,
                        frameIndex,
                        areaIndex,
                        summary);
                }

                if (request.SelectCreatedObjects)
                {
                    SelectCreatedObjects(sapModel, summary);
                }

                int refreshRet = sapModel.View.RefreshView(0, false);
                if (refreshRet != 0)
                {
                    summary.Warnings.Add("ETABS View.RefreshView failed after creation (return code " + refreshRet.ToString(CultureInfo.InvariantCulture) + ").");
                }

                return OperationResult<PileCapAssignmentSummaryDto>.Success(summary, FormatSummaryMessage(summary));
            }
            catch (Exception ex)
            {
                return OperationResult<PileCapAssignmentSummaryDto>.Failure("Quick Create Pile Cap and Pile failed: " + ex.Message);
            }
            finally
            {
                if (capturedUnits)
                {
                    sapModel.SetPresentUnits_2(originalForce, originalLength, originalTemperature);
                }
            }
        }

        public static OperationResult<IReadOnlyList<string>> GetConcreteMaterialNames(ETABSv1.cSapModel sapModel)
        {
            if (sapModel == null)
            {
                return OperationResult<IReadOnlyList<string>>.Failure("Active ETABS model is not available.");
            }

            int numberNames = 0;
            string[] names = null;
            int ret = sapModel.PropMaterial.GetNameList(ref numberNames, ref names, ETABSv1.eMatType.Concrete);
            if (ret != 0 || names == null)
            {
                return OperationResult<IReadOnlyList<string>>.Failure("Failed to get concrete material names from ETABS (PropMaterial.GetNameList return code " + ret.ToString(CultureInfo.InvariantCulture) + ").");
            }

            return OperationResult<IReadOnlyList<string>>.Success(new List<string>(names));
        }

        private static PileCapInputParameters CreateInputParameters(PileCapAssignmentRequestDto request)
        {
            var input = new PileCapInputParameters
            {
                ArrangementType = request.ArrangementType,
                PileDiameterMillimeters = request.PileDiameterMillimeters,
                PileLengthMillimeters = request.PileLengthMillimeters,
                PileMaterial = request.PileMaterial,
                RotationDegrees = request.RotationDegrees,
                AutoSpacing = request.AutoSpacing,
                PileSpacingMillimeters = request.PileSpacingMillimeters,
                SpacingXMillimeters = request.SpacingXMillimeters,
                SpacingYMillimeters = request.SpacingYMillimeters,
                PileCapThicknessMillimeters = request.PileCapThicknessMillimeters,
                EdgeDistanceMillimeters = request.EdgeDistanceMillimeters,
                PileCapMaterial = request.PileCapMaterial
            };

            if (input.AutoSpacing)
            {
                double defaultSpacing = input.PileDiameterMillimeters * 3.0;
                input.PileSpacingMillimeters = defaultSpacing;
                input.SpacingXMillimeters = defaultSpacing;
                input.SpacingYMillimeters = defaultSpacing;
            }

            return input;
        }

        private static HashSet<string> GetConcreteMaterialNameSet(ETABSv1.cSapModel sapModel)
        {
            OperationResult<IReadOnlyList<string>> materialResult = GetConcreteMaterialNames(sapModel);
            var names = new HashSet<string>(StringComparer.Ordinal);
            if (materialResult.IsSuccess && materialResult.Data != null)
            {
                foreach (string name in materialResult.Data)
                {
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        names.Add(name);
                    }
                }
            }

            return names;
        }

        private static OperationResult EnsurePileFrameProperty(
            ETABSv1.cSapModel sapModel,
            string propertyName,
            string materialName,
            double diameterMillimeters)
        {
            ETABSv1.eFramePropType propType = ETABSv1.eFramePropType.I;
            int typeRet = sapModel.PropFrame.GetTypeOAPI(propertyName, ref propType);
            if (typeRet != 0)
            {
                int createRet = sapModel.PropFrame.SetCircle(propertyName, materialName, diameterMillimeters, -1, string.Empty, string.Empty);
                return createRet == 0
                    ? OperationResult.Success()
                    : OperationResult.Failure("Failed to create pile frame section '" + propertyName + "' with PropFrame.SetCircle (return code " + createRet.ToString(CultureInfo.InvariantCulture) + ").");
            }

            if (propType != ETABSv1.eFramePropType.Circle)
            {
                return OperationResult.Failure("Property conflict: existing frame section '" + propertyName + "' is " + propType + ", not circular.");
            }

            string fileName = string.Empty;
            string existingMaterial = string.Empty;
            double existingDiameter = 0;
            int color = 0;
            string notes = string.Empty;
            string guid = string.Empty;
            int circleRet = sapModel.PropFrame.GetCircle(propertyName, ref fileName, ref existingMaterial, ref existingDiameter, ref color, ref notes, ref guid);
            if (circleRet != 0)
            {
                return OperationResult.Failure("Failed to read existing pile frame section '" + propertyName + "' with PropFrame.GetCircle (return code " + circleRet.ToString(CultureInfo.InvariantCulture) + ").");
            }

            if (!string.Equals(existingMaterial, materialName, StringComparison.Ordinal) ||
                Math.Abs(existingDiameter - diameterMillimeters) > SectionTolerance)
            {
                return OperationResult.Failure("Property conflict: existing frame section '" + propertyName + "' does not match the requested pile diameter/material.");
            }

            return OperationResult.Success();
        }

        private static OperationResult EnsurePileCapAreaProperty(
            ETABSv1.cSapModel sapModel,
            string propertyName,
            string materialName,
            double thicknessMillimeters)
        {
            int numberNames = 0;
            string[] names = null;
            int listRet = sapModel.PropArea.GetNameList(ref numberNames, ref names);
            bool exists = false;
            if (listRet == 0 && names != null)
            {
                foreach (string name in names)
                {
                    if (string.Equals(name, propertyName, StringComparison.Ordinal))
                    {
                        exists = true;
                        break;
                    }
                }
            }

            if (!exists)
            {
                int createRet = sapModel.PropArea.SetSlab(
                    propertyName,
                    ETABSv1.eSlabType.Footing,
                    ETABSv1.eShellType.ShellThick,
                    materialName,
                    thicknessMillimeters,
                    -1,
                    string.Empty,
                    string.Empty);
                return createRet == 0
                    ? OperationResult.Success()
                    : OperationResult.Failure("Failed to create pile-cap area section '" + propertyName + "' with PropArea.SetSlab (return code " + createRet.ToString(CultureInfo.InvariantCulture) + ").");
            }

            ETABSv1.eSlabType slabType = ETABSv1.eSlabType.Slab;
            ETABSv1.eShellType shellType = ETABSv1.eShellType.ShellThin;
            string existingMaterial = string.Empty;
            double existingThickness = 0;
            int color = 0;
            string notes = string.Empty;
            string guid = string.Empty;
            int slabRet = sapModel.PropArea.GetSlab(propertyName, ref slabType, ref shellType, ref existingMaterial, ref existingThickness, ref color, ref notes, ref guid);
            if (slabRet != 0)
            {
                return OperationResult.Failure("Property conflict: existing area section '" + propertyName + "' is not a slab/footing property that can be verified.");
            }

            if (shellType != ETABSv1.eShellType.ShellThick ||
                !string.Equals(existingMaterial, materialName, StringComparison.Ordinal) ||
                Math.Abs(existingThickness - thicknessMillimeters) > SectionTolerance)
            {
                return OperationResult.Failure("Property conflict: existing area section '" + propertyName + "' does not match the requested Shell-Thick pile-cap thickness/material.");
            }

            return OperationResult.Success();
        }

        private static IReadOnlyList<string> GetSelectedPointNames(ETABSv1.cSapModel sapModel, out int ignoredNonPointCount)
        {
            ignoredNonPointCount = 0;
            int numberItems = 0;
            int[] objectTypes = null;
            string[] objectNames = null;
            int ret = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
            if (ret != 0 || objectTypes == null || objectNames == null)
            {
                return new List<string>();
            }

            var pointNames = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < numberItems && i < objectTypes.Length && i < objectNames.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(objectNames[i]))
                {
                    continue;
                }

                if (objectTypes[i] == CSISapModelObjectTypeIds.Point)
                {
                    if (seen.Add(objectNames[i]))
                    {
                        pointNames.Add(objectNames[i]);
                    }
                }
                else
                {
                    ignoredNonPointCount++;
                }
            }

            return pointNames;
        }

        private static void ProcessSelectedPoint(
            ETABSv1.cSapModel sapModel,
            string selectedPointName,
            PileCapInputParameters input,
            PileCapGeometry geometry,
            double tolerance,
            PointIndex pointIndex,
            FrameIndex frameIndex,
            AreaIndex areaIndex,
            PileCapAssignmentSummaryDto summary)
        {
            var created = new CreatedObjectTracker();
            try
            {
                double originX = 0;
                double originY = 0;
                double originZ = 0;
                int coordRet = sapModel.PointObj.GetCoordCartesian(selectedPointName, ref originX, ref originY, ref originZ, "Global");
                if (coordRet != 0)
                {
                    throw new InvalidOperationException("PointObj.GetCoordCartesian failed for selected point '" + selectedPointName + "' (return code " + coordRet.ToString(CultureInfo.InvariantCulture) + ").");
                }

                pointIndex.AddOrUpdate(selectedPointName, new Point3D(originX, originY, originZ));

                int createdPileCount = CreatePilesForPoint(
                    sapModel,
                    selectedPointName,
                    input,
                    geometry,
                    originX,
                    originY,
                    originZ,
                    tolerance,
                    pointIndex,
                    frameIndex,
                    summary.PilePropertyName,
                    summary,
                    created);

                int createdAreaCount = CreatePileCapMeshForPoint(
                    sapModel,
                    selectedPointName,
                    input,
                    geometry,
                    originX,
                    originY,
                    originZ,
                    tolerance,
                    pointIndex,
                    areaIndex,
                    summary.PileCapPropertyName,
                    summary,
                    created);

                summary.CreatedPileCount += createdPileCount;
                if (createdAreaCount > 0)
                {
                    summary.CreatedPileCapCount++;
                }

                summary.SuccessfullyProcessedPointCount++;
                if (createdPileCount == 0 && createdAreaCount == 0)
                {
                    summary.SkippedPointCount++;
                    summary.Warnings.Add("Selected point '" + selectedPointName + "' was skipped because equivalent piles and pile-cap mesh areas already exist.");
                }
            }
            catch (Exception ex)
            {
                RollBackCreatedObjects(sapModel, created, summary, pointIndex, frameIndex, areaIndex);
                summary.FailedPointCount++;
                summary.Errors.Add("Selected point '" + selectedPointName + "', arrangement " + input.ArrangementType + ": " + ex.Message);
            }
        }

        private static int CreatePilesForPoint(
            ETABSv1.cSapModel sapModel,
            string selectedPointName,
            PileCapInputParameters input,
            PileCapGeometry geometry,
            double originX,
            double originY,
            double originZ,
            double tolerance,
            PointIndex pointIndex,
            FrameIndex frameIndex,
            string pilePropertyName,
            PileCapAssignmentSummaryDto summary,
            CreatedObjectTracker created)
        {
            int createdPileCount = 0;
            for (int i = 0; i < geometry.PileCenters.Count; i++)
            {
                PileCapPoint2D local = geometry.PileCenters[i];
                Point3D head = TransformToGlobal(local, originX, originY, originZ, input.RotationDegrees);
                Point3D toe = new Point3D(head.X, head.Y, originZ - input.PileLengthMillimeters);
                string headPointName = GetOrCreatePoint(sapModel, pointIndex, head, tolerance, created);
                string toePointName = GetOrCreatePoint(sapModel, pointIndex, toe, tolerance, created);

                string existingFrameName = frameIndex.FindEquivalent(headPointName, toePointName);
                if (!string.IsNullOrWhiteSpace(existingFrameName))
                {
                    summary.Warnings.Add("Selected point '" + selectedPointName + "': equivalent pile frame '" + existingFrameName + "' already exists and was not duplicated.");
                    continue;
                }

                string frameName = string.Empty;
                string userName = "QCP_Pile_" + selectedPointName + "_" + (i + 1).ToString(CultureInfo.InvariantCulture);
                int frameRet = sapModel.FrameObj.AddByPoint(headPointName, toePointName, ref frameName, pilePropertyName, userName);
                if (frameRet != 0 || string.IsNullOrWhiteSpace(frameName))
                {
                    throw new InvalidOperationException("FrameObj.AddByPoint failed while creating pile " + (i + 1).ToString(CultureInfo.InvariantCulture) + " with property '" + pilePropertyName + "' (return code " + frameRet.ToString(CultureInfo.InvariantCulture) + ").");
                }

                created.FrameNames.Add(frameName);
                summary.CreatedFrameNames.Add(frameName);
                frameIndex.AddOrUpdate(frameName, headPointName, toePointName);
                createdPileCount++;
            }

            return createdPileCount;
        }

        private static int CreatePileCapMeshForPoint(
            ETABSv1.cSapModel sapModel,
            string selectedPointName,
            PileCapInputParameters input,
            PileCapGeometry geometry,
            double originX,
            double originY,
            double originZ,
            double tolerance,
            PointIndex pointIndex,
            AreaIndex areaIndex,
            string pileCapPropertyName,
            PileCapAssignmentSummaryDto summary,
            CreatedObjectTracker created)
        {
            int createdAreaCount = 0;
            for (int areaIndexValue = 0; areaIndexValue < geometry.MeshAreas.Count; areaIndexValue++)
            {
                PileCapMeshArea meshArea = geometry.MeshAreas[areaIndexValue];
                var pointNames = new List<string>();
                var globalPoints = new List<Point3D>();
                foreach (PileCapPoint2D localPoint in meshArea.Points)
                {
                    Point3D globalPoint = TransformToGlobal(localPoint, originX, originY, originZ, input.RotationDegrees);
                    string pointName = GetOrCreatePoint(sapModel, pointIndex, globalPoint, tolerance, created);
                    pointNames.Add(pointName);
                    globalPoints.Add(globalPoint);
                }

                string existingAreaName = areaIndex.FindEquivalent(globalPoints, tolerance);
                if (!string.IsNullOrWhiteSpace(existingAreaName))
                {
                    summary.Warnings.Add("Selected point '" + selectedPointName + "': equivalent pile-cap mesh area '" + existingAreaName + "' already exists and was not duplicated.");
                    continue;
                }

                string[] pointArray = pointNames.ToArray();
                string areaName = string.Empty;
                string userName = "QCP_Cap_" + selectedPointName + "_" + (areaIndexValue + 1).ToString(CultureInfo.InvariantCulture);
                int areaRet = sapModel.AreaObj.AddByPoint(pointArray.Length, ref pointArray, ref areaName, pileCapPropertyName, userName);
                if (areaRet != 0 || string.IsNullOrWhiteSpace(areaName))
                {
                    throw new InvalidOperationException("AreaObj.AddByPoint failed while creating pile-cap mesh area " + (areaIndexValue + 1).ToString(CultureInfo.InvariantCulture) + " with property '" + pileCapPropertyName + "' (return code " + areaRet.ToString(CultureInfo.InvariantCulture) + ").");
                }

                created.AreaNames.Add(areaName);
                summary.CreatedAreaNames.Add(areaName);
                areaIndex.AddOrUpdate(areaName, globalPoints);
                createdAreaCount++;
            }

            return createdAreaCount;
        }

        private static string GetOrCreatePoint(
            ETABSv1.cSapModel sapModel,
            PointIndex pointIndex,
            Point3D point,
            double tolerance,
            CreatedObjectTracker created)
        {
            string existingName = pointIndex.FindByCoordinate(point, tolerance);
            if (!string.IsNullOrWhiteSpace(existingName))
            {
                return existingName;
            }

            string pointName = string.Empty;
            int pointRet = sapModel.PointObj.AddCartesian(point.X, point.Y, point.Z, ref pointName, string.Empty, "Global", true, 0);
            if (pointRet != 0 || string.IsNullOrWhiteSpace(pointName))
            {
                throw new InvalidOperationException("PointObj.AddCartesian failed at (" + Format(point.X) + ", " + Format(point.Y) + ", " + Format(point.Z) + ") (return code " + pointRet.ToString(CultureInfo.InvariantCulture) + ").");
            }

            pointIndex.AddOrUpdate(pointName, point);
            created.PointNames.Add(pointName);
            return pointName;
        }

        private static Point3D TransformToGlobal(
            PileCapPoint2D localPoint,
            double originX,
            double originY,
            double originZ,
            double rotationDegrees)
        {
            double theta = rotationDegrees * Math.PI / 180.0;
            double cos = Math.Cos(theta);
            double sin = Math.Sin(theta);
            double x = originX + localPoint.X * cos - localPoint.Y * sin;
            double y = originY + localPoint.X * sin + localPoint.Y * cos;
            return new Point3D(x, y, originZ);
        }

        private static PointIndex LoadPointIndex(ETABSv1.cSapModel sapModel)
        {
            var index = new PointIndex();
            int numberNames = 0;
            string[] pointNames = null;
            int listRet = sapModel.PointObj.GetNameList(ref numberNames, ref pointNames);
            if (listRet != 0 || pointNames == null)
            {
                return index;
            }

            for (int i = 0; i < numberNames && i < pointNames.Length; i++)
            {
                string pointName = pointNames[i];
                if (string.IsNullOrWhiteSpace(pointName))
                {
                    continue;
                }

                double x = 0;
                double y = 0;
                double z = 0;
                if (sapModel.PointObj.GetCoordCartesian(pointName, ref x, ref y, ref z, "Global") == 0)
                {
                    index.AddOrUpdate(pointName, new Point3D(x, y, z));
                }
            }

            return index;
        }

        private static FrameIndex LoadFrameIndex(ETABSv1.cSapModel sapModel, PointIndex pointIndex)
        {
            var index = new FrameIndex();
            int numberNames = 0;
            string[] frameNames = null;
            int listRet = sapModel.FrameObj.GetNameList(ref numberNames, ref frameNames);
            if (listRet != 0 || frameNames == null)
            {
                return index;
            }

            for (int i = 0; i < numberNames && i < frameNames.Length; i++)
            {
                string frameName = frameNames[i];
                string pointI = string.Empty;
                string pointJ = string.Empty;
                if (!string.IsNullOrWhiteSpace(frameName) &&
                    sapModel.FrameObj.GetPoints(frameName, ref pointI, ref pointJ) == 0 &&
                    pointIndex.Contains(pointI) &&
                    pointIndex.Contains(pointJ))
                {
                    index.AddOrUpdate(frameName, pointI, pointJ);
                }
            }

            return index;
        }

        private static AreaIndex LoadAreaIndex(ETABSv1.cSapModel sapModel, PointIndex pointIndex)
        {
            var index = new AreaIndex();
            int numberNames = 0;
            string[] areaNames = null;
            int listRet = sapModel.AreaObj.GetNameList(ref numberNames, ref areaNames);
            if (listRet != 0 || areaNames == null)
            {
                return index;
            }

            for (int i = 0; i < numberNames && i < areaNames.Length; i++)
            {
                string areaName = areaNames[i];
                int pointCount = 0;
                string[] pointNames = null;
                if (string.IsNullOrWhiteSpace(areaName) ||
                    sapModel.AreaObj.GetPoints(areaName, ref pointCount, ref pointNames) != 0 ||
                    pointNames == null)
                {
                    continue;
                }

                var points = new List<Point3D>();
                for (int pointIndexValue = 0; pointIndexValue < pointCount && pointIndexValue < pointNames.Length; pointIndexValue++)
                {
                    Point3D point;
                    if (pointIndex.TryGet(pointNames[pointIndexValue], out point))
                    {
                        points.Add(point);
                    }
                }

                if (points.Count == pointCount)
                {
                    index.AddOrUpdate(areaName, points);
                }
            }

            return index;
        }

        private static double GetModelMergeTolerance(ETABSv1.cSapModel sapModel)
        {
            double tolerance = 0;
            int ret = sapModel.GetMergeTol(ref tolerance);
            if (ret != 0 || tolerance <= 0 || double.IsNaN(tolerance) || double.IsInfinity(tolerance))
            {
                return CoordinateToleranceFallback;
            }

            return Math.Max(tolerance, 0.001);
        }

        private static void SelectCreatedObjects(ETABSv1.cSapModel sapModel, PileCapAssignmentSummaryDto summary)
        {
            sapModel.SelectObj.ClearSelection();
            foreach (string frameName in summary.CreatedFrameNames)
            {
                if (!string.IsNullOrWhiteSpace(frameName))
                {
                    sapModel.FrameObj.SetSelected(frameName, true, ETABSv1.eItemType.Objects);
                }
            }

            foreach (string areaName in summary.CreatedAreaNames)
            {
                if (!string.IsNullOrWhiteSpace(areaName))
                {
                    sapModel.AreaObj.SetSelected(areaName, true, ETABSv1.eItemType.Objects);
                }
            }
        }

        private static void RollBackCreatedObjects(
            ETABSv1.cSapModel sapModel,
            CreatedObjectTracker created,
            PileCapAssignmentSummaryDto summary,
            PointIndex pointIndex,
            FrameIndex frameIndex,
            AreaIndex areaIndex)
        {
            for (int i = created.AreaNames.Count - 1; i >= 0; i--)
            {
                string areaName = created.AreaNames[i];
                int deleteRet = sapModel.AreaObj.Delete(areaName, ETABSv1.eItemType.Objects);
                if (deleteRet != 0)
                {
                    summary.Warnings.Add("Rollback warning: AreaObj.Delete failed for '" + areaName + "' (return code " + deleteRet.ToString(CultureInfo.InvariantCulture) + ").");
                }

                summary.CreatedAreaNames.Remove(areaName);
                areaIndex.Remove(areaName);
            }

            for (int i = created.FrameNames.Count - 1; i >= 0; i--)
            {
                string frameName = created.FrameNames[i];
                int deleteRet = sapModel.FrameObj.Delete(frameName, ETABSv1.eItemType.Objects);
                if (deleteRet != 0)
                {
                    summary.Warnings.Add("Rollback warning: FrameObj.Delete failed for '" + frameName + "' (return code " + deleteRet.ToString(CultureInfo.InvariantCulture) + ").");
                }

                summary.CreatedFrameNames.Remove(frameName);
                frameIndex.Remove(frameName);
            }

            for (int i = created.PointNames.Count - 1; i >= 0; i--)
            {
                string pointName = created.PointNames[i];
                int deleteRet = sapModel.PointObj.DeleteSpecialPoint(pointName, ETABSv1.eItemType.Objects);
                if (deleteRet != 0)
                {
                    summary.Warnings.Add("Rollback warning: PointObj.DeleteSpecialPoint failed for '" + pointName + "' (return code " + deleteRet.ToString(CultureInfo.InvariantCulture) + ").");
                }

                if (deleteRet == 0)
                {
                    pointIndex.Remove(pointName);
                }
            }
        }

        private static string FormatSummaryMessage(PileCapAssignmentSummaryDto summary)
        {
            return "Quick Create Pile Cap and Pile completed. Selected points: " +
                   summary.SelectedPointCount.ToString(CultureInfo.InvariantCulture) +
                   ", processed: " +
                   summary.SuccessfullyProcessedPointCount.ToString(CultureInfo.InvariantCulture) +
                   ", pile caps: " +
                   summary.CreatedPileCapCount.ToString(CultureInfo.InvariantCulture) +
                   ", piles: " +
                   summary.CreatedPileCount.ToString(CultureInfo.InvariantCulture) +
                   ", failed: " +
                   summary.FailedPointCount.ToString(CultureInfo.InvariantCulture) +
                   ", ignored non-point objects: " +
                   summary.IgnoredNonPointObjectCount.ToString(CultureInfo.InvariantCulture) +
                   ".";
        }

        private static string Format(double value)
        {
            return value.ToString("G10", CultureInfo.InvariantCulture);
        }

        private static double Distance(Point3D left, Point3D right)
        {
            double dx = left.X - right.X;
            double dy = left.Y - right.Y;
            double dz = left.Z - right.Z;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        private sealed class CreatedObjectTracker
        {
            public CreatedObjectTracker()
            {
                PointNames = new List<string>();
                FrameNames = new List<string>();
                AreaNames = new List<string>();
            }

            public List<string> PointNames { get; private set; }

            public List<string> FrameNames { get; private set; }

            public List<string> AreaNames { get; private set; }
        }

        private sealed class PointIndex
        {
            private readonly Dictionary<string, Point3D> _points = new Dictionary<string, Point3D>(StringComparer.Ordinal);

            public void AddOrUpdate(string pointName, Point3D point)
            {
                if (!string.IsNullOrWhiteSpace(pointName))
                {
                    _points[pointName] = point;
                }
            }

            public bool Contains(string pointName)
            {
                return !string.IsNullOrWhiteSpace(pointName) && _points.ContainsKey(pointName);
            }

            public bool TryGet(string pointName, out Point3D point)
            {
                return _points.TryGetValue(pointName, out point);
            }

            public void Remove(string pointName)
            {
                if (!string.IsNullOrWhiteSpace(pointName))
                {
                    _points.Remove(pointName);
                }
            }

            public string FindByCoordinate(Point3D point, double tolerance)
            {
                foreach (KeyValuePair<string, Point3D> pair in _points)
                {
                    if (Distance(pair.Value, point) <= tolerance)
                    {
                        return pair.Key;
                    }
                }

                return string.Empty;
            }
        }

        private sealed class FrameIndex
        {
            private readonly Dictionary<string, FrameEndpoints> _frames = new Dictionary<string, FrameEndpoints>(StringComparer.Ordinal);

            public void AddOrUpdate(string frameName, string pointI, string pointJ)
            {
                if (!string.IsNullOrWhiteSpace(frameName))
                {
                    _frames[frameName] = new FrameEndpoints(pointI, pointJ);
                }
            }

            public void Remove(string frameName)
            {
                if (!string.IsNullOrWhiteSpace(frameName))
                {
                    _frames.Remove(frameName);
                }
            }

            public string FindEquivalent(string pointI, string pointJ)
            {
                foreach (KeyValuePair<string, FrameEndpoints> pair in _frames)
                {
                    if ((string.Equals(pair.Value.PointI, pointI, StringComparison.Ordinal) &&
                         string.Equals(pair.Value.PointJ, pointJ, StringComparison.Ordinal)) ||
                        (string.Equals(pair.Value.PointI, pointJ, StringComparison.Ordinal) &&
                         string.Equals(pair.Value.PointJ, pointI, StringComparison.Ordinal)))
                    {
                        return pair.Key;
                    }
                }

                return string.Empty;
            }
        }

        private sealed class AreaIndex
        {
            private readonly Dictionary<string, IReadOnlyList<Point3D>> _areas = new Dictionary<string, IReadOnlyList<Point3D>>(StringComparer.Ordinal);

            public void AddOrUpdate(string areaName, IReadOnlyList<Point3D> points)
            {
                if (!string.IsNullOrWhiteSpace(areaName) && points != null)
                {
                    _areas[areaName] = new List<Point3D>(points);
                }
            }

            public void Remove(string areaName)
            {
                if (!string.IsNullOrWhiteSpace(areaName))
                {
                    _areas.Remove(areaName);
                }
            }

            public string FindEquivalent(IReadOnlyList<Point3D> candidate, double tolerance)
            {
                if (candidate == null || candidate.Count == 0)
                {
                    return string.Empty;
                }

                foreach (KeyValuePair<string, IReadOnlyList<Point3D>> pair in _areas)
                {
                    if (HasSamePointSet(pair.Value, candidate, tolerance))
                    {
                        return pair.Key;
                    }
                }

                return string.Empty;
            }

            private static bool HasSamePointSet(IReadOnlyList<Point3D> existing, IReadOnlyList<Point3D> candidate, double tolerance)
            {
                if (existing == null || existing.Count != candidate.Count)
                {
                    return false;
                }

                var matched = new bool[existing.Count];
                for (int candidateIndex = 0; candidateIndex < candidate.Count; candidateIndex++)
                {
                    bool found = false;
                    for (int existingIndex = 0; existingIndex < existing.Count; existingIndex++)
                    {
                        if (!matched[existingIndex] && Distance(existing[existingIndex], candidate[candidateIndex]) <= tolerance)
                        {
                            matched[existingIndex] = true;
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        return false;
                    }
                }

                return true;
            }
        }

        private sealed class FrameEndpoints
        {
            public FrameEndpoints(string pointI, string pointJ)
            {
                PointI = pointI;
                PointJ = pointJ;
            }

            public string PointI { get; private set; }

            public string PointJ { get; private set; }
        }

        private struct Point3D
        {
            public Point3D(double x, double y, double z)
            {
                X = x;
                Y = y;
                Z = z;
            }

            public double X { get; private set; }

            public double Y { get; private set; }

            public double Z { get; private set; }
        }
    }
}
