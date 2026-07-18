using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ETABSv1;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Modelling.DropPanels;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Core.Tabular;
using ExcelCSIToolBox.Infrastructure.CSI.Etabs.Modelling;

namespace ExcelCSIToolBox.Infrastructure.CSI.Etabs.Modelling.DropPanels
{
    public sealed class EtabsDropPanelService : IDropPanelEtabsService
    {
        private const string LoadSetAssignmentTableKey = "Area Load Assignments - Uniform Load Sets";
        private const int ConnectivityAreaObjectType = 5;
        private const double NumericTolerance = 1e-7;

        private static readonly string[] MeshTableAliases =
        {
            "Area Assignments - Floor Auto Mesh Options",
            "Area Assignments - Floor Auto Mesh Option",
            "Area Assignments - Auto Mesh Options",
            "Area Auto Mesh Assignments",
            "Area Mesh Option Assignments"
        };

        private readonly IEtabsConnectionService _connectionService;
        private readonly ICsiApiDispatcher _dispatcher;
        private readonly ICsiOperationLogger _operationLogger;

        public EtabsDropPanelService(
            IEtabsConnectionService connectionService,
            ICsiApiDispatcher dispatcher,
            ICsiOperationLogger operationLogger)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException(nameof(connectionService));
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
            _operationLogger = operationLogger ?? throw new ArgumentNullException(nameof(operationLogger));
        }

        public OperationResult<DropPanelModelContext> GetModelContext()
        {
            return _dispatcher.Invoke(GetModelContextCore);
        }

        public OperationResult<IReadOnlyList<string>> GetConcreteMaterialNames()
        {
            return _dispatcher.Invoke(GetConcreteMaterialNamesCore);
        }

        public OperationResult<IReadOnlyDictionary<string, string>> GetFrameLabels(
            IReadOnlyList<string> frameUniqueNames)
        {
            return _dispatcher.Invoke(() => GetFrameLabelsCore(frameUniqueNames));
        }

        public OperationResult<IReadOnlyList<DropPanelColumnInfo>> ReadColumns(
            IReadOnlyList<string> frameNames,
            double verticalRatioTolerance)
        {
            return _dispatcher.Invoke(() => ReadColumnsCore(frameNames, verticalRatioTolerance));
        }

        public OperationResult<DropPanelPreparationSnapshot> PrepareSnapshot(
            IReadOnlyList<DropPanelColumnInfo> columns,
            IReadOnlyList<DropPanelRequest> requests,
            DropPanelOptions options)
        {
            return _dispatcher.Invoke(() => PrepareSnapshotCore(columns, requests, options));
        }

        public OperationResult<DropPanelApplyResult> Apply(DropPanelOperationPlan plan, DropPanelOptions options)
        {
            return _dispatcher.Invoke(() => ApplyCore(plan, options));
        }

        private OperationResult<DropPanelModelContext> GetModelContextCore()
        {
            cSapModel sapModel;
            OperationResult modelResult = TryGetSapModel(out sapModel);
            if (!modelResult.IsSuccess)
            {
                return OperationResult<DropPanelModelContext>.Failure(modelResult.Message);
            }

            string version = string.Empty;
            double versionNumber = 0.0;
            int versionReturn = sapModel.GetVersion(ref version, ref versionNumber);
            if (versionReturn != 0)
            {
                return OperationResult<DropPanelModelContext>.Failure(ReturnCodeMessage("SapModel.GetVersion", versionReturn));
            }

            string modelPath = sapModel.GetModelFilename(true) ?? string.Empty;
            bool isLocked = sapModel.GetModelIsLocked();
            eForce forceUnit = eForce.kN;
            eLength lengthUnit = eLength.m;
            eTemperature temperatureUnit = eTemperature.C;
            int unitsReturn = sapModel.GetPresentUnits_2(ref forceUnit, ref lengthUnit, ref temperatureUnit);
            if (unitsReturn != 0)
            {
                return OperationResult<DropPanelModelContext>.Failure(ReturnCodeMessage("SapModel.GetPresentUnits_2", unitsReturn));
            }

            string units = sapModel.GetPresentUnits().ToString();
            return OperationResult<DropPanelModelContext>.Success(new DropPanelModelContext
            {
                Version = version,
                ModelFileName = Path.GetFileName(modelPath),
                ModelPath = modelPath,
                PresentUnits = units,
                LengthUnit = FormatLengthUnit(lengthUnit),
                IsLocked = isLocked
            });
        }

        private OperationResult<IReadOnlyList<string>> GetConcreteMaterialNamesCore()
        {
            cSapModel sapModel;
            OperationResult modelResult = TryGetSapModel(out sapModel);
            if (!modelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(modelResult.Message);
            }

            OperationResult<IReadOnlyList<string>> materialResult =
                EtabsPileCapCreationService.GetConcreteMaterialNames(sapModel);
            if (!materialResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(materialResult.Message);
            }

            List<string> result = (materialResult.Data ?? new string[0])
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            return result.Count > 0
                ? OperationResult<IReadOnlyList<string>>.Success(result)
                : OperationResult<IReadOnlyList<string>>.Failure("The active ETABS model contains no concrete materials.");
        }

        private OperationResult<IReadOnlyDictionary<string, string>> GetFrameLabelsCore(
            IReadOnlyList<string> frameUniqueNames)
        {
            cSapModel sapModel;
            OperationResult modelResult = TryGetSapModel(out sapModel);
            if (!modelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyDictionary<string, string>>.Failure(modelResult.Message);
            }

            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (frameUniqueNames == null)
            {
                return OperationResult<IReadOnlyDictionary<string, string>>.Success(result);
            }

            foreach (string uniqueName in frameUniqueNames)
            {
                if (string.IsNullOrWhiteSpace(uniqueName) || result.ContainsKey(uniqueName))
                {
                    continue;
                }

                string label = string.Empty;
                string story = string.Empty;
                int ret = sapModel.FrameObj.GetLabelFromName(uniqueName, ref label, ref story);
                result[uniqueName] = ret == 0 && !string.IsNullOrWhiteSpace(label)
                    ? label
                    : string.Empty;
            }

            return OperationResult<IReadOnlyDictionary<string, string>>.Success(result);
        }

        private OperationResult<IReadOnlyList<DropPanelColumnInfo>> ReadColumnsCore(
            IReadOnlyList<string> frameNames,
            double verticalRatioTolerance)
        {
            cSapModel sapModel;
            OperationResult modelResult = TryGetSapModel(out sapModel);
            if (!modelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<DropPanelColumnInfo>>.Failure(modelResult.Message);
            }

            if (frameNames == null || frameNames.Count == 0)
            {
                return OperationResult<IReadOnlyList<DropPanelColumnInfo>>.Failure("Select at least one ETABS column.");
            }

            if (verticalRatioTolerance <= 0.0)
            {
                return OperationResult<IReadOnlyList<DropPanelColumnInfo>>.Failure("Vertical ratio tolerance must be greater than zero.");
            }

            var columns = new List<DropPanelColumnInfo>();
            var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string frameName in frameNames)
            {
                if (string.IsNullOrWhiteSpace(frameName) || !seenNames.Add(frameName))
                {
                    continue;
                }

                DropPanelColumnInfo column;
                OperationResult columnResult = TryReadColumn(sapModel, frameName, verticalRatioTolerance, out column);
                if (!columnResult.IsSuccess)
                {
                    column = column ?? new DropPanelColumnInfo { FrameName = frameName };
                    column.IsValid = false;
                    column.ValidationMessage = columnResult.Message;
                }

                columns.Add(column);
            }

            return columns.Count > 0
                ? OperationResult<IReadOnlyList<DropPanelColumnInfo>>.Success(columns)
                : OperationResult<IReadOnlyList<DropPanelColumnInfo>>.Failure("No unique ETABS column names were provided.");
        }
        private static OperationResult TryReadColumn(
            cSapModel sapModel,
            string frameName,
            double verticalRatioTolerance,
            out DropPanelColumnInfo column)
        {
            column = new DropPanelColumnInfo { FrameName = frameName };
            string pointI = string.Empty;
            string pointJ = string.Empty;
            int pointsReturn = sapModel.FrameObj.GetPoints(frameName, ref pointI, ref pointJ);
            if (pointsReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("FrameObj.GetPoints", pointsReturn, frameName));
            }

            DropPanelPoint3D iCoordinates;
            OperationResult pointIResult = TryReadPoint(sapModel, pointI, out iCoordinates);
            if (!pointIResult.IsSuccess)
            {
                return pointIResult;
            }

            DropPanelPoint3D jCoordinates;
            OperationResult pointJResult = TryReadPoint(sapModel, pointJ, out jCoordinates);
            if (!pointJResult.IsSuccess)
            {
                return pointJResult;
            }

            DropPanelPoint3D top = iCoordinates.Z >= jCoordinates.Z ? iCoordinates : jCoordinates;
            string topName = iCoordinates.Z >= jCoordinates.Z ? pointI : pointJ;
            string bottomName = iCoordinates.Z >= jCoordinates.Z ? pointJ : pointI;
            double dx = jCoordinates.X - iCoordinates.X;
            double dy = jCoordinates.Y - iCoordinates.Y;
            double dz = jCoordinates.Z - iCoordinates.Z;
            double horizontalLength = Math.Sqrt(dx * dx + dy * dy);

            string section = string.Empty;
            string autoSection = string.Empty;
            int sectionReturn = sapModel.FrameObj.GetSection(frameName, ref section, ref autoSection);
            if (sectionReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("FrameObj.GetSection", sectionReturn, frameName));
            }

            string label = string.Empty;
            string story = string.Empty;
            int labelReturn = sapModel.FrameObj.GetLabelFromName(frameName, ref label, ref story);
            if (labelReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("FrameObj.GetLabelFromName", labelReturn, frameName));
            }

            double[] transformation = null;
            int transformationReturn = sapModel.FrameObj.GetTransformationMatrix(frameName, ref transformation, true);
            if (transformationReturn != 0 || transformation == null || transformation.Length < 6)
            {
                return OperationResult.Failure(ReturnCodeMessage("FrameObj.GetTransformationMatrix", transformationReturn, frameName));
            }

            double localAxisRotation = Math.Atan2(transformation[4], transformation[3]) * 180.0 / Math.PI;
            column.BottomPointName = bottomName;
            column.TopPointName = topName;
            column.StoryName = story;
            column.X = top.X;
            column.Y = top.Y;
            column.Z = top.Z;
            column.SectionProperty = section;
            column.LocalAxisRotationDegrees = localAxisRotation;
            column.IsValid = Math.Abs(dz) > horizontalLength * verticalRatioTolerance;
            column.ValidationMessage = column.IsValid
                ? "Column head detected at the higher frame endpoint."
                : "The selected frame is horizontal or insufficiently vertical to be treated as a column.";
            return OperationResult.Success();
        }

        private OperationResult<DropPanelPreparationSnapshot> PrepareSnapshotCore(
            IReadOnlyList<DropPanelColumnInfo> columns,
            IReadOnlyList<DropPanelRequest> requests,
            DropPanelOptions options)
        {
            if (columns == null || columns.Count == 0 || requests == null || requests.Count == 0 || options == null)
            {
                return OperationResult<DropPanelPreparationSnapshot>.Failure("Columns, drop requests, and options are required.");
            }

            cSapModel sapModel;
            OperationResult modelResult = TryGetSapModel(out sapModel);
            if (!modelResult.IsSuccess)
            {
                return OperationResult<DropPanelPreparationSnapshot>.Failure(modelResult.Message);
            }

            Dictionary<string, HashSet<string>> connectedColumnsByArea;
            OperationResult connectivityResult = ReadConnectedAreas(sapModel, columns, out connectedColumnsByArea);
            if (!connectivityResult.IsSuccess)
            {
                return OperationResult<DropPanelPreparationSnapshot>.Failure(connectivityResult.Message);
            }

            Dictionary<string, List<string>> loadSetsByArea = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            if (options.PreserveShellUniformLoadSetAssignments)
            {
                bool loadSetAssignmentTableAvailable;
                OperationResult loadSetResult = ReadLoadSetAssignments(
                    sapModel, out loadSetsByArea, out loadSetAssignmentTableAvailable);
                if (!loadSetResult.IsSuccess)
                {
                    return OperationResult<DropPanelPreparationSnapshot>.Failure(loadSetResult.Message);
                }

                if (loadSetAssignmentTableAvailable)
                {
                    TableSnapshot editableLoadSetTable;
                    OperationResult editableResult = ReadEditingTable(sapModel, LoadSetAssignmentTableKey, out editableLoadSetTable);
                    TryCancelTableEditing(sapModel);
                    if (!editableResult.IsSuccess)
                    {
                        return OperationResult<DropPanelPreparationSnapshot>.Failure(
                            "Shell Uniform Load Set assignments can be read, but cannot be restored through the ETABS database table: " + editableResult.Message);
                    }

                    if (FindLoadSetFieldIndex(editableLoadSetTable.FieldKeys) < 0 ||
                        !FindObjectFieldIndexes(editableLoadSetTable.FieldKeys).CanResolveObject)
                    {
                        return OperationResult<DropPanelPreparationSnapshot>.Failure(
                            "The editable Shell Uniform Load Set assignment table schema is not recognized, so assignments cannot be restored safely.");
                    }
                }
            }

            TableSnapshot meshTable = null;
            if (options.PreserveMeshAssignments)
            {
                string meshTableKey;
                OperationResult meshKeyResult = ResolveMeshTableKey(sapModel, out meshTableKey);
                if (!meshKeyResult.IsSuccess)
                {
                    return OperationResult<DropPanelPreparationSnapshot>.Failure(meshKeyResult.Message);
                }

                OperationResult meshTableResult = ReadEditingTable(sapModel, meshTableKey, out meshTable);
                TryCancelTableEditing(sapModel);
                if (!meshTableResult.IsSuccess)
                {
                    return OperationResult<DropPanelPreparationSnapshot>.Failure(meshTableResult.Message);
                }

                if (!FindObjectFieldIndexes(meshTable.FieldKeys).CanResolveObject)
                {
                    return OperationResult<DropPanelPreparationSnapshot>.Failure(
                        "The editable area mesh assignment table does not expose recognizable area identity fields.");
                }
            }

            int numberNames = 0;
            string[] areaNames = null;
            int namesReturn = sapModel.AreaObj.GetNameList(ref numberNames, ref areaNames);
            if (namesReturn != 0)
            {
                return OperationResult<DropPanelPreparationSnapshot>.Failure(ReturnCodeMessage("AreaObj.GetNameList", namesReturn));
            }
            if (numberNames > 0 && (areaNames == null || areaNames.Length < numberNames))
            {
                return OperationResult<DropPanelPreparationSnapshot>.Failure("AreaObj.GetNameList returned an incomplete area-name array.");
            }

            DropPanelPreparationSnapshot snapshot = new DropPanelPreparationSnapshot
            {
                ModelPath = sapModel.GetModelFilename(true) ?? string.Empty,
                PresentUnits = sapModel.GetPresentUnits().ToString()
            };
            List<DropPanelAreaInfo> readableAreas = new List<DropPanelAreaInfo>();
            foreach (string areaName in (areaNames ?? new string[0]).Take(numberNames))
            {
                DropPanelAreaInfo area;
                OperationResult areaResult = TryReadAreaGeometry(sapModel, areaName, options.ElevationTolerance, out area);
                if (!areaResult.IsSuccess || area == null)
                {
                    if (connectedColumnsByArea.ContainsKey(areaName))
                    {
                        string connectedProperty = string.Empty;
                        int connectedPropertyReturn = sapModel.AreaObj.GetProperty(areaName, ref connectedProperty);
                        if (connectedPropertyReturn != 0)
                        {
                            return OperationResult<DropPanelPreparationSnapshot>.Failure(
                                ReturnCodeMessage("AreaObj.GetProperty", connectedPropertyReturn, areaName));
                        }

                        if (IsSlabProperty(sapModel, connectedProperty))
                        {
                            return OperationResult<DropPanelPreparationSnapshot>.Failure(
                                "A connected slab area object could not be read completely: " + areaResult.Message);
                        }
                    }

                    continue;
                }

                readableAreas.Add(area);
            }

            foreach (DropPanelAreaInfo area in readableAreas.Where(item => !item.IsOpening))
            {
                if (!IsAreaDirectlyConnectedToRequest(area, requests, connectedColumnsByArea, options) ||
                    !IsSlabProperty(sapModel, area.SectionProperty))
                {
                    continue;
                }

                HashSet<string> connectedColumns;
                if (connectedColumnsByArea.TryGetValue(area.AreaName, out connectedColumns))
                {
                    area.ConnectedColumnNames.AddRange(connectedColumns.OrderBy(name => name, StringComparer.OrdinalIgnoreCase));
                }

                DropPanelAreaAssignmentBackup assignment;
                OperationResult assignmentResult = ReadAreaAssignment(
                    sapModel,
                    area,
                    options,
                    loadSetsByArea,
                    meshTable,
                    out assignment);
                if (!assignmentResult.IsSuccess)
                {
                    return OperationResult<DropPanelPreparationSnapshot>.Failure(assignmentResult.Message);
                }

                area.Assignment = assignment;
                snapshot.Areas.Add(area);
            }

            foreach (DropPanelAreaInfo opening in readableAreas.Where(item => item.IsOpening))
            {
                if (snapshot.Areas.Any(source =>
                        Math.Abs(source.Elevation - opening.Elevation) <= options.ElevationTolerance &&
                        BoundingBoxesIntersect(source.Points, opening.Points, options.GeometryTolerance)))
                {
                    snapshot.Openings.Add(opening);
                }
            }

            if (snapshot.Areas.Count == 0)
            {
                return OperationResult<DropPanelPreparationSnapshot>.Failure(
                    "No slab area objects are directly connected to the selected column heads. Only areas returned by ETABS point connectivity are eligible for splitting.");
            }

            return OperationResult<DropPanelPreparationSnapshot>.Success(snapshot, "ETABS slab geometry and assignments were backed up.");
        }

        private static bool BoundingBoxesIntersect(
            IReadOnlyList<DropPanelPoint3D> left,
            IReadOnlyList<DropPanelPoint3D> right,
            double tolerance)
        {
            if (left == null || right == null || left.Count == 0 || right.Count == 0)
            {
                return false;
            }

            double leftMinX = left.Min(point => point.X) - tolerance;
            double leftMaxX = left.Max(point => point.X) + tolerance;
            double leftMinY = left.Min(point => point.Y) - tolerance;
            double leftMaxY = left.Max(point => point.Y) + tolerance;
            double rightMinX = right.Min(point => point.X);
            double rightMaxX = right.Max(point => point.X);
            double rightMinY = right.Min(point => point.Y);
            double rightMaxY = right.Max(point => point.Y);
            return leftMinX <= rightMaxX && leftMaxX >= rightMinX &&
                   leftMinY <= rightMaxY && leftMaxY >= rightMinY;
        }

        private static OperationResult ReadConnectedAreas(
            cSapModel sapModel,
            IReadOnlyList<DropPanelColumnInfo> columns,
            out Dictionary<string, HashSet<string>> connectedColumnsByArea)
        {
            connectedColumnsByArea = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (DropPanelColumnInfo column in columns.Where(item => item != null && item.IsValid))
            {
                int numberItems = 0;
                int[] objectTypes = null;
                string[] objectNames = null;
                int[] pointNumbers = null;
                int returnCode = sapModel.PointObj.GetConnectivity(
                    column.TopPointName,
                    ref numberItems,
                    ref objectTypes,
                    ref objectNames,
                    ref pointNumbers);
                if (returnCode != 0)
                {
                    return OperationResult.Failure(
                        ReturnCodeMessage("PointObj.GetConnectivity", returnCode, column.TopPointName));
                }
                if (numberItems > 0 &&
                    (objectTypes == null || objectNames == null ||
                     objectTypes.Length < numberItems || objectNames.Length < numberItems))
                {
                    return OperationResult.Failure(
                        "PointObj.GetConnectivity returned incomplete object data for column head point '" +
                        column.TopPointName + "'.");
                }

                bool foundConnectedArea = false;
                for (int index = 0; index < numberItems; index++)
                {
                    if (objectTypes[index] != ConnectivityAreaObjectType || string.IsNullOrWhiteSpace(objectNames[index]))
                    {
                        continue;
                    }

                    foundConnectedArea = true;
                    HashSet<string> connectedColumns;
                    if (!connectedColumnsByArea.TryGetValue(objectNames[index], out connectedColumns))
                    {
                        connectedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        connectedColumnsByArea[objectNames[index]] = connectedColumns;
                    }

                    connectedColumns.Add(column.FrameName);
                }

                if (!foundConnectedArea)
                {
                    return OperationResult.Failure(
                        "No area object is directly connected to top point '" + column.TopPointName +
                        "' of column '" + column.FrameName + "'.");
                }
            }

            return OperationResult.Success();
        }

        private static OperationResult TryReadAreaGeometry(
            cSapModel sapModel,
            string areaName,
            double elevationTolerance,
            out DropPanelAreaInfo area)
        {
            area = null;
            int numberPoints = 0;
            string[] pointNames = null;
            int pointsReturn = sapModel.AreaObj.GetPoints(areaName, ref numberPoints, ref pointNames);
            if (pointsReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("AreaObj.GetPoints", pointsReturn, areaName));
            }
            if (numberPoints < 3 || pointNames == null || pointNames.Length < numberPoints)
            {
                return OperationResult.Failure("AreaObj.GetPoints returned incomplete geometry for area '" + areaName + "'.");
            }

            List<DropPanelPoint3D> points = new List<DropPanelPoint3D>();
            foreach (string pointName in pointNames.Take(numberPoints))
            {
                DropPanelPoint3D point;
                OperationResult pointResult = TryReadPoint(sapModel, pointName, out point);
                if (!pointResult.IsSuccess)
                {
                    return pointResult;
                }

                points.Add(point);
            }

            double minimumZ = points.Min(point => point.Z);
            double maximumZ = points.Max(point => point.Z);
            if (maximumZ - minimumZ > Math.Max(elevationTolerance, 1e-8))
            {
                return OperationResult.Failure("Area '" + areaName + "' is not horizontal within the elevation tolerance.");
            }

            string property = string.Empty;
            int propertyReturn = sapModel.AreaObj.GetProperty(areaName, ref property);
            if (propertyReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("AreaObj.GetProperty", propertyReturn, areaName));
            }

            string label = string.Empty;
            string story = string.Empty;
            int labelReturn = sapModel.AreaObj.GetLabelFromName(areaName, ref label, ref story);
            if (labelReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("AreaObj.GetLabelFromName", labelReturn, areaName));
            }

            bool isOpening = false;
            int openingReturn = sapModel.AreaObj.GetOpening(areaName, ref isOpening);
            if (openingReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("AreaObj.GetOpening", openingReturn, areaName));
            }

            area = new DropPanelAreaInfo
            {
                AreaName = areaName,
                StoryName = story,
                SectionProperty = property,
                Elevation = points.Average(point => point.Z),
                IsOpening = isOpening,
                Points = points
            };
            return OperationResult.Success();
        }

        private static bool IsAreaDirectlyConnectedToRequest(
            DropPanelAreaInfo area,
            IReadOnlyList<DropPanelRequest> requests,
            IDictionary<string, HashSet<string>> connectedColumnsByArea,
            DropPanelOptions options)
        {
            HashSet<string> connectedColumns;
            return connectedColumnsByArea.TryGetValue(area.AreaName, out connectedColumns) &&
                   requests.Any(request => connectedColumns.Contains(request.ColumnName) &&
                                           Math.Abs(request.Elevation - area.Elevation) <= options.ElevationTolerance);
        }

        private static bool IsSlabProperty(cSapModel sapModel, string propertyName)
        {
            eSlabType slabType = eSlabType.Slab;
            eShellType shellType = eShellType.ShellThin;
            string material = string.Empty;
            double thickness = 0.0;
            int color = 0;
            string notes = string.Empty;
            string guid = string.Empty;
            if (sapModel.PropArea.GetSlab(propertyName, ref slabType, ref shellType, ref material, ref thickness, ref color, ref notes, ref guid) == 0)
            {
                return true;
            }

            eDeckType deckType = eDeckType.Filled;
            return sapModel.PropArea.GetDeck(propertyName, ref deckType, ref shellType, ref material, ref thickness, ref color, ref notes, ref guid) == 0;
        }

        private OperationResult<DropPanelPropertyProvisionResult> EnsureDropAreaProperty(
            cSapModel sapModel,
            DropPanelOptions options)
        {
            if (options == null)
            {
                return OperationResult<DropPanelPropertyProvisionResult>.Failure("Drop panel options are required.");
            }

            eForce forceUnit = eForce.kN;
            eLength lengthUnit = eLength.m;
            eTemperature temperatureUnit = eTemperature.C;
            int unitsReturn = sapModel.GetPresentUnits_2(ref forceUnit, ref lengthUnit, ref temperatureUnit);
            if (unitsReturn != 0)
            {
                return OperationResult<DropPanelPropertyProvisionResult>.Failure(
                    ReturnCodeMessage("SapModel.GetPresentUnits_2", unitsReturn));
            }

            string currentLengthUnit = FormatLengthUnit(lengthUnit);
            if (!string.Equals(currentLengthUnit, options.LengthUnit, StringComparison.Ordinal))
            {
                return OperationResult<DropPanelPropertyProvisionResult>.Failure(
                    "The ETABS length unit changed from '" + options.LengthUnit + "' to '" + currentLengthUnit +
                    "'. Review the drop thickness and run the operation again.");
            }

            OperationResult<IReadOnlyList<string>> materialsResult = GetConcreteMaterialNamesCore();
            if (!materialsResult.IsSuccess)
            {
                return OperationResult<DropPanelPropertyProvisionResult>.Failure(materialsResult.Message);
            }

            string materialName = materialsResult.Data.FirstOrDefault(name =>
                string.Equals(name, options.DropMaterial, StringComparison.Ordinal));
            if (string.IsNullOrWhiteSpace(materialName))
            {
                return OperationResult<DropPanelPropertyProvisionResult>.Failure(
                    "Concrete material '" + (options.DropMaterial ?? string.Empty) + "' no longer exists in ETABS.");
            }

            var converter = new ExcelCSIToolBox.Application.Modelling.PileCaps.EtabsUnitConverter();
            double thicknessInMm = Math.Round(options.DropThickness * converter.GetMillimetersPerUnit((int)lengthUnit), 4);

            OperationResult<string> nameResult = DropPanelPropertyNameBuilder.Build(thicknessInMm, materialName);
            if (!nameResult.IsSuccess)
            {
                return OperationResult<DropPanelPropertyProvisionResult>.Failure(nameResult.Message);
            }

            string requestedName = nameResult.Data;
            int numberNames = 0;
            string[] propertyNames = null;
            int listReturn = sapModel.PropArea.GetNameList(ref numberNames, ref propertyNames);
            if (listReturn != 0)
            {
                return OperationResult<DropPanelPropertyProvisionResult>.Failure(
                    ReturnCodeMessage("PropArea.GetNameList", listReturn));
            }

            if (numberNames > 0 && (propertyNames == null || propertyNames.Length < numberNames))
            {
                return OperationResult<DropPanelPropertyProvisionResult>.Failure(
                    "PropArea.GetNameList returned an incomplete area-property array.");
            }

            string existingName = (propertyNames ?? new string[0])
                .Take(numberNames)
                .FirstOrDefault(name => string.Equals(name, requestedName, StringComparison.OrdinalIgnoreCase));
            if (string.IsNullOrWhiteSpace(existingName))
            {
                int createReturn = sapModel.PropArea.SetSlab(
                    requestedName,
                    eSlabType.Drop,
                    eShellType.ShellThick,
                    materialName,
                    options.DropThickness,
                    -1,
                    string.Empty,
                    string.Empty);
                if (createReturn != 0)
                {
                    return OperationResult<DropPanelPropertyProvisionResult>.Failure(
                        ReturnCodeMessage("PropArea.SetSlab", createReturn, requestedName));
                }

                return OperationResult<DropPanelPropertyProvisionResult>.Success(
                    new DropPanelPropertyProvisionResult
                    {
                        PropertyName = requestedName,
                        WasCreated = true
                    });
            }

            eSlabType slabType = eSlabType.Slab;
            eShellType existingShellType = eShellType.ShellThin;
            string existingMaterial = string.Empty;
            double existingThickness = 0.0;
            int color = 0;
            string notes = string.Empty;
            string guid = string.Empty;
            int slabReturn = sapModel.PropArea.GetSlab(
                existingName,
                ref slabType,
                ref existingShellType,
                ref existingMaterial,
                ref existingThickness,
                ref color,
                ref notes,
                ref guid);
            if (slabReturn != 0)
            {
                return OperationResult<DropPanelPropertyProvisionResult>.Failure(
                    "Property conflict: existing area property '" + existingName +
                    "' is not a readable ETABS slab/drop property.");
            }

            double thicknessTolerance = Math.Max(1e-9, Math.Abs(options.DropThickness) * 1e-9);
            if (slabType != eSlabType.Drop ||
                existingShellType != eShellType.ShellThick ||
                !string.Equals(existingMaterial, materialName, StringComparison.Ordinal) ||
                Math.Abs(existingThickness - options.DropThickness) > thicknessTolerance)
            {
                return OperationResult<DropPanelPropertyProvisionResult>.Failure(
                    "Property conflict: existing area property '" + existingName +
                    "' has slab type '" + slabType + "', shell type '" + existingShellType +
                    "', material '" + existingMaterial + "', and thickness " +
                    existingThickness.ToString("0.################", CultureInfo.InvariantCulture) +
                    ". The requested definition is Drop, ShellThick, material '" + materialName +
                    "', and thickness " + options.DropThickness.ToString("0.################", CultureInfo.InvariantCulture) +
                    " " + currentLengthUnit + ". The existing property was not modified.");
            }

            return OperationResult<DropPanelPropertyProvisionResult>.Success(
                new DropPanelPropertyProvisionResult
                {
                    PropertyName = existingName,
                    WasCreated = false
                });
        }

        private static OperationResult ReadAreaAssignment(
            cSapModel sapModel,
            DropPanelAreaInfo area,
            DropPanelOptions options,
            IDictionary<string, List<string>> loadSetsByArea,
            TableSnapshot meshTable,
            out DropPanelAreaAssignmentBackup assignment)
        {
            assignment = new DropPanelAreaAssignmentBackup
            {
                SourceAreaName = area.AreaName,
                StoryName = area.StoryName,
                SectionProperty = area.SectionProperty,
                IsOpening = area.IsOpening,
                OriginalWindingIsCounterClockwise = SignedArea(area.Points) > 0.0
            };

            string label = string.Empty;
            string story = string.Empty;
            int labelReturn = sapModel.AreaObj.GetLabelFromName(area.AreaName, ref label, ref story);
            if (labelReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("AreaObj.GetLabelFromName", labelReturn, area.AreaName));
            }

            assignment.SourceAreaLabel = label;
            assignment.StoryName = story;

            double axisAngle = 0.0;
            bool advanced = false;
            int axesReturn = sapModel.AreaObj.GetLocalAxes(area.AreaName, ref axisAngle, ref advanced);
            if (axesReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("AreaObj.GetLocalAxes", axesReturn, area.AreaName));
            }

            assignment.LocalAxisAngle = axisAngle;
            assignment.UsesAdvancedLocalAxes = advanced;

            double[] transformation = null;
            int transformationReturn = sapModel.AreaObj.GetTransformationMatrix(area.AreaName, ref transformation, true);
            if (transformationReturn != 0 || transformation == null || transformation.Length < 9)
            {
                return OperationResult.Failure(ReturnCodeMessage("AreaObj.GetTransformationMatrix", transformationReturn, area.AreaName));
            }

            assignment.Local3Direction = NormalizeVector(new DropPanelVector3D(
                transformation[6], transformation[7], transformation[8]));

            if (options.PreserveDirectAreaLoads)
            {
                OperationResult loadsResult = ReadDirectAreaLoads(sapModel, area.AreaName, assignment.DirectAreaLoads);
                if (!loadsResult.IsSuccess)
                {
                    return loadsResult;
                }
            }

            if (options.PreserveShellUniformLoadSetAssignments)
            {
                List<string> loadSetNames;
                if (loadSetsByArea.TryGetValue(area.AreaName, out loadSetNames))
                {
                    assignment.ShellUniformLoadSetNames.AddRange(loadSetNames);
                }
            }

            if (options.PreserveDiaphragm)
            {
                string diaphragm = string.Empty;
                int diaphragmReturn = sapModel.AreaObj.GetDiaphragm(area.AreaName, ref diaphragm);
                if (diaphragmReturn != 0)
                {
                    return OperationResult.Failure(ReturnCodeMessage("AreaObj.GetDiaphragm", diaphragmReturn, area.AreaName));
                }

                assignment.Diaphragm = diaphragm;
            }

            if (options.PreserveAreaModifiers)
            {
                double[] modifiers = null;
                int modifiersReturn = sapModel.AreaObj.GetModifiers(area.AreaName, ref modifiers);
                if (modifiersReturn != 0)
                {
                    return OperationResult.Failure(ReturnCodeMessage("AreaObj.GetModifiers", modifiersReturn, area.AreaName));
                }

                assignment.Modifiers = modifiers ?? new double[0];
            }

            if (options.PreserveGroupAssignments)
            {
                int numberGroups = 0;
                string[] groups = null;
                int groupsReturn = sapModel.AreaObj.GetGroupAssign(area.AreaName, ref numberGroups, ref groups);
                if (groupsReturn != 0)
                {
                    return OperationResult.Failure(ReturnCodeMessage("AreaObj.GetGroupAssign", groupsReturn, area.AreaName));
                }

                assignment.Groups.AddRange((groups ?? new string[0]).Where(group => !IsImplicitAllGroup(group)));
            }

            if (options.PreservePierAndSpandrelLabels)
            {
                string pier = string.Empty;
                int pierReturn = sapModel.AreaObj.GetPier(area.AreaName, ref pier);
                if (pierReturn != 0)
                {
                    return OperationResult.Failure(ReturnCodeMessage("AreaObj.GetPier", pierReturn, area.AreaName));
                }

                string spandrel = string.Empty;
                int spandrelReturn = sapModel.AreaObj.GetSpandrel(area.AreaName, ref spandrel);
                if (spandrelReturn != 0)
                {
                    return OperationResult.Failure(ReturnCodeMessage("AreaObj.GetSpandrel", spandrelReturn, area.AreaName));
                }

                assignment.PierLabel = pier;
                assignment.SpandrelLabel = spandrel;
            }

            if (options.PreserveMeshAssignments)
            {
                assignment.MeshAssignment = CreateMeshAssignment(meshTable, area.AreaName, label, story);
            }

            return OperationResult.Success();
        }

        private static OperationResult ReadDirectAreaLoads(
            cSapModel sapModel,
            string areaName,
            ICollection<DropPanelDirectAreaLoad> destination)
        {
            int numberItems = 0;
            string[] areaNames = null;
            string[] loadPatterns = null;
            string[] coordinateSystems = null;
            int[] directions = null;
            double[] values = null;
            int returnCode = sapModel.AreaObj.GetLoadUniform(
                areaName,
                ref numberItems,
                ref areaNames,
                ref loadPatterns,
                ref coordinateSystems,
                ref directions,
                ref values,
                eItemType.Objects);
            if (returnCode != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("AreaObj.GetLoadUniform", returnCode, areaName));
            }

            for (int index = 0; index < numberItems; index++)
            {
                if (loadPatterns == null || coordinateSystems == null || directions == null || values == null ||
                    index >= loadPatterns.Length || index >= coordinateSystems.Length || index >= directions.Length || index >= values.Length)
                {
                    return OperationResult.Failure("AreaObj.GetLoadUniform returned incomplete arrays for area '" + areaName + "'.");
                }

                destination.Add(new DropPanelDirectAreaLoad
                {
                    LoadPattern = loadPatterns[index],
                    LoadType = "Uniform",
                    CoordinateSystem = coordinateSystems[index],
                    Direction = directions[index],
                    Value = values[index],
                    ReplaceExistingAssignments = !destination.Any(load =>
                        string.Equals(load.LoadPattern, loadPatterns[index], StringComparison.OrdinalIgnoreCase))
                });
            }

            return OperationResult.Success();
        }

        private OperationResult<DropPanelApplyResult> ApplyCore(DropPanelOperationPlan plan, DropPanelOptions options)
        {
            if (plan == null || !plan.IsValid || options == null)
            {
                return OperationResult<DropPanelApplyResult>.Failure("Valid drop panel regions are required before applying changes.");
            }

            cSapModel sapModel;
            OperationResult modelResult = TryGetSapModel(out sapModel);
            if (!modelResult.IsSuccess)
            {
                return OperationResult<DropPanelApplyResult>.Failure(modelResult.Message);
            }

            if (sapModel.GetModelIsLocked())
            {
                return OperationResult<DropPanelApplyResult>.Failure("The ETABS model is locked. Unlock it before applying drop panels.");
            }

            OperationResult revalidation = RevalidateSources(sapModel, plan, options);
            if (!revalidation.IsSuccess)
            {
                return OperationResult<DropPanelApplyResult>.Failure(revalidation.Message);
            }

            OperationResult<DropPanelPropertyProvisionResult> propertyResult =
                EnsureDropAreaProperty(sapModel, options);
            if (!propertyResult.IsSuccess)
            {
                return OperationResult<DropPanelApplyResult>.Failure(propertyResult.Message);
            }

            options.DropProperty = propertyResult.Data.PropertyName;

            List<string> sourceAreaNames = plan.SourceAreas.Select(area => area.AreaName).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            List<CreatedRegion> createdRegions = new List<CreatedRegion>();
            bool deletionStarted = false;
            try
            {
                int regionIndex = 0;
                string areaNamePrefix = "DP_" + DateTime.Now.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture);
                foreach (DropPanelRegion region in plan.Regions)
                {
                    regionIndex++;
                    string createdAreaName = CreateAreaRegion(sapModel, region, regionIndex, options, areaNamePrefix);
                    createdRegions.Add(new CreatedRegion(region, createdAreaName));
                }

                foreach (CreatedRegion createdRegion in createdRegions)
                {
                    RestoreDirectAreaLoadsAndDiaphragm(sapModel, createdRegion, options);
                }

                foreach (CreatedRegion createdRegion in createdRegions)
                {
                    RestoreModifiersGroupsAndLabels(sapModel, createdRegion, options);
                }

                // Create and populate every replacement first. The source shells are deleted only
                // after all API-based assignments have been copied successfully.
                foreach (string sourceAreaName in sourceAreaNames)
                {
                    int deleteReturn = sapModel.AreaObj.Delete(sourceAreaName, eItemType.Objects);
                    if (deleteReturn != 0)
                    {
                        throw new InvalidOperationException(ReturnCodeMessage("AreaObj.Delete", deleteReturn, sourceAreaName));
                    }

                    deletionStarted = true;
                }

                if (options.PreserveShellUniformLoadSetAssignments &&
                    createdRegions.Any(item => item.Region.Assignment.ShellUniformLoadSetNames.Count > 0))
                {
                    OperationResult loadSetRestore = RestoreLoadSetAssignments(sapModel, sourceAreaNames, createdRegions);
                    if (!loadSetRestore.IsSuccess)
                    {
                        throw new InvalidOperationException(loadSetRestore.Message);
                    }
                }

                if (options.PreserveMeshAssignments)
                {
                    OperationResult meshRestore = RestoreMeshAssignments(sapModel, sourceAreaNames, createdRegions);
                    if (!meshRestore.IsSuccess)
                    {
                        throw new InvalidOperationException(meshRestore.Message);
                    }
                }

                DropPanelApplyResult result = BuildResult(
                    sapModel,
                    plan,
                    createdRegions,
                    options,
                    propertyResult.Data.WasCreated);
                int refreshReturn = sapModel.View.RefreshView(0, false);

                _operationLogger.Log(
                    "ETABS",
                    "Drop Panel",
                    "Shells / Areas",
                    "Batch Replacement",
                    CsiMethodRiskLevel.High,
                    "Created " + createdRegions.Count.ToString(CultureInfo.InvariantCulture) + " region(s) from " + sourceAreaNames.Count.ToString(CultureInfo.InvariantCulture) + " source area(s).",
                    sourceAreaNames,
                    true,
                    true,
                    refreshReturn == 0
                        ? "Drop panel shells were split and assignments were copied."
                        : "Drop panel shells were split and assignments were copied; the ETABS view refresh returned code " +
                          refreshReturn.ToString(CultureInfo.InvariantCulture) + ".");

                return OperationResult<DropPanelApplyResult>.Success(
                    result,
                    "Drop panels were created. Inside regions use the drop property; outside regions keep their source property.");
            }
            catch (Exception ex)
            {
                string cleanupMessage = string.Empty;
                if (!deletionStarted && createdRegions.Count > 0)
                {
                    bool cleanupFailed = false;
                    foreach (CreatedRegion createdRegion in createdRegions)
                    {
                        cleanupFailed |= sapModel.AreaObj.Delete(createdRegion.AreaName, eItemType.Objects) != 0;
                    }

                    cleanupMessage = cleanupFailed
                        ? " Some temporary replacement shells could not be removed."
                        : " Temporary replacement shells were removed; the source shells were not changed.";
                }

                _operationLogger.Log(
                    "ETABS",
                    "Drop Panel",
                    "Shells / Areas",
                    "Batch Replacement",
                    CsiMethodRiskLevel.High,
                    "Apply failed after validation.",
                    sourceAreaNames,
                    true,
                    false,
                    ex.Message + cleanupMessage);
                return OperationResult<DropPanelApplyResult>.Failure("Drop panel apply failed: " + ex.Message + cleanupMessage);
            }
        }

        private static OperationResult RevalidateSources(
            cSapModel sapModel,
            DropPanelOperationPlan plan,
            DropPanelOptions options)
        {
            string currentModelPath = sapModel.GetModelFilename(true) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(plan.ModelPath) &&
                !string.Equals(Path.GetFullPath(plan.ModelPath), Path.GetFullPath(currentModelPath), StringComparison.OrdinalIgnoreCase))
            {
                return OperationResult.Failure("The active ETABS model changed while preparing the drop panels. Run the operation again.");
            }

            string currentUnits = sapModel.GetPresentUnits().ToString();
            if (!string.IsNullOrWhiteSpace(plan.PresentUnits) &&
                !string.Equals(plan.PresentUnits, currentUnits, StringComparison.Ordinal))
            {
                return OperationResult.Failure(
                    "ETABS present units changed from '" + plan.PresentUnits + "' to '" + currentUnits + "' while preparing the drop panels. Run the operation again.");
            }

            foreach (DropPanelAreaInfo sourceArea in plan.SourceAreas)
            {
                string property = string.Empty;
                int propertyReturn = sapModel.AreaObj.GetProperty(sourceArea.AreaName, ref property);
                if (propertyReturn != 0)
                {
                    return OperationResult.Failure("Source area '" + sourceArea.AreaName + "' no longer exists or cannot be read.");
                }

                if (!string.Equals(property, sourceArea.SectionProperty, StringComparison.OrdinalIgnoreCase))
                {
                    return OperationResult.Failure("Source area '" + sourceArea.AreaName + "' changed while preparing the drop panels. Run the operation again.");
                }

                DropPanelAreaInfo currentArea;
                OperationResult areaResult = TryReadAreaGeometry(
                    sapModel, sourceArea.AreaName, options.ElevationTolerance, out currentArea);
                if (!areaResult.IsSuccess || currentArea == null ||
                    !PolygonPointsEqual(sourceArea.Points, currentArea.Points, options.GeometryTolerance))
                {
                    return OperationResult.Failure(
                        "Source area '" + sourceArea.AreaName + "' geometry changed while preparing the drop panels. Run the operation again.");
                }
            }

            return OperationResult.Success();
        }

        private static bool PolygonPointsEqual(
            IReadOnlyList<DropPanelPoint3D> expected,
            IReadOnlyList<DropPanelPoint3D> actual,
            double tolerance)
        {
            if (expected == null || actual == null || expected.Count != actual.Count || expected.Count == 0)
            {
                return false;
            }

            for (int start = 0; start < actual.Count; start++)
            {
                if (!PointsEqual(expected[0], actual[start], tolerance))
                {
                    continue;
                }

                bool forwardMatches = true;
                bool reverseMatches = true;
                for (int index = 0; index < expected.Count; index++)
                {
                    forwardMatches &= PointsEqual(expected[index], actual[(start + index) % actual.Count], tolerance);
                    reverseMatches &= PointsEqual(expected[index], actual[(start - index + actual.Count) % actual.Count], tolerance);
                }

                if (forwardMatches || reverseMatches)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool PointsEqual(DropPanelPoint3D left, DropPanelPoint3D right, double tolerance)
        {
            return left != null && right != null &&
                   Math.Abs(left.X - right.X) <= tolerance &&
                   Math.Abs(left.Y - right.Y) <= tolerance &&
                   Math.Abs(left.Z - right.Z) <= tolerance;
        }

        private static string CreateAreaRegion(
            cSapModel sapModel,
            DropPanelRegion region,
            int regionIndex,
            DropPanelOptions options,
            string areaNamePrefix)
        {
            if (region == null || region.Assignment == null || region.Points == null || region.Points.Count < 3)
            {
                throw new InvalidOperationException("A generated region has incomplete source mapping or geometry.");
            }

            List<DropPanelPoint3D> points = new List<DropPanelPoint3D>(region.Points);
            DropPanelVector3D normal = ComputeNormal(points);
            if (Dot(normal, region.Assignment.Local3Direction) < 0.0)
            {
                points.Reverse();
            }

            int pointCount = points.Count;
            double[] x = points.Select(point => point.X).ToArray();
            double[] y = points.Select(point => point.Y).ToArray();
            double[] z = points.Select(point => point.Z).ToArray();
            string areaName = string.Empty;
            string requestedName = areaNamePrefix + "_" + regionIndex.ToString(CultureInfo.InvariantCulture);
            int addReturn = sapModel.AreaObj.AddByCoord(
                pointCount,
                ref x,
                ref y,
                ref z,
                ref areaName,
                region.ResultingSectionProperty,
                requestedName,
                "Global");
            if (addReturn != 0 || string.IsNullOrWhiteSpace(areaName))
            {
                throw new InvalidOperationException(ReturnCodeMessage("AreaObj.AddByCoord", addReturn, requestedName));
            }

            if (options.PreserveLocalAxes)
            {
                int axesReturn = sapModel.AreaObj.SetLocalAxes(areaName, region.Assignment.LocalAxisAngle, eItemType.Objects);
                if (axesReturn != 0)
                {
                    throw new InvalidOperationException(ReturnCodeMessage("AreaObj.SetLocalAxes", axesReturn, areaName));
                }
            }

            double[] transformation = null;
            int transformationReturn = sapModel.AreaObj.GetTransformationMatrix(areaName, ref transformation, true);
            if (transformationReturn != 0 || transformation == null || transformation.Length < 9)
            {
                throw new InvalidOperationException(ReturnCodeMessage("AreaObj.GetTransformationMatrix", transformationReturn, areaName));
            }

            DropPanelVector3D createdLocal3 = NormalizeVector(new DropPanelVector3D(transformation[6], transformation[7], transformation[8]));
            if (options.PreserveLocal3Orientation && Dot(createdLocal3, region.Assignment.Local3Direction) <= 0.0)
            {
                throw new InvalidOperationException("Created area '" + areaName + "' has the opposite local 3 orientation from source area '" + region.SourceAreaName + "'.");
            }

            return areaName;
        }

        private static void RestoreDirectAreaLoadsAndDiaphragm(
            cSapModel sapModel,
            CreatedRegion createdRegion,
            DropPanelOptions options)
        {
            DropPanelAreaAssignmentBackup assignment = createdRegion.Region.Assignment;
            string areaName = createdRegion.AreaName;
            if (options.PreserveDirectAreaLoads)
            {
                foreach (IGrouping<string, DropPanelDirectAreaLoad> group in assignment.DirectAreaLoads
                    .GroupBy(load => load.LoadPattern ?? string.Empty, StringComparer.OrdinalIgnoreCase))
                {
                    foreach (DropPanelDirectAreaLoad load in group)
                    {
                        int loadReturn = sapModel.AreaObj.SetLoadUniform(
                            areaName,
                            load.LoadPattern,
                            load.Value,
                            load.Direction,
                            load.ReplaceExistingAssignments,
                            load.CoordinateSystem,
                            eItemType.Objects);
                        if (loadReturn != 0)
                        {
                            throw new InvalidOperationException(ReturnCodeMessage("AreaObj.SetLoadUniform", loadReturn, areaName));
                        }
                    }
                }
            }

            if (options.PreserveDiaphragm && !string.IsNullOrWhiteSpace(assignment.Diaphragm))
            {
                int diaphragmReturn = sapModel.AreaObj.SetDiaphragm(areaName, assignment.Diaphragm);
                if (diaphragmReturn != 0)
                {
                    throw new InvalidOperationException(ReturnCodeMessage("AreaObj.SetDiaphragm", diaphragmReturn, areaName));
                }
            }
        }

        private static void RestoreModifiersGroupsAndLabels(
            cSapModel sapModel,
            CreatedRegion createdRegion,
            DropPanelOptions options)
        {
            DropPanelAreaAssignmentBackup assignment = createdRegion.Region.Assignment;
            string areaName = createdRegion.AreaName;

            if (options.PreserveAreaModifiers && assignment.Modifiers != null && assignment.Modifiers.Length > 0)
            {
                double[] modifiers = (double[])assignment.Modifiers.Clone();
                int modifiersReturn = sapModel.AreaObj.SetModifiers(areaName, ref modifiers, eItemType.Objects);
                if (modifiersReturn != 0)
                {
                    throw new InvalidOperationException(ReturnCodeMessage("AreaObj.SetModifiers", modifiersReturn, areaName));
                }
            }

            if (options.PreserveGroupAssignments)
            {
                foreach (string groupName in assignment.Groups.Where(group => !IsImplicitAllGroup(group)))
                {
                    int groupReturn = sapModel.AreaObj.SetGroupAssign(areaName, groupName, false, eItemType.Objects);
                    if (groupReturn != 0)
                    {
                        throw new InvalidOperationException(ReturnCodeMessage("AreaObj.SetGroupAssign", groupReturn, areaName + "/" + groupName));
                    }
                }
            }

            if (options.PreservePierAndSpandrelLabels)
            {
                if (!string.IsNullOrWhiteSpace(assignment.PierLabel))
                {
                    int pierReturn = sapModel.AreaObj.SetPier(areaName, assignment.PierLabel, eItemType.Objects);
                    if (pierReturn != 0)
                    {
                        throw new InvalidOperationException(ReturnCodeMessage("AreaObj.SetPier", pierReturn, areaName));
                    }
                }

                if (!string.IsNullOrWhiteSpace(assignment.SpandrelLabel))
                {
                    int spandrelReturn = sapModel.AreaObj.SetSpandrel(areaName, assignment.SpandrelLabel, eItemType.Objects);
                    if (spandrelReturn != 0)
                    {
                        throw new InvalidOperationException(ReturnCodeMessage("AreaObj.SetSpandrel", spandrelReturn, areaName));
                    }
                }
            }
        }

        private static OperationResult RestoreLoadSetAssignments(
            cSapModel sapModel,
            IReadOnlyList<string> deletedSourceAreaNames,
            IReadOnlyList<CreatedRegion> createdRegions)
        {
            TableSnapshot snapshot;
            OperationResult readResult = ReadEditingTable(sapModel, LoadSetAssignmentTableKey, out snapshot);
            if (!readResult.IsSuccess)
            {
                return readResult;
            }

            int loadSetIndex = FindLoadSetFieldIndex(snapshot.FieldKeys);
            ObjectFieldIndexes fields = FindObjectFieldIndexes(snapshot.FieldKeys);
            if (loadSetIndex < 0 || !fields.CanResolveObject)
            {
                TryCancelTableEditing(sapModel);
                return OperationResult.Failure("The ETABS Shell Uniform Load Set assignment table schema is not recognized.");
            }

            List<Dictionary<string, string>> records = FilterDeletedSourceRecords(
                snapshot.Records, snapshot.FieldKeys, fields, deletedSourceAreaNames, createdRegions);
            foreach (CreatedRegion createdRegion in createdRegions)
            {
                foreach (string loadSetName in createdRegion.Region.Assignment.ShellUniformLoadSetNames)
                {
                    Dictionary<string, string> record = EmptyRecord(snapshot.FieldKeys);
                    SetObjectIdentity(sapModel, createdRegion.AreaName, record, snapshot.FieldKeys, fields);
                    record[snapshot.FieldKeys[loadSetIndex]] = loadSetName;
                    records.Add(record);
                }
            }

            return WriteEditingTable(sapModel, snapshot, records);
        }

        private static OperationResult RestoreMeshAssignments(
            cSapModel sapModel,
            IReadOnlyList<string> deletedSourceAreaNames,
            IReadOnlyList<CreatedRegion> createdRegions)
        {
            DropPanelMeshAssignment firstAssignment = createdRegions
                .Select(item => item.Region.Assignment.MeshAssignment)
                .FirstOrDefault(item => item != null && !string.IsNullOrWhiteSpace(item.TableKey));
            if (firstAssignment == null)
            {
                return OperationResult.Failure("Mesh assignment backup is missing.");
            }

            TableSnapshot snapshot;
            OperationResult readResult = ReadEditingTable(sapModel, firstAssignment.TableKey, out snapshot);
            if (!readResult.IsSuccess)
            {
                return readResult;
            }

            ObjectFieldIndexes fields = FindObjectFieldIndexes(snapshot.FieldKeys);
            if (!fields.CanResolveObject)
            {
                TryCancelTableEditing(sapModel);
                return OperationResult.Failure("The ETABS mesh assignment table does not expose recognizable area identity fields.");
            }

            List<Dictionary<string, string>> records = FilterDeletedSourceRecords(
                snapshot.Records, snapshot.FieldKeys, fields, deletedSourceAreaNames, createdRegions);
            foreach (CreatedRegion createdRegion in createdRegions)
            {
                DropPanelMeshAssignment mesh = createdRegion.Region.Assignment.MeshAssignment;
                if (mesh == null)
                {
                    TryCancelTableEditing(sapModel);
                    return OperationResult.Failure("Mesh assignment backup is missing for source area '" + createdRegion.Region.SourceAreaName + "'.");
                }

                foreach (DropPanelTableRecord sourceRecord in mesh.Records)
                {
                    Dictionary<string, string> record = EmptyRecord(snapshot.FieldKeys);
                    foreach (string fieldKey in snapshot.FieldKeys)
                    {
                        string value;
                        if (sourceRecord.Values.TryGetValue(fieldKey, out value))
                        {
                            record[fieldKey] = value;
                        }
                    }

                    SetObjectIdentity(sapModel, createdRegion.AreaName, record, snapshot.FieldKeys, fields);
                    records.Add(record);
                }
            }

            return WriteEditingTable(sapModel, snapshot, records);
        }

        private static DropPanelApplyResult BuildResult(
            cSapModel sapModel,
            DropPanelOperationPlan plan,
            IReadOnlyList<CreatedRegion> createdRegions,
            DropPanelOptions options,
            bool propertyWasCreated)
        {
            DropPanelApplyResult result = new DropPanelApplyResult
            {
                ProcessedColumnCount = plan.Columns.Count(column => column != null && column.IsValid),
                CreatedDropAreaCount = createdRegions.Count(item => item.Region.IsDrop),
                DropPropertyName = options.DropProperty,
                DropPropertyCreated = propertyWasCreated,
                DropThickness = options.DropThickness,
                LengthUnit = options.LengthUnit,
                MaterialName = options.DropMaterial
            };
            result.CreatedAreaNames.AddRange(createdRegions.Select(item => item.AreaName));

            foreach (CreatedRegion createdRegion in createdRegions)
            {
                DropPanelRegion region = createdRegion.Region;
                DropPanelAreaAssignmentBackup expected = region.Assignment;
                string areaName = createdRegion.AreaName;
                DropPanelLogEntry entry = new DropPanelLogEntry
                {
                    Timestamp = DateTimeOffset.Now.ToString("o"),
                    EtabsModel = Path.GetFileName(sapModel.GetModelFilename(true)),
                    Story = expected.StoryName,
                    Column = string.Join(", ", region.ColumnNames),
                    SourceArea = region.SourceAreaName,
                    NewArea = areaName,
                    RegionType = region.IsDrop ? "Drop" : "Normal slab",
                    OriginalProperty = expected.SectionProperty,
                    NewProperty = region.ResultingSectionProperty,
                    DirectLoadStatus = options.PreserveDirectAreaLoads ? "Copied" : "Skipped",
                    ShellLoadSetStatus = options.PreserveShellUniformLoadSetAssignments ? "Copied" : "Skipped",
                    LocalAxisStatus = options.PreserveLocalAxes ? "Copied" : "Skipped",
                    Local3Status = options.PreserveLocal3Orientation ? "Copied" : "Skipped",
                    DiaphragmStatus = options.PreserveDiaphragm ? "Copied" : "Skipped",
                    Message = region.IsDrop
                        ? "Inside boundary; drop property assigned."
                        : "Outside boundary; source property retained."
                };
                result.LogEntries.Add(entry);
            }

            return result;
        }

        private OperationResult TryGetSapModel(out cSapModel sapModel)
        {
            sapModel = _connectionService.SapModel as cSapModel;
            if (sapModel == null)
            {
                return OperationResult.Failure("ETABS is not attached or no model is open. Attach to ETABS and try again.");
            }

            return OperationResult.Success();
        }

        private static OperationResult TryReadPoint(cSapModel sapModel, string pointName, out DropPanelPoint3D point)
        {
            point = null;
            double x = 0.0;
            double y = 0.0;
            double z = 0.0;
            int returnCode = sapModel.PointObj.GetCoordCartesian(pointName, ref x, ref y, ref z, "Global");
            if (returnCode != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("PointObj.GetCoordCartesian", returnCode, pointName));
            }

            point = new DropPanelPoint3D(x, y, z);
            return OperationResult.Success();
        }

        private static OperationResult ReadLoadSetAssignments(
            cSapModel sapModel,
            out Dictionary<string, List<string>> assignments,
            out bool tableAvailable)
        {
            assignments = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            tableAvailable = false;
            OperationResult<bool> availabilityResult = IsTableAvailable(sapModel, LoadSetAssignmentTableKey);
            if (!availabilityResult.IsSuccess)
            {
                return OperationResult.Failure(availabilityResult.Message);
            }

            tableAvailable = availabilityResult.Data;
            if (!tableAvailable)
            {
                return OperationResult.Success(
                    "The Shell Uniform Load Set assignment table is unavailable because the model contains no such assignments.");
            }

            TableSnapshot snapshot;
            OperationResult readResult = ReadDisplayTable(sapModel, LoadSetAssignmentTableKey, out snapshot);
            if (!readResult.IsSuccess)
            {
                return readResult;
            }

            int loadSetIndex = FindLoadSetFieldIndex(snapshot.FieldKeys);
            ObjectFieldIndexes fields = FindObjectFieldIndexes(snapshot.FieldKeys);
            if (loadSetIndex < 0 || !fields.CanResolveObject)
            {
                return OperationResult.Failure("The ETABS table '" + LoadSetAssignmentTableKey + "' has an unsupported schema.");
            }

            foreach (Dictionary<string, string> record in snapshot.Records)
            {
                OperationResult<string> areaNameResult = ResolveObjectName(sapModel, record, snapshot.FieldKeys, fields);
                if (!areaNameResult.IsSuccess)
                {
                    return OperationResult.Failure(areaNameResult.Message);
                }

                string areaName = areaNameResult.Data;
                string loadSetName = ReadField(record, snapshot.FieldKeys, loadSetIndex);
                if (string.IsNullOrWhiteSpace(areaName) || string.IsNullOrWhiteSpace(loadSetName))
                {
                    continue;
                }

                List<string> names;
                if (!assignments.TryGetValue(areaName, out names))
                {
                    names = new List<string>();
                    assignments[areaName] = names;
                }

                if (!names.Contains(loadSetName, StringComparer.OrdinalIgnoreCase))
                {
                    names.Add(loadSetName);
                }
            }

            return OperationResult.Success();
        }

        private static OperationResult<bool> IsTableAvailable(cSapModel sapModel, string targetTableKey)
        {
            int numberTables = 0;
            string[] tableKeys = null;
            string[] tableNames = null;
            int[] importTypes = null;
            int returnCode = sapModel.DatabaseTables.GetAvailableTables(
                ref numberTables,
                ref tableKeys,
                ref tableNames,
                ref importTypes);
            if (returnCode != 0)
            {
                return OperationResult<bool>.Failure(ReturnCodeMessage("DatabaseTables.GetAvailableTables", returnCode));
            }

            bool isAvailable = (tableKeys ?? new string[0])
                .Take(numberTables)
                .Any(key => string.Equals(key, targetTableKey, StringComparison.OrdinalIgnoreCase));
            return OperationResult<bool>.Success(isAvailable);
        }

        private static OperationResult ResolveMeshTableKey(cSapModel sapModel, out string tableKey)
        {
            tableKey = string.Empty;
            int numberTables = 0;
            string[] tableKeys = null;
            string[] tableNames = null;
            int[] importTypes = null;
            int returnCode = sapModel.DatabaseTables.GetAvailableTables(ref numberTables, ref tableKeys, ref tableNames, ref importTypes);
            if (returnCode != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("DatabaseTables.GetAvailableTables", returnCode));
            }

            for (int index = 0; index < numberTables; index++)
            {
                string key = tableKeys != null && index < tableKeys.Length ? tableKeys[index] : string.Empty;
                if (MeshTableAliases.Any(alias => string.Equals(alias, key, StringComparison.OrdinalIgnoreCase)))
                {
                    tableKey = key;
                    return OperationResult.Success();
                }
            }

            for (int index = 0; index < numberTables; index++)
            {
                string key = tableKeys != null && index < tableKeys.Length ? tableKeys[index] : string.Empty;
                string name = tableNames != null && index < tableNames.Length ? tableNames[index] : string.Empty;
                string combined = (key + " " + name).ToUpperInvariant();
                if (combined.Contains("AREA") && combined.Contains("MESH") &&
                    (combined.Contains("ASSIGN") || combined.Contains("OPTION")))
                {
                    tableKey = key;
                    return OperationResult.Success();
                }
            }

            return OperationResult.Failure(
                "The referenced ETABS API does not expose AreaObj.GetAutoMesh, and no importable area mesh assignment database table was found. " +
                "Clear 'Preserve Mesh Assignments' only if losing explicit mesh assignments is acceptable.");
        }

        private static DropPanelMeshAssignment CreateMeshAssignment(
            TableSnapshot table,
            string areaName,
            string label,
            string story)
        {
            DropPanelMeshAssignment assignment = new DropPanelMeshAssignment
            {
                TableKey = table.TableKey,
                TableVersion = table.TableVersion,
                FieldKeys = new List<string>(table.FieldKeys)
            };
            ObjectFieldIndexes fields = FindObjectFieldIndexes(table.FieldKeys);
            foreach (Dictionary<string, string> record in table.Records)
            {
                if (RecordMatchesObject(record, table.FieldKeys, fields, areaName, label, story))
                {
                    DropPanelTableRecord copied = new DropPanelTableRecord();
                    foreach (KeyValuePair<string, string> value in record)
                    {
                        copied.Values[value.Key] = value.Value;
                    }

                    assignment.Records.Add(copied);
                }
            }

            return assignment;
        }

        private static OperationResult ReadDisplayTable(cSapModel sapModel, string tableKey, out TableSnapshot snapshot)
        {
            snapshot = null;
            int tableVersion = 0;
            string[] requestedFields = null;
            string[] includedFields = null;
            int numberRecords = 0;
            string[] data = null;
            int returnCode = sapModel.DatabaseTables.GetTableForDisplayArray(
                tableKey,
                ref requestedFields,
                string.Empty,
                ref tableVersion,
                ref includedFields,
                ref numberRecords,
                ref data);
            if (returnCode != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("DatabaseTables.GetTableForDisplayArray", returnCode, tableKey));
            }

            return CreateTableSnapshot(tableKey, tableVersion, includedFields ?? requestedFields, numberRecords, data, out snapshot);
        }

        private static OperationResult ReadEditingTable(cSapModel sapModel, string tableKey, out TableSnapshot snapshot)
        {
            snapshot = null;
            int tableVersion = 0;
            string[] includedFields = null;
            int numberRecords = 0;
            string[] data = null;
            int returnCode = sapModel.DatabaseTables.GetTableForEditingArray(
                tableKey,
                string.Empty,
                ref tableVersion,
                ref includedFields,
                ref numberRecords,
                ref data);
            if (returnCode != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("DatabaseTables.GetTableForEditingArray", returnCode, tableKey));
            }

            return CreateTableSnapshot(tableKey, tableVersion, includedFields, numberRecords, data, out snapshot);
        }

        private static OperationResult CreateTableSnapshot(
            string tableKey,
            int tableVersion,
            string[] fields,
            int numberRecords,
            string[] data,
            out TableSnapshot snapshot)
        {
            snapshot = null;
            if (fields == null || fields.Length == 0)
            {
                return OperationResult.Failure("ETABS table '" + tableKey + "' returned no field keys.");
            }

            data = data ?? new string[0];
            int expectedLength = numberRecords * fields.Length;
            if (data.Length != expectedLength)
            {
                return OperationResult.Failure("ETABS table '" + tableKey + "' returned inconsistent record data.");
            }

            snapshot = new TableSnapshot
            {
                TableKey = tableKey,
                TableVersion = tableVersion,
                FieldKeys = fields.ToList()
            };
            for (int recordIndex = 0; recordIndex < numberRecords; recordIndex++)
            {
                Dictionary<string, string> record = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int fieldIndex = 0; fieldIndex < fields.Length; fieldIndex++)
                {
                    record[fields[fieldIndex]] = data[recordIndex * fields.Length + fieldIndex] ?? string.Empty;
                }

                snapshot.Records.Add(record);
            }

            return OperationResult.Success();
        }

        private static OperationResult WriteEditingTable(
            cSapModel sapModel,
            TableSnapshot snapshot,
            IReadOnlyList<Dictionary<string, string>> records)
        {
            int tableVersion = snapshot.TableVersion;
            string[] fields = snapshot.FieldKeys.ToArray();
            string[] data = Flatten(records, fields);
            int stageReturn = sapModel.DatabaseTables.SetTableForEditingArray(
                snapshot.TableKey,
                ref tableVersion,
                ref fields,
                records.Count,
                ref data);
            if (stageReturn != 0)
            {
                TryCancelTableEditing(sapModel);
                return OperationResult.Failure(ReturnCodeMessage("DatabaseTables.SetTableForEditingArray", stageReturn, snapshot.TableKey));
            }

            bool fillImportLog = true;
            int fatalErrors = 0;
            int errors = 0;
            int warnings = 0;
            int information = 0;
            string importLog = string.Empty;
            int applyReturn = sapModel.DatabaseTables.ApplyEditedTables(
                fillImportLog,
                ref fatalErrors,
                ref errors,
                ref warnings,
                ref information,
                ref importLog);
            if (applyReturn != 0 || fatalErrors != 0 || errors != 0)
            {
                TryCancelTableEditing(sapModel);
                return OperationResult.Failure(
                    "ETABS failed to apply table '" + snapshot.TableKey + "'. Return code " + applyReturn.ToString(CultureInfo.InvariantCulture) +
                    ", fatal errors " + fatalErrors.ToString(CultureInfo.InvariantCulture) +
                    ", errors " + errors.ToString(CultureInfo.InvariantCulture) +
                    ". Import log: " + importLog);
            }

            return OperationResult.Success();
        }

        private static string[] Flatten(IReadOnlyList<Dictionary<string, string>> records, IReadOnlyList<string> fields)
        {
            List<string> data = new List<string>(records.Count * fields.Count);
            foreach (Dictionary<string, string> record in records)
            {
                foreach (string field in fields)
                {
                    string value;
                    data.Add(record.TryGetValue(field, out value) ? value ?? string.Empty : string.Empty);
                }
            }

            return data.ToArray();
        }

        private static void TryCancelTableEditing(cSapModel sapModel)
        {
            try
            {
                int cancelReturn = sapModel.DatabaseTables.CancelTableEditing();
                if (cancelReturn != 0)
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceWarning(
                    "ETABS database table edit cleanup failed after the edit session had already completed or failed: " + ex.Message);
            }
        }

        private static int FindLoadSetFieldIndex(IReadOnlyList<string> fields)
        {
            return CsiTableFieldAliasResolver.FindFirstIndex(
                fields,
                "LoadSet", "Load Set", "Load Set Name", "LoadSetName",
                "UniformLoadSet", "Uniform Load Set", "UniformLoadSetName", "Uniform Load Set Name");
        }

        private static ObjectFieldIndexes FindObjectFieldIndexes(IReadOnlyList<string> fields)
        {
            return new ObjectFieldIndexes
            {
                UniqueNameIndex = CsiTableFieldAliasResolver.FindFirstIndex(fields, "UniqueName", "Unique Name", "ObjectName", "Object Name"),
                StoryIndex = CsiTableFieldAliasResolver.FindFirstIndex(fields, "Story", "StoryName", "Story Name"),
                LabelIndex = CsiTableFieldAliasResolver.FindFirstIndex(fields, "Label", "LabelName", "Label Name", "Area", "AreaName", "Area Name", "Shell", "ShellName", "Shell Name")
            };
        }

        private static OperationResult<string> ResolveObjectName(
            cSapModel sapModel,
            IDictionary<string, string> record,
            IReadOnlyList<string> fields,
            ObjectFieldIndexes indexes)
        {
            string uniqueName = ReadField(record, fields, indexes.UniqueNameIndex);
            if (!string.IsNullOrWhiteSpace(uniqueName))
            {
                return OperationResult<string>.Success(uniqueName.Trim());
            }

            string label = ReadField(record, fields, indexes.LabelIndex);
            string story = ReadField(record, fields, indexes.StoryIndex);
            if (string.IsNullOrWhiteSpace(label) || string.IsNullOrWhiteSpace(story))
            {
                return OperationResult<string>.Success(string.Empty);
            }

            string areaName = string.Empty;
            int returnCode = sapModel.AreaObj.GetNameFromLabel(label, story, ref areaName);
            return returnCode == 0
                ? OperationResult<string>.Success(areaName)
                : OperationResult<string>.Failure(ReturnCodeMessage("AreaObj.GetNameFromLabel", returnCode, label + "/" + story));
        }

        private static bool RecordMatchesObject(
            IDictionary<string, string> record,
            IReadOnlyList<string> fields,
            ObjectFieldIndexes indexes,
            string areaName,
            string label,
            string story)
        {
            string recordUniqueName = ReadField(record, fields, indexes.UniqueNameIndex);
            if (!string.IsNullOrWhiteSpace(recordUniqueName) && string.Equals(recordUniqueName, areaName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return string.Equals(ReadField(record, fields, indexes.LabelIndex), label, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(ReadField(record, fields, indexes.StoryIndex), story, StringComparison.OrdinalIgnoreCase);
        }

        private static List<Dictionary<string, string>> FilterDeletedSourceRecords(
            IEnumerable<Dictionary<string, string>> records,
            IReadOnlyList<string> fields,
            ObjectFieldIndexes indexes,
            IReadOnlyList<string> deletedNames,
            IReadOnlyList<CreatedRegion> createdRegions)
        {
            HashSet<string> deleted = new HashSet<string>(deletedNames ?? new string[0], StringComparer.OrdinalIgnoreCase);
            HashSet<string> deletedLabelsAndStories = new HashSet<string>(
                (createdRegions ?? new CreatedRegion[0])
                    .Where(item => item != null && item.Region != null && item.Region.Assignment != null &&
                                   !string.IsNullOrWhiteSpace(item.Region.Assignment.SourceAreaLabel) &&
                                   !string.IsNullOrWhiteSpace(item.Region.Assignment.StoryName))
                    .Select(item => IdentityKey(item.Region.Assignment.SourceAreaLabel, item.Region.Assignment.StoryName)),
                StringComparer.OrdinalIgnoreCase);
            List<Dictionary<string, string>> retained = new List<Dictionary<string, string>>();
            foreach (Dictionary<string, string> record in records)
            {
                string uniqueName = ReadField(record, fields, indexes.UniqueNameIndex);
                string labelAndStory = IdentityKey(
                    ReadField(record, fields, indexes.LabelIndex),
                    ReadField(record, fields, indexes.StoryIndex));
                bool isDeleted = (!string.IsNullOrWhiteSpace(uniqueName) && deleted.Contains(uniqueName)) ||
                                 deletedLabelsAndStories.Contains(labelAndStory);
                if (!isDeleted)
                {
                    retained.Add(new Dictionary<string, string>(record, StringComparer.OrdinalIgnoreCase));
                }
            }

            return retained;
        }

        private static string IdentityKey(string label, string story)
        {
            return (label ?? string.Empty).Trim() + "\u001f" + (story ?? string.Empty).Trim();
        }

        private static void SetObjectIdentity(
            cSapModel sapModel,
            string areaName,
            IDictionary<string, string> record,
            IReadOnlyList<string> fields,
            ObjectFieldIndexes indexes)
        {
            string label = string.Empty;
            string story = string.Empty;
            int returnCode = sapModel.AreaObj.GetLabelFromName(areaName, ref label, ref story);
            if (returnCode != 0)
            {
                throw new InvalidOperationException(ReturnCodeMessage("AreaObj.GetLabelFromName", returnCode, areaName));
            }

            if (indexes.UniqueNameIndex >= 0)
            {
                record[fields[indexes.UniqueNameIndex]] = areaName;
            }

            if (indexes.LabelIndex >= 0)
            {
                record[fields[indexes.LabelIndex]] = label;
            }

            if (indexes.StoryIndex >= 0)
            {
                record[fields[indexes.StoryIndex]] = story;
            }
        }

        private static Dictionary<string, string> EmptyRecord(IReadOnlyList<string> fields)
        {
            Dictionary<string, string> record = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string field in fields)
            {
                record[field] = string.Empty;
            }

            return record;
        }

        private static string ReadField(IDictionary<string, string> record, IReadOnlyList<string> fields, int index)
        {
            if (index < 0 || index >= fields.Count)
            {
                return string.Empty;
            }

            string value;
            return record.TryGetValue(fields[index], out value) ? value ?? string.Empty : string.Empty;
        }

        private static DropPanelVector3D ComputeNormal(IReadOnlyList<DropPanelPoint3D> points)
        {
            double x = 0.0;
            double y = 0.0;
            double z = 0.0;
            for (int index = 0; index < points.Count; index++)
            {
                DropPanelPoint3D current = points[index];
                DropPanelPoint3D next = points[(index + 1) % points.Count];
                x += (current.Y - next.Y) * (current.Z + next.Z);
                y += (current.Z - next.Z) * (current.X + next.X);
                z += (current.X - next.X) * (current.Y + next.Y);
            }

            return NormalizeVector(new DropPanelVector3D(x, y, z));
        }

        private static DropPanelVector3D NormalizeVector(DropPanelVector3D vector)
        {
            double length = Math.Sqrt(vector.X * vector.X + vector.Y * vector.Y + vector.Z * vector.Z);
            return length <= 1e-12
                ? new DropPanelVector3D(0.0, 0.0, 0.0)
                : new DropPanelVector3D(vector.X / length, vector.Y / length, vector.Z / length);
        }

        private static double Dot(DropPanelVector3D left, DropPanelVector3D right)
        {
            if (left == null || right == null)
            {
                return 0.0;
            }

            return left.X * right.X + left.Y * right.Y + left.Z * right.Z;
        }

        private static double SignedArea(IReadOnlyList<DropPanelPoint3D> points)
        {
            double area = 0.0;
            for (int index = 0; index < points.Count; index++)
            {
                DropPanelPoint3D current = points[index];
                DropPanelPoint3D next = points[(index + 1) % points.Count];
                area += current.X * next.Y - next.X * current.Y;
            }

            return area / 2.0;
        }

        private static bool IsImplicitAllGroup(string groupName)
        {
            return string.Equals((groupName ?? string.Empty).Trim(), "ALL", StringComparison.OrdinalIgnoreCase);
        }

        private static string ReturnCodeMessage(string method, int returnCode, string objectName = null)
        {
            string target = string.IsNullOrWhiteSpace(objectName) ? string.Empty : " for '" + objectName + "'";
            return method + " failed" + target + " (return code " + returnCode.ToString(CultureInfo.InvariantCulture) + ").";
        }

        private static string FormatLengthUnit(eLength lengthUnit)
        {
            switch (lengthUnit)
            {
                case eLength.inch:
                    return "in";
                case eLength.ft:
                    return "ft";
                case eLength.micron:
                    return "micron";
                case eLength.mm:
                    return "mm";
                case eLength.cm:
                    return "cm";
                case eLength.m:
                    return "m";
                default:
                    return lengthUnit.ToString();
            }
        }

        private static string Format(double value)
        {
            return value.ToString("G17", CultureInfo.InvariantCulture);
        }

        private sealed class CreatedRegion
        {
            public CreatedRegion(DropPanelRegion region, string areaName)
            {
                Region = region;
                AreaName = areaName;
            }

            public DropPanelRegion Region { get; private set; }

            public string AreaName { get; private set; }
        }

        private sealed class TableSnapshot
        {
            public TableSnapshot()
            {
                FieldKeys = new List<string>();
                Records = new List<Dictionary<string, string>>();
            }

            public string TableKey { get; set; }

            public int TableVersion { get; set; }

            public List<string> FieldKeys { get; set; }

            public List<Dictionary<string, string>> Records { get; set; }
        }

        private sealed class ObjectFieldIndexes
        {
            public int UniqueNameIndex { get; set; }

            public int StoryIndex { get; set; }

            public int LabelIndex { get; set; }

            public bool CanResolveObject
            {
                get { return UniqueNameIndex >= 0 || (StoryIndex >= 0 && LabelIndex >= 0); }
            }
        }
    }
}
