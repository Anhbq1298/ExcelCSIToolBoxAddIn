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
        private string _lastBackupPath;
        private string _lastOriginalModelPath;

        public EtabsDropPanelService(
            IEtabsConnectionService connectionService,
            ICsiApiDispatcher dispatcher,
            ICsiOperationLogger operationLogger)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException(nameof(connectionService));
            _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
            _operationLogger = operationLogger ?? throw new ArgumentNullException(nameof(operationLogger));
        }

        public bool IsRollbackAvailable
        {
            get
            {
                return !string.IsNullOrWhiteSpace(_lastBackupPath) &&
                       !string.IsNullOrWhiteSpace(_lastOriginalModelPath) &&
                       File.Exists(_lastBackupPath);
            }
        }

        public OperationResult<DropPanelModelContext> GetModelContext()
        {
            return _dispatcher.Invoke(GetModelContextCore);
        }

        public OperationResult<IReadOnlyList<string>> GetDropPropertyNames()
        {
            return _dispatcher.Invoke(GetDropPropertyNamesCore);
        }

        public OperationResult<IReadOnlyList<DropPanelColumnInfo>> ReadSelectedColumns(double verticalRatioTolerance)
        {
            return _dispatcher.Invoke(() => ReadSelectedColumnsCore(verticalRatioTolerance));
        }

        public OperationResult<DropPanelPreparationSnapshot> PrepareSnapshot(
            IReadOnlyList<DropPanelColumnInfo> columns,
            IReadOnlyList<DropPanelRequest> requests,
            DropPanelOptions options)
        {
            return _dispatcher.Invoke(() => PrepareSnapshotCore(columns, requests, options));
        }

        public OperationResult HighlightAreas(IReadOnlyList<string> areaNames)
        {
            return _dispatcher.Invoke(() => HighlightAreasCore(areaNames));
        }

        public OperationResult<DropPanelApplyResult> Apply(DropPanelOperationPlan plan, DropPanelOptions options)
        {
            return _dispatcher.Invoke(() => ApplyCore(plan, options));
        }

        public OperationResult Rollback()
        {
            return _dispatcher.Invoke(RollbackCore);
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
            string units = sapModel.GetPresentUnits().ToString();
            return OperationResult<DropPanelModelContext>.Success(new DropPanelModelContext
            {
                Version = version,
                ModelFileName = Path.GetFileName(modelPath),
                ModelPath = modelPath,
                PresentUnits = units,
                IsLocked = isLocked
            });
        }

        private OperationResult<IReadOnlyList<string>> GetDropPropertyNamesCore()
        {
            cSapModel sapModel;
            OperationResult modelResult = TryGetSapModel(out sapModel);
            if (!modelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(modelResult.Message);
            }

            int numberNames = 0;
            string[] names = null;
            int returnCode = sapModel.PropArea.GetNameList(ref numberNames, ref names);
            if (returnCode != 0)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(ReturnCodeMessage("PropArea.GetNameList", returnCode));
            }

            if (numberNames > 0 && (names == null || names.Length < numberNames))
            {
                return OperationResult<IReadOnlyList<string>>.Failure("PropArea.GetNameList returned an incomplete property-name array.");
            }

            List<string> result = (names ?? new string[0])
                .Take(numberNames)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Where(name => IsSlabProperty(sapModel, name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            return OperationResult<IReadOnlyList<string>>.Success(result);
        }

        private OperationResult<IReadOnlyList<DropPanelColumnInfo>> ReadSelectedColumnsCore(double verticalRatioTolerance)
        {
            cSapModel sapModel;
            OperationResult modelResult = TryGetSapModel(out sapModel);
            if (!modelResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<DropPanelColumnInfo>>.Failure(modelResult.Message);
            }

            if (verticalRatioTolerance <= 0.0)
            {
                return OperationResult<IReadOnlyList<DropPanelColumnInfo>>.Failure("Vertical ratio tolerance must be greater than zero.");
            }

            int numberItems = 0;
            int[] objectTypes = null;
            string[] objectNames = null;
            int selectedReturn = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
            if (selectedReturn != 0)
            {
                return OperationResult<IReadOnlyList<DropPanelColumnInfo>>.Failure(ReturnCodeMessage("SelectObj.GetSelected", selectedReturn));
            }

            if (numberItems == 0)
            {
                return OperationResult<IReadOnlyList<DropPanelColumnInfo>>.Failure("No ETABS objects are selected.");
            }

            if (objectTypes == null || objectNames == null ||
                objectTypes.Length < numberItems || objectNames.Length < numberItems)
            {
                return OperationResult<IReadOnlyList<DropPanelColumnInfo>>.Failure(
                    "SelectObj.GetSelected returned incomplete object arrays.");
            }

            List<DropPanelColumnInfo> columns = new List<DropPanelColumnInfo>();
            for (int index = 0; index < numberItems; index++)
            {
                if (string.IsNullOrWhiteSpace(objectNames[index]))
                {
                    columns.Add(new DropPanelColumnInfo
                    {
                        FrameName = string.Empty,
                        IsValid = false,
                        ValidationMessage = "ETABS returned an empty selected-object name."
                    });
                    continue;
                }

                if (objectTypes[index] != CSISapModelObjectTypeIds.Frame)
                {
                    columns.Add(new DropPanelColumnInfo
                    {
                        FrameName = objectNames[index],
                        IsValid = false,
                        ValidationMessage = "The selected object is not a frame object."
                    });
                    continue;
                }

                DropPanelColumnInfo column;
                OperationResult columnResult = TryReadColumn(sapModel, objectNames[index], verticalRatioTolerance, out column);
                if (!columnResult.IsSuccess)
                {
                    column = column ?? new DropPanelColumnInfo { FrameName = objectNames[index] };
                    column.IsValid = false;
                    column.ValidationMessage = columnResult.Message;
                }

                columns.Add(column);
            }

            return OperationResult<IReadOnlyList<DropPanelColumnInfo>>.Success(columns);
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
                OperationResult loadSetResult = ReadLoadSetAssignments(sapModel, out loadSetsByArea);
                if (!loadSetResult.IsSuccess)
                {
                    return OperationResult<DropPanelPreparationSnapshot>.Failure(loadSetResult.Message);
                }

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
                if (!IsAreaRelevant(area, requests, connectedColumnsByArea, options) ||
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
                    "No slab area candidates were found near the selected column heads. Verify connectivity, elevations, and drop dimensions.");
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
                    continue;
                }

                for (int index = 0; index < numberItems; index++)
                {
                    if (objectTypes == null || objectNames == null || index >= objectTypes.Length || index >= objectNames.Length ||
                        objectTypes[index] != ConnectivityAreaObjectType || string.IsNullOrWhiteSpace(objectNames[index]))
                    {
                        continue;
                    }

                    HashSet<string> connectedColumns;
                    if (!connectedColumnsByArea.TryGetValue(objectNames[index], out connectedColumns))
                    {
                        connectedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        connectedColumnsByArea[objectNames[index]] = connectedColumns;
                    }

                    connectedColumns.Add(column.FrameName);
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

        private static bool IsAreaRelevant(
            DropPanelAreaInfo area,
            IReadOnlyList<DropPanelRequest> requests,
            IDictionary<string, HashSet<string>> connectedColumnsByArea,
            DropPanelOptions options)
        {
            HashSet<string> connectedColumns;
            if (connectedColumnsByArea.TryGetValue(area.AreaName, out connectedColumns) &&
                requests.Any(request => connectedColumns.Contains(request.ColumnName) &&
                                        Math.Abs(request.Elevation - area.Elevation) <= options.ElevationTolerance))
            {
                return true;
            }

            double minX = area.Points.Min(point => point.X) - options.GeometryTolerance;
            double maxX = area.Points.Max(point => point.X) + options.GeometryTolerance;
            double minY = area.Points.Min(point => point.Y) - options.GeometryTolerance;
            double maxY = area.Points.Max(point => point.Y) + options.GeometryTolerance;
            foreach (DropPanelRequest request in requests)
            {
                if (Math.Abs(request.Elevation - area.Elevation) > options.ElevationTolerance || request.Points.Count == 0)
                {
                    continue;
                }

                double requestMinX = request.Points.Min(point => point.X);
                double requestMaxX = request.Points.Max(point => point.X);
                double requestMinY = request.Points.Min(point => point.Y);
                double requestMaxY = request.Points.Max(point => point.Y);
                DropPanelPoint3D requestCenter = new DropPanelPoint3D(
                    request.Points.Average(point => point.X),
                    request.Points.Average(point => point.Y),
                    request.Elevation);
                if (PointInPolygonOrNearBoundary(requestCenter, area.Points, options.GeometryTolerance))
                {
                    return true;
                }

                if (minX <= requestMaxX && maxX >= requestMinX && minY <= requestMaxY && maxY >= requestMinY)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool PointInPolygonOrNearBoundary(
            DropPanelPoint3D point,
            IReadOnlyList<DropPanelPoint3D> polygon,
            double tolerance)
        {
            if (point == null || polygon == null || polygon.Count < 3)
            {
                return false;
            }

            bool inside = false;
            for (int index = 0, previousIndex = polygon.Count - 1; index < polygon.Count; previousIndex = index++)
            {
                DropPanelPoint3D current = polygon[index];
                DropPanelPoint3D previous = polygon[previousIndex];
                if (DistanceToSegment(point, previous, current) <= tolerance)
                {
                    return true;
                }

                bool crosses = (current.Y > point.Y) != (previous.Y > point.Y) &&
                               point.X < (previous.X - current.X) * (point.Y - current.Y) /
                               (previous.Y - current.Y) + current.X;
                if (crosses)
                {
                    inside = !inside;
                }
            }

            return inside;
        }

        private static double DistanceToSegment(
            DropPanelPoint3D point,
            DropPanelPoint3D start,
            DropPanelPoint3D end)
        {
            double dx = end.X - start.X;
            double dy = end.Y - start.Y;
            double lengthSquared = dx * dx + dy * dy;
            if (lengthSquared <= NumericTolerance * NumericTolerance)
            {
                double pointDx = point.X - start.X;
                double pointDy = point.Y - start.Y;
                return Math.Sqrt(pointDx * pointDx + pointDy * pointDy);
            }

            double parameter = ((point.X - start.X) * dx + (point.Y - start.Y) * dy) / lengthSquared;
            parameter = Math.Max(0.0, Math.Min(1.0, parameter));
            double closestX = start.X + parameter * dx;
            double closestY = start.Y + parameter * dy;
            double closestDx = point.X - closestX;
            double closestDy = point.Y - closestY;
            return Math.Sqrt(closestDx * closestDx + closestDy * closestDy);
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

        private OperationResult HighlightAreasCore(IReadOnlyList<string> areaNames)
        {
            cSapModel sapModel;
            OperationResult modelResult = TryGetSapModel(out sapModel);
            if (!modelResult.IsSuccess)
            {
                return modelResult;
            }

            int clearReturn = sapModel.SelectObj.ClearSelection();
            if (clearReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("SelectObj.ClearSelection", clearReturn));
            }

            foreach (string areaName in areaNames ?? new string[0])
            {
                int selectReturn = sapModel.AreaObj.SetSelected(areaName, true, eItemType.Objects);
                if (selectReturn != 0)
                {
                    return OperationResult.Failure(ReturnCodeMessage("AreaObj.SetSelected", selectReturn, areaName));
                }
            }

            int refreshReturn = sapModel.View.RefreshView(0, false);
            return refreshReturn == 0
                ? OperationResult.Success("Affected ETABS areas highlighted.")
                : OperationResult.Failure(ReturnCodeMessage("View.RefreshView", refreshReturn));
        }

        private OperationResult<DropPanelApplyResult> ApplyCore(DropPanelOperationPlan plan, DropPanelOptions options)
        {
            if (plan == null || !plan.IsValid || options == null)
            {
                return OperationResult<DropPanelApplyResult>.Failure("A valid preview plan is required before applying changes.");
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

            // A new Apply must never use a backup created by an earlier operation.
            _lastBackupPath = null;
            _lastOriginalModelPath = null;

            OperationResult revalidation = RevalidateSources(sapModel, plan, options);
            if (!revalidation.IsSuccess)
            {
                return OperationResult<DropPanelApplyResult>.Failure(revalidation.Message);
            }

            string backupPath = string.Empty;
            if (options.SaveEtabsBackupBeforeApply)
            {
                OperationResult backupResult = CreateBackup(sapModel, out backupPath);
                if (!backupResult.IsSuccess)
                {
                    return OperationResult<DropPanelApplyResult>.Failure(backupResult.Message);
                }
            }

            List<string> sourceAreaNames = plan.SourceAreas.Select(area => area.AreaName).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            List<CreatedRegion> createdRegions = new List<CreatedRegion>();
            bool deletionStarted = false;
            try
            {
                foreach (string sourceAreaName in sourceAreaNames)
                {
                    int deleteReturn = sapModel.AreaObj.Delete(sourceAreaName, eItemType.Objects);
                    if (deleteReturn != 0)
                    {
                        throw new InvalidOperationException(ReturnCodeMessage("AreaObj.Delete", deleteReturn, sourceAreaName));
                    }

                    deletionStarted = true;
                }

                int regionIndex = 0;
                string areaNamePrefix = "DP_" + DateTime.Now.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture);
                foreach (DropPanelRegion region in plan.Regions)
                {
                    regionIndex++;
                    string createdAreaName = CreateAreaRegion(sapModel, region, regionIndex, options, areaNamePrefix);
                    createdRegions.Add(new CreatedRegion(region, createdAreaName));
                }

                if (options.PreserveShellUniformLoadSetAssignments)
                {
                    OperationResult loadSetRestore = RestoreLoadSetAssignments(sapModel, sourceAreaNames, createdRegions);
                    if (!loadSetRestore.IsSuccess)
                    {
                        throw new InvalidOperationException(loadSetRestore.Message);
                    }
                }

                foreach (CreatedRegion createdRegion in createdRegions)
                {
                    RestoreDirectAreaLoadsAndDiaphragm(sapModel, createdRegion, options);
                }

                if (options.PreserveMeshAssignments)
                {
                    OperationResult meshRestore = RestoreMeshAssignments(sapModel, sourceAreaNames, createdRegions);
                    if (!meshRestore.IsSuccess)
                    {
                        throw new InvalidOperationException(meshRestore.Message);
                    }
                }

                foreach (CreatedRegion createdRegion in createdRegions)
                {
                    RestoreModifiersGroupsAndLabels(sapModel, createdRegion, options);
                }

                DropPanelApplyResult result = VerifyAndBuildResult(sapModel, plan, createdRegions, options, backupPath);
                int refreshReturn = sapModel.View.RefreshView(0, false);
                if (refreshReturn != 0)
                {
                    AddIssue(result, string.Empty, string.Empty, "ETABS View", "Refreshed", refreshReturn.ToString(CultureInfo.InvariantCulture), ReturnCodeMessage("View.RefreshView", refreshReturn));
                }

                _operationLogger.Log(
                    "ETABS",
                    "Drop Panel",
                    "Shells / Areas",
                    "Batch Replacement",
                    CsiMethodRiskLevel.High,
                    "Created " + createdRegions.Count.ToString(CultureInfo.InvariantCulture) + " region(s) from " + sourceAreaNames.Count.ToString(CultureInfo.InvariantCulture) + " source area(s).",
                    sourceAreaNames,
                    true,
                    result.VerificationPassed,
                    result.VerificationPassed ? "Drop panel batch applied and verified." : "Drop panel batch applied, but verification failed.");

                return OperationResult<DropPanelApplyResult>.Success(
                    result,
                    result.VerificationPassed
                        ? "Drop panels were applied and verified."
                        : "Drop panels were applied, but read-back verification failed. Use Rollback or review the verification log.");
            }
            catch (Exception ex)
            {
                string rollbackMessage = string.Empty;
                if (deletionStarted && IsRollbackAvailable)
                {
                    OperationResult rollbackResult = RollbackCore();
                    rollbackMessage = rollbackResult.IsSuccess
                        ? " The saved model backup was restored automatically."
                        : " Automatic rollback failed: " + rollbackResult.Message;
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
                    ex.Message + rollbackMessage);
                return OperationResult<DropPanelApplyResult>.Failure("Drop panel apply failed: " + ex.Message + rollbackMessage);
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
                return OperationResult.Failure("The active ETABS model changed after Preview. Run Preview again.");
            }

            string currentUnits = sapModel.GetPresentUnits().ToString();
            if (!string.IsNullOrWhiteSpace(plan.PresentUnits) &&
                !string.Equals(plan.PresentUnits, currentUnits, StringComparison.Ordinal))
            {
                return OperationResult.Failure(
                    "ETABS present units changed from '" + plan.PresentUnits + "' to '" + currentUnits + "' after Preview. Restore the units or run Preview again.");
            }

            int propertyCount = 0;
            string[] propertyNames = null;
            int propertiesReturn = sapModel.PropArea.GetNameList(ref propertyCount, ref propertyNames);
            if (propertiesReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("PropArea.GetNameList", propertiesReturn));
            }

            if (!(propertyNames ?? new string[0]).Contains(options.DropProperty, StringComparer.OrdinalIgnoreCase))
            {
                return OperationResult.Failure("The selected drop property no longer exists. Run Preview again.");
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
                    return OperationResult.Failure("Source area '" + sourceArea.AreaName + "' changed after Preview. Run Preview again.");
                }

                DropPanelAreaInfo currentArea;
                OperationResult areaResult = TryReadAreaGeometry(
                    sapModel, sourceArea.AreaName, options.ElevationTolerance, out currentArea);
                if (!areaResult.IsSuccess || currentArea == null ||
                    !PolygonPointsEqual(sourceArea.Points, currentArea.Points, options.GeometryTolerance))
                {
                    return OperationResult.Failure(
                        "Source area '" + sourceArea.AreaName + "' geometry changed after Preview. Run Preview again.");
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

        private OperationResult CreateBackup(cSapModel sapModel, out string backupPath)
        {
            backupPath = string.Empty;
            string modelPath = sapModel.GetModelFilename(true) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(modelPath))
            {
                return OperationResult.Failure("Save the ETABS model before applying drop panels so a backup can be created.");
            }

            int saveReturn = sapModel.File.Save(modelPath);
            if (saveReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("File.Save", saveReturn, modelPath));
            }

            try
            {
                string directory = Path.GetDirectoryName(modelPath);
                string fileName = Path.GetFileNameWithoutExtension(modelPath);
                string extension = Path.GetExtension(modelPath);
                string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
                backupPath = Path.Combine(directory, fileName + ".drop-panel-backup-" + timestamp + extension);
                int suffix = 2;
                while (File.Exists(backupPath))
                {
                    backupPath = Path.Combine(directory, fileName + ".drop-panel-backup-" + timestamp + "-" + suffix.ToString(CultureInfo.InvariantCulture) + extension);
                    suffix++;
                }

                File.Copy(modelPath, backupPath, false);
                _lastOriginalModelPath = modelPath;
                _lastBackupPath = backupPath;
                return OperationResult.Success("ETABS backup created at " + backupPath + ".");
            }
            catch (Exception ex)
            {
                return OperationResult.Failure("Could not create the ETABS backup: " + ex.Message);
            }
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

        private static DropPanelApplyResult VerifyAndBuildResult(
            cSapModel sapModel,
            DropPanelOperationPlan plan,
            IReadOnlyList<CreatedRegion> createdRegions,
            DropPanelOptions options,
            string backupPath)
        {
            DropPanelApplyResult result = new DropPanelApplyResult { BackupFilePath = backupPath };
            result.CreatedAreaNames.AddRange(createdRegions.Select(item => item.AreaName));

            Dictionary<string, List<string>> loadSetsByArea = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            if (options.VerifyAssignmentsAfterApply && options.PreserveShellUniformLoadSetAssignments)
            {
                OperationResult loadSetRead = ReadLoadSetAssignments(sapModel, out loadSetsByArea);
                if (!loadSetRead.IsSuccess)
                {
                    AddIssue(result, string.Empty, string.Empty, "Shell Uniform Load Set", "Readable", "Unreadable", loadSetRead.Message);
                }
            }


            TableSnapshot meshVerificationTable = null;
            string meshVerificationError = string.Empty;
            if (options.VerifyAssignmentsAfterApply && options.PreserveMeshAssignments)
            {
                DropPanelMeshAssignment firstMesh = createdRegions
                    .Select(item => item.Region.Assignment.MeshAssignment)
                    .FirstOrDefault(item => item != null && !string.IsNullOrWhiteSpace(item.TableKey));
                if (firstMesh == null)
                {
                    meshVerificationError = "Mesh assignment backup is missing.";
                }
                else
                {
                    OperationResult meshRead = ReadDisplayTable(sapModel, firstMesh.TableKey, out meshVerificationTable);
                    if (!meshRead.IsSuccess)
                    {
                        meshVerificationError = meshRead.Message;
                    }
                }
            }

            foreach (CreatedRegion createdRegion in createdRegions)
            {
                DropPanelRegion region = createdRegion.Region;
                DropPanelAreaAssignmentBackup expected = region.Assignment;
                string areaName = createdRegion.AreaName;
                if (options.VerifyAssignmentsAfterApply)
                {
                    VerifyProperty(sapModel, result, region, areaName);
                    VerifyLocalAxes(sapModel, result, region, areaName, options);
                    VerifyDirectLoads(sapModel, result, region, areaName, options);
                    VerifyDiaphragm(sapModel, result, region, areaName, options);
                    VerifyModifiers(sapModel, result, region, areaName, options);
                    VerifyGroups(sapModel, result, region, areaName, options);
                    VerifyLabels(sapModel, result, region, areaName, options);
                    VerifyMeshAssignment(
                        sapModel, result, region, areaName, options, meshVerificationTable, meshVerificationError);

                    if (options.PreserveShellUniformLoadSetAssignments)
                    {
                        List<string> actualLoadSets;
                        loadSetsByArea.TryGetValue(areaName, out actualLoadSets);
                        if (!SetEquals(expected.ShellUniformLoadSetNames, actualLoadSets))
                        {
                            AddIssue(result, region.SourceAreaName, areaName, "Shell Uniform Load Set",
                                Join(expected.ShellUniformLoadSetNames), Join(actualLoadSets), "Shell Uniform Load Set assignments do not match.");
                        }
                    }
                }

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
                    DirectLoadStatus = HasIssue(result, areaName, "Direct Area Load") ? "Failed" : "Passed",
                    ShellLoadSetStatus = HasIssue(result, areaName, "Shell Uniform Load Set") ? "Failed" : "Passed",
                    LocalAxisStatus = HasIssue(result, areaName, "Local Axis") ? "Failed" : "Passed",
                    Local3Status = HasIssue(result, areaName, "Local 3") ? "Failed" : "Passed",
                    DiaphragmStatus = HasIssue(result, areaName, "Diaphragm") ? "Failed" : "Passed",
                    VerificationStatus = result.VerificationIssues.Any(issue => string.Equals(issue.NewAreaName, areaName, StringComparison.OrdinalIgnoreCase)) ? "Failed" : "Passed",
                    Message = result.VerificationIssues.Any(issue => string.Equals(issue.NewAreaName, areaName, StringComparison.OrdinalIgnoreCase))
                        ? "One or more read-back checks failed."
                        : "Assignments restored and verified."
                };
                result.LogEntries.Add(entry);
            }

            return result;
        }

        private OperationResult RollbackCore()
        {
            if (!IsRollbackAvailable)
            {
                return OperationResult.Failure("No Drop Panel backup is available for rollback.");
            }

            cSapModel sapModel;
            OperationResult modelResult = TryGetSapModel(out sapModel);
            if (!modelResult.IsSuccess)
            {
                return modelResult;
            }

            if (sapModel.GetModelIsLocked())
            {
                return OperationResult.Failure("The ETABS model is locked. Unlock it before rollback.");
            }

            int openReturn = sapModel.File.OpenFile(_lastBackupPath);
            if (openReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("File.OpenFile", openReturn, _lastBackupPath));
            }

            int saveReturn = sapModel.File.Save(_lastOriginalModelPath);
            if (saveReturn != 0)
            {
                return OperationResult.Failure(ReturnCodeMessage("File.Save", saveReturn, _lastOriginalModelPath));
            }

            _operationLogger.Log(
                "ETABS", "Drop Panel Rollback", "Model", "Backup Restore", CsiMethodRiskLevel.High,
                "Restored " + _lastBackupPath + " to " + _lastOriginalModelPath + ".",
                new string[0], true, true, "Drop panel rollback completed.");
            return OperationResult.Success("The ETABS model was restored from " + _lastBackupPath + ".");
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

        private static OperationResult ReadLoadSetAssignments(cSapModel sapModel, out Dictionary<string, List<string>> assignments)
        {
            assignments = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
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

        private static void VerifyProperty(cSapModel sapModel, DropPanelApplyResult result, DropPanelRegion region, string areaName)
        {
            string actual = string.Empty;
            int returnCode = sapModel.AreaObj.GetProperty(areaName, ref actual);
            if (returnCode != 0 || !string.Equals(actual, region.ResultingSectionProperty, StringComparison.OrdinalIgnoreCase))
            {
                AddIssue(result, region.SourceAreaName, areaName, "Area Property", region.ResultingSectionProperty, actual,
                    returnCode == 0 ? "Area property does not match." : ReturnCodeMessage("AreaObj.GetProperty", returnCode, areaName));
            }
        }

        private static void VerifyLocalAxes(cSapModel sapModel, DropPanelApplyResult result, DropPanelRegion region, string areaName, DropPanelOptions options)
        {
            if (options.PreserveLocalAxes)
            {
                double angle = 0.0;
                bool advanced = false;
                int axesReturn = sapModel.AreaObj.GetLocalAxes(areaName, ref angle, ref advanced);
                if (axesReturn != 0 || !NearlyEqual(angle, region.Assignment.LocalAxisAngle))
                {
                    AddIssue(result, region.SourceAreaName, areaName, "Local Axis",
                        Format(region.Assignment.LocalAxisAngle), Format(angle),
                        axesReturn == 0 ? "Local axis angle does not match." : ReturnCodeMessage("AreaObj.GetLocalAxes", axesReturn, areaName));
                }
            }

            if (options.PreserveLocal3Orientation)
            {
                double[] transformation = null;
                int returnCode = sapModel.AreaObj.GetTransformationMatrix(areaName, ref transformation, true);
                DropPanelVector3D actual = transformation != null && transformation.Length >= 9
                    ? NormalizeVector(new DropPanelVector3D(transformation[6], transformation[7], transformation[8]))
                    : new DropPanelVector3D();
                if (returnCode != 0 || Dot(actual, region.Assignment.Local3Direction) < 1.0 - 1e-5)
                {
                    AddIssue(result, region.SourceAreaName, areaName, "Local 3",
                        VectorText(region.Assignment.Local3Direction), VectorText(actual),
                        returnCode == 0 ? "Local 3 orientation does not match." : ReturnCodeMessage("AreaObj.GetTransformationMatrix", returnCode, areaName));
                }
            }
        }

        private static void VerifyDirectLoads(cSapModel sapModel, DropPanelApplyResult result, DropPanelRegion region, string areaName, DropPanelOptions options)
        {
            if (!options.PreserveDirectAreaLoads)
            {
                return;
            }

            List<DropPanelDirectAreaLoad> actual = new List<DropPanelDirectAreaLoad>();
            OperationResult readResult = ReadDirectAreaLoads(sapModel, areaName, actual);
            if (!readResult.IsSuccess || !LoadsEqual(region.Assignment.DirectAreaLoads, actual))
            {
                AddIssue(result, region.SourceAreaName, areaName, "Direct Area Load",
                    LoadsText(region.Assignment.DirectAreaLoads), LoadsText(actual),
                    readResult.IsSuccess ? "Direct area loads do not match." : readResult.Message);
            }
        }

        private static void VerifyDiaphragm(cSapModel sapModel, DropPanelApplyResult result, DropPanelRegion region, string areaName, DropPanelOptions options)
        {
            if (!options.PreserveDiaphragm)
            {
                return;
            }

            string actual = string.Empty;
            int returnCode = sapModel.AreaObj.GetDiaphragm(areaName, ref actual);
            if (returnCode != 0 || !string.Equals(actual ?? string.Empty, region.Assignment.Diaphragm ?? string.Empty, StringComparison.OrdinalIgnoreCase))
            {
                AddIssue(result, region.SourceAreaName, areaName, "Diaphragm", region.Assignment.Diaphragm, actual,
                    returnCode == 0 ? "Diaphragm assignment does not match." : ReturnCodeMessage("AreaObj.GetDiaphragm", returnCode, areaName));
            }
        }

        private static void VerifyModifiers(cSapModel sapModel, DropPanelApplyResult result, DropPanelRegion region, string areaName, DropPanelOptions options)
        {
            if (!options.PreserveAreaModifiers)
            {
                return;
            }

            double[] actual = null;
            int returnCode = sapModel.AreaObj.GetModifiers(areaName, ref actual);
            if (returnCode != 0 || !NumbersEqual(region.Assignment.Modifiers, actual))
            {
                AddIssue(result, region.SourceAreaName, areaName, "Area Modifiers",
                    NumbersText(region.Assignment.Modifiers), NumbersText(actual),
                    returnCode == 0 ? "Area modifiers do not match." : ReturnCodeMessage("AreaObj.GetModifiers", returnCode, areaName));
            }
        }

        private static void VerifyGroups(cSapModel sapModel, DropPanelApplyResult result, DropPanelRegion region, string areaName, DropPanelOptions options)
        {
            if (!options.PreserveGroupAssignments)
            {
                return;
            }

            int count = 0;
            string[] groups = null;
            int returnCode = sapModel.AreaObj.GetGroupAssign(areaName, ref count, ref groups);
            List<string> actual = (groups ?? new string[0]).Where(group => !IsImplicitAllGroup(group)).ToList();
            if (returnCode != 0 || !SetEquals(region.Assignment.Groups, actual))
            {
                AddIssue(result, region.SourceAreaName, areaName, "Group Assignment", Join(region.Assignment.Groups), Join(actual),
                    returnCode == 0 ? "Group assignments do not match." : ReturnCodeMessage("AreaObj.GetGroupAssign", returnCode, areaName));
            }
        }

        private static void VerifyLabels(cSapModel sapModel, DropPanelApplyResult result, DropPanelRegion region, string areaName, DropPanelOptions options)
        {
            if (!options.PreservePierAndSpandrelLabels)
            {
                return;
            }

            string pier = string.Empty;
            int pierReturn = sapModel.AreaObj.GetPier(areaName, ref pier);
            if (pierReturn != 0 || !string.Equals(pier ?? string.Empty, region.Assignment.PierLabel ?? string.Empty, StringComparison.OrdinalIgnoreCase))
            {
                AddIssue(result, region.SourceAreaName, areaName, "Pier Label", region.Assignment.PierLabel, pier,
                    pierReturn == 0 ? "Pier label does not match." : ReturnCodeMessage("AreaObj.GetPier", pierReturn, areaName));
            }

            string spandrel = string.Empty;
            int spandrelReturn = sapModel.AreaObj.GetSpandrel(areaName, ref spandrel);
            if (spandrelReturn != 0 || !string.Equals(spandrel ?? string.Empty, region.Assignment.SpandrelLabel ?? string.Empty, StringComparison.OrdinalIgnoreCase))
            {
                AddIssue(result, region.SourceAreaName, areaName, "Spandrel Label", region.Assignment.SpandrelLabel, spandrel,
                    spandrelReturn == 0 ? "Spandrel label does not match." : ReturnCodeMessage("AreaObj.GetSpandrel", spandrelReturn, areaName));
            }
        }

        private static void VerifyMeshAssignment(
            cSapModel sapModel,
            DropPanelApplyResult result,
            DropPanelRegion region,
            string areaName,
            DropPanelOptions options,
            TableSnapshot actualTable,
            string readError)
        {
            if (!options.PreserveMeshAssignments)
            {
                return;
            }

            DropPanelMeshAssignment expected = region.Assignment.MeshAssignment;
            if (expected == null)
            {
                AddIssue(result, region.SourceAreaName, areaName, "Mesh Assignment", "Backed up", "Missing",
                    "Mesh assignment backup is missing.");
                return;
            }

            if (!string.IsNullOrWhiteSpace(readError) || actualTable == null)
            {
                AddIssue(result, region.SourceAreaName, areaName, "Mesh Assignment", "Readable", "Unreadable",
                    string.IsNullOrWhiteSpace(readError) ? "The mesh assignment table could not be read." : readError);
                return;
            }

            ObjectFieldIndexes indexes = FindObjectFieldIndexes(actualTable.FieldKeys);
            if (!indexes.CanResolveObject)
            {
                AddIssue(result, region.SourceAreaName, areaName, "Mesh Assignment", "Recognized schema", "Unsupported schema",
                    "The ETABS mesh assignment table does not expose recognizable area identity fields.");
                return;
            }

            string label = string.Empty;
            string story = string.Empty;
            int labelReturn = sapModel.AreaObj.GetLabelFromName(areaName, ref label, ref story);
            if (labelReturn != 0)
            {
                AddIssue(result, region.SourceAreaName, areaName, "Mesh Assignment", "Readable identity", "Unreadable identity",
                    ReturnCodeMessage("AreaObj.GetLabelFromName", labelReturn, areaName));
                return;
            }

            List<string> expectedRecords = expected.Records
                .Select(record => CanonicalizeTableRecord(record.Values, actualTable.FieldKeys, indexes))
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList();
            List<string> actualRecords = actualTable.Records
                .Where(record => RecordMatchesObject(record, actualTable.FieldKeys, indexes, areaName, label, story))
                .Select(record => CanonicalizeTableRecord(record, actualTable.FieldKeys, indexes))
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (!expectedRecords.SequenceEqual(actualRecords, StringComparer.OrdinalIgnoreCase))
            {
                AddIssue(result, region.SourceAreaName, areaName, "Mesh Assignment",
                    expectedRecords.Count.ToString(CultureInfo.InvariantCulture) + " matching row(s)",
                    actualRecords.Count.ToString(CultureInfo.InvariantCulture) + " matching row(s)",
                    "Mesh assignment table rows do not match the source assignment.");
            }
        }

        private static string CanonicalizeTableRecord(
            IDictionary<string, string> record,
            IReadOnlyList<string> fields,
            ObjectFieldIndexes indexes)
        {
            List<string> values = new List<string>();
            for (int index = 0; index < fields.Count; index++)
            {
                if (index == indexes.UniqueNameIndex || index == indexes.LabelIndex || index == indexes.StoryIndex)
                {
                    continue;
                }

                values.Add(fields[index].Trim().ToUpperInvariant() + "=" + NormalizeTableValue(ReadField(record, fields, index)));
            }

            return string.Join("\u001e", values);
        }

        private static string NormalizeTableValue(string value)
        {
            string trimmed = (value ?? string.Empty).Trim();
            double number;
            if (double.TryParse(trimmed, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number))
            {
                return number.ToString("G17", CultureInfo.InvariantCulture);
            }

            return trimmed.ToUpperInvariant();
        }

        private static bool LoadsEqual(IReadOnlyList<DropPanelDirectAreaLoad> expected, IReadOnlyList<DropPanelDirectAreaLoad> actual)
        {
            List<DropPanelDirectAreaLoad> remaining = new List<DropPanelDirectAreaLoad>(actual ?? new DropPanelDirectAreaLoad[0]);
            foreach (DropPanelDirectAreaLoad expectedLoad in expected ?? new DropPanelDirectAreaLoad[0])
            {
                int index = remaining.FindIndex(item =>
                    string.Equals(item.LoadPattern, expectedLoad.LoadPattern, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(item.LoadType, expectedLoad.LoadType, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(item.CoordinateSystem, expectedLoad.CoordinateSystem, StringComparison.OrdinalIgnoreCase) &&
                    item.Direction == expectedLoad.Direction && NearlyEqual(item.Value, expectedLoad.Value));
                if (index < 0)
                {
                    return false;
                }

                remaining.RemoveAt(index);
            }

            return remaining.Count == 0;
        }

        private static bool NumbersEqual(IReadOnlyList<double> expected, IReadOnlyList<double> actual)
        {
            if ((expected == null ? 0 : expected.Count) != (actual == null ? 0 : actual.Count))
            {
                return false;
            }

            for (int index = 0; index < (expected == null ? 0 : expected.Count); index++)
            {
                if (!NearlyEqual(expected[index], actual[index]))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool SetEquals(IEnumerable<string> expected, IEnumerable<string> actual)
        {
            HashSet<string> expectedSet = new HashSet<string>(expected ?? new string[0], StringComparer.OrdinalIgnoreCase);
            HashSet<string> actualSet = new HashSet<string>(actual ?? new string[0], StringComparer.OrdinalIgnoreCase);
            return expectedSet.SetEquals(actualSet);
        }

        private static bool NearlyEqual(double left, double right)
        {
            return Math.Abs(left - right) <= NumericTolerance * Math.Max(1.0, Math.Max(Math.Abs(left), Math.Abs(right)));
        }

        private static void AddIssue(
            DropPanelApplyResult result,
            string sourceArea,
            string newArea,
            string assignmentType,
            string expected,
            string actual,
            string message)
        {
            result.VerificationIssues.Add(new DropPanelVerificationIssue
            {
                SourceAreaName = sourceArea,
                NewAreaName = newArea,
                AssignmentType = assignmentType,
                ExpectedValue = expected,
                ActualValue = actual,
                ErrorMessage = message
            });
        }

        private static bool HasIssue(DropPanelApplyResult result, string areaName, string assignmentType)
        {
            return result.VerificationIssues.Any(issue =>
                string.Equals(issue.NewAreaName, areaName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(issue.AssignmentType, assignmentType, StringComparison.OrdinalIgnoreCase));
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

        private static string Format(double value)
        {
            return value.ToString("G17", CultureInfo.InvariantCulture);
        }

        private static string VectorText(DropPanelVector3D vector)
        {
            return vector == null ? string.Empty : Format(vector.X) + ", " + Format(vector.Y) + ", " + Format(vector.Z);
        }

        private static string Join(IEnumerable<string> values)
        {
            return string.Join(", ", (values ?? new string[0]).OrderBy(value => value, StringComparer.OrdinalIgnoreCase));
        }

        private static string NumbersText(IEnumerable<double> values)
        {
            return string.Join(", ", (values ?? new double[0]).Select(Format));
        }

        private static string LoadsText(IEnumerable<DropPanelDirectAreaLoad> loads)
        {
            return string.Join("; ", (loads ?? new DropPanelDirectAreaLoad[0])
                .Select(load => load.LoadPattern + "|" + load.LoadType + "|" + load.Direction.ToString(CultureInfo.InvariantCulture) + "|" + load.CoordinateSystem + "|" + Format(load.Value))
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase));
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
