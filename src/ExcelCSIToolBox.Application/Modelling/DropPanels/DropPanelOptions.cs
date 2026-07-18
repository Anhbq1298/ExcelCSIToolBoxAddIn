namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelOptions
    {
        public DropPanelOptions()
        {
            DropSizeX = 2.0;
            DropSizeY = 2.0;
            RotationMode = DropPanelRotationMode.GlobalXY;
            GeometryTolerance = 0.001;
            ElevationTolerance = 0.01;
            MinimumPolygonArea = 0.001;
            VerticalRatioTolerance = 4.0;
            PreserveDirectAreaLoads = true;
            PreserveShellUniformLoadSetAssignments = true;
            PreserveLocalAxes = true;
            PreserveLocal3Orientation = true;
            PreserveDiaphragm = true;
            PreserveMeshAssignments = true;
            PreserveAreaModifiers = true;
            PreserveGroupAssignments = true;
            PreservePierAndSpandrelLabels = true;
            SaveEtabsBackupBeforeApply = true;
            MergeAdjacentRegionsOnlyWhenAssignmentSignaturesMatch = true;
            VerifyAssignmentsAfterApply = true;
        }

        public string DropProperty { get; set; }

        public double DropSizeX { get; set; }

        public double DropSizeY { get; set; }

        public DropPanelRotationMode RotationMode { get; set; }

        public double UserDefinedRotationAngle { get; set; }

        public double GeometryTolerance { get; set; }

        public double ElevationTolerance { get; set; }

        public double MinimumPolygonArea { get; set; }

        public double VerticalRatioTolerance { get; set; }

        public bool PreserveDirectAreaLoads { get; set; }

        public bool PreserveShellUniformLoadSetAssignments { get; set; }

        public bool PreserveLocalAxes { get; set; }

        public bool PreserveLocal3Orientation { get; set; }

        public bool PreserveDiaphragm { get; set; }

        public bool PreserveMeshAssignments { get; set; }

        public bool PreserveAreaModifiers { get; set; }

        public bool PreserveGroupAssignments { get; set; }

        public bool PreservePierAndSpandrelLabels { get; set; }

        public bool SaveEtabsBackupBeforeApply { get; set; }

        public bool MergeAdjacentRegionsOnlyWhenAssignmentSignaturesMatch { get; set; }

        public bool VerifyAssignmentsAfterApply { get; set; }
    }
}
