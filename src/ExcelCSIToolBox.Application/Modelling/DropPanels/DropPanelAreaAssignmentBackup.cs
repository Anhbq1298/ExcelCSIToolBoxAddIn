using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelAreaAssignmentBackup
    {
        public DropPanelAreaAssignmentBackup()
        {
            DirectAreaLoads = new List<DropPanelDirectAreaLoad>();
            ShellUniformLoadSetNames = new List<string>();
            Groups = new List<string>();
            Modifiers = new double[0];
            Local3Direction = new DropPanelVector3D(0.0, 0.0, 1.0);
        }

        public string SourceAreaName { get; set; }

        public string SourceAreaLabel { get; set; }

        public string StoryName { get; set; }

        public string SectionProperty { get; set; }

        public double LocalAxisAngle { get; set; }

        public bool UsesAdvancedLocalAxes { get; set; }

        public DropPanelVector3D Local3Direction { get; set; }

        public bool OriginalWindingIsCounterClockwise { get; set; }

        public string Diaphragm { get; set; }

        public List<DropPanelDirectAreaLoad> DirectAreaLoads { get; set; }

        public List<string> ShellUniformLoadSetNames { get; set; }

        public List<string> Groups { get; set; }

        public DropPanelMeshAssignment MeshAssignment { get; set; }

        public double[] Modifiers { get; set; }

        public string PierLabel { get; set; }

        public string SpandrelLabel { get; set; }

        public bool IsOpening { get; set; }
    }
}
