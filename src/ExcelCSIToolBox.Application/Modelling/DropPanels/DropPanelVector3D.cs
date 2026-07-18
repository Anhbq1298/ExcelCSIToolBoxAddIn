namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelVector3D
    {
        public DropPanelVector3D()
        {
        }

        public DropPanelVector3D(double x, double y, double z)
        {
            X = x;
            Y = y;
            Z = z;
        }

        public double X { get; set; }

        public double Y { get; set; }

        public double Z { get; set; }
    }
}
