namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelDirectAreaLoad
    {
        public string LoadPattern { get; set; }

        public string LoadType { get; set; }

        public double Value { get; set; }

        public int Direction { get; set; }

        public string CoordinateSystem { get; set; }

        /// <summary>
        /// Indicates that this load starts a complete replacement set for its load pattern during restoration.
        /// ETABS does not expose this setter flag through GetLoadUniform, so it is derived deterministically on read.
        /// </summary>
        public bool ReplaceExistingAssignments { get; set; }
    }
}
