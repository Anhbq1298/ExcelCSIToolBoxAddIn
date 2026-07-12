using System.Collections.Generic;

namespace ExcelCSIToolBox.Core.GenerativeDesign
{
    public sealed class DesignConstraintSet
    {
        public double MinSpan { get; set; }
        public double MaxSpan { get; set; }
        public int MinStories { get; set; }
        public int MaxStories { get; set; }
        public double MaxDriftRatio { get; set; }
        public double MaxWeight { get; set; }
        public List<string> PreferredMaterials { get; set; }

        public DesignConstraintSet()
        {
            PreferredMaterials = new List<string>();
        }
    }
}
