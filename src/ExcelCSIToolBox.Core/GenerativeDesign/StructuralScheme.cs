using System.Collections.Generic;

namespace ExcelCSIToolBox.Core.GenerativeDesign
{
    public sealed class StructuralScheme
    {
        public string SchemeType { get; set; }
        public int BayCountX { get; set; }
        public int BayCountY { get; set; }
        public int StoryCount { get; set; }
        public double TypicalBayLength { get; set; }
        public double TypicalStoryHeight { get; set; }
        public List<string> PrimaryMaterials { get; set; }

        public StructuralScheme()
        {
            PrimaryMaterials = new List<string>();
        }
    }
}
