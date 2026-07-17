using System;
using System.Globalization;

namespace ExcelCSIToolBox.Application.Modelling.PileCaps
{
    public static class PileCapPropertyNameBuilder
    {
        public static string BuildPileFrameSectionName(double pileDiameterMillimeters, string materialName)
        {
            return "P_" + FormatIntegerMillimeters(pileDiameterMillimeters) + "D_" + (materialName ?? string.Empty);
        }

        public static string BuildPileCapAreaSectionName(double pileCapThicknessMillimeters, string materialName)
        {
            return "PC_" + FormatIntegerMillimeters(pileCapThicknessMillimeters) + "_" + (materialName ?? string.Empty);
        }

        private static string FormatIntegerMillimeters(double value)
        {
            double rounded = Math.Round(value, 0, MidpointRounding.AwayFromZero);
            return rounded.ToString("0", CultureInfo.InvariantCulture);
        }
    }
}
