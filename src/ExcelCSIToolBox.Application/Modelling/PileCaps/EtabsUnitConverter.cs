using System;

namespace ExcelCSIToolBox.Application.Modelling.PileCaps
{
    public class EtabsUnitConverter
    {
        private const int Inch = 1;
        private const int Foot = 2;
        private const int Micron = 3;
        private const int Millimeter = 4;
        private const int Centimeter = 5;
        private const int Meter = 6;

        public double MillimetersToModelLength(double millimeters, int etabsLengthUnit)
        {
            return millimeters / GetMillimetersPerUnit(etabsLengthUnit);
        }

        public double ModelLengthToMillimeters(double modelLength, int etabsLengthUnit)
        {
            return modelLength * GetMillimetersPerUnit(etabsLengthUnit);
        }

        public double GetMillimetersPerUnit(int etabsLengthUnit)
        {
            switch (etabsLengthUnit)
            {
                case Inch:
                    return 25.4;
                case Foot:
                    return 304.8;
                case Micron:
                    return 0.001;
                case Millimeter:
                    return 1.0;
                case Centimeter:
                    return 10.0;
                case Meter:
                    return 1000.0;
                default:
                    throw new ArgumentOutOfRangeException("etabsLengthUnit", "Unsupported ETABS length unit.");
            }
        }
    }
}
