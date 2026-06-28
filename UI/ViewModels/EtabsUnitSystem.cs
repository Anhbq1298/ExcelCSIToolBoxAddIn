using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public sealed class EtabsUnitSystem
    {
        public EtabsUnitSystem(
            string displayName,
            int forceUnit,
            int lengthUnit,
            int temperatureUnit,
            string forceUnitText,
            string momentUnitText,
            string lengthUnitText,
            int legacyUnitsCode)
        {
            DisplayName = displayName;
            ForceUnit = forceUnit;
            LengthUnit = lengthUnit;
            TemperatureUnit = temperatureUnit;
            ForceUnitText = forceUnitText;
            MomentUnitText = momentUnitText;
            LengthUnitText = lengthUnitText;
            LegacyUnitsCode = legacyUnitsCode;
        }

        public string DisplayName { get; private set; }

        public int ForceUnit { get; private set; }

        public int LengthUnit { get; private set; }

        public int TemperatureUnit { get; private set; }

        public string ForceUnitText { get; private set; }

        public string MomentUnitText { get; private set; }

        public string LengthUnitText { get; private set; }

        public int LegacyUnitsCode { get; private set; }

        public string PresentUnitsText
        {
            get { return DisplayName + "-C"; }
        }

        public CSISapModelPresentUnitSystemDTO ToDto()
        {
            return new CSISapModelPresentUnitSystemDTO
            {
                ForceUnit = ForceUnit,
                LengthUnit = LengthUnit,
                TemperatureUnit = TemperatureUnit
            };
        }

        public BaseReactionUnitOption ToExportUnitOption()
        {
            return new BaseReactionUnitOption(
                DisplayName,
                LegacyUnitsCode,
                ForceUnitText,
                MomentUnitText,
                LengthUnitText);
        }

        public bool Matches(CSISapModelPresentUnitSystemDTO unitSystem)
        {
            return unitSystem != null &&
                   unitSystem.ForceUnit == ForceUnit &&
                   unitSystem.LengthUnit == LengthUnit &&
                   unitSystem.TemperatureUnit == TemperatureUnit;
        }
    }
}
