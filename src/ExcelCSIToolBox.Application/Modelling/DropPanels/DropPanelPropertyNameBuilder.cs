using System;
using System.Globalization;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public static class DropPanelPropertyNameBuilder
    {
        public static OperationResult<string> Build(double thickness, string materialName)
        {
            if (double.IsNaN(thickness) || double.IsInfinity(thickness) || thickness <= 0.0)
            {
                return OperationResult<string>.Failure("Drop thickness must be greater than zero.");
            }

            if (string.IsNullOrWhiteSpace(materialName))
            {
                return OperationResult<string>.Failure("Select a concrete material.");
            }

            string thicknessText = thickness.ToString("0.################", CultureInfo.InvariantCulture);
            return OperationResult<string>.Success("Drop_" + thicknessText + "_" + materialName);
        }
    }
}
