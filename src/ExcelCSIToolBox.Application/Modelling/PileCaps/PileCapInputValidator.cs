using System.Collections.Generic;
using ExcelCSIToolBox.Core.Contracts.CSI.PileCap;

namespace ExcelCSIToolBox.Application.Modelling.PileCaps
{
    public class PileCapInputValidator
    {
        public IReadOnlyList<string> Validate(PileCapInputParameters input)
        {
            var messages = new List<string>();
            if (input == null)
            {
                messages.Add("Pile-cap inputs are required.");
                return messages;
            }

            if (input.PileDiameterMillimeters <= 0)
            {
                messages.Add("Pile diameter must be greater than zero.");
            }

            if (input.PileLengthMillimeters <= 0)
            {
                messages.Add("Pile length must be greater than zero.");
            }

            if (input.PileCapThicknessMillimeters <= 0)
            {
                messages.Add("Pile-cap thickness must be greater than zero.");
            }

            if (input.EdgeDistanceMillimeters < 0)
            {
                messages.Add("Edge distance must be zero or greater.");
            }

            if (string.IsNullOrWhiteSpace(input.PileMaterial))
            {
                messages.Add("Pile material is required.");
            }

            if (string.IsNullOrWhiteSpace(input.PileCapMaterial))
            {
                messages.Add("Pile-cap material is required.");
            }

            ValidateSpacing(input, messages);
            return messages;
        }

        private static void ValidateSpacing(PileCapInputParameters input, ICollection<string> messages)
        {
            if (input.ArrangementType == PileCapArrangementType.Mono)
            {
                return;
            }

            if (input.ArrangementType == PileCapArrangementType.FourPile)
            {
                if (input.SpacingXMillimeters <= input.PileDiameterMillimeters)
                {
                    messages.Add("Spacing X must be greater than pile diameter.");
                }

                if (input.SpacingYMillimeters <= input.PileDiameterMillimeters)
                {
                    messages.Add("Spacing Y must be greater than pile diameter.");
                }

                return;
            }

            if (input.PileSpacingMillimeters <= input.PileDiameterMillimeters)
            {
                messages.Add("Pile spacing must be greater than pile diameter.");
            }
        }
    }
}
