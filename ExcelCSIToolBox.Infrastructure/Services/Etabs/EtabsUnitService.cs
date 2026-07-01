using System;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs
{
    public class EtabsUnitService : IEtabsUnitService
    {
        private readonly IEtabsConnectionService _connectionService;

        public EtabsUnitService(IEtabsConnectionService connectionService)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException(nameof(connectionService));
        }

        public CSISapModelPresentUnitSystemDTO SelectedUnitSystem { get; private set; }

        public string SelectedUnitDisplayName { get; private set; }

        public string SelectedPresentUnitsText { get; private set; }

        public void SetSelectedUnitSystem(
            CSISapModelPresentUnitSystemDTO unitSystem,
            string displayName,
            string presentUnitsText)
        {
            SelectedUnitSystem = unitSystem;
            SelectedUnitDisplayName = displayName;
            SelectedPresentUnitsText = presentUnitsText;
        }

        public void SetPresentUnitsFromMainWindow()
        {
            if (_connectionService.SapModel == null)
            {
                throw new InvalidOperationException("Please attach to ETABS first.");
            }

            if (SelectedUnitSystem == null)
            {
                throw new InvalidOperationException("Please select a unit system first.");
            }

            OperationResult result = _connectionService.SetPresentUnitSystem(SelectedUnitSystem);
            if (!result.IsSuccess)
            {
                string message = string.IsNullOrWhiteSpace(result.Message)
                    ? "Failed to set ETABS unit system."
                    : result.Message;
                throw new InvalidOperationException(message);
            }
        }
    }
}
