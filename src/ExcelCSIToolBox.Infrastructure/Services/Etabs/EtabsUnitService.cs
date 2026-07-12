using System;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Services;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;

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

            OperationResult<CSISapModelPresentUnitSystemDTO> unitSystemResult = ResolveUnitSystemFromMainWindow();
            if (!unitSystemResult.IsSuccess)
            {
                throw new InvalidOperationException(unitSystemResult.Message);
            }

            OperationResult result = _connectionService.SetPresentUnitSystem(unitSystemResult.Data);
            if (!result.IsSuccess)
            {
                string message = string.IsNullOrWhiteSpace(result.Message)
                    ? "Failed to set ETABS unit system."
                    : result.Message;
                throw new InvalidOperationException(message);
            }
        }

        public OperationResult<CsiPresentUnitScope> CreatePresentUnitScopeFromMainWindow()
        {
            if (_connectionService.SapModel == null)
            {
                return OperationResult<CsiPresentUnitScope>.Failure("Please attach to ETABS first.");
            }

            OperationResult<CSISapModelPresentUnitSystemDTO> unitSystemResult = ResolveUnitSystemFromMainWindow();
            if (!unitSystemResult.IsSuccess)
            {
                return OperationResult<CsiPresentUnitScope>.Failure(unitSystemResult.Message);
            }

            return CsiPresentUnitScope.Apply(_connectionService, unitSystemResult.Data);
        }

        private OperationResult<CSISapModelPresentUnitSystemDTO> ResolveUnitSystemFromMainWindow()
        {
            if (SelectedUnitSystem != null)
            {
                return OperationResult<CSISapModelPresentUnitSystemDTO>.Success(SelectedUnitSystem);
            }

            OperationResult<CSISapModelPresentUnitSystemDTO> currentUnitResult = _connectionService.GetPresentUnitSystem();
            if (currentUnitResult != null && currentUnitResult.IsSuccess && currentUnitResult.Data != null)
            {
                return currentUnitResult;
            }

            string message = currentUnitResult == null || string.IsNullOrWhiteSpace(currentUnitResult.Message)
                ? "Please select a unit system first."
                : currentUnitResult.Message;
            return OperationResult<CSISapModelPresentUnitSystemDTO>.Failure(message);
        }
    }
}
