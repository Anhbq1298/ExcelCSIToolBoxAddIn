using System;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBox.Application.Services
{
    public sealed class CsiPresentUnitScope : IDisposable
    {
        private readonly Func<CSISapModelPresentUnitSystemDTO, OperationResult> _setPresentUnitSystem;
        private readonly CSISapModelPresentUnitSystemDTO _originalUnitSystem;
        private bool _disposed;

        private CsiPresentUnitScope(
            Func<CSISapModelPresentUnitSystemDTO, OperationResult> setPresentUnitSystem,
            CSISapModelPresentUnitSystemDTO originalUnitSystem)
        {
            _setPresentUnitSystem = setPresentUnitSystem ?? throw new ArgumentNullException(nameof(setPresentUnitSystem));
            _originalUnitSystem = Copy(originalUnitSystem);
            RestoreResult = OperationResult.Success();
        }

        public OperationResult RestoreResult { get; private set; }

        public static OperationResult<CsiPresentUnitScope> Apply(
            ICSISapModelConnectionService connectionService,
            CSISapModelPresentUnitSystemDTO requestedUnitSystem)
        {
            if (connectionService == null)
            {
                return OperationResult<CsiPresentUnitScope>.Failure("CSI connection service is not available.");
            }

            return Apply(
                connectionService.GetPresentUnitSystem,
                connectionService.SetPresentUnitSystem,
                requestedUnitSystem);
        }

        public static OperationResult<CsiPresentUnitScope> Apply(
            IEtabsConnectionService connectionService,
            CSISapModelPresentUnitSystemDTO requestedUnitSystem)
        {
            if (connectionService == null)
            {
                return OperationResult<CsiPresentUnitScope>.Failure("ETABS connection service is not available.");
            }

            return Apply(
                connectionService.GetPresentUnitSystem,
                connectionService.SetPresentUnitSystem,
                requestedUnitSystem);
        }

        private static OperationResult<CsiPresentUnitScope> Apply(
            Func<OperationResult<CSISapModelPresentUnitSystemDTO>> getPresentUnitSystem,
            Func<CSISapModelPresentUnitSystemDTO, OperationResult> setPresentUnitSystem,
            CSISapModelPresentUnitSystemDTO requestedUnitSystem)
        {
            if (requestedUnitSystem == null)
            {
                return OperationResult<CsiPresentUnitScope>.Failure("Select ETABS output units before running.");
            }

            OperationResult<CSISapModelPresentUnitSystemDTO> originalResult;
            try
            {
                originalResult = getPresentUnitSystem();
            }
            catch (Exception ex)
            {
                return OperationResult<CsiPresentUnitScope>.Failure("Failed to read current CSI present units: " + ex.Message);
            }

            if (originalResult == null || !originalResult.IsSuccess || originalResult.Data == null)
            {
                string message = originalResult == null || string.IsNullOrWhiteSpace(originalResult.Message)
                    ? "Failed to read current CSI present units."
                    : originalResult.Message;
                return OperationResult<CsiPresentUnitScope>.Failure(message);
            }

            try
            {
                OperationResult applyResult = setPresentUnitSystem(Copy(requestedUnitSystem));
                if (applyResult == null || !applyResult.IsSuccess)
                {
                    string message = applyResult == null || string.IsNullOrWhiteSpace(applyResult.Message)
                        ? "Failed to set CSI present units."
                        : applyResult.Message;
                    return OperationResult<CsiPresentUnitScope>.Failure(message);
                }
            }
            catch (Exception ex)
            {
                return OperationResult<CsiPresentUnitScope>.Failure("Failed to set CSI present units: " + ex.Message);
            }

            return OperationResult<CsiPresentUnitScope>.Success(
                new CsiPresentUnitScope(setPresentUnitSystem, originalResult.Data));
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            try
            {
                RestoreResult = _setPresentUnitSystem(Copy(_originalUnitSystem));
                if (RestoreResult == null)
                {
                    RestoreResult = OperationResult.Failure("Failed to restore CSI present units.");
                }
            }
            catch (Exception ex)
            {
                RestoreResult = OperationResult.Failure("Failed to restore CSI present units: " + ex.Message);
            }
        }

        private static CSISapModelPresentUnitSystemDTO Copy(CSISapModelPresentUnitSystemDTO unitSystem)
        {
            if (unitSystem == null)
            {
                return null;
            }

            return new CSISapModelPresentUnitSystemDTO
            {
                ForceUnit = unitSystem.ForceUnit,
                LengthUnit = unitSystem.LengthUnit,
                TemperatureUnit = unitSystem.TemperatureUnit
            };
        }
    }
}
