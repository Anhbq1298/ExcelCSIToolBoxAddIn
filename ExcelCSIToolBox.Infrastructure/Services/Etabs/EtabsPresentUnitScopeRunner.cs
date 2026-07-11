using System;
using System.Diagnostics;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Services;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs
{
    internal static class EtabsPresentUnitScopeRunner
    {
        public static CsiPresentUnitScope Begin(IEtabsUnitService unitService)
        {
            if (unitService == null)
            {
                throw new InvalidOperationException("ETABS unit service is not available.");
            }

            OperationResult<CsiPresentUnitScope> scopeResult = unitService.CreatePresentUnitScopeFromMainWindow();
            if (scopeResult == null || !scopeResult.IsSuccess || scopeResult.Data == null)
            {
                string message = scopeResult == null || string.IsNullOrWhiteSpace(scopeResult.Message)
                    ? "Failed to set ETABS present units."
                    : scopeResult.Message;
                throw new InvalidOperationException(message);
            }

            return scopeResult.Data;
        }

        public static void Restore(CsiPresentUnitScope scope, string context)
        {
            if (scope == null)
            {
                return;
            }

            scope.Dispose();
            if (scope.RestoreResult != null && !scope.RestoreResult.IsSuccess)
            {
                Trace.WriteLine("Failed to restore ETABS present units after " + Safe(context) + ": " + scope.RestoreResult.Message);
            }
        }

        private static string Safe(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "export" : value;
        }
    }
}
