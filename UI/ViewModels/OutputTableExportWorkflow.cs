using System;
using ExcelCSIToolBox.Application.UseCases;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public static class OutputTableExportWorkflow
    {
        public static OutputTableExportOptionsWindow Run(
            OutputTableExportConfig config,
            CsiToolboxUseCaseBundle useCases,
            ICSISapModelConnectionService csiConnectionService,
            IExcelOutputService excelOutputService)
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            var viewModel = new OutputTableExportOptionsViewModel(
                useCases,
                csiConnectionService,
                excelOutputService,
                config);
            var window = new OutputTableExportOptionsWindow(viewModel);
            window.Show();
            return window;
        }
    }
}
