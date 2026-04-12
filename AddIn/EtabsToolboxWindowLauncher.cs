using System;
using System.Windows;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using ExcelCSIToolBoxAddIn.UI.ViewModels;
using ExcelCSIToolBoxAddIn.UI.Views;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    /// <summary>
    /// Opens the ETABS toolbox shell window from the add-in layer.
    /// Keeps window creation logic out of ribbon click handlers.
    /// </summary>
    public class EtabsToolboxWindowLauncher
    {
        private readonly IEtabsConnectionService _etabsConnectionService;

        public EtabsToolboxWindowLauncher(IEtabsConnectionService etabsConnectionService)
        {
            _etabsConnectionService = etabsConnectionService ?? throw new ArgumentNullException(nameof(etabsConnectionService));
        }

        public void OpenWindow()
        {
            var viewModel = new EtabsToolboxViewModel(_etabsConnectionService);
            var window = new EtabsToolboxWindow
            {
                DataContext = viewModel
            };

            // Show as modeless so users can continue working in Excel.
            window.Show();
            window.Activate();
        }
    }
}
