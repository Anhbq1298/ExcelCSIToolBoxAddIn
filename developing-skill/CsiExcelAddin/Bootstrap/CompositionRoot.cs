using CsiExcelAddin.CsiAdapters.Etabs;
using CsiExcelAddin.Services.Excel;
using CsiExcelAddin.ViewModels;
using CsiExcelAddin.Views;
using Microsoft.Office.Interop.Excel;

namespace CsiExcelAddin.Bootstrap
{
    /// <summary>
    /// Composition root — the single place where all dependencies are wired together.
    /// Only this class knows which concrete implementations are in use.
    /// ViewModels, Services, and Adapters only see interfaces.
    ///
    /// To switch from ETABS to SAP2000: change the adapter here only.
    /// No other file needs to change.
    /// </summary>
    public static class CompositionRoot
    {
        /// <summary>
        /// Builds a fully wired MainWindow ready to be shown.
        /// Call this from the Ribbon callback — nowhere else.
        /// </summary>
        public static MainWindow BuildMainWindow()
        {
            // Resolve Excel application from the running VSTO host
            Application excelApp = Globals.ThisAddIn.Application;

            // Wire services
            var excelWriter = new ExcelRangeWriter(excelApp);

            // Wire adapter — swap this line to switch CSI product
            var adapter = new EtabsAdapter();

            // Wire ViewModel
            var viewModel = new MainViewModel(adapter, excelWriter);

            // Wire View
            var mainWindow = new MainWindow(viewModel);
            return mainWindow;
        }
    }
}
