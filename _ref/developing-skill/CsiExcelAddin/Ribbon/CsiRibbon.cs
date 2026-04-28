using Microsoft.Office.Core;
using CsiExcelAddin.Views;
using CsiExcelAddin.Bootstrap;

// RibbonX callback attribute required by Office — do not remove
[assembly: System.Runtime.InteropServices.ComVisible(true)]

namespace CsiExcelAddin.Ribbon
{
    /// <summary>
    /// Handles Excel Ribbon callbacks for this add-in.
    /// Responsibility is strictly limited to triggering the UI entry point.
    /// No business logic, no CSI calls, no Excel data processing here.
    /// </summary>
    public class CsiRibbon : IRibbonExtensibility
    {
        /// <summary>
        /// Returns the Ribbon XML that defines the custom tab and buttons.
        /// Called by Office when loading the add-in.
        /// </summary>
        public string GetCustomUI(string ribbonId)
        {
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='tabCsi' label='CSI Tools'>
        <group id='grpCsi' label='CSI Excel'>
          <button id='btnOpenMain'
                  label='Open CSI Panel'
                  imageMso='TableExcelSpreadsheetInsert'
                  size='large'
                  onAction='OnOpenMainPanel'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        /// <summary>
        /// Called when the user clicks the 'Open CSI Panel' Ribbon button.
        /// Delegates window creation to the composition root — no direct construction here.
        /// </summary>
        public void OnOpenMainPanel(IRibbonControl control)
        {
            // CompositionRoot wires up services, adapter, and ViewModel
            var window = CompositionRoot.BuildMainWindow();
            window.Show();
        }
    }
}
