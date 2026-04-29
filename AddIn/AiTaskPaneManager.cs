using System;
using ExcelCSIToolBoxAddIn.UI.Views;
using Microsoft.Office.Core;

namespace ExcelCSIToolBoxAddIn.AddIn
{
    internal static class AiTaskPaneManager
    {
        private static Microsoft.Office.Tools.CustomTaskPane _aiTaskPane;
        private static WpfTaskPaneHost _aiTaskPaneHost;

        internal static void TogglePane()
        {
            if (_aiTaskPane != null && _aiTaskPane.Visible)
            {
                _aiTaskPane.Visible = false;
                return;
            }

            ShowPane();
        }

        internal static void ShowPane()
        {
            EnsurePane();
            _aiTaskPane.Visible = true;
        }

        internal static void DisposePane()
        {
            if (_aiTaskPane != null)
            {
                Globals.ExcelCSIToolBoxAddin.CustomTaskPanes.Remove(_aiTaskPane);
                _aiTaskPane = null;
            }

            if (_aiTaskPaneHost != null)
            {
                _aiTaskPaneHost.Dispose();
                _aiTaskPaneHost = null;
            }
        }

        private static void EnsurePane()
        {
            if (_aiTaskPane != null)
            {
                return;
            }

            if (Globals.ExcelCSIToolBoxAddin == null)
            {
                throw new InvalidOperationException("The Excel add-in is not initialized.");
            }

            _aiTaskPaneHost = new WpfTaskPaneHost(new AiAgentChatControl());
            _aiTaskPane = Globals.ExcelCSIToolBoxAddin.CustomTaskPanes.Add(_aiTaskPaneHost, "MHT AI Assistant");
            _aiTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            _aiTaskPane.Width = 420;
            _aiTaskPane.Visible = false;
        }
    }
}
