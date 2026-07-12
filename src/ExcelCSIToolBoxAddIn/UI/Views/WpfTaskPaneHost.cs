using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    internal sealed class WpfTaskPaneHost : UserControl
    {
        public WpfTaskPaneHost(UIElement child)
        {
            Dock = DockStyle.Fill;

            Controls.Add(new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = child
            });
        }
    }
}
