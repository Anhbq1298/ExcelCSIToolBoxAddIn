using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace ExcelCSIToolBoxAddIn.UI.Views
{
    internal sealed class AiAgentTaskPaneHost : UserControl
    {
        public AiAgentTaskPaneHost()
        {
            Dock = DockStyle.Fill;

            var elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = new AiAgentChatTab()
            };

            Controls.Add(elementHost);
        }
    }
}
