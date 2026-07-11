using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ExcelCSIToolBoxAddIn.UI.Forms
{
    internal partial class LoadPatternPickerForm : Form
    {
        public LoadPatternPickerForm(IEnumerable<string> availableLoadPatterns)
        {
            InitializeComponent();
            foreach (string patternName in (availableLoadPatterns ?? new string[0]).Where(name => !string.IsNullOrWhiteSpace(name)))
            {
                cboLoadPatterns.Items.Add(patternName);
            }

            if (cboLoadPatterns.Items.Count > 0)
            {
                cboLoadPatterns.SelectedIndex = 0;
            }

            btnAdd.Enabled = cboLoadPatterns.Items.Count > 0;
        }

        public string SelectedLoadPattern
        {
            get { return cboLoadPatterns.SelectedItem == null ? null : cboLoadPatterns.SelectedItem.ToString(); }
        }
    }
}
