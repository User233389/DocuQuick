using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Document_Maker
{
    public partial class ClipBoradWindow : Form
    {
        public ClipBoradWindow()
        {
            InitializeComponent();
        }

        private void ClipBoradWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true; // Prevent the form from closing
            this.Hide(); // Hide the form instead of closing it
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            string selectedItem = listBox1.SelectedItem?.ToString();
            toolTip1.ToolTipTitle = "コピー完了"; // Set tooltip title
            toolTip1.Show(selectedItem + " をコピーしました。", listBox1, 1000); // Show tooltip for 1 second
            Clipboard.SetText(listBox1.SelectedItem.ToString()); // Copy the selected item to clipboard


        }
    }
}
