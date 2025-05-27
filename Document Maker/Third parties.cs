using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Krypton.Toolkit;

namespace Document_Maker
{
    public partial class Third_parties : KryptonForm
    {
        public Third_parties()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.nuget.org/packages/Microsoft.Office.Interop.Outlook");

        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.icons8.com/");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.nuget.org/packages/Microsoft.Office.Interop.Word");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.nuget.org/packages/WindowsAPICodePack/");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.nuget.org/packages/RibbonWinForms/5.1.0-beta");
        }

        private void button7_Click(object sender, EventArgs e)
        {
             System.Diagnostics.Process.Start("https://www.nuget.org/packages/AeroWizard/2.0.9");
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
