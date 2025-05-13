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
    public partial class AppExit : Form
    {
        public AppExit()
        {
            InitializeComponent();
        }

        private void AppExit_Shown(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
