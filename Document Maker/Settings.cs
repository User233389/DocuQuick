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
    public partial class Settings : KryptonForm
    {
        public Settings()
        {
            InitializeComponent();


            GeneralPanel.Width = 676;
            AlignmentPanel.Width = 676;
            AutomaticUpdatePanel.Width = 676;
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void Settings_Load(object sender, EventArgs e)
        {

        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {

        }

        private void kryptonCheckButton1_Click(object sender, EventArgs e)
        {
            kryptonHeaderGroup1.ValuesPrimary.Heading = "一般";
            kryptonCheckButton1.Checked = true;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = false;

            GeneralPanel.Visible = true;
            AlignmentPanel.Visible = false;
            AutomaticUpdatePanel.Visible = false;

            GeneralPanel.Dock = DockStyle.Fill;
        }

        private void kryptonCheckButton2_Click(object sender, EventArgs e)
        {
            kryptonHeaderGroup1.ValuesPrimary.Heading = "連携";
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = true;
            kryptonCheckButton3.Checked = false;

            GeneralPanel.Visible = false;
            AlignmentPanel.Visible = true;
            AutomaticUpdatePanel.Visible = false;

            AlignmentPanel.Dock = DockStyle.Fill;
        }

        private void kryptonCheckButton3_Click(object sender, EventArgs e)
        {
            kryptonHeaderGroup1.ValuesPrimary.Heading = "自動更新";
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = true;

            GeneralPanel.Visible = false;
            AlignmentPanel.Visible = false;
            AutomaticUpdatePanel.Visible = true;

            AutomaticUpdatePanel.Dock = DockStyle.Fill;
        }

        private void Settings_Shown(object sender, EventArgs e)
        {
            this.Height = 600;
            GeneralPanel.AutoScroll = true;
            AlignmentPanel.AutoScroll = true;
            AutomaticUpdatePanel.AutoScroll = true;

            kryptonHeaderGroup1.ValuesPrimary.Heading = "一般";

            GeneralPanel.Dock = DockStyle.Fill;
            GeneralPanel.Visible = true;
            AlignmentPanel.Visible = false;
            AutomaticUpdatePanel.Visible = false;
        }
    }
}
