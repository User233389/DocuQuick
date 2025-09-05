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
    public partial class UseSoftWareWindow : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public UseSoftWareWindow()
        {
            InitializeComponent();
        }

        private void kryptonLabel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void UseSoftWareWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
        }

        private void kryptonLabel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonLabel2_Click(object sender, EventArgs e)
        {

        }

        private void kryptonPage1_Click(object sender, EventArgs e)
        {

        }

        private void kryptonPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonLabel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonPage3_Click(object sender, EventArgs e)
        {

        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {

        }

        private void UseSoftWareWindow_Load(object sender, EventArgs e)
        {
            kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.Panel;


        }

        //ページを戻った場合
        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            if(kryptonNavigator1.SelectedPage == kryptonPage2)
            {
                kryptonNavigator1.SelectedPage = kryptonPage1;
                kryptonButton3.Enabled = false;
            }
            else if (kryptonNavigator1.SelectedPage == kryptonPage4)
            {
                kryptonNavigator1.SelectedPage = kryptonPage2;
            }
            else if (kryptonNavigator1.SelectedPage == kryptonPage5)
            {
                kryptonNavigator1.SelectedPage = kryptonPage4;
            }
            else if (kryptonNavigator1.SelectedPage == kryptonPage6)
            {
                kryptonNavigator1.SelectedPage = kryptonPage5;
                kryptonButton2.Enabled = true;
                kryptonButton1.Text = "キャンセル";
            }
        }

        //ページを進んだ場合
        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            if (kryptonNavigator1.SelectedPage == kryptonPage1)
            {
                kryptonNavigator1.SelectedPage = kryptonPage2;
                kryptonButton3.Enabled = true;
            }
            else if (kryptonNavigator1.SelectedPage == kryptonPage2)
            {
                kryptonNavigator1.SelectedPage = kryptonPage4;
            }
            else if (kryptonNavigator1.SelectedPage == kryptonPage4)
            {
                kryptonNavigator1.SelectedPage = kryptonPage5;
            }
            else if (kryptonNavigator1.SelectedPage == kryptonPage5)
            {
                kryptonNavigator1.SelectedPage = kryptonPage6;
                kryptonButton2.Enabled = false;
                kryptonButton1.Text = "完了";
            }
        }

        private void kryptonCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(kryptonCheckBox1.Checked == true)
            {
                this.TopMost = true;
            }
            else
            {
                this.TopMost = false;
            }
        }

        private void kryptonButton1_Click_1(object sender, EventArgs e)
        {
            this.Hide();
        }

        public void LoadView()
        {

        }
        private void UseSoftWareWindow_Shown(object sender, EventArgs e)
        {

        }

        private void UseSoftWareWindow_Paint(object sender, PaintEventArgs e)
        {
            //Office2007青色
            if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007黒色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

        }
    }
}
