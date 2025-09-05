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
    public partial class ThirdParty : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public ThirdParty()
        {
            InitializeComponent();
        }

        private void kryptonLabel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {

        }

        private void kryptonPanel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void ThirdParty_Load(object sender, EventArgs e)
        {
            //Office2007青色
            if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;



                kryptonPanel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonPanel3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonPanel3.StateCommon.Color1 = Color.FromArgb(225, 238, 255);
            }
            //Office2007銀色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;

                kryptonPanel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonPanel3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonPanel3.StateCommon.Color1 = Color.FromArgb(191, 194, 201);
            }
            //Office2007黒色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                kryptonLabel1.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel2.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel3.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel4.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel5.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel6.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel7.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel8.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel9.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel10.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel11.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel12.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel13.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel14.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel15.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel16.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel17.StateCommon.ShortText.Color1 = Color.White;

                kryptonPanel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonPanel3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonPanel3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonPanel3.StateCommon.Color1 = Color.FromArgb(30, 30, 30);
            }
            //Office2010青色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;

                kryptonPanel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonPanel3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;

                kryptonPanel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonPanel3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;

                kryptonLabel1.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel2.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel3.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel4.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel5.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel6.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel7.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel8.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel9.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel10.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel11.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel12.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel13.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel14.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel15.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel16.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel17.StateCommon.ShortText.Color1 = Color.White;

                kryptonPanel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonPanel3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
            }
        }

        private void kryptonPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.nuget.org/packages/Krypton.Toolkit/95.25.8.235?_src=template");
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.nuget.org/packages/Krypton.Toolkit.Suite.Extended.Ribbon");
        }

        private void kryptonButton4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.nuget.org/packages/Krypton.Components.Suite/4.5.8?_src=template");
        }

        private void kryptonButton5_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.nuget.org/packages/Microsoft.Office.Interop.Word/15.0.4797.1004?_src=template");
        }

        private void kryptonButton6_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.nuget.org/packages/Microsoft.Office.Interop.Word/15.0.4797.1004?_src=template");
        }

        private void kryptonButton7_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.nuget.org/packages/Microsoft.Office.Interop.Word/15.0.4797.1004?_src=template");
        }

        private void kryptonButton8_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.nuget.org/packages/FluentTransitions/2.0.1?_src=template");
        }

        private void kryptonButton9_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.nuget.org/packages/Microsoft.Windows.SDK.Contracts/10.0.26100.4948?_src=template");
        }
    }
}
