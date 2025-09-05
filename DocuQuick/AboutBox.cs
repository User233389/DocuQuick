using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Document_Maker
{
    partial class AboutBox : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public AboutBox()
        {
            InitializeComponent();
            this.Text = String.Format("{0} のバージョン情報", AssemblyTitle);
            this.labelProductName.Text = AssemblyProduct;
            this.labelVersion.Text = String.Format("バージョン {0}", AssemblyVersion);
            this.labelCopyright.Text = AssemblyCopyright;
            this.labelCompanyName.Text = AssemblyCompany;
        }
        
        #region アセンブリ属性アクセサー

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }


        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion

        private void AboutBox_Load(object sender, EventArgs e)
        {
            //Office2007青色
            if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonGroupBox1.StateCommon.Back.Color1 = Color.Empty;
                kryptonGroupBox1.StateCommon.Content.ShortText.Color1 = Color.Empty;
                kryptonLinkLabel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonPanel2.StateCommon.Color1 = Color.Empty;
            }
            //Office2007銀色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonGroupBox1.StateCommon.Back.Color1 = Color.Empty;
                kryptonGroupBox1.StateCommon.Content.ShortText.Color1 = Color.Empty;
                kryptonLinkLabel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonPanel2.StateCommon.Color1 = Color.Empty;
            }
            //Office2007黒色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                kryptonGroupBox1.StateCommon.Back.Color1 = Color.FromArgb(83, 83, 83);
                kryptonGroupBox1.StateCommon.Content.ShortText.Color1 = Color.White;
                labelProductName.StateCommon.ShortText.Color1 = Color.White;
                labelVersion.StateCommon.ShortText.Color1 = Color.White;
                labelCopyright.StateCommon.ShortText.Color1 = Color.White;
                labelCompanyName.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel1.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel2.StateCommon.ShortText.Color1 = Color.White;
                kryptonLinkLabel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;
                kryptonPanel2.StateCommon.Color1 = Color.FromArgb(83, 83, 83);
            }
            //Office2010青色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonGroupBox1.StateCommon.Back.Color1 = Color.Empty;
                kryptonGroupBox1.StateCommon.Content.ShortText.Color1 = Color.Empty;
                kryptonLinkLabel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonPanel2.StateCommon.Color1 = Color.Empty;
            }
            //Office2010銀色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonGroupBox1.StateCommon.Back.Color1 = Color.Empty;
                kryptonGroupBox1.StateCommon.Content.ShortText.Color1 = Color.Empty;
                kryptonLinkLabel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonPanel2.StateCommon.Color1 = Color.Empty;
            }
            //Office2010黒色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
                labelProductName.StateCommon.ShortText.Color1 = Color.White;
                labelVersion.StateCommon.ShortText.Color1 = Color.White;
                labelCopyright.StateCommon.ShortText.Color1 = Color.White;
                labelCompanyName.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel1.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel2.StateCommon.ShortText.Color1 = Color.White;
                kryptonLinkLabel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;
                kryptonPanel2.StateCommon.Color1 = Color.FromArgb(113, 113, 113);
            }

        }

        private void AboutBox_Shown(object sender, EventArgs e)
        {
            kryptonPanel1.AutoScroll = true;
            UpdateAvailableMessageAnimation();

        }
        
        async Task UpdateAvailableMessageAnimation()
        {
            await Task.Delay(5000);
            buttonSpecAny1.Text = "アップデートを行う...";
        }

        private void labelCopyright_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/User233389/Document-Maker/releases");
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void kryptonLinkLabel2_LinkClicked(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/User233389/Document-Maker");
        }
    }
}
