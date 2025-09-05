using ComponentFactory.Krypton.Navigator;
using ComponentFactory.Krypton.Ribbon;
using FluentTransitions;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Foundation.Metadata;
using Windows.Security.Credentials.UI;
using Windows.UI.Xaml.Documents;
using Windows.UI.Xaml.Shapes;

namespace Document_Maker
{

    public partial class Form1 : ComponentFactory.Krypton.Toolkit.KryptonForm
    {

        public Form1()
        {
            InitializeComponent();
        }


        TreeNode treeNode1 = new TreeNode();


        private void kryptonRibbonGroupCheckBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        //下準備
        public void SetTheme()
        {
            //2007
            if (Properties.Settings.Default.Theme == "Office2007Blue")
            {
                //テーマの設定
                of2007.Checked = true;
                of2010.Checked = false;
                //ラジオボタンの設定
                kryptonContextMenuRadioButton1.Checked = true;
                kryptonContextMenuRadioButton2.Checked = false;
                kryptonContextMenuRadioButton3.Checked = false;

                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(191, 219, 255);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(191, 219, 255);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007青色
                this.BackColor = Color.FromArgb(191, 219, 255);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(191, 219, 255);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;



                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            else if(Properties.Settings.Default.Theme == "Office2007Silver")
            {
                //テーマの設定
                of2007.Checked = true;
                of2010.Checked = false;
                //ラジオボタンの設定
                kryptonContextMenuRadioButton1.Checked = false;
                kryptonContextMenuRadioButton2.Checked = true;
                kryptonContextMenuRadioButton3.Checked = false;

                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(208, 212, 221);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(208, 212, 221);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007銀色
                this.BackColor = Color.FromArgb(208, 212, 221);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(208, 212, 221);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;


                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            else  if(Properties.Settings.Default.Theme == "Office2007Black")
            {
                //テーマの設定
                of2007.Checked = true;
                of2010.Checked = false;
                //ラジオボタンの設定
                kryptonContextMenuRadioButton1.Checked = false;
                kryptonContextMenuRadioButton2.Checked = false;
                kryptonContextMenuRadioButton3.Checked = true;

                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(83, 83, 83);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(83, 83, 83);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007黒色
                this.BackColor = Color.FromArgb(83, 83, 83);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.Black;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.White;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(83, 83, 83);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //2010
            else if(Properties.Settings.Default.Theme == "Office2010Blue")
            {
                //テーマの設定
                of2007.Checked = false;
                of2010.Checked = true;

                //ラジオボタンの設定
                kryptonContextMenuRadioButton1.Checked = true;
                kryptonContextMenuRadioButton2.Checked = false;
                kryptonContextMenuRadioButton3.Checked = false;

                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(187, 206, 230);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(187, 206, 230);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010青色
                this.BackColor = Color.FromArgb(187, 206, 230);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(187, 206, 230);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;



                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;


                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            else if(Properties.Settings.Default.Theme == "Office2010Silver")
            {
                //テーマの設定
                of2007.Checked = false;
                of2010.Checked = true;

                //ラジオボタンの設定
                kryptonContextMenuRadioButton1.Checked = false;
                kryptonContextMenuRadioButton2.Checked = true;
                kryptonContextMenuRadioButton3.Checked = false;

                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(227, 230, 232);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(227, 230, 232);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010銀色
                this.BackColor = Color.FromArgb(227, 230, 232);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(227, 230, 232);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            else if(Properties.Settings.Default.Theme == "Office2010Black")
            {
                //テーマの設定
                of2007.Checked = false;
                of2010.Checked = true;
                //ラジオボタンの設定
                kryptonContextMenuRadioButton1.Checked = false;
                kryptonContextMenuRadioButton2.Checked = false;
                kryptonContextMenuRadioButton3.Checked = true;

                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(113, 113, 113);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(113, 113, 113);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010黒色
                this.BackColor = Color.FromArgb(113, 113, 113);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.Black;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.White;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(113, 113, 113);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.White;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
            }
        }

        public void SetQAT()
        {
            if(Properties.Settings.Default.QAT1_Visible == true)
            {
                kryptonRibbonQATButton1.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton1.Visible = false;
            }

            if (Properties.Settings.Default.QAT2_Visible == true)
            {
                kryptonRibbonQATButton2.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton2.Visible = false;
            }

            if (Properties.Settings.Default.QAT3_Visible == true)
            {
                kryptonRibbonQATButton3.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton3.Visible = false;
            }

            if (Properties.Settings.Default.QAT4_Visible == true)
            {
                kryptonRibbonQATButton4.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton4.Visible = false;
            }

            if (Properties.Settings.Default.QAT5_Visible == true)
            {
                kryptonRibbonQATButton5.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton5.Visible = false;
            }

            if (Properties.Settings.Default.QAT6_Visible == true)
            {
                kryptonRibbonQATButton6.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton6.Visible = false;
            }

            if (Properties.Settings.Default.QAT7_Visible == true)
            {
                kryptonRibbonQATButton7.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton7.Visible = false;
            }

            if (Properties.Settings.Default.QAT8_Visible == true)
            {
                kryptonRibbonQATButton8.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton8.Visible = false;
            }
        }

        public void SetSheetSpace()
        {
            Sheets_TopPanel.Height = Properties.Settings.Default.Space_Top;
            Sheets_ButtomPanel.Height = Properties.Settings.Default.Space_Buttom;
            Sheets_LeftPanel.Width = Properties.Settings.Default.Space_Left;
            Sheets_RightPanel.Width = Properties.Settings.Default.Space_Right;

            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = Properties.Settings.Default.Space_Top;
            kryptonRibbonGroupNumericUpDown1.Value = Properties.Settings.Default.Space_Buttom;
            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = Properties.Settings.Default.Space_Left;
            kryptonRibbonGroupNumericUpDown2.Value = Properties.Settings.Default.Space_Right;
        }

        public void SetSheetText()
        {
            //発行元部署
            kryptonTextBox11.Text = Properties.Settings.Default.SendingDepartment;
            //宛先会社
            kryptonTextBox1.Text = Properties.Settings.Default.To_CompanyOrOrganizationName;
            //宛先肩書
            kryptonComboBox10.Text = Properties.Settings.Default.To_Title;
            //宛先氏名
            kryptonTextBox2.Text = Properties.Settings.Default.To_Name;

            //発信者会社
            kryptonTextBox3.Text = Properties.Settings.Default.Caller_CompanyOrOrganizationName;
            //発信者所在地
            kryptonTextBox4.Text = Properties.Settings.Default.Caller_Location;
            //発信者建物名
            kryptonTextBox5.Text = Properties.Settings.Default.Caller_BuildingName;
            //発信者階数
            kryptonNumericUpDown2.Value = Properties.Settings.Default.Caller_FloorNumber;
            //発信者肩書
            kryptonComboBox9.Text = Properties.Settings.Default.Caller_Title;
            //発信者氏名
            kryptonTextBox6.Text = Properties.Settings.Default.Caller_Name;
            //メールアドレス
            kryptonTextBox7.Text = Properties.Settings.Default.Caller_MailAddress_User;
            kryptonComboBox8.Text = Properties.Settings.Default.Caller_MailAddress_Domain;
            //電話番号1
            kryptonComboBox6.Text = Properties.Settings.Default.Caller_PhoneNumber1;
            kryptonTextBox14.Text = Properties.Settings.Default.Caller_PhoneNumber2;
            kryptonTextBox8.Text = Properties.Settings.Default.Caller_PhoneNumber3;
            //Fax番号
            kryptonComboBox7.Text = Properties.Settings.Default.Caller_FaxNumber1;
            kryptonTextBox9.Text = Properties.Settings.Default.Caller_FaxNumber2;
            kryptonTextBox15.Text = Properties.Settings.Default.Caller_FaxNumber3;


        }

        public void RunAppTask()
        {

            if (Properties.Settings.Default.ShowApplicationTask == 0)
            {
                kryptonRadioButton1.Checked = false;
                kryptonRadioButton2.Checked = false;
                kryptonRadioButton3.Checked = true;
            }
            else if (Properties.Settings.Default.ShowApplicationTask == 1)
            {
                kryptonRadioButton1.Checked = false;
                kryptonRadioButton2.Checked = true;
                kryptonRadioButton3.Checked = false;
                Transition
                    .With(kryptonPanel21, nameof(Height), 0)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                kryptonTrackBar1.Enabled = false;
                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = false;
                kryptonLabel42.Enabled = false;

                kryptonRibbon.Enabled = false;
                kryptonRibbon.MinimizedMode = true;
                kryptonPage9.Visible = true;
                kryptonNavigator_Workbench.SelectedPage = kryptonPage9;
                kryptonNavigator_Workbench.NavigatorMode = NavigatorMode.Panel;
                this.Text = "テンプレート - DoQuick Designer";

                kryptonLabel7.Enabled = false;
                kryptonCheckButton1.Enabled = false;
                kryptonCheckButton2.Enabled = false;
                kryptonLabel1.Enabled = false;
            }
            else if (Properties.Settings.Default.ShowApplicationTask == 2)
            {
                kryptonRadioButton1.Checked = true;
                kryptonRadioButton2.Checked = false;
                kryptonRadioButton3.Checked = false;
                ShowDcw();  

            }
        }

        async System.Threading.Tasks.Task ShowDcw()
        {
            await System.Threading.Tasks.Task.Delay(1000);
            DCW dCW = new DCW();

            Properties.Settings.Default.dCW_TopSpace = Sheets_TopPanel.Height;
            Properties.Settings.Default.dCW_ButtomSpace = Sheets_ButtomPanel.Height;
            Properties.Settings.Default.dCW_LeftSpace = Sheets_LeftPanel.Width;
            Properties.Settings.Default.dCW_RightSpace = Sheets_RightPanel.Width;
            Properties.Settings.Default.Save();

            // Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                dCW.BackColor = Color.FromArgb(191, 219, 255);
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                dCW.BackColor = Color.FromArgb(208, 212, 221);
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                dCW.BackColor = Color.FromArgb(83, 83, 83);
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                dCW.BackColor = Color.FromArgb(187, 206, 230);
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                dCW.BackColor = Color.FromArgb(227, 230, 232);
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                dCW.BackColor = Color.FromArgb(113, 113, 113);
            }

            dCW.ShowDialog();

            //trueの場合のみ実行
            //ウィザードの「完了」ボタンをクリックしたときに実行
            if (dCW.IsWizardFinished == true)
            {
                FontReset();
                //発行番号
                if (dCW.NoIssueNumber == false)
                {
                    kryptonCheckBox3.Checked = false;
                    kryptonTextBox11.Text = dCW.IssueNumber_Publisher;
                    kryptonNumericUpDown1.Value = dCW.IssueNumber;
                }
                else
                {
                    kryptonCheckBox3.Checked = true;
                }

                //日付
                if (dCW.NoDate == false)
                {
                    kryptonCheckBox2.Checked = false;
                    kryptonDateTimePicker1.Value = dCW.Date;
                    if (dCW.UseEraName == true)
                    {
                        kryptonCheckBox1.Checked = true;
                    }
                    else
                    {
                        kryptonCheckBox1.Checked = false;
                    }
                }
                else
                {
                    kryptonCheckBox2.Checked = true;
                }

                //発信者
                kryptonTextBox1.Text = dCW.AdCompany;
                kryptonComboBox10.Text = dCW.AdTitle;
                kryptonTextBox2.Text = dCW.AdName;

                kryptonTextBox3.Text = dCW.CaCampany;
                kryptonTextBox4.Text = dCW.CaLocation;
                kryptonTextBox5.Text = dCW.CaBuildingName;
                kryptonNumericUpDown2.Value = dCW.CaFloorNumber;
                kryptonComboBox9.Text = dCW.CaTitle;
                kryptonTextBox6.Text = dCW.CaName;
                kryptonTextBox7.Text = dCW.CaMailAddress;
                kryptonComboBox8.Text = dCW.CaMailAddress_Domain;
                //電話番号
                kryptonComboBox6.Text = dCW.CaPhoneNumber1;
                kryptonTextBox14.Text = dCW.CaPhoneNumber2;
                kryptonTextBox8.Text = dCW.CaPhoneNumber3;
                kryptonComboBox7.Text = dCW.CaFaxNumber1;
                kryptonTextBox9.Text = dCW.CaFaxNumber2;
                kryptonTextBox15.Text = dCW.CaFaxNumber3;

                //表題
                kryptonTextBox10.Text = dCW.title;
                kryptonTextBox10.StateCommon.Content.Color1 = dCW.titleColor;
                Sheets_TitleButton.ForeColor = dCW.titleColor;

                //表題のフォント
                kryptonRibbonGroupComboBox_Font.Text = dCW.ftName;
                kryptonRibbonGroupComboBox_FontSize.Text = dCW.ftSize.ToString();
                kryptonRibbonColorButton_TextColor.SelectedColor = dCW.titleColor;


                if (dCW.titleBold == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Bold);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;

                    kryptonRibbonButton_Bold.Checked = true;
                }
                else
                {
                    kryptonRibbonButton_Bold.Checked = false;
                }

                if (dCW.titleItalic == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Italic);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonRibbonButton_Italic.Checked = true;
                }
                else
                {
                    kryptonRibbonButton_Italic.Checked = false;
                }

                if (dCW.titleUnderline == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Underline);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonContextMenuItem15.Checked = true;
                }
                if (dCW.titleUnderline == false)
                {
                    kryptonContextMenuItem15.Checked = false;
                }

                if (dCW.titleStrikeout == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Strikeout);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonContextMenuItem16.Checked = true;
                }
                if (dCW.titleStrikeout == false)
                {
                    kryptonContextMenuItem16.Checked = false;
                }

                //あいさつ文
                //月
                kryptonComboBox1.Text = dCW.UseSourouBunDate;
                //頭語
                kryptonComboBox2.Text = dCW.acronym;
                //候文
                kryptonComboBox11.Text = dCW.souroubun;
                //前文
                kryptonComboBox3.Text = dCW.PreviousText;
                //感謝のあいさつ
                kryptonComboBox4.Text = dCW.ThankYouGreeting;
                //結語
                kryptonComboBox5.Text = dCW.Conclusion;

                //内容
                kryptonTextBox12.Text = dCW.Content;
                kryptonTextBox13.Text = dCW.Notetaking;
            }
            // falseの場合は何もしない
        }

        public void RunWordInstalled()
        {
            if(Properties.Settings.Default.IsAvailableDocumentCreationSoftware == true)
            {
                kryptonCheckBox4.Checked = true;
                if(IsWordInstalled() != true)
                {
                    MessageBox.Show("お使いのコンピューターには文書作成ソフトウェアが使用できる状態ではなく、十分な動作が期待できない可能性があります。");
                }
            }
            else
            {
                kryptonCheckBox4.Checked = false;
            }
        }

        private bool IsWordInstalled()
        {
            const string wordRegistryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE";

            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(wordRegistryKey))
            {
                return key != null;
            }
        }

        public void SetSettings()
        {

            if (Properties.Settings.Default.ShowApplicationTask == 0)
            {
                kryptonRadioButton1.Checked = false;
                kryptonRadioButton2.Checked = false;
                kryptonRadioButton3.Checked = true;
            }
            else if (Properties.Settings.Default.ShowApplicationTask == 1)
            {
                kryptonRadioButton1.Checked = false;
                kryptonRadioButton2.Checked = true;
                kryptonRadioButton3.Checked = false;
            }
            else if (Properties.Settings.Default.ShowApplicationTask == 2)
            {
                kryptonRadioButton1.Checked = true;
                kryptonRadioButton2.Checked = false;
                kryptonRadioButton3.Checked = false;

            }

            if (Properties.Settings.Default.IsAvailableDocumentCreationSoftware == true)
            {
                kryptonCheckBox4.Checked = true;
            }
            else
            {
                kryptonCheckBox4.Checked = false;
            }

            if (Properties.Settings.Default.IsWindowsStartUpRunForDCMK == true)
            {
                kryptonCheckBox5.Checked = true;
            }
            else
            {
                kryptonCheckBox5.Checked = false;
            }

            if (Properties.Settings.Default.IsUseEraName == true)
            {
                kryptonCheckBox1.Checked = true;
                kryptonCheckBox7.Checked = true;
            }
            else
            {
                kryptonCheckBox1.Checked = false;
                kryptonCheckBox7.Checked = false;
            }
            kryptonNumericUpDown4.Value = Properties.Settings.Default.Space_Top;
            kryptonNumericUpDown7.Value = Properties.Settings.Default.Space_Buttom;
            kryptonNumericUpDown5.Value = Properties.Settings.Default.Space_Left;
            kryptonNumericUpDown6.Value = Properties.Settings.Default.Space_Right;

            kryptonTextBox16.Text = Properties.Settings.Default.SendingDepartment;
            kryptonTextBox17.Text = Properties.Settings.Default.To_CompanyOrOrganizationName;
            kryptonComboBox12.Text = Properties.Settings.Default.To_Title;
            kryptonTextBox18.Text = Properties.Settings.Default.To_Name;
            kryptonTextBox19.Text = Properties.Settings.Default.Caller_CompanyOrOrganizationName;
            kryptonTextBox32.Text = Properties.Settings.Default.Caller_Location;
            kryptonTextBox20.Text = Properties.Settings.Default.Caller_BuildingName;
            kryptonNumericUpDown3.Value = Properties.Settings.Default.Caller_FloorNumber;
            kryptonComboBox13.Text = Properties.Settings.Default.Caller_Title;
            kryptonTextBox21.Text = Properties.Settings.Default.Caller_Name;
            kryptonTextBox22.Text = Properties.Settings.Default.Caller_MailAddress_User;
            kryptonComboBox14.Text = Properties.Settings.Default.Caller_MailAddress_Domain;
            kryptonComboBox15.Text = Properties.Settings.Default.Caller_PhoneNumber1;
            kryptonTextBox23.Text = Properties.Settings.Default.Caller_PhoneNumber2;
            kryptonTextBox24.Text = Properties.Settings.Default.Caller_PhoneNumber3;
            kryptonComboBox16.Text = Properties.Settings.Default.Caller_FaxNumber1;
            kryptonTextBox26.Text = Properties.Settings.Default.Caller_FaxNumber2;
            kryptonTextBox25.Text = Properties.Settings.Default.Caller_FaxNumber3;

        }
        #region TreeNodeの追加;

        TreeNode miniTreeNode1 = new TreeNode();
        TreeNode ultraMiniNode1 = new TreeNode();
        TreeNode ultraMiniNode2 = new TreeNode();
        TreeNode ultraMiniNode3 = new TreeNode();
        TreeNode ultraMiniNode4 = new TreeNode();
        TreeNode ultraMiniNode5 = new TreeNode();
        TreeNode miniTreeNode2 = new TreeNode();
        TreeNode ultraMiniNode6 = new TreeNode();
        TreeNode ultraMiniNode7 = new TreeNode();
        TreeNode ultraMiniNode8 = new TreeNode();
        TreeNode ultraMiniNode9 = new TreeNode();
        TreeNode ultraMiniNode10 = new TreeNode();
        TreeNode ultraMiniNode11 = new TreeNode();
        TreeNode treeNode2 = new TreeNode();
        TreeNode miniTreeNode3 = new TreeNode();
        TreeNode ultraMiniNode12 = new TreeNode();
        TreeNode ultraMiniNode13 = new TreeNode();
        TreeNode ultraMiniNode14 = new TreeNode();
        TreeNode ultraMiniNode15 = new TreeNode();
        TreeNode ultraMiniNode16 = new TreeNode();
        TreeNode ultraMiniNode17 = new TreeNode();
        TreeNode miniTreeNode4 = new TreeNode();
        TreeNode ultraMiniNode18 = new TreeNode();
        TreeNode ultraMiniNode19 = new TreeNode();
        TreeNode ultraMiniNode20 = new TreeNode();
        TreeNode hyperTreeNode1 = new TreeNode();
        TreeNode ultraMiniNode21 = new TreeNode();
        TreeNode treeNode3 = new TreeNode();
        TreeNode miniTreeNode22 = new TreeNode();
        TreeNode miniTreeNode23 = new TreeNode();
        TreeNode miniTreeNode24 = new TreeNode();
        TreeNode miniTreeNode25 = new TreeNode();
        TreeNode miniTreeNode26 = new TreeNode();
        TreeNode miniTreeNode27 = new TreeNode();
        TreeNode miniTreeNode28 = new TreeNode();
        TreeNode treeNode4 = new TreeNode();
        TreeNode miniTreeNode29 = new TreeNode();
        TreeNode miniTreeNode30 = new TreeNode();
        TreeNode miniTreeNode31 = new TreeNode();
        TreeNode miniTreeNode32 = new TreeNode();
        TreeNode miniTreeNode33 = new TreeNode();
        TreeNode miniTreeNode34 = new TreeNode();

        public void AddTreeNodes()
        {
            //TreeViewに各種ノードを追加する
            //ノード1

            treeNode1.Text = "取引文書";
            treeView1.Nodes.Add(treeNode1);
            //子ノード1

            miniTreeNode1.Text = "通常取引";
            treeNode1.Nodes.Add(miniTreeNode1);
            //孫ノード1

            ultraMiniNode1.Text = "注文書";
            miniTreeNode1.Nodes.Add(ultraMiniNode1);
            //孫ノード2

            ultraMiniNode2.Text = "承諾書";
            miniTreeNode1.Nodes.Add(ultraMiniNode2);
            //孫ノード3

            ultraMiniNode3.Text = "依頼文";
            miniTreeNode1.Nodes.Add(ultraMiniNode3);
            //孫ノード4

            ultraMiniNode4.Text = "照会文";
            miniTreeNode1.Nodes.Add(ultraMiniNode4);
            //孫ノード5

            ultraMiniNode5.Text = "回答文";
            miniTreeNode1.Nodes.Add(ultraMiniNode5);
            //子ノード2

            miniTreeNode2.Text = "例外的取引";
            treeNode1.Nodes.Add(miniTreeNode2);
            //孫ノード6

            ultraMiniNode6.Text = "催促文";
            miniTreeNode2.Nodes.Add(ultraMiniNode6);
            //孫ノード7

            ultraMiniNode7.Text = "断り文";
            miniTreeNode2.Nodes.Add(ultraMiniNode7);
            //孫ノード8

            ultraMiniNode8.Text = "交渉文";
            miniTreeNode2.Nodes.Add(ultraMiniNode8);
            //孫ノード9

            ultraMiniNode9.Text = "抗議文";
            miniTreeNode2.Nodes.Add(ultraMiniNode9);
            //孫ノード10

            ultraMiniNode10.Text = "お詫び文";
            miniTreeNode2.Nodes.Add(ultraMiniNode10);
            //孫ノード11

            ultraMiniNode11.Text = "取り消し文";
            miniTreeNode2.Nodes.Add(ultraMiniNode11);
            //ノード2

            treeNode2.Text = "社公文書";
            treeView1.Nodes.Add(treeNode2);
            //子ノード3

            miniTreeNode3.Text = "公的";
            treeNode2.Nodes.Add(miniTreeNode3);
            //孫ノード12

            ultraMiniNode12.Text = "あいさつ文";
            miniTreeNode3.Nodes.Add(ultraMiniNode12);
            //孫ノード13

            ultraMiniNode13.Text = "お祝い文";
            miniTreeNode3.Nodes.Add(ultraMiniNode13);
            //孫ノード14

            ultraMiniNode14.Text = "招待文";
            miniTreeNode3.Nodes.Add(ultraMiniNode14);
            //孫ノード15

            ultraMiniNode15.Text = "お礼文";
            miniTreeNode3.Nodes.Add(ultraMiniNode15);
            //孫ノード16

            ultraMiniNode16.Text = "案内文";
            miniTreeNode3.Nodes.Add(ultraMiniNode16);
            //孫ノード17

            ultraMiniNode17.Text = "通知文";
            miniTreeNode3.Nodes.Add(ultraMiniNode17);
            //子ノード4

            miniTreeNode4.Text = "私的";
            treeNode2.Nodes.Add(miniTreeNode4);
            //孫ノード18

            ultraMiniNode18.Text = "年賀文";
            miniTreeNode4.Nodes.Add(ultraMiniNode18);
            //孫ノード19

            ultraMiniNode19.Text = "季節のあいさつ文";
            miniTreeNode4.Nodes.Add(ultraMiniNode19);
            //孫ノード20

            ultraMiniNode20.Text = "見舞い文";
            miniTreeNode4.Nodes.Add(ultraMiniNode20);
            //赤子ノード1

            hyperTreeNode1.Text = "個人宛見舞い文";
            ultraMiniNode20.Nodes.Add(hyperTreeNode1);
            //孫ノード21

            ultraMiniNode21.Text = "お悔やみ文";
            miniTreeNode4.Nodes.Add(ultraMiniNode21);

            //ノード3

            treeNode3.Text = "連絡文書";
            treeView2.Nodes.Add(treeNode3);
            //孫ノード22

            miniTreeNode22.Text = "通達文";
            treeNode3.Nodes.Add(miniTreeNode22);
            //孫ノード23

            miniTreeNode23.Text = "指示文";
            treeNode3.Nodes.Add(miniTreeNode23);
            //孫ノード24

            miniTreeNode24.Text = "依頼文";
            treeNode3.Nodes.Add(miniTreeNode24);
            //孫ノード25

            miniTreeNode25.Text = "照会文";
            treeNode3.Nodes.Add(miniTreeNode25);
            //孫ノード26

            miniTreeNode26.Text = "回答文";
            treeNode3.Nodes.Add(miniTreeNode26);
            //孫ノード27

            miniTreeNode27.Text = "通知文";
            treeNode3.Nodes.Add(miniTreeNode27);
            //孫ノード28

            miniTreeNode28.Text = "案内文";
            treeNode3.Nodes.Add(miniTreeNode28);
            //ノード4

            treeNode4.Text = "報告文書";
            treeView2.Nodes.Add(treeNode4);
            //孫ノード29

            miniTreeNode29.Text = "参加報告書";
            treeNode4.Nodes.Add(miniTreeNode29);
            //孫ノード30

            miniTreeNode30.Text = "出張報告書";
            treeNode4.Nodes.Add(miniTreeNode30);
            //孫ノード31

            miniTreeNode31.Text = "上申書";
            treeNode4.Nodes.Add(miniTreeNode31);
            //孫ノード32

            miniTreeNode32.Text = "届出文";
            treeNode4.Nodes.Add(miniTreeNode32);
            //孫ノード33

            miniTreeNode33.Text = "始末書";
            treeNode4.Nodes.Add(miniTreeNode33);
            //孫ノード33

            miniTreeNode34.Text = "理由書";
            treeNode4.Nodes.Add(miniTreeNode34);
            //後にTreeViewをすべて展開
            treeView1.ExpandAll();
            treeView2.ExpandAll();

            kryptonNavigator1.NavigatorMode = NavigatorMode.HeaderGroup;

        }

        #endregion

        #region シートの設定
        public void SetSheets()
        {
            //置換コントロールを消す
            kryptonPanel21.Height = 0;
            //カレンダーコントロールを今日に設定する
            kryptonDateTimePicker1.CalendarTodayDate = DateTime.Now;

            //シート内の設定
            kryptonTextBox7.Text = string.Empty;
            kryptonComboBox8.Text = string.Empty;

            kryptonComboBox6.Text = string.Empty;
            kryptonTextBox14.Text = string.Empty;
            kryptonTextBox8.Text = string.Empty;

            kryptonComboBox7.Text = string.Empty;
            kryptonTextBox9.Text = string.Empty;
            kryptonTextBox15.Text = string.Empty;

            kryptonComboBox2.Text = "拝啓";

            DateTime date = DateTime.Now;
            kryptonComboBox1.Text = date.Month.ToString();

            kryptonComboBox3.Text = "貴社ますますご盛栄のこととお慶び申し上げます。";


            kryptonComboBox4.Text = "平素は格別のご高配を賜り、厚く御礼申し上げます。";


            Sheets_TitleButton.Dock = DockStyle.Fill;

            InstalledFontCollection fonts = new InstalledFontCollection();
            FontFamily[] fontFamilies = fonts.Families;


            foreach (FontFamily font in fontFamilies)
            {
                kryptonRibbonGroupComboBox_Font.Items.Add(font.Name);
                kryptonRibbonGroupComboBox_Font.AutoCompleteCustomSource.Add(font.Name);
                kryptonRibbonGroupComboBox_NotepadFont.Items.Add(font.Name);
                kryptonRibbonGroupComboBox_NotepadFont.AutoCompleteCustomSource.Add(font.Name);


                kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();
            }

            //シート内の設定を変更
            kryptonCheckBox1.Checked = true;
            kryptonCheckButton2.Checked = false;

            Sheets_NumberLabel.Visible = false;
            Sheets_DateLabel.Visible = false;
            Sheets_AddressCompanyLabel.Visible = false;
            Sheets_AddressTitleAndNameLabel.Visible = false;
            Sheets_CallerCompanyLabel.Visible = false;
            Sheets_CallerLocationLabel.Visible = false;
            Sheets_BuildingNameLabel.Visible = false;
            Sheets_CallerTitleAndNameLabel.Visible = false;
            Sheets_CallerMallAddressLabel.Visible = false;
            Sheets_CallerTelLabel.Visible = false;
            Sheets_CallerFaxTelLabel.Visible = false;
            Sheets_TitleButton.Visible = false;
            Sheets_ContentLabel.Visible = false;
            Sheet_ConclusionLabel.Visible = false;

            panel4.Height = 221;

            panel2.Visible = true;
            panel3.Visible = true;
            kryptonTextBox1.Visible = true;
            panel11.Visible = true;
            kryptonTextBox3.Visible = true;
            kryptonTextBox4.Visible = true;
            panel6.Visible = true;
            panel10.Visible = true;
            panel9.Visible = true;
            kryptonTextBox8.Visible = true;
            kryptonTextBox9.Visible = true;
            kryptonTextBox10.Visible = true;
            panel5.Visible = true;
            panel7.Visible = true;
            panel8.Visible = true;

            //編集用シートの初期化
            //日付を今日に設定する
            if (kryptonCheckBox1.Checked == true)
            {
                DateTime date1 = kryptonDateTimePicker1.Value.Date;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                //下記のように西暦ではなく和暦として表示するように設定する
                culturejp.DateTimeFormat.Calendar = new JapaneseCalendar();
                Sheets_DateLabel.Text = date.ToString("ggy年M月d日", culturejp);
            }
            else
            {
                DateTime date1 = kryptonDateTimePicker1.Value.Date;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                Sheets_DateLabel.Text = date.ToString("yyyy年M月d日");
            }

        }
        #endregion

        public void SetDialog()
        {
            if (Properties.Settings.Default.ShowNotepadWarningPanel == true)
            {
                WarningPanel1.Visible = true;
            }
            else
            {
                WarningPanel1.Visible = false;
            }

        }

        //アプリの読み込み処理
        private void Form1_Load(object sender, EventArgs e)
        {
            //ノードの追加処理
            AddTreeNodes();
            //シートの設定処理
            SetSheets();
            //テーマの復元
            SetTheme();
            //QATの表示状態の復元
            SetQAT();
            //シートの空白間隔の復元
            SetSheetSpace();
            //シートのテキストの復元
            SetSheetText();
            //設定画面の復元
            SetSettings();
            //DCMKの起動タスク
            RunAppTask();
            //ダイアログ表示の復元    
            SetDialog();

            //キーボードショートカットの初期化
            kryptonRibbon.SelectedContext = string.Empty;
            kryptonRibbonButton_Paste.ShortcutKeys = Keys.Control | Keys.V;
            kryptonRibbonGroupButton_NotepadPaste.ShortcutKeys = Keys.None;

            kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

            Sheets_Sheet.Anchor = AnchorStyles.Top;

            // 親コントロールのサイズを取得
            int parentWidth = this.ClientSize.Width;
            int parentHeight = this.ClientSize.Height;

            // パネルのサイズを取得
            int panelWidth = Sheets_Sheet.Width;
            int panelHeight = Sheets_Sheet.Height;

            // パネルの位置を中央に設定
            Sheets_Sheet.Location = new System.Drawing.Point(
                (parentWidth - panelWidth) / 2 - 10,
                90
            );

            //シート
            kryptonRibbonButton_Bold.ShortcutKeys = Keys.Control | Keys.B;
            kryptonRibbonButton_Italic.ShortcutKeys = Keys.Control | Keys.I;
            kryptonContextMenuItem15.ShortcutKeys = Keys.Control | Keys.U;
            kryptonContextMenuItem16.ShortcutKeys = Keys.Control | Keys.T;

            kryptonRibbonColorButton_TextColor.ShortcutKeys = Keys.Control | Keys.Shift | Keys.C;

            kryptonRibbonGroupClusterButton4.ShortcutKeys = Keys.Control | Keys.Shift | Keys.U;
            kryptonRibbonGroupClusterButton5.ShortcutKeys = Keys.Control | Keys.Shift | Keys.D;

            //メモ
            kryptonRibbonGroupClusterButton1.ShortcutKeys = Keys.None;
            kryptonRibbonGroupClusterButton2.ShortcutKeys = Keys.None;
            kryptonContextMenuItem35.ShortcutKeys = Keys.None;
            kryptonContextMenuItem36.ShortcutKeys = Keys.None;

            kryptonRibbonGroupColorButton2.ShortcutKeys = Keys.None;
            kryptonRibbonGroupColorButton3.ShortcutKeys = Keys.None;

            kryptonRibbonGroupClusterButton6.ShortcutKeys = Keys.None;
            kryptonRibbonGroupClusterButton7.ShortcutKeys = Keys.None;

            //メモの内容を復元
            String str = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DocuQuick\SaveFile.rtf";

            if (File.Exists(str))
            {
                Notepads_kryptonRichTextBox_Notepad.LoadFile(str);
            }
            else
            {
                Notepads_kryptonRichTextBox_Notepad.Text = "(ここにメモしたい文字を入力します)";
            }
        }

        private void kryptonCommandLinkButton1_Click(object sender, EventArgs e)
        {

        }

        private void kryptonSplitContainer1_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }

        private void kryptonSplitContainer1_SplitterMoving(object sender, SplitterCancelEventArgs e)
        {

        }


        private void DatePage_Click(object sender, EventArgs e)
        {

        }

        private void WarningPanel_CloseButton_Click(object sender, EventArgs e)
        {
            WarningPanel1.Dispose();
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
        }

        private void kryptonLabel24_Click(object sender, EventArgs e)
        {

        }



        private void kryptonNavigator_Workbench_Selecting(object sender, ComponentFactory.Krypton.Navigator.KryptonPageCancelEventArgs e)
        {

        }

        private void kryptonRibbon_SelectedTabChanged(object sender, EventArgs e)
        {

        }

        private void kryptonNavigator_Workbench_SelectedPageChanged(object sender, EventArgs e)
        {

        }




        private void kryptonRibbonButton_Content_Click(object sender, EventArgs e)
        {

        }

        #region テーマの切り替え処理
        private void of2007_Click(object sender, EventArgs e)
        {
            //青
            if (kryptonContextMenuRadioButton1.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2007Blue";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(191, 219, 255);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(191, 219, 255);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007青色
                this.BackColor = Color.FromArgb(191, 219, 255);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(191, 219, 255);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;



                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //シルバー
            else if (kryptonContextMenuRadioButton2.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2007Silver";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(208, 212, 221);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(208, 212, 221);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007銀色
                this.BackColor = Color.FromArgb(208, 212, 221);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(208, 212, 221);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;


                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //ブラック
            else if (kryptonContextMenuRadioButton3.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2007Black";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(83, 83, 83);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(83, 83, 83);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007黒色
                this.BackColor = Color.FromArgb(83, 83, 83);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.Black;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.White;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(83, 83, 83);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;

            }
            of2007.Checked = true;
            of2010.Checked = false;
        }

        private void of2010_Click(object sender, EventArgs e)
        {
            //青
            if (kryptonContextMenuRadioButton1.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2010Blue";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(187, 206, 230);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(187, 206, 230);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010青色
                this.BackColor = Color.FromArgb(187, 206, 230);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(187, 206, 230);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;



                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;


                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //シルバー
            else if (kryptonContextMenuRadioButton2.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2010Silver";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(227,230,232);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(227, 230, 232);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010銀色
                this.BackColor = Color.FromArgb(227, 230, 232);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(227, 230, 232);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //ブラック
            else if (kryptonContextMenuRadioButton3.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2010Black";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(113, 113, 113);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(113, 113, 113);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010黒色
                this.BackColor = Color.FromArgb(113, 113, 113);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.Black;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.White;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(113, 113, 113);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.White;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;

            }


            of2007.Checked = false;
            of2010.Checked = true;
        }
        #endregion

        #region テーマカラーの切り替えとそれに伴う処理
        private void kryptonContextMenuRadioButton1_Click(object sender, EventArgs e)
        {
            //青
            if (of2007.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2007Blue";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(191,219,255);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(191, 219, 255);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007青色
                this.BackColor = Color.FromArgb(191, 219, 255);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(191, 219, 255);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;


                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //青
            else if (of2010.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2010Blue";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(187,206,230);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(187, 206, 230);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010青色
                this.BackColor = Color.FromArgb(187, 206, 230);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(187, 206, 230);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;



                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
        }

        private void kryptonContextMenuRadioButton2_Click(object sender, EventArgs e)
        {
            //シルバー
            if (of2007.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2007Silver";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(208,212,221);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(208, 212, 221);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007銀色
                this.BackColor = Color.FromArgb(208, 212, 221);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(208, 212, 221);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;


                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            else if (of2010.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2010Silver";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(227,230,232);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(227, 230, 232);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010銀色
                this.BackColor = Color.FromArgb(227, 230, 232);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(227, 230, 232);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;


                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;


                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
        }

        private void kryptonContextMenuRadioButton3_Click(object sender, EventArgs e)
        {
            //ブラック
            if (of2007.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2007Black";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(83,83,83);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(83, 83, 83);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007黒色
                this.BackColor = Color.FromArgb(83, 83, 83);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.Black;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.White;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(83, 83, 83);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;


                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;


                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;

            }
            //ブラック
            else if (of2010.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2010Black";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(113,113,113);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(113, 113, 113);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010黒色
                this.BackColor = Color.FromArgb(113, 113, 113);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.Black;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.White;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(113, 113, 113);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.White;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.White;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;

            }
        }
        #endregion


        private void kryptonContextMenu1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void RibbonAppButtonContextMenu_AboutApp_Click(object sender, EventArgs e)
        {
            AboutBox about = new AboutBox();

            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                about.BackColor = Color.FromArgb(191, 219, 255);
                about.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                about.BackColor = Color.FromArgb(208, 212, 221);
                about.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                about.BackColor = Color.FromArgb(83, 83, 83);
                about.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                about.BackColor = Color.FromArgb(187, 206, 230);
                about.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                about.BackColor = Color.FromArgb(227, 230, 232);
                about.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                about.BackColor = Color.FromArgb(113, 113, 113);
                about.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
            about.ShowDialog();
        }

        private void buttonSpecAppMenu1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void kryptonContextMenuItem13_Click(object sender, EventArgs e)
        {
            ThirdParty thirdParty = new ThirdParty();

            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                thirdParty.BackColor = Color.FromArgb(191, 219, 255);
                thirdParty.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                thirdParty.BackColor = Color.FromArgb(208, 212, 221);
                thirdParty.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                thirdParty.BackColor = Color.FromArgb(83, 83, 83);
                thirdParty.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                thirdParty.BackColor = Color.FromArgb(187, 206, 230);
                thirdParty.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                thirdParty.BackColor = Color.FromArgb(227, 230, 232);
                thirdParty.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                thirdParty.BackColor = Color.FromArgb(113, 113, 113);
                thirdParty.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            thirdParty.ShowDialog();
        }





        private void fontGroup_DialogBoxLauncherClick(object sender, EventArgs e)
        {
            FontDialog fd = new FontDialog();
            fd.Font = Sheets_TitleButton.Font;
            fd.ShowColor = true;
            fd.Color = Sheets_TitleButton.ForeColor;
            kryptonRibbonColorButton_TextColor.SelectedColor = Sheets_TitleButton.ForeColor;

            if (fd.ShowDialog() == DialogResult.OK)
            {
                Sheets_TitleButton.Font = fd.Font;
                Sheets_TitleButton.ForeColor = fd.Color;
                kryptonRibbonColorButton_TextColor.SelectedColor = fd.Color;

                kryptonTextBox10.StateCommon.Content.Font = fd.Font;
                kryptonTextBox10.StateCommon.Content.Color1 = fd.Color;

                kryptonRibbonGroupComboBox_Font.Text = fd.Font.Name;
                kryptonRibbonGroupComboBox_FontSize.Text = fd.Font.Size.ToString();

                if (kryptonTextBox10.StateCommon.Content.Font.Style == FontStyle.Bold)
                {
                    kryptonRibbonButton_Bold.Checked = true;
                }
                else
                {
                    kryptonRibbonButton_Bold.Checked = false;
                }

                if (kryptonTextBox10.StateCommon.Content.Font.Style == FontStyle.Italic)
                {
                    kryptonRibbonButton_Italic.Checked = true;
                }
                else
                {
                    kryptonRibbonButton_Italic.Checked = false;
                }

                if (kryptonTextBox10.StateCommon.Content.Font.Style == FontStyle.Underline)
                {
                    kryptonContextMenuItem15.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem15.Checked = false;
                }

                if (kryptonTextBox10.StateCommon.Content.Font.Style == FontStyle.Strikeout)
                {
                    kryptonContextMenuItem16.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem16.Checked = false;
                }
            }
        }

        public void fd_ShowHelpReqest(Object sender, EventArgs e)
        {

        }

        public void fd_ShowHelpReqest2(Object sender, EventArgs e)
        {

        }

        private void kryptonRibbonGroupButton_NotepadFonts_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonColorButton_TextColor_SelectedColorChanged(object sender, ComponentFactory.Krypton.Toolkit.ColorEventArgs e)
        {
            Sheets_TitleButton.ForeColor = kryptonRibbonColorButton_TextColor.SelectedColor;
            kryptonTextBox10.StateCommon.Content.Color1 = kryptonRibbonColorButton_TextColor.SelectedColor;
        }


        private void Notepads_kryptonRichTextBox_Notepad_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void Notepads_kryptonRichTextBox_Notepad_TextChanged(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_Font_TextUpdate(object sender, EventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    kryptonRibbonGroupComboBox_Font.Text,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void Sheets_TitleLabel_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_Font_SelectionChangeCommitted(object sender, EventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    kryptonRibbonGroupComboBox_Font.Text,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void kryptonRibbonGroupComboBox_Font_DropDownClosed(object sender, EventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    kryptonRibbonGroupComboBox_Font.Text,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void kryptonRibbonGroupComboBox_Font_SelectedValueChanged(object sender, EventArgs e)
        {
        }

        private void kryptonContextMenuItem1_CheckStateChanged(object sender, EventArgs e)
        {

        }


        private void kryptonContextMenuItem1_Click(object sender, EventArgs e)
        {

            DCW dCW = new DCW();

            Properties.Settings.Default.dCW_TopSpace = Sheets_TopPanel.Height;
            Properties.Settings.Default.dCW_ButtomSpace = Sheets_ButtomPanel.Height;
            Properties.Settings.Default.dCW_LeftSpace = Sheets_LeftPanel.Width;
            Properties.Settings.Default.dCW_RightSpace = Sheets_RightPanel.Width;
            Properties.Settings.Default.Save();


            // Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                dCW.BackColor = Color.FromArgb(191, 219, 255);
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                dCW.BackColor = Color.FromArgb(208, 212, 221);
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                dCW.BackColor = Color.FromArgb(83, 83, 83);
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                dCW.BackColor = Color.FromArgb(187, 206, 230);
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                dCW.BackColor = Color.FromArgb(227, 230, 232);
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                dCW.BackColor = Color.FromArgb(113, 113, 113);
            }

            dCW.ShowDialog();

            //trueの場合のみ実行
            //ウィザードの「完了」ボタンをクリックしたときに実行
            if (dCW.IsWizardFinished == true)
            {
                FontReset();
                //発行番号
                if (dCW.NoIssueNumber == false)
                {
                    kryptonCheckBox3.Checked = false;
                    kryptonTextBox11.Text = dCW.IssueNumber_Publisher;
                    kryptonNumericUpDown1.Value = dCW.IssueNumber;
                }
                else
                {
                    kryptonCheckBox3.Checked = true;
                }

                //日付
                if (dCW.NoDate == false)
                {
                    kryptonCheckBox2.Checked = false;
                    kryptonDateTimePicker1.Value = dCW.Date;
                    if(dCW.UseEraName == true)
                    {
                        kryptonCheckBox1.Checked = true;
                    }
                    else
                    {
                        kryptonCheckBox1.Checked = false;
                    }
                }
                else
                {
                    kryptonCheckBox2.Checked = true;
                }

                //発信者
                kryptonTextBox1.Text = dCW.AdCompany;
                kryptonComboBox10.Text = dCW.AdTitle;
                kryptonTextBox2.Text = dCW.AdName;

                kryptonTextBox3.Text = dCW.CaCampany;
                kryptonTextBox4.Text = dCW.CaLocation;
                kryptonTextBox5.Text = dCW.CaBuildingName;
                kryptonNumericUpDown2.Value = dCW.CaFloorNumber;
                kryptonComboBox9.Text = dCW.CaTitle;
                kryptonTextBox6.Text = dCW.CaName;
                kryptonTextBox7.Text = dCW.CaMailAddress;
                kryptonComboBox8.Text = dCW.CaMailAddress_Domain;
                //電話番号
                kryptonComboBox6.Text = dCW.CaPhoneNumber1;
                kryptonTextBox14.Text = dCW.CaPhoneNumber2;
                kryptonTextBox8.Text = dCW.CaPhoneNumber3;
                kryptonComboBox7.Text = dCW.CaFaxNumber1;
                kryptonTextBox9.Text = dCW.CaFaxNumber2;
                kryptonTextBox15.Text = dCW.CaFaxNumber3;

                //表題
                kryptonTextBox10.Text = dCW.title;
                kryptonTextBox10.StateCommon.Content.Color1 = dCW.titleColor;
                Sheets_TitleButton.ForeColor = dCW.titleColor;

                //表題のフォント
                kryptonRibbonGroupComboBox_Font.Text = dCW.ftName;
                kryptonRibbonGroupComboBox_FontSize.Text = dCW.ftSize.ToString();
                kryptonRibbonColorButton_TextColor.SelectedColor = dCW.titleColor;


                if (dCW.titleBold == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style|FontStyle.Bold);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                   
                    kryptonRibbonButton_Bold.Checked = true;
                }
                else
                {
                    kryptonRibbonButton_Bold.Checked = false;
                }

                if (dCW.titleItalic == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Italic);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonRibbonButton_Italic.Checked = true;
                }
                else
                {
                    kryptonRibbonButton_Italic.Checked = false;
                }

                if(dCW.titleUnderline == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Underline);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonContextMenuItem15.Checked = true;
                }
                if (dCW.titleUnderline == false)
                {
                    kryptonContextMenuItem15.Checked = false;
                }

                if (dCW.titleStrikeout == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Strikeout);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonContextMenuItem16.Checked = true;
                }
                if (dCW.titleStrikeout == false)
                {
                    kryptonContextMenuItem16.Checked = false;
                }

                //あいさつ文
                //月
                kryptonComboBox1.Text = dCW.UseSourouBunDate;
                //頭語
                kryptonComboBox2.Text = dCW.acronym;
                //候文
                kryptonComboBox11.Text = dCW.souroubun;
                //前文
                kryptonComboBox3.Text = dCW.PreviousText;
                //感謝のあいさつ
                kryptonComboBox4.Text = dCW.ThankYouGreeting;
                //結語
                kryptonComboBox5.Text = dCW.Conclusion;

                //内容
                kryptonTextBox12.Text = dCW.Content;
                kryptonTextBox13.Text = dCW.Notetaking;
            }
            // falseの場合は何もしない
        }

        private void kryptonRibbonGroupComboBox_Font_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    kryptonRibbonGroupComboBox_Font.Text,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void kryptonRibbonGroupComboBox_Font_KeyDown(object sender, KeyEventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    kryptonRibbonGroupComboBox_Font.Text,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void kryptonRibbonGroupButton1_Click(object sender, EventArgs e)
        {

        }

        private void Editpanel_NoEditCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void kryptonContextMenuItem12_Click(object sender, EventArgs e)
        {

        }



        UseSoftWareWindow useSoftWareWindow = new UseSoftWareWindow();
        KeboradShortCut keboradShortCut = new KeboradShortCut();
        public void Form1_Activated(object sender, EventArgs e)
        {

            if (useSoftWareWindow.Visible == true)
            {
                kryptonRibbonGroupButton_Tutorial.Enabled = false;
                kryptonContextMenuItem12.Enabled = false;
            }
            else
            {
                kryptonRibbonGroupButton_Tutorial.Enabled = true;
                kryptonContextMenuItem12.Enabled = true;
            }

            if (keboradShortCut.Visible == true)
            {
                buttonSpecAppMenu2.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.False;
            }
            else
            {
                buttonSpecAppMenu2.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.True;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //不要なオブジェクトを破棄してから終了する
            useSoftWareWindow.Dispose();
            keboradShortCut.Dispose();

            //ファイル保存
            AutoSave();

            //QAT状態確認・保存
            if(kryptonRibbonQATButton1.Visible == true)
            {
                Properties.Settings.Default.QAT1_Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT1_Visible = false;
            }

            if (kryptonRibbonQATButton2.Visible == true)
            {
                Properties.Settings.Default.QAT2_Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT2_Visible = false;
            }

            if (kryptonRibbonQATButton3.Visible == true)
            {
                Properties.Settings.Default.QAT3_Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT3_Visible = false;
            }

            if (kryptonRibbonQATButton4.Visible == true)
            {
                Properties.Settings.Default.QAT4_Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT4_Visible = false;
            }

            if (kryptonRibbonQATButton5.Visible == true)
            {
                Properties.Settings.Default.QAT5_Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT5_Visible = false;
            }

            if (kryptonRibbonQATButton6.Visible == true)
            {
                Properties.Settings.Default.QAT6_Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT6_Visible = false;
            }

            if (kryptonRibbonQATButton7.Visible == true)
            {
                Properties.Settings.Default.QAT7_Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT7_Visible = false;
            }

            if (kryptonRibbonQATButton8.Visible == true)
            {
                Properties.Settings.Default.QAT8_Visible = true;
            }
            else
            {
                Properties.Settings.Default.QAT8_Visible = false;
            }

            Properties.Settings.Default.Save();

        }

        #region リボンコントロールのホーム「文書作成ソフトウェアで編集」をクリックしたときの処理

        private void SetWordRangeColor(Range range, Color color)
        {
            // Word の RGB 値は Red + (Green << 8) + (Blue << 16)
            int rgb = color.R | (color.G << 8) | (color.B << 16);
            range.Font.Color = (WdColor)rgb;
        }


        private void kryptonRibbonButton_OpenWSoft_Click(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "出力中...";
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = true;
            Document doc = word.Documents.Add();


            //外枠の余白を設定
            doc.PageSetup.TopMargin = Sheets_TopPanel.Height;
            doc.PageSetup.BottomMargin = Sheets_ButtomPanel.Height;
            doc.PageSetup.LeftMargin = Sheets_LeftPanel.Width;
            doc.PageSetup.RightMargin = Sheets_RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }


            //発行番号
            if (Sheets_NumberLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                paragraph1.Range.Text = Sheets_NumberLabel.Text;
                paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph1.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph1.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph1.Range.InsertParagraphAfter();
            }
            //日付
            if (Sheets_DateLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                paragraph2.Range.Text = Sheets_DateLabel.Text;
                paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph2.Range.InsertParagraphAfter();
            }
            //相手先会社名
            if (Sheets_AddressCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph3 = doc.Paragraphs.Add();
                paragraph3.Range.Text = Sheets_AddressCompanyLabel.Text;
                paragraph3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph3.Range.InsertParagraphAfter();
            }
            //相手先氏名
            if (Sheets_AddressTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                paragraph4.Range.Text = Sheets_AddressTitleAndNameLabel.Text;
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph4.Range.InsertParagraphAfter();
            }
            //発信者会社名
            if (Sheets_CallerCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph5 = doc.Paragraphs.Add();
                paragraph5.Range.Text = Sheets_CallerCompanyLabel.Text;
                paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph5.Range.InsertParagraphAfter();
            }
            //発信者所在地
            if (Sheets_CallerLocationLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph6 = doc.Paragraphs.Add();
                paragraph6.Range.Text = Sheets_CallerLocationLabel.Text;
                paragraph6.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph6.Range.InsertParagraphAfter();
            }
            //発信者建物名と階数
            if (Sheets_BuildingNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph7 = doc.Paragraphs.Add();
                paragraph7.Range.Text = Sheets_BuildingNameLabel.Text;
                paragraph7.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph7.Range.InsertParagraphAfter();
            }
            //発信者氏名
            if (Sheets_CallerTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                paragraph8.Range.Text = Sheets_CallerTitleAndNameLabel.Text;
                paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph8.Range.InsertParagraphAfter();
            }
            //メールアドレス
            if (Sheets_CallerMallAddressLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph9 = doc.Paragraphs.Add();
                paragraph9.Range.Text = "メールアドレス:" + Sheets_CallerMallAddressLabel.Text;
                paragraph9.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph9.Range.InsertParagraphAfter();
            }
            //電話番号
            if (Sheets_CallerTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph10 = doc.Paragraphs.Add();
                paragraph10.Range.Text = "電話番号:" + Sheets_CallerTelLabel.Text;
                paragraph10.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph10.Range.InsertParagraphAfter();
            }
            //Fax番号
            if (Sheets_CallerFaxTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph11 = doc.Paragraphs.Add();
                paragraph11.Range.Text = "Fax番号:" + Sheets_CallerFaxTelLabel.Text;
                paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph11.Range.InsertParagraphAfter();
            }
            //表題
            if (Sheets_TitleButton.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph12 = doc.Paragraphs.Add();
                if(kryptonRibbonButton_Bold.Checked == true)
                {
                    paragraph12.Range.Bold = 1;
                }
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    paragraph12.Range.Italic = 1;
                }
                if (kryptonContextMenuItem15.Checked == true)
                {
                    paragraph12.Range.Underline = WdUnderline.wdUnderlineSingle;
                }
                if (kryptonContextMenuItem16.Checked == true)
                {
                    paragraph12.Range.Font.StrikeThrough = 1;
                }
                paragraph12.Range.Font.Name = Sheets_TitleButton.Font.Name;
                paragraph12.Range.Font.Size = (int)kryptonTextBox10.StateCommon.Content.Font.Size;
                paragraph12.Range.Text = Sheets_TitleButton.Text;
                SetWordRangeColor(paragraph12.Range, Sheets_TitleButton.ForeColor);
                paragraph12.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph12.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph12.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph12.Range.InsertParagraphAfter();

            }
            //あいさつ文
            if (Sheets_ContentLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph13 = doc.Paragraphs.Add();
                paragraph13.Range.Bold = 0;
                paragraph13.Range.Italic = 0;
                paragraph13.Range.Underline = WdUnderline.wdUnderlineNone;
                paragraph13.Range.Font.StrikeThrough = 0;
                paragraph13.Range.Font.Name = "游明朝";
                paragraph13.Range.Font.Size = 10;
                paragraph13.Range.Text = Sheets_ContentLabel.Text;
                paragraph13.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph13.Range.Font.Color = WdColor.wdColorBlack;
                paragraph13.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph13.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph13.Range.InsertParagraphAfter();
            }
            //内容
            try
            {
                int LinesCount = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox12.Lines[LinesCount];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount = LinesCount + 1;
                    if (LinesCount == kryptonTextBox12.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //結語
            //kryptonRibbonGroupCheckBoxにチェックがない場合に動作する
            if (kryptonRibbonGroupCheckBox1.Checked != true)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph14 = doc.Paragraphs.Add();
                paragraph14.Range.Text = Sheet_ConclusionLabel.Text;
                paragraph14.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph14.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph14.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph14.Range.InsertParagraphAfter();
            }
            kryptonLabel1.Text = "出力完了";
            stausUpdate();
            //記
            if (Sheet_NoteLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph15 = doc.Paragraphs.Add();
                paragraph15.Range.Text = Sheet_NoteLabel.Text;
                paragraph15.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph15.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph15.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph15.Range.InsertParagraphAfter();
            }
            //記し書き
            try
            {
                int LinesCount2 = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox13.Lines[LinesCount2];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount2 = LinesCount2 + 1;
                    if (LinesCount2 == kryptonTextBox13.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //以上
            if (Sheets_EndLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph16 = doc.Paragraphs.Add();
                paragraph16.Range.Text = Sheets_EndLabel.Text;
                paragraph16.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph16.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph16.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph16.Range.InsertParagraphAfter();
            }

            GC.Collect();





        }
        #endregion

        #region リボンコントロールのホーム「Docx  形式で保存」をクリックしたときの処理
        private void kryptonRibbonGroupButton10_Click(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "出力中...";
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            //バックグラウンド上でWordを起動する
            word.Visible = false;
            Document doc = word.Documents.Add();


            //外枠の余白を設定
            doc.PageSetup.TopMargin = Sheets_TopPanel.Height;
            doc.PageSetup.BottomMargin = Sheets_ButtomPanel.Height;
            doc.PageSetup.LeftMargin = Sheets_LeftPanel.Width;
            doc.PageSetup.RightMargin = Sheets_RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //発行番号
            if (Sheets_NumberLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                paragraph1.Range.Text = Sheets_NumberLabel.Text;
                paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph1.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph1.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph1.Range.InsertParagraphAfter();
            }
            //日付
            if (Sheets_DateLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                paragraph2.Range.Text = Sheets_DateLabel.Text;
                paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph2.Range.InsertParagraphAfter();
            }
            //相手先会社名
            if (Sheets_AddressCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph3 = doc.Paragraphs.Add();
                paragraph3.Range.Text = Sheets_AddressCompanyLabel.Text;
                paragraph3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph3.Range.InsertParagraphAfter();
            }
            //相手先氏名
            if (Sheets_AddressTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                paragraph4.Range.Text = Sheets_AddressTitleAndNameLabel.Text;
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph4.Range.InsertParagraphAfter();
            }
            //発信者会社名
            if (Sheets_CallerCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph5 = doc.Paragraphs.Add();
                paragraph5.Range.Text = Sheets_CallerCompanyLabel.Text;
                paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph5.Range.InsertParagraphAfter();
            }
            //発信者所在地
            if (Sheets_CallerLocationLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph6 = doc.Paragraphs.Add();
                paragraph6.Range.Text = Sheets_CallerLocationLabel.Text;
                paragraph6.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph6.Range.InsertParagraphAfter();
            }
            //発信者建物名と階数
            if (Sheets_BuildingNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph7 = doc.Paragraphs.Add();
                paragraph7.Range.Text = Sheets_BuildingNameLabel.Text;
                paragraph7.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph7.Range.InsertParagraphAfter();
            }
            //発信者氏名
            if (Sheets_CallerTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                paragraph8.Range.Text = Sheets_CallerTitleAndNameLabel.Text;
                paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph8.Range.InsertParagraphAfter();
            }
            //メールアドレス
            if (Sheets_CallerMallAddressLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph9 = doc.Paragraphs.Add();
                paragraph9.Range.Text = "メールアドレス:" + Sheets_CallerMallAddressLabel.Text;
                paragraph9.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph9.Range.InsertParagraphAfter();
            }
            //電話番号
            if (Sheets_CallerTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph10 = doc.Paragraphs.Add();
                paragraph10.Range.Text = "電話番号:" + Sheets_CallerTelLabel.Text;
                paragraph10.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph10.Range.InsertParagraphAfter();
            }
            //Fax番号
            if (Sheets_CallerFaxTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph11 = doc.Paragraphs.Add();
                paragraph11.Range.Text = "Fax番号:" + Sheets_CallerFaxTelLabel.Text;
                paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph11.Range.InsertParagraphAfter();
            }
            //表題
            if (Sheets_TitleButton.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph12 = doc.Paragraphs.Add();
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    paragraph12.Range.Bold = 1;
                }
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    paragraph12.Range.Italic = 1;
                }
                if (kryptonContextMenuItem15.Checked == true)
                {
                    paragraph12.Range.Underline = WdUnderline.wdUnderlineSingle;
                }
                if (kryptonContextMenuItem16.Checked == true)
                {
                    paragraph12.Range.Font.StrikeThrough = 1;
                }
                paragraph12.Range.Font.Size = (int)kryptonTextBox10.StateCommon.Content.Font.Size;
                paragraph12.Range.Text = Sheets_TitleButton.Text;
                paragraph12.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                SetWordRangeColor(paragraph12.Range, Sheets_TitleButton.ForeColor);
                paragraph12.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph12.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph12.Range.InsertParagraphAfter();

            }
            //あいさつ文
            if (Sheets_ContentLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph13 = doc.Paragraphs.Add();
                paragraph13.Range.Bold = 0;
                paragraph13.Range.Italic = 0;
                paragraph13.Range.Underline = WdUnderline.wdUnderlineNone;
                paragraph13.Range.Font.StrikeThrough = 0;
                paragraph13.Range.Font.Size = 10;
                paragraph13.Range.Text = Sheets_ContentLabel.Text;
                paragraph13.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph13.Range.Font.Color = WdColor.wdColorBlack;
                paragraph13.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph13.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph13.Range.InsertParagraphAfter();
            }
            //内容
            try
            {
                int LinesCount = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox12.Lines[LinesCount];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount = LinesCount + 1;
                    if (LinesCount == kryptonTextBox12.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //結語
            //kryptonRibbonGroupCheckBoxにチェックがない場合に動作する
            if (kryptonRibbonGroupCheckBox1.Checked != true)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph14 = doc.Paragraphs.Add();
                paragraph14.Range.Text = Sheet_ConclusionLabel.Text;
                paragraph14.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph14.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph14.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph14.Range.InsertParagraphAfter();
            }
            kryptonLabel1.Text = "出力完了";
            stausUpdate();
            //記
            if (Sheet_NoteLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph15 = doc.Paragraphs.Add();
                paragraph15.Range.Text = Sheet_NoteLabel.Text;
                paragraph15.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph15.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph15.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph15.Range.InsertParagraphAfter();
            }
            //記し書き
            try
            {
                int LinesCount2 = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox13.Lines[LinesCount2];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount2 = LinesCount2 + 1;
                    if (LinesCount2 == kryptonTextBox13.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //以上
            if (Sheets_EndLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph16 = doc.Paragraphs.Add();
                paragraph16.Range.Text = Sheets_EndLabel.Text;
                paragraph16.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph16.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph16.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph16.Range.InsertParagraphAfter();
            }

            //保存処理
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "ドキュメントファイルを保存する場所を選択";
            sd.Filter = "Word 文書 (*.docx)|*.docx";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    doc.SaveAs2(sd.FileName);
                    MessageBox.Show("ファイルが以下の場所に正しく保存されました。\r\n" + sd.FileName, "ファイル保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("ファイルが正しく保存されませんでした。保存するファイルの場所が適切か文書作成ソフトウェアがインストールされているか確認してください。\r\n\r\nエラー内容:\r\n" + ex.Message, "ファイル保存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

            //保存を確認せず閉じる
            try
            {
                doc.Close(false);
                word.Quit();
            }
            catch { }


            GC.Collect();
        }
        #endregion

        //連携完了後の処理
        async System.Threading.Tasks.Task stausUpdate()
        {
            await System.Threading.Tasks.Task.Delay(5000);
            kryptonLabel1.Text = "準備完了";
        }

        public void Timer_Tick(object sender, EventArgs e)
        {




        }

        #region 表示モード切り替え処理
        private void kryptonCheckButton1_Click(object sender, EventArgs e)
        {
            kryptonRibbonGroupButton_ViewMode1.Checked = true;
            kryptonRibbonGroupButton_ViewMode2.Checked = false;

            kryptonCheckButton1.Checked = true;
            kryptonCheckButton2.Checked = false;

            Sheets_NumberLabel.Visible = false;
            Sheets_DateLabel.Visible = false;
            Sheets_AddressCompanyLabel.Visible = false;
            Sheets_AddressTitleAndNameLabel.Visible = false;
            Sheets_CallerCompanyLabel.Visible = false;
            Sheets_CallerLocationLabel.Visible = false;
            Sheets_BuildingNameLabel.Visible = false;
            Sheets_CallerTitleAndNameLabel.Visible = false;
            Sheets_CallerMallAddressLabel.Visible = false;
            Sheets_CallerTelLabel.Visible = false;
            Sheets_CallerFaxTelLabel.Visible = false;
            Sheets_TitleButton.Visible = false;
            Sheets_ContentLabel.Visible = false;
            Sheet_ConclusionLabel.Visible = false;

            panel4.Height = 221;

            panel2.Visible = true;
            panel3.Visible = true;
            kryptonTextBox1.Visible = true;
            panel11.Visible = true;
            kryptonTextBox3.Visible = true;
            kryptonTextBox4.Visible = true;
            panel6.Visible = true;
            panel10.Visible = true;
            panel9.Visible = true;
            kryptonTextBox8.Visible = true;
            kryptonTextBox9.Visible = true;
            kryptonTextBox10.Visible = true;
            panel5.Visible = true;
            panel7.Visible = true;
            panel8.Visible = true;
            label9.Visible = true;

            kryptonComboBox5.Visible = true;

            label11.Visible = true;
            label12.Visible = true;

            Sheets_NumberPanel.Visible = true;
            Sheets_DatePanel.Visible = true;
            Sheets_AddressCompanyPanel.Visible = true;
            Sheets_AddressTitleAndNamePanel.Visible = true;
            Sheets_CallerCompanyPanel.Visible = true;
            Sheets_CallerLocationPanel.Visible = true;
            Sheets_BuildingNamePanel.Visible = true;
            Sheets_CallerTitleAndNamePanel.Visible = true;
            Sheets_CallerMallAddressPanel.Visible = true;
            Sheets_CallerTelPanel.Visible = true;
            Sheets_CallerFaxTelPanel.Visible = true;
        }

        private void panel5_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonCheckButton2_Click(object sender, EventArgs e)
        {
            kryptonRibbonGroupButton_ViewMode1.Checked = false;
            kryptonRibbonGroupButton_ViewMode2.Checked = true;

            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = true;

            Sheets_NumberLabel.Visible = true;
            Sheets_DateLabel.Visible = true;
            Sheets_AddressCompanyLabel.Visible = true;
            Sheets_AddressTitleAndNameLabel.Visible = true;
            Sheets_CallerCompanyLabel.Visible = true;
            Sheets_CallerLocationLabel.Visible = true;
            Sheets_BuildingNameLabel.Visible = true;
            Sheets_CallerTitleAndNameLabel.Visible = true;
            Sheets_CallerMallAddressLabel.Visible = true;
            Sheets_CallerTelLabel.Visible = true;
            Sheets_CallerFaxTelLabel.Visible = true;
            Sheets_TitleButton.Visible = true;
            Sheets_ContentLabel.Visible = true;
            Sheet_ConclusionLabel.Visible = true;

            panel4.Height = 221;

            panel2.Visible = false;
            panel3.Visible = false;
            kryptonTextBox1.Visible = false;
            panel11.Visible = false;
            kryptonTextBox3.Visible = false;
            kryptonTextBox4.Visible = false;
            panel6.Visible = false;
            panel10.Visible = false;
            panel9.Visible = false;
            kryptonTextBox8.Visible = false;
            kryptonTextBox9.Visible = false;
            kryptonTextBox10.Visible = false;
            panel5.Visible = false;
            panel7.Visible = false;
            panel8.Visible = false;

            kryptonComboBox5.Visible = false;

            //Number
            if (kryptonCheckBox3.Checked == true)
            {
                Sheets_NumberPanel.Visible = false;
            }
            else
            {
                Sheets_NumberPanel.Visible = true;
            }

            //Date
            if (kryptonCheckBox2.Checked == true)
            {
                Sheets_DatePanel.Visible = false;
            }
            else
            {
                Sheets_DatePanel.Visible = true;
            }

            //AdCompany
            if (Sheets_AddressCompanyLabel.Text == string.Empty)
            {
                Sheets_AddressCompanyPanel.Visible = false;
            }
            else
            {
                Sheets_AddressCompanyPanel.Visible = true;
            }

            //AdName
            if (Sheets_AddressTitleAndNameLabel.Text == string.Empty)
            {
                Sheets_AddressTitleAndNamePanel.Visible = false;
            }
            else
            {
                Sheets_AddressTitleAndNamePanel.Visible = true;
            }

            //Company
            if (Sheets_CallerCompanyLabel.Text == string.Empty)
            {
                Sheets_CallerCompanyPanel.Visible = false;
            }
            else
            {
                Sheets_CallerCompanyPanel.Visible = true;
            }

            //Location
            if (Sheets_CallerLocationLabel.Text == string.Empty)
            {
                Sheets_CallerLocationPanel.Visible = false;
            }
            else
            {
                Sheets_CallerLocationPanel.Visible = true;
            }

            //Buiding Name
            if (Sheets_BuildingNameLabel.Text == string.Empty)
            {
                Sheets_BuildingNamePanel.Visible = false;
            }
            else
            {
                Sheets_BuildingNamePanel.Visible = true;
            }

            //Name
            if (Sheets_CallerTitleAndNameLabel.Text == string.Empty)
            {
                Sheets_CallerTitleAndNamePanel.Visible = false;
            }
            else
            {
                Sheets_CallerTitleAndNamePanel.Visible = true;
            }

            //Mail
            if (label11.Font.Strikeout == true)
            {
                label11.Visible = false;
                label12.Visible = false;
                Sheets_CallerMallAddressPanel.Visible = false;
            }
            else
            {
                label11.Visible = true;
                label12.Visible = true;
                Sheets_CallerMallAddressPanel.Visible = true;
            }

            //Tel
            if (label9.Font.Strikeout == true)
            {
                label9.Visible = false;
                Sheets_CallerTelPanel.Visible = false;
            }
            else
            {
                label9.Visible = true;
                Sheets_CallerTelPanel.Visible = true;
            }

            if (label10.Font.Strikeout == true)
            {
                label10.Visible = true;
                Sheets_CallerFaxTelPanel.Visible = false;
            }
            else
            {
                label10.Visible = true;
                Sheets_CallerFaxTelPanel.Visible = true;
            }
        }
        #endregion

        #region キーによる編集項目切り替え処理
        private void kryptonTextBox11_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonLabel7.Text = "シート内の項目を移動するにはFunction+Endキーを押してください。Function+Homeキーを押すと前の項目に戻ります。";
                kryptonNumericUpDown1.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonLabel7.Text = "シート内の項目を移動するにはFunction+Endキーを押してください。Function+Homeキーを押すと前の項目に戻ります。";
                kryptonComboBox5.Focus();
            }
        }

        private void kryptonNumericUpDown1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonDateTimePicker1.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox11.Focus();
            }
        }

        private void kryptonDateTimePicker1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox1.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonNumericUpDown1.Focus();
            }
        }

        private void kryptonTextBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox10.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonDateTimePicker1.Focus();
            }
        }

        private void kryptonComboBox10_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox2.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox1.Focus();
            }
        }

        private void kryptonTextBox2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox3.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox10.Focus();
            }
        }

        private void kryptonTextBox3_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox4.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox2.Focus();
            }
        }

        private void kryptonTextBox4_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox5.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox3.Focus();
            }
        }

        //修正用
        private void kryptonTextBox5_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonNumericUpDown2.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox4.Focus();
            }
        }

        private void kryptonNumericUpDown2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox9.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox5.Focus();
            }
        }

        private void kryptonComboBox9_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox6.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonNumericUpDown2.Focus();
            }
        }

        private void kryptonTextBox6_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox7.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox9.Focus();
            }
        }
        //完了

        private void kryptonTextBox7_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox8.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox6.Focus();
            }
        }

        private void kryptonComboBox8_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox6.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox7.Focus();
            }
        }

        private void kryptonComboBox6_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox14.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox8.Focus();
            }
        }

        private void kryptonTextBox14_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox8.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox6.Focus();
            }
        }

        private void kryptonTextBox8_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox7.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox14.Focus();
            }
        }

        private void kryptonComboBox7_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox9.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox8.Focus();
            }
        }

        private void kryptonTextBox9_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox15.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox7.Focus();
            }
        }

        private void kryptonTextBox15_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox10.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox9.Focus();
            }
        }

        private void kryptonTextBox10_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox1.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox15.Focus();
            }
        }

        private void kryptonComboBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox2.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox10.Focus();
            }
        }

        private void kryptonComboBox2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox11.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox1.Focus();
            }
        }



        private void kryptonComboBox11_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox3.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox2.Focus();
            }
        }


        private void kryptonComboBox3_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox4.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox11.Focus();
            }
        }

        private void kryptonComboBox4_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox5.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox3.Focus();
            }
        }

        private void kryptonComboBox5_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox11.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox4.Focus();
            }
        }
        #endregion

        private void kryptonTextBox5_MouseUp(object sender, MouseEventArgs e)
        {

        }

        private void kryptonNumericUpDown2_MouseUp(object sender, MouseEventArgs e)
        {

        }

        private void kryptonTextBox11_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void kryptonTextBox11_Click(object sender, EventArgs e)
        {

        }

        private void kryptonTextBox11_TextChanged(object sender, EventArgs e)
        {
            if (kryptonTextBox11.Text != string.Empty)
            {
                label2.Text = "発第";
                Sheets_NumberLabel.Text = kryptonTextBox11.Text + "発第" + kryptonNumericUpDown1.Value + "号";
            }
            else
            {
                label2.Text = "　第";
                Sheets_NumberLabel.Text = "第" + kryptonNumericUpDown1.Value + "号";
            }

        }

        private void kryptonDateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (kryptonCheckBox1.Checked == true)
            {
                DateTime date = kryptonDateTimePicker1.Value.Date;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                //下記のように西暦ではなく和暦として表示するように設定する
                culturejp.DateTimeFormat.Calendar = new JapaneseCalendar();
                Sheets_DateLabel.Text = date.ToString("ggy年M月d日", culturejp);
            }
            else
            {
                DateTime date = kryptonDateTimePicker1.Value.Date;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                Sheets_DateLabel.Text = date.ToString("yyyy年M月d日");
            }

        }

        private void kryptonTextBox1_TextChanged(object sender, EventArgs e)
        {
            Sheets_AddressCompanyLabel.Text = kryptonTextBox1.Text;
        }

        private void kryptonComboBox10_TextChanged(object sender, EventArgs e)
        {
            if (kryptonComboBox10.Text != string.Empty)
            {

                if (kryptonTextBox2.Text != string.Empty)
                {
                    Sheets_AddressTitleAndNameLabel.Text = kryptonComboBox10.Text + "　" + kryptonTextBox2.Text + "様";
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                }
                else
                {
                    Sheets_AddressTitleAndNameLabel.Text = string.Empty;
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                }

            }
            else
            {
                if (kryptonTextBox2.Text != string.Empty)
                {
                    Sheets_AddressTitleAndNameLabel.Text = kryptonTextBox2.Text + "様";
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                }
                else
                {
                    Sheets_AddressTitleAndNameLabel.Text = string.Empty;
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                }
            }

            if (kryptonComboBox10.Text == "お客様各位")
            {
                kryptonTextBox1.Text = string.Empty;
                kryptonTextBox1.Enabled = false;
                Sheets_AddressTitleAndNameLabel.Text = "お客様各位";
                kryptonTextBox2.Text = string.Empty;
                kryptonTextBox2.Enabled = false;
                label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }
            else if (kryptonComboBox10.Text == "従業員各位")
            {
                kryptonTextBox1.Text = string.Empty;
                kryptonTextBox1.Enabled = false;
                Sheets_AddressTitleAndNameLabel.Text = "従業員各位";
                kryptonTextBox2.Text = string.Empty;
                kryptonTextBox2.Enabled = false;
                label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }
            else
            {
                kryptonTextBox1.Enabled = true;
                kryptonTextBox2.Enabled = true;
                label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
            }

        }


        private void kryptonTextBox2_TextChanged(object sender, EventArgs e)
        {

            if (kryptonComboBox10.Text != string.Empty)
            {

                if (kryptonTextBox2.Text != string.Empty)
                {
                    Sheets_AddressTitleAndNameLabel.Text = kryptonComboBox10.Text + "　" + kryptonTextBox2.Text + "様";
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                }
                else
                {
                    Sheets_AddressTitleAndNameLabel.Text = string.Empty;
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                }
            }
            else
            {
                if (kryptonTextBox2.Text != string.Empty)
                {

                    Sheets_AddressTitleAndNameLabel.Text = kryptonTextBox2.Text + "様";
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                }
                else
                {
                    Sheets_AddressTitleAndNameLabel.Text = string.Empty;
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                }
            }

            if (kryptonComboBox10.Text == "お客様各位")
            {
                kryptonTextBox1.Text = string.Empty;
                kryptonTextBox1.Enabled = false;
                Sheets_AddressTitleAndNameLabel.Text = "お客様各位";
                kryptonTextBox2.Text = string.Empty;
                kryptonTextBox2.Enabled = false;
                label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }
            else if (kryptonComboBox10.Text == "従業員各位")
            {
                kryptonTextBox1.Text = string.Empty;
                kryptonTextBox1.Enabled = false;
                Sheets_AddressTitleAndNameLabel.Text = "従業員各位";
                kryptonTextBox2.Text = string.Empty;
                kryptonTextBox2.Enabled = false;
                label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }


        }

        private void kryptonTextBox3_TextChanged(object sender, EventArgs e)
        {
            Sheets_CallerCompanyLabel.Text = kryptonTextBox3.Text;
        }

        private void kryptonTextBox4_TextChanged(object sender, EventArgs e)
        {
            Sheets_CallerLocationLabel.Text = kryptonTextBox4.Text;
        }

        private void kryptonTextBox5_TextChanged(object sender, EventArgs e)
        {
            if (kryptonTextBox5.Text != string.Empty)
            {
                kryptonNumericUpDown2.Enabled = true;
                Sheets_BuildingNameLabel.Text = kryptonTextBox5.Text + "　" + kryptonNumericUpDown2.Value + "階";
                label5.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
            }
            else
            {
                kryptonNumericUpDown2.Enabled = false;
                Sheets_BuildingNameLabel.Text = string.Empty;
                label5.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }

        }

        private void kryptonNumericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            if (kryptonTextBox5.Text != string.Empty)
            {
                kryptonNumericUpDown2.Enabled = true;
                Sheets_BuildingNameLabel.Text = kryptonTextBox5.Text + "　" + kryptonNumericUpDown2.Value + "階";
                label5.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                if (kryptonNumericUpDown2.Value <= 0)
                {
                    int negativeNumber = (int)kryptonNumericUpDown2.Value;
                    int positiveNumber = Math.Abs(negativeNumber);

                    Sheets_BuildingNameLabel.Text = kryptonTextBox5.Text + "　" + "地下" + positiveNumber + "階";
                }
            }

            //0の値を入力しないようにする
            if (kryptonNumericUpDown2.Value == 0)
            {
                kryptonNumericUpDown2.Value = 1;

            }
        }

        private void kryptonComboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonTextBox6_TextChanged(object sender, EventArgs e)
        {
            if (kryptonComboBox9.Text != string.Empty)
            {
                Sheets_CallerTitleAndNameLabel.Text = kryptonComboBox9.Text + "　" + kryptonTextBox6.Text;
                if (kryptonTextBox6.Text == string.Empty)
                {
                    Sheets_CallerTitleAndNameLabel.Text = string.Empty;
                }
            }
            else
            {
                Sheets_CallerTitleAndNameLabel.Text = kryptonTextBox6.Text;
                if (kryptonTextBox6.Text == string.Empty)
                {
                    Sheets_CallerTitleAndNameLabel.Text = string.Empty;
                }
            }
        }

        private void kryptonComboBox9_TextChanged(object sender, EventArgs e)
        {
            if (kryptonComboBox9.Text != string.Empty)
            {
                Sheets_CallerTitleAndNameLabel.Text = kryptonComboBox9.Text + "　" + kryptonTextBox6.Text;
                if (kryptonTextBox6.Text == string.Empty)
                {
                    Sheets_CallerTitleAndNameLabel.Text = string.Empty;
                }

            }
            else
            {
                Sheets_CallerTitleAndNameLabel.Text = kryptonTextBox6.Text;
                if (kryptonTextBox6.Text == string.Empty)
                {
                    Sheets_CallerTitleAndNameLabel.Text = string.Empty;
                }
            }

        }

        private void kryptonTextBox7_TextChanged(object sender, EventArgs e)
        {
            if (kryptonTextBox7.Text != string.Empty)
            {
                Sheets_CallerMallAddressLabel.Text = kryptonTextBox7.Text + "@" + kryptonComboBox8.Text;
                label11.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                label12.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
            }
            else
            {
                Sheets_CallerMallAddressLabel.Text = string.Empty;
                label11.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                label12.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }
        }

        private void kryptonComboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonComboBox8_TextChanged(object sender, EventArgs e)
        {
            if (kryptonComboBox8.Text != string.Empty)
            {
                Sheets_CallerMallAddressLabel.Text = kryptonTextBox7.Text + "@" + kryptonComboBox8.Text;
                label11.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                label12.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
            }
            else
            {
                Sheets_CallerMallAddressLabel.Text = string.Empty;
                label11.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                label12.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }
        }

        private void kryptonComboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonCheckBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (kryptonCheckBox3.Checked == true)
            {
                kryptonTextBox11.Enabled = false;
                kryptonNumericUpDown1.Enabled = false;
                label1.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                label2.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_NumberLabel.Text = string.Empty;
            }
            else
            {
                kryptonTextBox11.Enabled = true;
                kryptonNumericUpDown1.Enabled = true;
                label1.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                label2.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);

                if (kryptonTextBox11.Text != string.Empty)
                {
                    label2.Text = "発第";
                    Sheets_NumberLabel.Text = kryptonTextBox11.Text + "発第" + kryptonNumericUpDown1.Value + "号";
                }
                else
                {
                    label2.Text = "　第";
                    Sheets_NumberLabel.Text = "第" + kryptonNumericUpDown1.Value + "号";
                }

            }
        }

        private void kryptonCheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (kryptonCheckBox2.Checked == true)
            {
                kryptonCheckBox1.Enabled = false;
                kryptonDateTimePicker1.Enabled = false;
                Sheets_DateLabel.Text = string.Empty;
            }
            else
            {
                kryptonCheckBox1.Enabled = true;
                kryptonDateTimePicker1.Enabled = true;

                if (kryptonCheckBox1.Checked == true)
                {
                    DateTime date = kryptonDateTimePicker1.Value.Date;
                    CultureInfo culturejp = new CultureInfo("ja-Jp");
                    //下記のように西暦ではなく和暦として表示するように設定する
                    culturejp.DateTimeFormat.Calendar = new JapaneseCalendar();
                    Sheets_DateLabel.Text = date.ToString("ggy年M月d日", culturejp);
                }
                else
                {
                    DateTime date = kryptonDateTimePicker1.Value.Date;
                    CultureInfo culturejp = new CultureInfo("ja-Jp");
                    Sheets_DateLabel.Text = date.ToString("yyyy年M月d日");
                }
            }
        }

        private void kryptonComboBox6_TextChanged(object sender, EventArgs e)
        {
            //6
            if (kryptonComboBox6.Text != string.Empty)
            {
                label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                //14
                if (kryptonTextBox14.Text != string.Empty)
                {
                    label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                    Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                    //8
                    if (kryptonTextBox8.Text != string.Empty)
                    {
                        label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                        Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                    }
                    else
                    {
                        label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                        Sheets_CallerTelLabel.Text = string.Empty;
                    }
                }
                else
                {
                    label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                    Sheets_CallerTelLabel.Text = string.Empty;
                }
            }
            else
            {
                label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_CallerTelLabel.Text = string.Empty;
            }

        }

        private void kryptonTextBox14_TextChanged(object sender, EventArgs e)
        {
            //6
            if (kryptonComboBox6.Text != string.Empty)
            {
                label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                //14
                if (kryptonTextBox14.Text != string.Empty)
                {
                    label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                    Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                    //8
                    if (kryptonTextBox8.Text != string.Empty)
                    {
                        label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                        Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                    }
                    else
                    {
                        label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                        Sheets_CallerTelLabel.Text = string.Empty;
                    }
                }
                else
                {
                    label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                    Sheets_CallerTelLabel.Text = string.Empty;
                }
            }
            else
            {
                label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_CallerTelLabel.Text = string.Empty;
            }
        }

        private void kryptonTextBox8_TextChanged(object sender, EventArgs e)
        {
            //6
            if (kryptonComboBox6.Text != string.Empty)
            {
                label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                //14
                if (kryptonTextBox14.Text != string.Empty)
                {
                    label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                    Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                    //8
                    if (kryptonTextBox8.Text != string.Empty)
                    {
                        label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                        Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                    }
                    else
                    {
                        label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                        Sheets_CallerTelLabel.Text = string.Empty;
                    }
                }
                else
                {
                    label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                    Sheets_CallerTelLabel.Text = string.Empty;
                }
            }
            else
            {
                label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_CallerTelLabel.Text = string.Empty;
            }
        }


        private void kryptonComboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonComboBox7_TextChanged(object sender, EventArgs e)
        {
            //7
            if (kryptonComboBox7.Text != string.Empty)
            {
                label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                //15
                if (kryptonTextBox15.Text != string.Empty)
                {
                    label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                    Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                    //9
                    if (kryptonComboBox9.Text != string.Empty)
                    {
                        label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                        Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                    }
                    else
                    {
                        label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                        Sheets_CallerFaxTelLabel.Text = string.Empty;
                    }
                }
                else
                {
                    label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                    Sheets_CallerFaxTelLabel.Text = string.Empty;
                }
            }
            else
            {
                label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_CallerFaxTelLabel.Text = string.Empty;

            }
        }

        private void kryptonTextBox9_TextChanged(object sender, EventArgs e)
        {
            //9
            if (kryptonTextBox9.Text != string.Empty)
            {
                label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                //15
                if (kryptonTextBox15.Text != string.Empty)
                {
                    label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                    Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                    //7
                    if (kryptonComboBox7.Text != string.Empty)
                    {
                        label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                        Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                    }
                    else
                    {
                        label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                        Sheets_CallerFaxTelLabel.Text = string.Empty;
                    }
                }
                else
                {
                    label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                    Sheets_CallerFaxTelLabel.Text = string.Empty;
                }
            }
            else
            {
                label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_CallerFaxTelLabel.Text = string.Empty;
            }
        }

        private void kryptonTextBox15_TextChanged(object sender, EventArgs e)
        {
            //15
            if (kryptonTextBox15.Text != string.Empty)
            {
                label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                //9
                if (kryptonTextBox9.Text != string.Empty)
                {
                    label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                    Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                    //7
                    if (kryptonComboBox7.Text != string.Empty)
                    {
                        label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                        Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                    }
                    else
                    {
                        label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                        Sheets_CallerFaxTelLabel.Text = string.Empty;
                    }
                }
                else
                {
                    label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                    Sheets_CallerFaxTelLabel.Text = string.Empty;
                }
            }
            else
            {
                label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_CallerFaxTelLabel.Text = string.Empty;
            }
        }

        private void Notepads_kryptonRichTextBox_Notepad_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;
            }
            else
            {
                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;
            }

            if (this.Width <= 902)
            {
                Sheets_Sheet.Top = 59;
                Sheets_Sheet.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            }
            else
            {

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

            }
        }

        private void kryptonComboBox2_TextChanged(object sender, EventArgs e)
        {
            Sheets_ContentLabel.Text = kryptonComboBox2.Text + "　" + kryptonComboBox11.Text + kryptonComboBox3.Text + kryptonComboBox4.Text;
            #region 頭語の選択による結語候補の切り替え処理
            //一般的
            if (kryptonComboBox2.Text == "拝啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "拝呈")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "啓上")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "敬白")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "拝進")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            //丁寧さ
            else if (kryptonComboBox2.Text == "謹啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
                kryptonComboBox5.Text = "謹言";
            }
            else if (kryptonComboBox2.Text == "謹呈")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
                kryptonComboBox5.Text = "謹言";
            }
            else if (kryptonComboBox2.Text == "粛啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
                kryptonComboBox5.Text = "謹言";
            }
            else if (kryptonComboBox2.Text == "慕啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
                kryptonComboBox5.Text = "謹言";
            }
            else if (kryptonComboBox2.Text == "謹白")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
                kryptonComboBox5.Text = "謹言";
            }
            //急ぎ
            else if (kryptonComboBox2.Text == "急啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "不二",
                    "早々",
                    "不備",
                });
                kryptonComboBox5.Text = "草々";
            }
            else if (kryptonComboBox2.Text == "急呈")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "不二",
                    "早々",
                    "不備",
                });
                kryptonComboBox5.Text = "草々";
            }
            else if (kryptonComboBox2.Text == "急白")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "不二",
                    "早々",
                    "不備",
                });
                kryptonComboBox5.Text = "草々";
            }
            //略式
            else if (kryptonComboBox2.Text == "前略")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "早々",
                    "不二",
                });
                kryptonComboBox5.Text = "草々";
            }
            else if (kryptonComboBox2.Text == "冠省")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "早々",
                    "不二",
                });
                kryptonComboBox5.Text = "草々";
            }
            else if (kryptonComboBox2.Text == "略啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "早々",
                    "不二",
                });
                kryptonComboBox5.Text = "草々";
            }
            else if (kryptonComboBox2.Text == "寸啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "早々",
                    "不二",
                });
                kryptonComboBox5.Text = "草々";
            }
            else if (kryptonComboBox2.Text == "草啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "早々",
                    "不二",
                });
                kryptonComboBox5.Text = "草々";
            }
            //初めて
            else if (kryptonComboBox2.Text == "初めてお手紙を差し上げます")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "突然お手紙を差し上げますご無礼お許しください")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            //重ねて
            else if (kryptonComboBox2.Text == "拝復")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝答",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "複啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝答",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "謹復")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝答",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            //お悔み
            else if (kryptonComboBox2.Text == "合掌")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "合掌",
                });
                kryptonComboBox5.Text = "合掌";
            }
            else if (kryptonComboBox2.Text == "敬具")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                });
                kryptonComboBox5.Text = "敬具";
            }
            #endregion
        }

        private void kryptonComboBox1_TextChanged(object sender, EventArgs e)
        {
            #region 月の選択による候文候補の切り替え処理
            if (kryptonComboBox1.Text == "1")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "新春の候、",
                    "初春の候、",
                    "頌春の候、",
                    "厳寒の候、",
                    "厳冬の候、",
                    "中冬の候、",
                    "寒冷の候、",
                    "麗春の候、",
                    "大寒のみぎり、",
                    "酷寒のみぎり、",
                    "寒さ厳しき季節、",
                });
                kryptonComboBox11.Text = "新春の候、";
            }
            else if (kryptonComboBox1.Text == "2")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "余寒の候、",
                    "春寒の候、",
                    "晩冬の候、",
                    "向春の候、",
                    "解氷の候、",
                    "梅花の候、",
                    "余寒なお厳しき折、",
                });
                kryptonComboBox11.Text = "余寒の候、";
            }
            else if (kryptonComboBox1.Text == "3")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "早春の候、",
                    "春寒の候、",
                    "孟春の候、",
                    "春雨降りやまぬ候、",
                    "浅春のみぎり、",
                    "春寒しだいに緩むころ、",
                    "冬の名残のまだ去りやらぬ時候、",
                    "春光天地に満ちて快い時候、",
                    "春分の季節、",
                    "春色のなごやかな季節、",
                });
                kryptonComboBox11.Text = "早春の候、";
            }
            else if (kryptonComboBox1.Text == "4")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "陽春の候、",
                    "春暖の候、",
                    "軽暖の候、",
                    "麗春の候、",
                    "春暖快適の候、",
                    "桜花爛漫の候、",
                    "花信相次ぐ候、",
                    "春眠暁を覚えずの候、",
                    "仲春四月、",
                    "春たけなわの今日この頃、",
                });
                kryptonComboBox11.Text = "早春の候、";
            }
            else if (kryptonComboBox1.Text == "5")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "新緑の候、",
                    "薫風の候、",
                    "初夏の候、",
                    "立夏の候、",
                    "暮春の候、",
                    "老春の候、",
                    "軽暑の候、",
                    "惜春のみぎり、",
                    "若葉の鮮やかな季節、",
                });
                kryptonComboBox11.Text = "新緑の候、";
            }
            else if (kryptonComboBox1.Text == "6")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "梅雨の候、",
                    "初夏の候、",
                    "短夜の候、",
                    "五月雨の候、",
                    "長雨の候、",
                    "薄暑の候、",
                    "向夏の候、",
                    "麦秋の候、",
                    "向暑のみぎり、",
                    "若鮎おどる季節、",
                });
                kryptonComboBox11.Text = "梅雨の候、";
            }
            else if (kryptonComboBox1.Text == "7")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "猛暑の候、",
                    "酷暑の候、",
                    "炎暑の候、",
                    "盛夏の候、",
                    "大暑の候、",
                    "灼熱の候、",
                    "炎熱のみぎり、",
                    "甚暑のみぎり、",
                    "三伏のみぎり、",
                    "暑さ厳しき折から、",
                });
                kryptonComboBox11.Text = "猛暑の候、";
            }
            else if (kryptonComboBox1.Text == "8")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "残暑の候、",
                    "残炎の候、",
                    "残夏の候、",
                    "暮夏の候、",
                    "季夏の候、",
                    "新涼の候、",
                    "秋暑厳しき候、",
                    "晩夏のみぎり、",
                    "処暑のみぎり、",
                    "処暑のみぎり、",
                });
                kryptonComboBox11.Text = "残暑の候、";
            }
            else if (kryptonComboBox1.Text == "9")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "初秋の候、",
                    "仲秋の候、",
                    "錦秋の候、",
                    "寒露の候、",
                    "黄葉の候、",
                    "秋雨の候、",
                    "金風の候、",
                    "秋晴れの候、",
                    "菊薫る候、",
                    "秋たけなわの候、",
                    "紅葉の季節、",
                    "秋冷の心地よい季節、",
                });
                kryptonComboBox11.Text = "初秋の候、";
            }
            else if (kryptonComboBox1.Text == "10")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "秋冷の候、",
                    "仲秋の候、",
                    "錦秋の候、",
                    "寒露の候、",
                    "黄葉の候、",
                    "秋雨の候、",
                    "金風の候、",
                    "秋晴れの候、",
                    "菊薫る候、",
                    "秋たけなわの候、",
                    "紅葉の季節、",
                    "秋冷の心地よい季節、",
                });
                kryptonComboBox11.Text = "初秋の候、";
            }
            else if (kryptonComboBox1.Text == "11")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "晩秋の候、",
                    "暮秋の候、",
                    "向寒の候、",
                    "深冷の候、",
                    "菊花の候、",
                    "紅葉の候、",
                    "初霜の候、",
                    "氷雨の候、",
                    "枯れ葉舞う季節、",
                });
                kryptonComboBox11.Text = "晩秋の候、";
            }
            else if (kryptonComboBox1.Text == "12")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "寒冷の候、",
                    "師走の候、",
                    "初冬の候、",
                    "寒気の候、",
                    "霜気の候、",
                    "霜寒の候、",
                    "季冬の候、",
                    "歳晩の候、",
                    "歳末ご多忙の折、",
                    "心せわしい年の暮れ、",
                });
                kryptonComboBox11.Text = "寒冷の候、";
            }
            #endregion
        }

        #region ナビゲーションバーのサイズ切り替え処理
        private void buttonSpecNavigator1_Click(object sender, EventArgs e)
        {
            if (kryptonSplitContainer2.SplitterDistance <= 100)
            {
                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookFull;

                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;

                buttonSpecNavigator1.Orientation = ComponentFactory.Krypton.Toolkit.PaletteButtonOrientation.Inherit;
                buttonSpecNavigator1.Text = "縮小";

                kryptonPage4.MaximumSize = new Size(0, 0);
                kryptonPage4.MinimumSize = new Size(0, 0);

                Transition
                    .With(kryptonSplitContainer2, nameof(kryptonSplitContainer2.SplitterDistance), 302)
                    .CriticalDamp(TimeSpan.FromSeconds(0.6));
            }
            else
            {
                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookMini;

                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;

                buttonSpecNavigator1.Orientation = ComponentFactory.Krypton.Toolkit.PaletteButtonOrientation.FixedLeft;
                buttonSpecNavigator1.Text = "広げる";

                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;
                kryptonPage4.MaximumSize = new Size(270, 0);
                kryptonPage4.MinimumSize = new Size(270, 0);

                Transition
                    .With(kryptonSplitContainer2, nameof(kryptonSplitContainer2.SplitterDistance), 42)
                    .CriticalDamp(TimeSpan.FromSeconds(0.6));
            }

        }

        private void kryptonSplitContainer2_SplitterMoving(object sender, SplitterCancelEventArgs e)
        {
            if (kryptonSplitContainer2.SplitterDistance <= 100)
            {
                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookMini;

                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;

                buttonSpecNavigator1.Orientation = ComponentFactory.Krypton.Toolkit.PaletteButtonOrientation.FixedLeft;
                buttonSpecNavigator1.Text = "広げる";

                kryptonPage4.MaximumSize = new Size(270, 0);
                kryptonPage4.MinimumSize = new Size(270, 0);

            }
            else
            {
                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookFull;


                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;

                buttonSpecNavigator1.Orientation = ComponentFactory.Krypton.Toolkit.PaletteButtonOrientation.Inherit;
                buttonSpecNavigator1.Text = "縮小";

                kryptonPage4.MaximumSize = new Size(0, 0);
                kryptonPage4.MinimumSize = new Size(0, 0);
            }
        }

        private void kryptonSplitContainer2_SplitterMoved(object sender, SplitterEventArgs e)
        {
            if (kryptonSplitContainer2.SplitterDistance <= 100)
            {
                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookMini;

                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;

                buttonSpecNavigator1.Orientation = ComponentFactory.Krypton.Toolkit.PaletteButtonOrientation.FixedLeft;
                buttonSpecNavigator1.Text = "広げる";

                kryptonPage4.MaximumSize = new Size(270, 0);
                kryptonPage4.MinimumSize = new Size(270, 0);
            }
            else
            {
                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookFull;

                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookFull;


                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;

                buttonSpecNavigator1.Orientation = ComponentFactory.Krypton.Toolkit.PaletteButtonOrientation.Inherit;
                buttonSpecNavigator1.Text = "縮小";

                kryptonPage4.MaximumSize = new Size(0, 0);
                kryptonPage4.MinimumSize = new Size(0, 0);
            }
        }
        #endregion

        private void kryptonRibbonGroupButton7_Click(object sender, EventArgs e)
        {

        }



        private void kryptonComboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonComboBox11_TextChanged(object sender, EventArgs e)
        {
            Sheets_ContentLabel.Text = kryptonComboBox2.Text + "　" + kryptonComboBox11.Text + kryptonComboBox3.Text + kryptonComboBox4.Text;
        }

        private void kryptonComboBox3_TextChanged(object sender, EventArgs e)
        {
            Sheets_ContentLabel.Text = kryptonComboBox2.Text + "　" + kryptonComboBox11.Text + kryptonComboBox3.Text + kryptonComboBox4.Text;
        }

        private void kryptonComboBox4_TextChanged(object sender, EventArgs e)
        {
            Sheets_ContentLabel.Text = kryptonComboBox2.Text + "　" + kryptonComboBox11.Text + kryptonComboBox3.Text + kryptonComboBox4.Text;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            Sheets_Sheet.Top = 59;
            Sheets_Sheet.Left = 0;
        }

        private void kryptonComboBox5_TextChanged(object sender, EventArgs e)
        {
            Sheet_ConclusionLabel.Text = kryptonComboBox5.Text;
        }

        private void kryptonTextBox10_TextChanged(object sender, EventArgs e)
        {
            if (kryptonTextBox10.Text != string.Empty)
            {
                Sheets_TitleButton.Text = kryptonTextBox10.Text;
                this.Text = kryptonTextBox10.Text + " - DocuQuick Designer";
            }
            else
            {
                this.Text = "無題 - DocuQuick Designer";
            }

        }

        private void kryptonTextBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void kryptonTextBox13_TextChanged(object sender, EventArgs e)
        {

        }



        #region クリップボード
        //コピー
        private void kryptonRibbonGroupButton1_Click_1(object sender, EventArgs e)
        {
            if (kryptonTextBox11.Focused == true)
            {
                kryptonTextBox11.Copy();
            }
            else if (kryptonNumericUpDown1.Focused == true)
            {
                //kryptonNumericUpDownにはCopyメソッドがないためstring＆SetText経由でクリップボードにコピーする
                string Clip = kryptonNumericUpDown1.Value.ToString();
                Clipboard.SetText(Clip);
            }
            else if (kryptonDateTimePicker1.Focused == true)
            {
                //kryptonDateTimePickerにはCopyメソッドがないためSheets_DateLabelをSetText経由でクリップボードにコピーする
                Clipboard.SetText(Sheets_DateLabel.Text);
            }
            else if (kryptonTextBox1.Focused == true)
            {
                kryptonTextBox1.Copy();
            }
            else if (kryptonComboBox10.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox10.SelectedText);
            }
            else if (kryptonTextBox2.Focused == true)
            {
                kryptonTextBox2.Copy();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Copy();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Copy();
            }
            else if (kryptonTextBox4.Focused == true)
            {
                kryptonTextBox4.Copy();
            }
            else if (kryptonTextBox5.Focused == true)
            {
                kryptonTextBox5.Copy();
            }
            else if (kryptonNumericUpDown2.Focused == true)
            {
                //kryptonNumericUpDownにはCopyメソッドがないためstring＆SetText経由でクリップボードにコピーする
                string Clip = kryptonNumericUpDown2.Value.ToString();
                Clipboard.SetText(Clip);
            }
            else if (kryptonComboBox9.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox9.SelectedText);
            }
            else if (kryptonTextBox6.Focused == true)
            {
                kryptonTextBox6.Copy();
            }
            else if (kryptonTextBox7.Focused == true)
            {
                kryptonTextBox7.Copy();
            }
            else if (kryptonComboBox8.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox8.SelectedText);
            }
            else if (kryptonComboBox6.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox6.SelectedText);
            }
            else if (kryptonTextBox14.Focused == true)
            {
                kryptonTextBox14.Copy();
            }
            else if (kryptonTextBox8.Focused == true)
            {
                kryptonTextBox8.Copy();
            }
            else if (kryptonComboBox7.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox7.SelectedText);
            }
            else if (kryptonTextBox9.Focused == true)
            {
                kryptonTextBox9.Copy();
            }
            else if (kryptonTextBox15.Focused == true)
            {
                kryptonTextBox15.Copy();
            }
            else if (kryptonTextBox10.Focused == true)
            {
                kryptonTextBox10.Copy();
            }
            else if (kryptonComboBox2.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox2.SelectedText);
            }
            else if (kryptonComboBox11.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox11.SelectedText);
            }
            else if (kryptonComboBox3.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox3.SelectedText);
            }
            else if (kryptonComboBox4.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox4.SelectedText);
            }
            else if (kryptonTextBox12.Focused == true)
            {
                kryptonTextBox12.Copy();
            }
            else if (kryptonComboBox5.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox5.SelectedText);
            }
            else if (kryptonTextBox13.Focused == true)
            {
                kryptonTextBox13.Copy();
            }
        }

        //切り取り
        private void kryptonRibbonGroupButton11_Click(object sender, EventArgs e)
        {
            if (kryptonTextBox11.Focused == true)
            {
                kryptonTextBox11.Cut();
            }
            else if (kryptonNumericUpDown1.Focused == true)
            {
                //kryptonNumericUpDownにはCutメソッドがないためstring＆SetText経由でクリップボードにコピーする
                string Clip = kryptonNumericUpDown1.Value.ToString();
                Clipboard.SetText(Clip);
                //kryptonNumericUpDownは値を削除できないためかわりに値を1に変更する
                kryptonNumericUpDown1.Value = 1;
            }
            else if (kryptonDateTimePicker1.Focused == true)
            {
                //kryptonDateTimePickerにはCutメソッドがないためSheets_DateLabelをSetText経由でクリップボードにコピーする
                Clipboard.SetText(Sheets_DateLabel.Text);
                //kryptonDateTimePicker1は値を削除できないためかわりに今日の日付に変更する
                kryptonDateTimePicker1.Value = kryptonDateTimePicker1.CalendarTodayDate;
            }
            else if (kryptonTextBox1.Focused == true)
            {
                kryptonTextBox1.Cut();
            }
            else if (kryptonComboBox10.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox10.SelectedText);
                // 修正: '-=' 演算子はstring型に使えません。選択されたテキストを除去するにはReplaceを使います。
                if (!string.IsNullOrEmpty(kryptonComboBox10.SelectedText))
                {
                    kryptonComboBox10.Text = kryptonComboBox10.Text.Replace(kryptonComboBox10.SelectedText, "");
                }
            }
            else if (kryptonTextBox2.Focused == true)
            {
                kryptonTextBox2.Cut();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Cut();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Cut();
            }
            else if (kryptonTextBox4.Focused == true)
            {
                kryptonTextBox4.Cut();
            }
            else if (kryptonTextBox5.Focused == true)
            {
                kryptonTextBox5.Cut();
            }
            else if (kryptonNumericUpDown2.Focused == true)
            {
                //kryptonNumericUpDownにはCutメソッドがないためstring＆SetText経由でクリップボードにコピーする
                string Clip = kryptonNumericUpDown2.Value.ToString();
                Clipboard.SetText(Clip);
                //kryptonNumericUpDownは値を削除できないためかわりに値を1に変更する
                kryptonNumericUpDown2.Value = 1;
            }
            else if (kryptonComboBox9.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox9.SelectedText);
                // 修正: '-=' 演算子はstring型に使えません。選択されたテキストを除去するにはReplaceを使います。
                if (!string.IsNullOrEmpty(kryptonComboBox9.SelectedText))
                {
                    kryptonComboBox9.Text = kryptonComboBox9.Text.Replace(kryptonComboBox9.SelectedText, "");
                }
            }
            else if (kryptonTextBox6.Focused == true)
            {
                kryptonTextBox6.Cut();
            }
            else if (kryptonTextBox7.Focused == true)
            {
                kryptonTextBox7.Cut();
            }
            else if (kryptonComboBox8.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox8.SelectedText);
                // 修正: '-=' 演算子はstring型に使えません。選択されたテキストを除去するにはReplaceを使います。
                if (!string.IsNullOrEmpty(kryptonComboBox8.SelectedText))
                {
                    kryptonComboBox8.Text = kryptonComboBox8.Text.Replace(kryptonComboBox8.SelectedText, "");
                }
            }
            else if (kryptonComboBox6.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox6.SelectedText);
                // 修正: '-=' 演算子はstring型に使えません。選択されたテキストを除去するにはReplaceを使います。
                if (!string.IsNullOrEmpty(kryptonComboBox6.SelectedText))
                {
                    kryptonComboBox6.Text = kryptonComboBox6.Text.Replace(kryptonComboBox6.SelectedText, "");
                }
            }
            else if (kryptonTextBox14.Focused == true)
            {
                kryptonTextBox14.Cut();
            }
            else if (kryptonTextBox8.Focused == true)
            {
                kryptonTextBox8.Cut();
            }
            else if (kryptonComboBox7.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox7.SelectedText);
                // 修正: '-=' 演算子はstring型に使えません。選択されたテキストを除去するにはReplaceを使います。
                if (!string.IsNullOrEmpty(kryptonComboBox7.SelectedText))
                {
                    kryptonComboBox7.Text = kryptonComboBox7.Text.Replace(kryptonComboBox7.SelectedText, "");
                }

            }
            else if (kryptonTextBox9.Focused == true)
            {
                kryptonTextBox9.Cut();
            }
            else if (kryptonTextBox15.Focused == true)
            {
                kryptonTextBox15.Cut();
            }
            else if (kryptonTextBox10.Focused == true)
            {
                kryptonTextBox10.Cut();
            }
            else if (kryptonComboBox2.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox2.SelectedText);
                // 修正: '-=' 演算子はstring型に使えません。選択されたテキストを除去するにはReplaceを使います。
                if (!string.IsNullOrEmpty(kryptonComboBox2.SelectedText))
                {
                    kryptonComboBox2.Text = kryptonComboBox2.Text.Replace(kryptonComboBox2.SelectedText, "");
                }
            }
            else if (kryptonComboBox11.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox11.SelectedText);
                if (!string.IsNullOrEmpty(kryptonComboBox11.SelectedText))
                {
                    kryptonComboBox11.Text = kryptonComboBox11.Text.Replace(kryptonComboBox11.SelectedText, "");
                }
            }
            else if (kryptonComboBox3.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox3.SelectedText);
                if (!string.IsNullOrEmpty(kryptonComboBox3.SelectedText))
                {
                    kryptonComboBox3.Text = kryptonComboBox3.Text.Replace(kryptonComboBox3.SelectedText, "");
                }
            }
            else if (kryptonComboBox4.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox4.SelectedText);
                if (!string.IsNullOrEmpty(kryptonComboBox4.SelectedText))
                {
                    kryptonComboBox4.Text = kryptonComboBox4.Text.Replace(kryptonComboBox4.SelectedText, "");
                }
            }
            else if (kryptonTextBox12.Focused == true)
            {
                kryptonTextBox12.Cut();
            }
            else if (kryptonComboBox5.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox5.SelectedText);
                if (!string.IsNullOrEmpty(kryptonComboBox5.SelectedText))
                {
                    kryptonComboBox5.Text = kryptonComboBox5.Text.Replace(kryptonComboBox5.SelectedText, "");
                }
            }
            else if (kryptonTextBox13.Focused == true)
            {
                kryptonTextBox13.Cut();
            }
        }

        //削除
        private void kryptonRibbonGroupButton12_Click(object sender, EventArgs e)
        {
            if (kryptonTextBox11.Focused == true)
            {
                kryptonTextBox11.Clear();
            }
            else if (kryptonNumericUpDown1.Focused == true)
            {
                //kryptonNumericUpDownは値を削除できないためかわりに値を1に変更する
                kryptonNumericUpDown1.Value = 1;
            }
            else if (kryptonDateTimePicker1.Focused == true)
            {
                //kryptonDateTimePicker1は値を削除できないためかわりに今日の日付に変更する
                kryptonDateTimePicker1.Value = kryptonDateTimePicker1.CalendarTodayDate;
            }
            else if (kryptonTextBox1.Focused == true)
            {
                kryptonTextBox1.Clear();
            }
            else if (kryptonComboBox10.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox10.SelectedText))
                {
                    kryptonComboBox10.Text = kryptonComboBox10.Text.Replace(kryptonComboBox10.SelectedText, "");
                }
            }
            else if (kryptonTextBox2.Focused == true)
            {
                kryptonTextBox2.Clear();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Clear();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Clear();
            }
            else if (kryptonTextBox4.Focused == true)
            {
                kryptonTextBox4.Clear();
            }
            else if (kryptonTextBox5.Focused == true)
            {
                kryptonTextBox5.Clear();
            }
            else if (kryptonNumericUpDown2.Focused == true)
            {
                //kryptonNumericUpDownは値を削除できないためかわりに値を1に変更する
                kryptonNumericUpDown2.Value = 1;
            }
            else if (kryptonComboBox9.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox8.SelectedText))
                {
                    kryptonComboBox8.Text = kryptonComboBox8.Text.Replace(kryptonComboBox9.SelectedText, "");
                }
            }
            else if (kryptonTextBox6.Focused == true)
            {
                kryptonTextBox6.Clear();
            }
            else if (kryptonTextBox7.Focused == true)
            {
                kryptonTextBox7.Clear();
            }
            else if (kryptonComboBox8.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox8.SelectedText))
                {
                    kryptonComboBox8.Text = kryptonComboBox8.Text.Replace(kryptonComboBox8.SelectedText, "");
                }
            }
            else if (kryptonComboBox6.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox6.SelectedText))
                {
                    kryptonComboBox6.Text = kryptonComboBox6.Text.Replace(kryptonComboBox6.SelectedText, "");
                }
            }
            else if (kryptonTextBox14.Focused == true)
            {
                kryptonTextBox14.Clear();
            }
            else if (kryptonTextBox8.Focused == true)
            {
                kryptonTextBox8.Clear();
            }
            else if (kryptonComboBox7.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox7.SelectedText))
                {
                    kryptonComboBox7.Text = kryptonComboBox7.Text.Replace(kryptonComboBox7.SelectedText, "");
                }
            }
            else if (kryptonTextBox9.Focused == true)
            {
                kryptonTextBox9.Clear();
            }
            else if (kryptonTextBox15.Focused == true)
            {
                kryptonTextBox15.Clear();
            }
            else if (kryptonTextBox10.Focused == true)
            {
                kryptonTextBox10.Clear();
            }
            else if (kryptonComboBox2.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox2.SelectedText))
                {
                    kryptonComboBox2.Text = kryptonComboBox2.Text.Replace(kryptonComboBox2.SelectedText, "");
                }
            }
            else if (kryptonComboBox11.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox11.SelectedText))
                {
                    kryptonComboBox11.Text = kryptonComboBox11.Text.Replace(kryptonComboBox11.SelectedText, "");
                }
            }
            else if (kryptonComboBox3.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox3.SelectedText))
                {
                    kryptonComboBox3.Text = kryptonComboBox3.Text.Replace(kryptonComboBox3.SelectedText, "");
                }
            }
            else if (kryptonComboBox4.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox4.SelectedText))
                {
                    kryptonComboBox4.Text = kryptonComboBox4.Text.Replace(kryptonComboBox4.SelectedText, "");
                }
            }
            else if (kryptonTextBox12.Focused == true)
            {
                kryptonTextBox12.Clear();
            }
            else if (kryptonComboBox5.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox5.SelectedText))
                {
                    kryptonComboBox5.Text = kryptonComboBox5.Text.Replace(kryptonComboBox5.SelectedText, "");
                }
            }
            else if (kryptonTextBox13.Focused == true)
            {
                kryptonTextBox13.Clear();
            }
        }

        //貼り付け
        private void kryptonRibbonButton_Paste_Click(object sender, EventArgs e)
        {
            if (kryptonTextBox11.Focused == true)
            {
                kryptonTextBox11.Paste();
            }
            //NnumricUpDawnも無視
            //DateTimePickarは無視
            else if (kryptonTextBox1.Focused == true)
            {
                kryptonTextBox1.Paste();
            }
            else if (kryptonComboBox10.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox10.Text += Clipboard.GetText();
            }
            else if (kryptonTextBox2.Focused == true)
            {
                kryptonTextBox2.Paste();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Paste();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Paste();
            }
            else if (kryptonTextBox4.Focused == true)
            {
                kryptonTextBox4.Paste();
            }
            else if (kryptonTextBox5.Focused == true)
            {
                kryptonTextBox5.Paste();
            }
            else if (kryptonNumericUpDown2.Focused == true)
            {
                //最後にコピーした文字をkryptonNumericUpDownにペーストする
                kryptonNumericUpDown2.Value = Clipboard.GetText().Length;
            }
            else if (kryptonComboBox9.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox9.Text += Clipboard.GetText().Length;
            }
            else if (kryptonTextBox6.Focused == true)
            {
                kryptonTextBox6.Paste();
            }
            else if (kryptonTextBox7.Focused == true)
            {
                kryptonTextBox7.Paste();
            }
            else if (kryptonComboBox8.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox8.Text += Clipboard.GetText();
            }
            else if (kryptonComboBox6.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox6.Text += Clipboard.GetText();
            }
            else if (kryptonTextBox14.Focused == true)
            {
                kryptonTextBox14.Paste();
            }
            else if (kryptonTextBox8.Focused == true)
            {
                kryptonTextBox8.Paste();
            }
            else if (kryptonComboBox7.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox7.Text += Clipboard.GetText();
            }
            else if (kryptonTextBox9.Focused == true)
            {
                kryptonTextBox9.Paste();
            }
            else if (kryptonTextBox15.Focused == true)
            {
                kryptonTextBox15.Paste();
            }
            else if (kryptonTextBox10.Focused == true)
            {
                kryptonTextBox10.Paste();
            }
            else if (kryptonComboBox2.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox2.Text += Clipboard.GetText();
            }
            else if (kryptonComboBox11.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox11.Text += Clipboard.GetText();
            }
            else if (kryptonComboBox3.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox3.Text += Clipboard.GetText();
            }
            else if (kryptonComboBox4.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox3.Text += Clipboard.GetText();
            }
            else if (kryptonTextBox12.Focused == true)
            {
                kryptonTextBox12.Paste();
            }
            else if (kryptonComboBox5.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox3.Text += Clipboard.GetText();
            }
            else if (kryptonTextBox13.Focused == true)
            {
                kryptonTextBox13.Paste();
            }
        }
        #endregion

        private void kryptonRibbonGroupButton12_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextMenuArgs e)
        {

        }

        #region 設定画面表示処理
        private void kryptonContextMenuItem4_Click(object sender, EventArgs e)
        {
            Transition
                .With(kryptonPanel21, nameof(Height), 0)
                .CriticalDamp(TimeSpan.FromSeconds(0.4));
            kryptonTrackBar1.Enabled = false;
            kryptonButton15.Enabled = false;
            kryptonButton14.Enabled = false;
            kryptonLabel42.Enabled = false;

            kryptonPage2.Visible = true;
            kryptonNavigator_Workbench.NavigatorMode = NavigatorMode.Panel;
            kryptonNavigator_Workbench.SelectedPage = kryptonPage2;

            kryptonRibbon.MinimizedMode = true;
            kryptonRibbon.Enabled = false;

            kryptonLabel7.Enabled = false;
            kryptonCheckButton1.Enabled = false;
            kryptonCheckButton2.Enabled = false;
            kryptonLabel1.Enabled = false;

            this.Text = "設定 - DocuQuick Designer";

            kryptonCheckButton3.Checked = true;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage6;
        }



        private void kryptonGroupBox1_Panel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonCheckButton3_Click(object sender, EventArgs e)
        {
            kryptonCheckButton3.Checked = true;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage6;
        }

        private void kryptonCheckButton4_Click(object sender, EventArgs e)
        {
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = true;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage7;
        }

        private void kryptonCheckButton5_Click(object sender, EventArgs e)
        {
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = true;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage10;
        }

        private void kryptonCheckButton6_Click(object sender, EventArgs e)
        {
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = true;
            kryptonCheckButton7.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage11;
        }

        private void kryptonCheckButton7_Click(object sender, EventArgs e)
        {
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = true;

            kryptonNavigator1.SelectedPage = kryptonPage12;
        }

        #endregion

        #region テンプレート選択画面表示処理
        private void kryptonRibbonRecentDoc10_Click(object sender, EventArgs e)
        {
            Transition
                .With(kryptonPanel21, nameof(Height), 0)
                .CriticalDamp(TimeSpan.FromSeconds(0.4));
            kryptonTrackBar1.Enabled = false;
            kryptonButton15.Enabled = false;
            kryptonButton14.Enabled = false;
            kryptonLabel42.Enabled = false;

            kryptonRibbon.Enabled = false;
            kryptonRibbon.MinimizedMode = true;
            kryptonPage9.Visible = true;
            kryptonNavigator_Workbench.SelectedPage = kryptonPage9;
            kryptonNavigator_Workbench.NavigatorMode = NavigatorMode.Panel;
            this.Text = "テンプレート - DocuQuick Designer";

            kryptonLabel7.Enabled = false;
            kryptonCheckButton1.Enabled = false;
            kryptonCheckButton2.Enabled = false;
            kryptonLabel1.Enabled = false;
        }

        private void kryptonButton7_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton16.Checked == true)
            {
                Transition
                    .With(kryptonPanel21, nameof(Height), 36)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                kryptonRibbonGroupButton16.Checked = true;
            }
            else
            {
                Transition
                    .With(kryptonPanel21, nameof(Height), 0)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                kryptonRibbonGroupButton16.Checked = false;
            }

            if (kryptonTrackBar1.Value == 0)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }
            else if (kryptonTrackBar1.Value == 10)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = false;
                kryptonLabel42.Enabled = true;
            }
            else
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }

            kryptonPage9.Visible = false;
            kryptonNavigator_Workbench.NavigatorMode = NavigatorMode.BarTabGroup;
            kryptonNavigator_Workbench.SelectedPage = kryptonPage1;

            kryptonRibbon.MinimizedMode = false;
            kryptonRibbon.Enabled = true;

            kryptonLabel7.Enabled = true;
            kryptonCheckButton1.Enabled = true;
            kryptonCheckButton2.Enabled = true;
            kryptonLabel1.Enabled = true;

            if (kryptonTextBox10.Text != string.Empty)
            {
                Sheets_TitleButton.Text = kryptonTextBox10.Text;
                this.Text = kryptonTextBox10.Text + " - DocuQuick Designer";
            }
            else
            {
                this.Text = "無題 - DocuQuick Designer";
            }

            kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

            Sheets_Sheet.Anchor = AnchorStyles.Top;

            // 親コントロールのサイズを取得
            int parentWidth = this.ClientSize.Width;
            int parentHeight = this.ClientSize.Height;

            // パネルのサイズを取得
            int panelWidth = Sheets_Sheet.Width;
            int panelHeight = Sheets_Sheet.Height;

            // パネルの位置を中央に設定
            Sheets_Sheet.Location = new System.Drawing.Point(
                (parentWidth - panelWidth) / 2 - 10,
                90
            );

            Sheets_Sheet.Top = 59;
        }
        #endregion

        //連絡帳表示処理
        private void kryptonRibbonGroupButton18_Click(object sender, EventArgs e)
        {
            kryptonNavigator_Workbench.SelectedPage = AddressTab;
        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {

            //変更した設定の保存処理を行う
            if(kryptonRadioButton3.Checked == true)
            {
                Properties.Settings.Default.ShowApplicationTask = 0;
            }
            else if (kryptonRadioButton2.Checked == true)
            {
                Properties.Settings.Default.ShowApplicationTask = 1;
            }
            else if (kryptonRadioButton1.Checked == true)
            {
                Properties.Settings.Default.ShowApplicationTask = 2;
            }

            if(kryptonCheckBox4.Checked == true)
            {
                Properties.Settings.Default.IsAvailableDocumentCreationSoftware = true;
            }
            else
            {
                Properties.Settings.Default.IsAvailableDocumentCreationSoftware = false;
            }

            if (kryptonCheckBox7.Checked == true)
            {
                Properties.Settings.Default.IsUseEraName = true;
            }
            else
            {
                Properties.Settings.Default.IsUseEraName = false;
            }

            if(kryptonCheckBox5.Checked == true)
            {
                try
                {
                    Microsoft.Win32.RegistryKey regkey =
                        Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                        @"Software\Microsoft\Windows\CurrentVersion\RunOnce", true);
                    regkey.SetValue(System.Windows.Forms.Application.ProductName, System.Windows.Forms.Application.ExecutablePath);
                    regkey.Close();
                }
                catch { }

                Properties.Settings.Default.IsWindowsStartUpRunForDCMK = true;
            }
            else
            {
                try
                {
                    Microsoft.Win32.RegistryKey regkey =
                        Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                        @"Software\Microsoft\Windows\CurrentVersion\RunOnce", true);
                    regkey.DeleteValue(System.Windows.Forms.Application.ProductName, false);
                    regkey.Close();
                }
                catch { }

                Properties.Settings.Default.IsWindowsStartUpRunForDCMK = false;
            }
            //シートの空白間隔
            Properties.Settings.Default.Space_Top = (int)kryptonNumericUpDown4.Value;
            Properties.Settings.Default.Space_Buttom = (int)kryptonNumericUpDown7.Value;
            Properties.Settings.Default.Space_Left = (int)kryptonNumericUpDown5.Value;
            Properties.Settings.Default.Space_Right = (int)kryptonNumericUpDown6.Value;

            //内容
            Properties.Settings.Default.SendingDepartment = kryptonTextBox16.Text;
            Properties.Settings.Default.To_CompanyOrOrganizationName = kryptonTextBox17.Text;
            Properties.Settings.Default.To_Title = kryptonComboBox12.Text;
            Properties.Settings.Default.To_Name = kryptonTextBox18.Text;
            Properties.Settings.Default.Caller_CompanyOrOrganizationName = kryptonTextBox19.Text;
            Properties.Settings.Default.Caller_Location = kryptonTextBox32.Text;
            Properties.Settings.Default.Caller_BuildingName = kryptonTextBox20.Text;
            Properties.Settings.Default.Caller_FloorNumber = (int)kryptonNumericUpDown3.Value;
            Properties.Settings.Default.Caller_Title = kryptonComboBox13.Text;
            Properties.Settings.Default.Caller_Name = kryptonTextBox21.Text;
            Properties.Settings.Default.Caller_MailAddress_User = kryptonTextBox22.Text;
            Properties.Settings.Default.Caller_MailAddress_Domain = kryptonComboBox14.Text;
            Properties.Settings.Default.Caller_PhoneNumber1 = kryptonComboBox15.Text;
            Properties.Settings.Default.Caller_PhoneNumber2 = kryptonTextBox23.Text;
            Properties.Settings.Default.Caller_PhoneNumber3 = kryptonTextBox24.Text;
            Properties.Settings.Default.Caller_FaxNumber1 = kryptonComboBox16.Text;
            Properties.Settings.Default.Caller_FaxNumber2 = kryptonTextBox26.Text;
            Properties.Settings.Default.Caller_FaxNumber3 = kryptonTextBox25.Text;

            Properties.Settings.Default.Save();
            if (kryptonTrackBar1.Value == 0)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }
            else if (kryptonTrackBar1.Value == 10)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = false;
                kryptonLabel42.Enabled = true;
            }
            else
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }

            kryptonPage2.Visible = false;
            kryptonNavigator_Workbench.NavigatorMode = NavigatorMode.BarTabGroup;
            kryptonNavigator_Workbench.SelectedPage = kryptonPage1;

            kryptonRibbon.MinimizedMode = false;
            kryptonRibbon.Enabled = true;

            kryptonLabel7.Enabled = true;
            kryptonCheckButton1.Enabled = true;
            kryptonCheckButton2.Enabled = true;
            kryptonLabel1.Enabled = true;

            Transition
                .With(kryptonPanel21, nameof(Height), 0)
                .CriticalDamp(TimeSpan.FromSeconds(0.4));

            if (kryptonTextBox10.Text != string.Empty)
            {
                Sheets_TitleButton.Text = kryptonTextBox10.Text;
                this.Text = kryptonTextBox10.Text + " - DocuQuick Designer";
            }
            else
            {
                this.Text = "無題 - DocuQuick Designer";
            }

            kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

            Sheets_Sheet.Anchor = AnchorStyles.Top;

            // 親コントロールのサイズを取得
            int parentWidth = this.ClientSize.Width;
            int parentHeight = this.ClientSize.Height;

            // パネルのサイズを取得
            int panelWidth = Sheets_Sheet.Width;
            int panelHeight = Sheets_Sheet.Height;

            // パネルの位置を中央に設定
            Sheets_Sheet.Location = new System.Drawing.Point(
                (parentWidth - panelWidth) / 2 - 10,
                90
            );

            Sheets_Sheet.Top = 59;
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void kryptonRibbonColorButton_TextColor_SelectedColorChanged(object sender, EventArgs e)
        {
            Sheets_TitleButton.ForeColor = kryptonRibbonColorButton_TextColor.SelectedColor;
            kryptonTextBox10.StateCommon.Content.Color1 = kryptonRibbonColorButton_TextColor.SelectedColor;
        }

        private void kryptonRibbonGroupComboBox_FontSize_TextUpdate(object sender, EventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    Sheets_TitleButton.Font.Name,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void kryptonRibbonGroupComboBox_FontSize_TextUpdate(object sender, KeyEventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    Sheets_TitleButton.Font.Name,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void kryptonRibbonGroupComboBox_FontSize_TextUpdate(object sender, PropertyChangedEventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    Sheets_TitleButton.Font.Name,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        public void FontReset()
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    Sheets_TitleButton.Font.Name,
                    fontSize,
                    FontStyle.Regular
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;

            }
        }

        //太字
        private void kryptonRibbonButton_Bold_Click(object sender, EventArgs e)
        {

            if (kryptonRibbonButton_Bold.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Sheets_TitleButton.Font = new System.Drawing.Font(
                        Sheets_TitleButton.Font.Name,
                        fontSize,
                        Sheets_TitleButton.Font.Style | FontStyle.Bold
                    );
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonRibbonButton_Bold.Checked = true;
                }
            }
            else if (kryptonRibbonButton_Bold.Checked == false)
            {
                //太字ボタンをチェックをオフにする
                kryptonRibbonButton_Bold.Checked = false;
                //フォントスタイルのみ初期化
                FontReset();
                float fontSize;

                //斜体が有効な場合
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Italic
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonRibbonButton_Italic.Checked = true;
                    }
                }

                //下線が有効な場合
                if (kryptonContextMenuItem15.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Underline
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonContextMenuItem15.Checked = true;
                    }
                }

                //打ち消し線が有効な場合
                if (kryptonContextMenuItem16.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Strikeout
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonContextMenuItem16.Checked = true;
                    }
                }
            }
        }

        //斜体
        private void kryptonRibbonButton_Italic_Click(object sender, EventArgs e)
        {

            if (kryptonRibbonButton_Italic.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Sheets_TitleButton.Font = new System.Drawing.Font(
                        Sheets_TitleButton.Font.Name,
                        fontSize,
                        Sheets_TitleButton.Font.Style | FontStyle.Italic
                    );
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonRibbonButton_Italic.Checked = true;
                }
            }
            else if (kryptonRibbonButton_Italic.Checked == false)
            {
                //斜体ボタンをチェックをオフにする
                kryptonRibbonButton_Italic.Checked = false;
                //フォントスタイルのみ初期化
                FontReset();
                float fontSize;

                //太字が有効な場合
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Bold
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonRibbonButton_Bold.Checked = true;
                    }
                }

                //下線が有効な場合
                if (kryptonContextMenuItem15.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Underline
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonContextMenuItem15.Checked = true;
                    }
                }

                //打ち消し線が有効な場合
                if (kryptonContextMenuItem16.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Strikeout
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonContextMenuItem16.Checked = true;
                    }
                }
            }
        }



        private void kryptonRibbonButton_TextLine_Click(object sender, EventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    Sheets_TitleButton.Font.Name,
                    fontSize,
                    Sheets_TitleButton.Font.Style | FontStyle.Underline
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        //下線
        private void kryptonContextMenuItem15_CheckedChanged(object sender, EventArgs e)
        {

            if (kryptonContextMenuItem15.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Sheets_TitleButton.Font = new System.Drawing.Font(
                        Sheets_TitleButton.Font.Name,
                        fontSize,
                        Sheets_TitleButton.Font.Style | FontStyle.Underline
                    );
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonContextMenuItem15.Checked = true;
                }
            }
            else if (kryptonContextMenuItem15.Checked == false)
            {
                //下線メニューアイテムをチェックをオフにする
                kryptonContextMenuItem15.Checked = false;
                //フォントスタイルのみ初期化
                FontReset();
                float fontSize;

                //太字が有効な場合
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Bold
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonRibbonButton_Bold.Checked = true;
                    }
                }

                //斜体が有効な場合
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Italic
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonRibbonButton_Italic.Checked = true;
                    }
                }

                //打ち消し線が有効な場合
                if (kryptonContextMenuItem16.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Strikeout
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonContextMenuItem16.Checked = true;
                    }
                }
            }
        }

        //打ち消し線
        private void kryptonContextMenuItem16_Click(object sender, EventArgs e)
        {

            //打ち消し線が有効な場合
            if (kryptonContextMenuItem16.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Sheets_TitleButton.Font = new System.Drawing.Font(
                        Sheets_TitleButton.Font.Name,
                        fontSize,
                        Sheets_TitleButton.Font.Style | FontStyle.Strikeout
                    );
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonContextMenuItem16.Checked = true;
                }
            }
            else if (kryptonContextMenuItem16.Checked == false)
            {
                //下線メニューアイテムをチェックをオフにする
                kryptonContextMenuItem16.Checked = false;
                //フォントスタイルのみ初期化
                FontReset();
                float fontSize;

                //太字が有効な場合
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Bold
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonRibbonButton_Bold.Checked = true;
                    }
                }

                //斜体が有効な場合
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Italic
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonRibbonButton_Italic.Checked = true;
                    }
                }

                //下線が有効な場合
                if (kryptonContextMenuItem15.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Underline
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonContextMenuItem15.Checked = true;
                    }
                }
            }
        }

        private void kryptonRibbonGroupClusterButton4_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupComboBox_FontSize.Text == "8")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "9";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "9")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "10";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "9")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "10";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "10")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "10.5";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "10.5")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "11";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "11")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "12";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "12")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "14";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "14")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "16";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "16")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "18";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "18")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "20";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "20")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "22";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "22")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "24";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "24")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "26";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "26")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "28";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "28")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "36";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "36")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "48";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "48")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "72";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "72")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "72";
            }
        }

        private void kryptonRibbonGroupClusterButton5_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupComboBox_FontSize.Text == "8")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "8";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "9")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "8";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "10")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "9";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "10.5")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "10";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "11")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "10.5";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "12")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "11";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "14")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "12";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "16")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "14";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "18")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "16";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "20")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "18";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "22")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "20";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "24")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "22";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "26")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "24";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "28")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "26";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "36")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "28";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "48")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "36";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "72")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "48";
            }
        }

        private void kryptonContextMenuItem28_Click(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "出力中...";
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = true;
            Document doc = word.Documents.Add();


            //外枠の余白を設定
            doc.PageSetup.TopMargin = Sheets_TopPanel.Height;
            doc.PageSetup.BottomMargin = Sheets_ButtomPanel.Height;
            doc.PageSetup.LeftMargin = Sheets_LeftPanel.Width;
            doc.PageSetup.RightMargin = Sheets_RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //発行番号
            if (Sheets_NumberLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                paragraph1.Range.Text = Sheets_NumberLabel.Text;
                paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph1.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph1.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph1.Range.InsertParagraphAfter();
            }
            //日付
            if (Sheets_DateLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                paragraph2.Range.Text = Sheets_DateLabel.Text;
                paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph2.Range.InsertParagraphAfter();
            }
            //相手先会社名
            if (Sheets_AddressCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph3 = doc.Paragraphs.Add();
                paragraph3.Range.Text = Sheets_AddressCompanyLabel.Text;
                paragraph3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph3.Range.InsertParagraphAfter();
            }
            //相手先氏名
            if (Sheets_AddressTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                paragraph4.Range.Text = Sheets_AddressTitleAndNameLabel.Text;
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph4.Range.InsertParagraphAfter();
            }
            //発信者会社名
            if (Sheets_CallerCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph5 = doc.Paragraphs.Add();
                paragraph5.Range.Text = Sheets_CallerCompanyLabel.Text;
                paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph5.Range.InsertParagraphAfter();
            }
            //発信者所在地
            if (Sheets_CallerLocationLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph6 = doc.Paragraphs.Add();
                paragraph6.Range.Text = Sheets_CallerLocationLabel.Text;
                paragraph6.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph6.Range.InsertParagraphAfter();
            }
            //発信者建物名と階数
            if (Sheets_BuildingNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph7 = doc.Paragraphs.Add();
                paragraph7.Range.Text = Sheets_BuildingNameLabel.Text;
                paragraph7.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph7.Range.InsertParagraphAfter();
            }
            //発信者氏名
            if (Sheets_CallerTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                paragraph8.Range.Text = Sheets_CallerTitleAndNameLabel.Text;
                paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph8.Range.InsertParagraphAfter();
            }
            //メールアドレス
            if (Sheets_CallerMallAddressLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph9 = doc.Paragraphs.Add();
                paragraph9.Range.Text = "メールアドレス:" + Sheets_CallerMallAddressLabel.Text;
                paragraph9.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph9.Range.InsertParagraphAfter();
            }
            //電話番号
            if (Sheets_CallerTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph10 = doc.Paragraphs.Add();
                paragraph10.Range.Text = "電話番号:" + Sheets_CallerTelLabel.Text;
                paragraph10.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph10.Range.InsertParagraphAfter();
            }
            //Fax番号
            if (Sheets_CallerFaxTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph11 = doc.Paragraphs.Add();
                paragraph11.Range.Text = "Fax番号:" + Sheets_CallerFaxTelLabel.Text;
                paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph11.Range.InsertParagraphAfter();
            }
            //表題
            if (Sheets_TitleButton.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph12 = doc.Paragraphs.Add();
                if(kryptonRibbonButton_Bold.Checked == true)
                {
                    paragraph12.Range.Bold = 1;
                }
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    paragraph12.Range.Italic = 1;
                }
                if (kryptonContextMenuItem15.Checked == true)
                {
                    paragraph12.Range.Underline = WdUnderline.wdUnderlineSingle;
                }
                if (kryptonContextMenuItem16.Checked == true)
                {
                    paragraph12.Range.Font.StrikeThrough = 1;
                }
                paragraph12.Range.Font.Name = Sheets_TitleButton.Font.Name;
                paragraph12.Range.Font.Size = (int)kryptonTextBox10.StateCommon.Content.Font.Size;
                paragraph12.Range.Text = Sheets_TitleButton.Text;
                paragraph12.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                SetWordRangeColor(paragraph12.Range, Sheets_TitleButton.ForeColor);
                paragraph12.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph12.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph12.Range.InsertParagraphAfter();

            }
            //あいさつ文
            if (Sheets_ContentLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph13 = doc.Paragraphs.Add();
                paragraph13.Range.Bold = 0;
                paragraph13.Range.Italic = 0;
                paragraph13.Range.Underline = WdUnderline.wdUnderlineNone;
                paragraph13.Range.Font.StrikeThrough = 0;
                paragraph13.Range.Font.Name = "游明朝";
                paragraph13.Range.Font.Size = 10;
                paragraph13.Range.Text = Sheets_ContentLabel.Text;
                paragraph13.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph13.Range.Font.Color = WdColor.wdColorBlack;
                paragraph13.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph13.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph13.Range.InsertParagraphAfter();
            }
            //内容
            try
            {
                int LinesCount = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox12.Lines[LinesCount];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount = LinesCount + 1;
                    if (LinesCount == kryptonTextBox12.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //結語
            //kryptonRibbonGroupCheckBoxにチェックがない場合に動作する
            if (kryptonRibbonGroupCheckBox1.Checked != true)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph14 = doc.Paragraphs.Add();
                paragraph14.Range.Text = Sheet_ConclusionLabel.Text;
                paragraph14.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph14.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph14.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph14.Range.InsertParagraphAfter();
            }
            kryptonLabel1.Text = "出力完了";
            stausUpdate();
            //記
            if (Sheet_NoteLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph15 = doc.Paragraphs.Add();
                paragraph15.Range.Text = Sheet_NoteLabel.Text;
                paragraph15.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph15.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph15.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph15.Range.InsertParagraphAfter();
            }
            //記し書き
            try
            {
                int LinesCount2 = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox13.Lines[LinesCount2];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount2 = LinesCount2 + 1;
                    if (LinesCount2 == kryptonTextBox13.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //以上
            if (Sheets_EndLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph16 = doc.Paragraphs.Add();
                paragraph16.Range.Text = Sheets_EndLabel.Text;
                paragraph16.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph16.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph16.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph16.Range.InsertParagraphAfter();
            }
            GC.Collect();

            doc.PrintPreview();
        }

        private void kryptonContextMenuItem30_Click(object sender, EventArgs e)
        {
            //Wordを起動
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = true;

            //新しい文書を作成
            Document doc = word.Documents.Add();

            GC.Collect();
        }

        //上
        private void kryptonRibbonGroupNumericUpDown_VerticalSpace_ValueChanged(object sender, EventArgs e)
        {
            Sheets_TopPanel.Height = (int)kryptonRibbonGroupNumericUpDown_VerticalSpace.Value;
        }

        //左
        private void kryptonRibbonGroupNumericUpDown_WidthSpace_ValueChanged(object sender, EventArgs e)
        {
            Sheets_RightPanel.Height = (int)kryptonRibbonGroupNumericUpDown_WidthSpace.Value;
        }

        //下
        private void kryptonRibbonGroupNumericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            Sheets_ButtomPanel.Height = (int)kryptonRibbonGroupNumericUpDown1.Value;
        }

        //右
        private void kryptonRibbonGroupNumericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            Sheets_LeftPanel.Width = (int)kryptonRibbonGroupNumericUpDown2.Value;
        }

        //広い
        private void kryptonContextMenuItem21_Click(object sender, EventArgs e)
        {
            Sheets_TopPanel.Height = 100;
            Sheets_ButtomPanel.Height = 100;
            Sheets_LeftPanel.Width = 200;
            Sheets_RightPanel.Width = 200;
            //上
            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = (int)Sheets_TopPanel.Height;
            //右
            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = (int)Sheets_RightPanel.Width;
            //下
            kryptonRibbonGroupNumericUpDown1.Value = (int)Sheets_ButtomPanel.Height;
            //左」
            kryptonRibbonGroupNumericUpDown2.Value = (int)Sheets_LeftPanel.Width;
        }

        //やや狭い
        private void kryptonContextMenuItem20_Click(object sender, EventArgs e)
        {
            //上
            Sheets_TopPanel.Height = 100;
            //下
            Sheets_ButtomPanel.Height = 100;
            //右
            Sheets_LeftPanel.Width = 75;
            //左
            Sheets_RightPanel.Width = 75;

            //上
            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = (int)Sheets_TopPanel.Height;
            //右
            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = (int)Sheets_RightPanel.Width;
            //下
            kryptonRibbonGroupNumericUpDown1.Value = (int)Sheets_ButtomPanel.Height;
            //左」
            kryptonRibbonGroupNumericUpDown2.Value = (int)Sheets_LeftPanel.Width;
        }

        //標準
        private void kryptonContextMenuItem18_Click(object sender, EventArgs e)
        {
            Sheets_TopPanel.Height = 138;
            Sheets_ButtomPanel.Height = 118;
            Sheets_LeftPanel.Width = 118;
            Sheets_RightPanel.Width = 118;
            //上
            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = (int)Sheets_TopPanel.Height;
            //右
            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = (int)Sheets_RightPanel.Width;
            //下
            kryptonRibbonGroupNumericUpDown1.Value = (int)Sheets_ButtomPanel.Height;
            //左」
            kryptonRibbonGroupNumericUpDown2.Value = (int)Sheets_LeftPanel.Width;
        }

        //狭い
        private void kryptonContextMenuItem19_Click(object sender, EventArgs e)
        {
            //上
            Sheets_TopPanel.Height = 50;
            //下
            Sheets_ButtomPanel.Height = 50;
            //右
            Sheets_LeftPanel.Width = 50;
            //左
            Sheets_RightPanel.Width = 50;

            //上
            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = (int)Sheets_TopPanel.Height;
            //右
            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = (int)Sheets_RightPanel.Width;
            //下
            kryptonRibbonGroupNumericUpDown1.Value = (int)Sheets_ButtomPanel.Height;
            //左」
            kryptonRibbonGroupNumericUpDown2.Value = (int)Sheets_LeftPanel.Width;
        }


        private void kryptonRibbonGroupButton2_Click(object sender, EventArgs e)
        {
        }

        private void kryptonRibbonGroupButton13_Click(object sender, EventArgs e)
        {

            this.Hide();
            kryptonLabel1.Text = "出力中...";
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            //バックグラウンド上でWordを起動する
            word.Visible = false;
            Document doc = word.Documents.Add();


            //外枠の余白を設定
            doc.PageSetup.TopMargin = Sheets_TopPanel.Height;
            doc.PageSetup.BottomMargin = Sheets_ButtomPanel.Height;
            doc.PageSetup.LeftMargin = Sheets_LeftPanel.Width;
            doc.PageSetup.RightMargin = Sheets_RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //発行番号
            if (Sheets_NumberLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                paragraph1.Range.Text = Sheets_NumberLabel.Text;
                paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph1.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph1.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph1.Range.InsertParagraphAfter();
            }
            //日付
            if (Sheets_DateLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                paragraph2.Range.Text = Sheets_DateLabel.Text;
                paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph2.Range.InsertParagraphAfter();
            }
            //相手先会社名
            if (Sheets_AddressCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph3 = doc.Paragraphs.Add();
                paragraph3.Range.Text = Sheets_AddressCompanyLabel.Text;
                paragraph3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph3.Range.InsertParagraphAfter();
            }
            //相手先氏名
            if (Sheets_AddressTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                paragraph4.Range.Text = Sheets_AddressTitleAndNameLabel.Text;
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph4.Range.InsertParagraphAfter();
            }
            //発信者会社名
            if (Sheets_CallerCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph5 = doc.Paragraphs.Add();
                paragraph5.Range.Text = Sheets_CallerCompanyLabel.Text;
                paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph5.Range.InsertParagraphAfter();
            }
            //発信者所在地
            if (Sheets_CallerLocationLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph6 = doc.Paragraphs.Add();
                paragraph6.Range.Text = Sheets_CallerLocationLabel.Text;
                paragraph6.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph6.Range.InsertParagraphAfter();
            }
            //発信者建物名と階数
            if (Sheets_BuildingNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph7 = doc.Paragraphs.Add();
                paragraph7.Range.Text = Sheets_BuildingNameLabel.Text;
                paragraph7.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph7.Range.InsertParagraphAfter();
            }
            //発信者氏名
            if (Sheets_CallerTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                paragraph8.Range.Text = Sheets_CallerTitleAndNameLabel.Text;
                paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph8.Range.InsertParagraphAfter();
            }
            //メールアドレス
            if (Sheets_CallerMallAddressLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph9 = doc.Paragraphs.Add();
                paragraph9.Range.Text = "メールアドレス:" + Sheets_CallerMallAddressLabel.Text;
                paragraph9.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph9.Range.InsertParagraphAfter();
            }
            //電話番号
            if (Sheets_CallerTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph10 = doc.Paragraphs.Add();
                paragraph10.Range.Text = "電話番号:" + Sheets_CallerTelLabel.Text;
                paragraph10.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph10.Range.InsertParagraphAfter();
            }
            //Fax番号
            if (Sheets_CallerFaxTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph11 = doc.Paragraphs.Add();
                paragraph11.Range.Text = "Fax番号:" + Sheets_CallerFaxTelLabel.Text;
                paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph11.Range.InsertParagraphAfter();
            }
            //表題
            if (Sheets_TitleButton.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph12 = doc.Paragraphs.Add();
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    paragraph12.Range.Bold = 1;
                }
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    paragraph12.Range.Italic = 1;
                }
                if (kryptonContextMenuItem15.Checked == true)
                {
                    paragraph12.Range.Underline = WdUnderline.wdUnderlineSingle;
                }
                if (kryptonContextMenuItem16.Checked == true)
                {
                    paragraph12.Range.Font.StrikeThrough = 1;
                }
                paragraph12.Range.Font.Size = (int)kryptonTextBox10.StateCommon.Content.Font.Size;
                paragraph12.Range.Text = Sheets_TitleButton.Text;
                paragraph12.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                SetWordRangeColor(paragraph12.Range, Sheets_TitleButton.ForeColor);
                paragraph12.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph12.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph12.Range.InsertParagraphAfter();

            }
            //あいさつ文
            if (Sheets_ContentLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph13 = doc.Paragraphs.Add();
                paragraph13.Range.Bold = 0;
                paragraph13.Range.Italic = 0;
                paragraph13.Range.Underline = WdUnderline.wdUnderlineNone;
                paragraph13.Range.Font.StrikeThrough = 0;
                paragraph13.Range.Font.Size = 10;
                paragraph13.Range.Text = Sheets_ContentLabel.Text;
                paragraph13.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph13.Range.Font.Color = WdColor.wdColorBlack;
                paragraph13.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph13.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph13.Range.InsertParagraphAfter();
            }
            //内容
            try
            {
                int LinesCount = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox12.Lines[LinesCount];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount = LinesCount + 1;
                    if (LinesCount == kryptonTextBox12.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //結語
            //kryptonRibbonGroupCheckBoxにチェックがない場合に動作する
            if (kryptonRibbonGroupCheckBox1.Checked != true)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph14 = doc.Paragraphs.Add();
                paragraph14.Range.Text = Sheet_ConclusionLabel.Text;
                paragraph14.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph14.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph14.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph14.Range.InsertParagraphAfter();
            }
            kryptonLabel1.Text = "出力完了";
            stausUpdate();
            //記
            if (Sheet_NoteLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph15 = doc.Paragraphs.Add();
                paragraph15.Range.Text = Sheet_NoteLabel.Text;
                paragraph15.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph15.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph15.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph15.Range.InsertParagraphAfter();
            }
            //記し書き
            try
            {
                int LinesCount2 = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox13.Lines[LinesCount2];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount2 = LinesCount2 + 1;
                    if (LinesCount2 == kryptonTextBox13.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //以上
            if (Sheets_EndLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph16 = doc.Paragraphs.Add();
                paragraph16.Range.Text = Sheets_EndLabel.Text;
                paragraph16.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph16.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph16.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph16.Range.InsertParagraphAfter();
            }

            //保存処理
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "ドキュメントファイルを保存する場所を選択 - DocuQuick";
            sd.Filter = "Word 文書 (*.docx)|*.docx";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    doc.SaveAs2(sd.FileName);
                    //Outlook連携処理
                    try
                    {
                        Microsoft.Office.Interop.Outlook.Application outlook = new Microsoft.Office.Interop.Outlook.Application();
                        MailItem mailItem = (MailItem)outlook.CreateItem(OlItemType.olMailItem);
                        mailItem.Subject = "文書送信のご案内";
                        mailItem.Attachments.Add(sd.FileName);
                        mailItem.Display(true);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show("Microsoft Outlook を正しく動作しませんでした。 Microsoft Outlook が正しくインストールされているか確認してください。\r\n\r\nエラー内容:\r\n" + ex.Message, "共有失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("ファイルが正しく保存されませんでした。保存するファイルの場所が適切か文書作成ソフトウェアがインストールされているか確認してください。\r\n\r\nエラー内容:\r\n" + ex.Message, "ファイル保存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

            //保存を確認せず閉じる
            try
            {
                doc.Close(false);
                word.Quit();
            }
            catch { }

            this.Show();


            GC.Collect();
        }

        private void kryptonNumericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            panel15.Height = (int)kryptonNumericUpDown4.Value / 10;
        }

        private void kryptonNumericUpDown7_ValueChanged(object sender, EventArgs e)
        {
            panel16.Height = (int)kryptonNumericUpDown7.Value / 10;
        }

        private void kryptonNumericUpDown5_ValueChanged(object sender, EventArgs e)
        {
            panel13.Width = (int)kryptonNumericUpDown5.Value / 10;
        }

        private void kryptonNumericUpDown6_ValueChanged(object sender, EventArgs e)
        {
            panel14.Width = (int)kryptonNumericUpDown6.Value / 10;
        }

        private void kryptonRibbonGroupButton_Tutorial_Click(object sender, EventArgs e)
        {

            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                useSoftWareWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                useSoftWareWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                useSoftWareWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                useSoftWareWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                useSoftWareWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                useSoftWareWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            useSoftWareWindow.Show();
            kryptonRibbonGroupButton_Tutorial.Enabled = false;
            kryptonContextMenuItem12.Enabled = false;
        }

        public void kryptonContextMenuItem11_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/User233389/Document-Maker/wiki/Docuemt-Maker-%E3%83%A6%E3%83%BC%E3%82%B6%E3%83%BC%E3%82%AC%E3%82%A4%E3%83%89");
        }

        private void kryptonCommandLinkButton2_Click(object sender, EventArgs e)
        {
            ResetWarningTaskDialog resetWarningTaskDialog = new ResetWarningTaskDialog();
            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                resetWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                resetWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                resetWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                resetWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                resetWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                resetWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            resetWarningTaskDialog.ShowDialog();

            if(resetWarningTaskDialog.DialogResult == DialogResult.Yes)
            {
                Properties.Settings.Default.Reset();
                System.Windows.Forms.Application.Restart();
            }
        }

        private void kryptonCommandLinkButton1_Click_1(object sender, EventArgs e)
        {

            Properties.Settings.Default.ShowResetDialog = true;
            Properties.Settings.Default.ShowNotepadWarningPanel = true;
            Properties.Settings.Default.Save();

            DialogResetMessagebox dialogResetMessagebox = new DialogResetMessagebox();
            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                dialogResetMessagebox.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                dialogResetMessagebox.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                dialogResetMessagebox.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                dialogResetMessagebox.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                dialogResetMessagebox.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                dialogResetMessagebox.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            dialogResetMessagebox.ShowDialog();
        }

        private void kryptonRibbonGroupButton_Support_Click(object sender, EventArgs e)
        {
            //GitHubのWebサイトに移動
            System.Diagnostics.Process.Start("https://github.com/User233389/DocuQuick");
        }

        private void kryptonTrackBar1_ValueChanged(object sender, EventArgs e)
        {
            //10の目盛りに合わせてサイズを+50上げる
            if (kryptonTrackBar1.Value == 0)
            {
                Sheets_Sheet.Size = new Size(842, 999);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                //縮小ボタンを無効化
                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "10%";
                kryptonRibbonGroupComboBox1.Text = "10";
            }
            else if (kryptonTrackBar1.Value == 1)
            {
                Sheets_Sheet.Size = new Size(892, 1049);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "20%";
                kryptonRibbonGroupComboBox1.Text = "20";
            }
            else if (kryptonTrackBar1.Value == 2)
            {
                Sheets_Sheet.Size = new Size(942, 1099);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "30%";
                kryptonRibbonGroupComboBox1.Text = "30";
            }
            else if (kryptonTrackBar1.Value == 3)
            {
                Sheets_Sheet.Size = new Size(992, 1149);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "40%";
                kryptonRibbonGroupComboBox1.Text = "40";
            }
            else if (kryptonTrackBar1.Value == 4)
            {
                Sheets_Sheet.Size = new Size(1042, 1199);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "50%";
                kryptonRibbonGroupComboBox1.Text = "50";
            }
            else if (kryptonTrackBar1.Value == 5)
            {
                Sheets_Sheet.Size = new Size(1092, 1249);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "60%";
                kryptonRibbonGroupComboBox1.Text = "60";
            }
            else if (kryptonTrackBar1.Value == 6)
            {
                Sheets_Sheet.Size = new Size(1142, 1299);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "70%";
                kryptonRibbonGroupComboBox1.Text = "70";
            }
            else if (kryptonTrackBar1.Value == 7)
            {
                Sheets_Sheet.Size = new Size(1192, 1349);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "80%";
                kryptonRibbonGroupComboBox1.Text = "80";
            }
            else if (kryptonTrackBar1.Value == 8)
            {
                Sheets_Sheet.Size = new Size(1242, 1399);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "90%";
                kryptonRibbonGroupComboBox1.Text = "90";
            }
            else if (kryptonTrackBar1.Value == 9)
            {
                Sheets_Sheet.Size = new Size(1292, 1449);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "100%";
                kryptonRibbonGroupComboBox1.Text = "100";
            }
            else if (kryptonTrackBar1.Value == 10)
            {
                Sheets_Sheet.Size = new Size(1342, 1499);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                //拡大ボタンを無効化
                kryptonButton14.Enabled = false;

                kryptonLabel42.Text = "110%";
                kryptonRibbonGroupComboBox1.Text = "110";
            }
        }

        private void kryptonButton14_Click(object sender, EventArgs e)
        {
            if (kryptonTrackBar1.Value == kryptonTrackBar1.Value)
            {
                kryptonTrackBar1.Value = kryptonTrackBar1.Value + 1;
            }
        }

        private void kryptonButton15_Click(object sender, EventArgs e)
        {
            if (kryptonTrackBar1.Value == kryptonTrackBar1.Value)
            {
                kryptonTrackBar1.Value = kryptonTrackBar1.Value - 1;
            }
        }

        private void kryptonRibbonGroupComboBox1_TextUpdate(object sender, EventArgs e)
        {
            //10の目盛りに合わせてサイズを+50上げる
            if (kryptonRibbonGroupComboBox1.Text == "10")
            {
                Sheets_Sheet.Size = new Size(842, 999);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                //縮小ボタンを無効化
                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "10%";
                kryptonTrackBar1.Value = 0;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "20")
            {
                Sheets_Sheet.Size = new Size(892, 1049);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "20%";
                kryptonTrackBar1.Value = 1;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "30")
            {
                Sheets_Sheet.Size = new Size(942, 1099);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "30%";
                kryptonTrackBar1.Value = 2;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "40")
            {
                Sheets_Sheet.Size = new Size(992, 1149);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "40%";
                kryptonTrackBar1.Value = 3;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "50")
            {
                Sheets_Sheet.Size = new Size(1042, 1199);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "50%";
                kryptonTrackBar1.Value = 4;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "60")
            {
                Sheets_Sheet.Size = new Size(1092, 1249);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "60%";
                kryptonTrackBar1.Value = 5;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "70")
            {
                Sheets_Sheet.Size = new Size(1142, 1299);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "70%";
                kryptonTrackBar1.Value = 6;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "80")
            {
                Sheets_Sheet.Size = new Size(1192, 1349);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "80%";
                kryptonTrackBar1.Value = 7;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "90")
            {
                Sheets_Sheet.Size = new Size(1242, 1399);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "90%";
                kryptonTrackBar1.Value = 8;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "100")
            {
                Sheets_Sheet.Size = new Size(1292, 1449);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "100%";
                kryptonTrackBar1.Value = 9;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "110")
            {
                Sheets_Sheet.Size = new Size(1342, 1499);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                //拡大ボタンを無効化
                kryptonButton14.Enabled = false;

                kryptonLabel42.Text = "100%";
                kryptonTrackBar1.Value = 10;
            }

        }

        private void kryptonRibbonGroupButton2_Click_1(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton2.Checked == true)
            {
                this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Maximized;
                this.AllowFormChrome = false;
            }
            else
            {
                this.FormBorderStyle = FormBorderStyle.Sizable;
                this.WindowState = FormWindowState.Normal;
                this.AllowFormChrome = true;
            }
        }

        //戻る
        private void kryptonRibbonQATButton7_Click(object sender, EventArgs e)
        {
            if (kryptonNavigator_Workbench.SelectedPage == kryptonPage3)
            {
                kryptonNavigator_Workbench.SelectedPage = AddressTab;
                kryptonRibbonQATButton7.Enabled = true;
            }
            else if (kryptonNavigator_Workbench.SelectedPage == AddressTab)
            {
                kryptonNavigator_Workbench.SelectedPage = kryptonPage1;
                kryptonRibbonQATButton7.Enabled = false;
            }
        }

        //進む
        private void kryptonRibbonQATButton8_Click(object sender, EventArgs e)
        {
            if (kryptonNavigator_Workbench.SelectedPage == kryptonPage1)
            {
                kryptonNavigator_Workbench.SelectedPage = AddressTab;
                kryptonRibbonQATButton7.Enabled = true;
            }
            else if (kryptonNavigator_Workbench.SelectedPage == AddressTab)
            {
                kryptonNavigator_Workbench.SelectedPage = kryptonPage3;
                kryptonRibbonQATButton8.Enabled = false;
            }
        }

        private void kryptonNavigator_Workbench_Selected(object sender, ComponentFactory.Krypton.Navigator.KryptonPageEventArgs e)
        {
            if(kryptonRibbon.Enabled == true)
            {
                if (kryptonNavigator_Workbench.SelectedPage == AddressTab)
                {
                    kryptonRibbon.SelectedContext = "Address";
                    kryptonRibbon.SelectedTab = AddressTab1;
                }
                else if (kryptonNavigator_Workbench.SelectedPage == kryptonPage3)
                {
                    kryptonRibbon.SelectedContext = "Notepad";
                    kryptonRibbon.SelectedTab = NotepadTab;
                    kryptonRibbonButton_Paste.ShortcutKeys = Keys.None;
                    kryptonRibbonGroupButton_NotepadPaste.ShortcutKeys = Keys.Control | Keys.V;

                    //シート
                    kryptonRibbonButton_Bold.ShortcutKeys = Keys.None;
                    kryptonRibbonButton_Italic.ShortcutKeys = Keys.None;
                    kryptonContextMenuItem15.ShortcutKeys = Keys.None;
                    kryptonContextMenuItem16.ShortcutKeys = Keys.None;

                    kryptonRibbonColorButton_TextColor.ShortcutKeys = Keys.None;

                    kryptonRibbonGroupClusterButton4.ShortcutKeys = Keys.None;
                    kryptonRibbonGroupClusterButton5.ShortcutKeys = Keys.None;

                    //メモ
                    kryptonRibbonGroupClusterButton1.ShortcutKeys = Keys.Control | Keys.B;
                    kryptonRibbonGroupClusterButton2.ShortcutKeys = Keys.Control | Keys.I;
                    kryptonContextMenuItem35.ShortcutKeys = Keys.Control | Keys.U;
                    kryptonContextMenuItem36.ShortcutKeys = Keys.Control | Keys.T;

                    kryptonRibbonGroupColorButton2.ShortcutKeys = Keys.Control | Keys.Shift | Keys.C;
                    kryptonRibbonGroupColorButton3.ShortcutKeys = Keys.Control | Keys.Shift | Keys.M;

                    kryptonRibbonGroupClusterButton6.ShortcutKeys = Keys.Control | Keys.Shift | Keys.U;
                    kryptonRibbonGroupClusterButton7.ShortcutKeys = Keys.Control | Keys.Shift | Keys.D;
                }
                else
                {
                    kryptonRibbon.SelectedContext = string.Empty;
                    kryptonRibbonButton_Paste.ShortcutKeys = Keys.Control | Keys.V;
                    kryptonRibbonGroupButton_NotepadPaste.ShortcutKeys = Keys.None;

                    kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                    Sheets_Sheet.Anchor = AnchorStyles.Top;

                    // 親コントロールのサイズを取得
                    int parentWidth = this.ClientSize.Width;
                    int parentHeight = this.ClientSize.Height;

                    // パネルのサイズを取得
                    int panelWidth = Sheets_Sheet.Width;
                    int panelHeight = Sheets_Sheet.Height;

                    // パネルの位置を中央に設定
                    Sheets_Sheet.Location = new System.Drawing.Point(
                        (parentWidth - panelWidth) / 2 - 10,
                        90
                    );

                    //シート
                    kryptonRibbonButton_Bold.ShortcutKeys = Keys.Control | Keys.B;
                    kryptonRibbonButton_Italic.ShortcutKeys = Keys.Control | Keys.I;
                    kryptonContextMenuItem15.ShortcutKeys = Keys.Control | Keys.U;
                    kryptonContextMenuItem16.ShortcutKeys = Keys.Control | Keys.T;

                    kryptonRibbonColorButton_TextColor.ShortcutKeys = Keys.Control | Keys.Shift | Keys.C;

                    kryptonRibbonGroupClusterButton4.ShortcutKeys = Keys.Control | Keys.Shift | Keys.U;
                    kryptonRibbonGroupClusterButton5.ShortcutKeys = Keys.Control | Keys.Shift | Keys.D;

                    //メモ
                    kryptonRibbonGroupClusterButton1.ShortcutKeys = Keys.None;
                    kryptonRibbonGroupClusterButton2.ShortcutKeys = Keys.None;
                    kryptonContextMenuItem35.ShortcutKeys = Keys.None;
                    kryptonContextMenuItem36.ShortcutKeys = Keys.None;

                    kryptonRibbonGroupColorButton2.ShortcutKeys = Keys.None;
                    kryptonRibbonGroupColorButton3.ShortcutKeys = Keys.None;

                    kryptonRibbonGroupClusterButton6.ShortcutKeys = Keys.None;
                    kryptonRibbonGroupClusterButton7.ShortcutKeys = Keys.None;


                }


            }

            Sheets_Sheet.Top = 59;

            if (kryptonNavigator_Workbench.SelectedPage == kryptonPage1)
            {
                kryptonRibbonQATButton7.Enabled = false;
                kryptonRibbonQATButton8.Enabled = true;

                kryptonRibbonQATButton4.Enabled = false;
                kryptonRibbonQATButton5.Enabled = false;
            }
            else if (kryptonNavigator_Workbench.SelectedPage == AddressTab)
            {
                kryptonRibbonQATButton7.Enabled = true;
                kryptonRibbonQATButton8.Enabled = true;

                kryptonRibbonQATButton4.Enabled = false;
                kryptonRibbonQATButton5.Enabled = false;
            }
            else if (kryptonNavigator_Workbench.SelectedPage == kryptonPage3)
            {
                kryptonRibbonQATButton7.Enabled = true;
                kryptonRibbonQATButton8.Enabled = false;

                kryptonRibbonQATButton4.Enabled = true;
                kryptonRibbonQATButton5.Enabled = true;
            }
        }


        //メモ帳
        //元に戻す
        private void kryptonRibbonQATButton4_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Undo();
        }

        //やり直す
        private void kryptonRibbonQATButton5_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Redo();
        }

        //貼り付け
        private void kryptonRibbonGroupButton_NotepadPaste_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Paste();
        }

        private void kryptonRibbonGroupButton_NotepadCopy_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Copy();
        }

        private void kryptonRibbonGroupButton_NotepadCut_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Cut();
        }


        public int rtbLangth { get; set; }
        public int rtbStart { get; set; }

        //フォント変更
        private void kryptonRibbonGroupComboBox_NotepadFont_DropDownClosed(object sender, EventArgs e)
        {

            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);
            // 現在のフォント名を変更する
            Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                kryptonRibbonGroupComboBox_NotepadFont.Text,
                Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size,
                Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style
            );

        }


        //フォントサイズを変更
        private void kryptonRibbonGroupComboBox_NotepadFontSize_DropDownClosed(object sender, EventArgs e)
        {

            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                    fontSize,
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style
                );
            }

        }

        private void Notepads_kryptonRichTextBox_Notepad_SelectionChanged(object sender, EventArgs e)
        {
            //文字選択数取得
            rtbLangth = Notepads_kryptonRichTextBox_Notepad.SelectionLength;
            rtbStart = Notepads_kryptonRichTextBox_Notepad.SelectionStart;

            try
            {
                //文字色の確認
                kryptonRibbonGroupColorButton2.SelectedColor = Notepads_kryptonRichTextBox_Notepad.SelectionColor;
                //マーカー色の確認
                kryptonRibbonGroupColorButton3.SelectedColor = Notepads_kryptonRichTextBox_Notepad.SelectionBackColor;
                //フォントスタイルの確認
                //太字
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Bold)
                {
                    kryptonRibbonGroupClusterButton1.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton1.Checked = false;
                }

                //斜体
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Italic)
                {
                    kryptonRibbonGroupClusterButton2.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton2.Checked = false;
                }

                //下線
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Underline)
                {
                    kryptonContextMenuItem35.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem35.Checked = false;
                }

                //打ち消し線
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Strikeout)
                {
                    kryptonContextMenuItem36.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem35.Checked = false;
                }

                // 選択範囲のフォントを取得（同一フォントでない場合は null を返す）
                System.Drawing.Font selFont = Notepads_kryptonRichTextBox_Notepad.SelectionFont;

                if (selFont != null)
                {
                    // 選択範囲が単一フォントの場合はそのまま反映
                    kryptonRibbonGroupComboBox_NotepadFont.Text = selFont.Name;
                    kryptonRibbonGroupComboBox_NotepadFontSize.Text = selFont.Size.ToString();
                }
                else
                {
                    // フォントが混在している場合は、選択範囲の先頭文字のフォントを取得して表示（元の選択は復元）
                    int selStart = Notepads_kryptonRichTextBox_Notepad.SelectionStart;
                    int selLen = Notepads_kryptonRichTextBox_Notepad.SelectionLength;

                    if (selStart < Notepads_kryptonRichTextBox_Notepad.TextLength && selLen > 0)
                    {
                        // 一時的に先頭1文字を選択してフォントを調べる
                        Notepads_kryptonRichTextBox_Notepad.Select(selStart, 1);
                        System.Drawing.Font firstCharFont = Notepads_kryptonRichTextBox_Notepad.SelectionFont;

                        // 元の選択範囲を復元
                        Notepads_kryptonRichTextBox_Notepad.Select(selStart, selLen);

                        if (firstCharFont != null)
                        {
                            // 「混在」表記を付けて先頭のフォント情報を表示
                            kryptonRibbonGroupComboBox_NotepadFont.Text = firstCharFont.Name + " (混在)";
                            kryptonRibbonGroupComboBox_NotepadFontSize.Text = firstCharFont.Size.ToString();
                            return;
                        }
                    }

                    // 選択なしやフォント情報が取れない場合は空にする
                    kryptonRibbonGroupComboBox_NotepadFont.Text = string.Empty;
                    kryptonRibbonGroupComboBox_NotepadFontSize.Text = string.Empty;
                }
            }
            catch(System.Exception)
            {
                kryptonRibbonGroupComboBox_NotepadFont.Text = "(混在)";
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "(混在)";
            }


        }

        private void Notepads_kryptonRichTextBox_Notepad_MouseUp(object sender, MouseEventArgs e)
        {


        
        }

        private void kryptonRibbonGroupComboBox_NotepadFont_DropDown(object sender, EventArgs e)
        {

        }

        private void Notepads_kryptonRichTextBox_Notepad_FontChanged(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);
        }

        private void kryptonRibbonGroupComboBox_NotepadFont_DropDownClosed(object sender, KeyEventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_NotepadFont_DropDownClosed(object sender, KeyPressEventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_NotepadFont_DropDownClosed(object sender, PropertyChangedEventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_NotepadFontSize_DropDownClosed(object sender, KeyEventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_NotepadFontSize_DropDownClosed(object sender, KeyPressEventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_NotepadFontSize_DropDownClosed(object sender, PropertyChangedEventArgs e)
        {

        }
        public void FontReset2()
        {
            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                    fontSize,
                    FontStyle.Regular
                );
            }
        }

        //太字
        private void kryptonRibbonGroupClusterButton1_Click(object sender, EventArgs e)
        {
            if(kryptonRibbonGroupClusterButton1.Checked == true)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                        fontSize,
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Bold
                    );
                }

                //完了後他のフォントスタイルを確認
                kryptonRibbonGroupClusterButton1.Checked = true;
                if(Notepads_kryptonRichTextBox_Notepad.SelectionFont.Italic)
                {
                    kryptonRibbonGroupClusterButton2.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Underline)
                {
                    kryptonContextMenuItem35.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Strikeout)
                {
                    kryptonContextMenuItem36.Checked = true;
                }

            }
            else if(kryptonRibbonGroupClusterButton1.Checked == false)
            {
                kryptonRibbonGroupClusterButton1.Checked = false;
                FontReset2();

                //斜体が有効な場合
                if (kryptonRibbonGroupClusterButton2.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Italic
                        );
                    }

                    kryptonRibbonGroupClusterButton2.Checked = true;
                }

                //下線が有効な場合
                if (kryptonContextMenuItem35.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Underline
                        );
                    }

                    kryptonContextMenuItem35.Checked = true;
                }

                //打ち消し線が有効な場合
                if (kryptonContextMenuItem36.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Strikeout
                        );
                    }

                    kryptonContextMenuItem36.Checked = true;
                }
            }

        }

        //斜体
        private void kryptonRibbonGroupClusterButton2_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupClusterButton2.Checked == true)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                        fontSize,
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Italic
                    );
                }

                //完了後他のフォントスタイルを確認
                kryptonRibbonGroupClusterButton2.Checked = true;
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Bold)
                {
                    kryptonRibbonGroupClusterButton1.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Underline)
                {
                    kryptonContextMenuItem35.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Strikeout)
                {
                    kryptonContextMenuItem36.Checked = true;
                }
            }
            else if (kryptonRibbonGroupClusterButton2.Checked == false)
            {
                FontReset2();

                //太字が有効な場合
                if (kryptonRibbonGroupClusterButton1.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Bold
                        );
                    }

                    kryptonRibbonGroupClusterButton1.Checked = true;
                }

                //下線が有効な場合
                if (kryptonContextMenuItem35.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Underline
                        );
                    }

                    kryptonContextMenuItem35.Checked = true;
                }

                //打ち消し線が有効な場合
                if (kryptonContextMenuItem36.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Strikeout
                        );
                    }

                    kryptonContextMenuItem36.Checked = true;
                }

            }
        }

        //下線
        private void kryptonContextMenuItem35_Click(object sender, EventArgs e)
        {
            if (kryptonContextMenuItem35.Checked == true)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                        fontSize,
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Underline
                    );
                }

                //完了後他のフォントスタイルを確認
                kryptonContextMenuItem35.Checked = true;
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Bold)
                {
                    kryptonRibbonGroupClusterButton1.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Italic)
                {
                    kryptonRibbonGroupClusterButton2.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Strikeout)
                {
                    kryptonContextMenuItem36.Checked = true;
                }


            }
            else if (kryptonContextMenuItem35.Checked == false)
            {
                FontReset2();
                //太字
                if (kryptonRibbonGroupClusterButton1.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Bold
                        );
                    }

                    kryptonRibbonGroupClusterButton1.Checked = true;
                }

                //斜体
                if (kryptonRibbonGroupClusterButton2.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Italic
                        );
                    }

                    kryptonRibbonGroupClusterButton2.Checked = true;
                }



                //打ち消し線
                if (kryptonContextMenuItem36.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Strikeout
                        );
                    }

                    kryptonContextMenuItem36.Checked = true;
                }
            }
        }

        //打ち消し線
        private void kryptonContextMenuItem36_Click(object sender, EventArgs e)
        {
            if (kryptonContextMenuItem36.Checked == true)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                        fontSize,
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Strikeout
                    );
                }

                //完了後他のフォントスタイルを確認
                kryptonContextMenuItem36.Checked = true;
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Bold)
                {
                    kryptonRibbonGroupClusterButton1.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Italic)
                {
                    kryptonRibbonGroupClusterButton2.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Underline)
                {
                    kryptonContextMenuItem35.Checked = true;
                }
            }
            else if (kryptonContextMenuItem36.Checked == false)
            {
                FontReset2();
                //太字
                if (kryptonRibbonGroupClusterButton1.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Bold
                        );
                    }

                    kryptonRibbonGroupClusterButton1.Checked = true;
                }

                //斜体
                if (kryptonRibbonGroupClusterButton2.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Italic
                        );
                    }

                    kryptonRibbonGroupClusterButton2.Checked = true;
                }

                //下線
                if (kryptonContextMenuItem35.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Underline
                        );
                    }

                    kryptonContextMenuItem35.Checked = true;
                }

            }
        }

        //文字色
        private void kryptonRibbonGroupColorButton2_SelectedColorChanged(object sender, ComponentFactory.Krypton.Toolkit.ColorEventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.SelectionColor = kryptonRibbonGroupColorButton2.SelectedColor;
        }

        //マーカー色
        private void kryptonRibbonGroupColorButton3_SelectedColorChanged(object sender, ComponentFactory.Krypton.Toolkit.ColorEventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.SelectionBackColor = kryptonRibbonGroupColorButton3.SelectedColor;
        }

        private void kryptonRibbonGroupButton_NotepadSaveAs_Click(object sender, EventArgs e)
        {
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "メモを保存する場所を選択...";
            sd.Filter = "リッチテキストファイル (*.rtf)|*.rtf|書式なしテキストファイル(*.txt)|*.txt";
            if(sd.ShowDialog() == DialogResult.OK)
            {
                //rtfファイルだった場合
                if(sd.FilterIndex == 1)
                {
                    Notepads_kryptonRichTextBox_Notepad.SaveFile(sd.FileName);
                }
                //txtファイルだった場合
                else if (sd.FilterIndex == 2)
                {
                    StreamWriter writer = new StreamWriter(sd.FileName);
                    string str = Notepads_kryptonRichTextBox_Notepad.Text;
                    writer.WriteLine(str);
                    writer.Close();
                    writer.Dispose();
                }
            }
        }



        private void kryptonRibbonGroup2_DialogBoxLauncherClick(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

            FontDialog fd = new FontDialog();
            fd.Font = Notepads_kryptonRichTextBox_Notepad.SelectionFont;
            fd.ShowColor = true;
            fd.Color = Notepads_kryptonRichTextBox_Notepad.SelectionColor;

            if (fd.ShowDialog() == DialogResult.OK)
            {
                // 現在のフォント名を変更する
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    fd.Font.Name,
                    fd.Font.Size,
                    fd.Font.Style
                );

                //フォント名を確認する
                kryptonRibbonGroupComboBox_NotepadFont.Text = fd.Font.Name;
                //フォントサイズを確認する
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = fd.Font.Size.ToString();
                //フォントスタイルを確認する
                //太字
                if(Notepads_kryptonRichTextBox_Notepad.SelectionFont.Bold)
                {
                    kryptonRibbonGroupClusterButton1.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton1.Checked = false;
                }
                //斜体
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Italic)
                {
                    kryptonRibbonGroupClusterButton2.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton2.Checked = false;
                }
                //下線
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Underline)
                {
                    kryptonContextMenuItem35.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem35.Checked = false;
                }
                //打ち消し線
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Strikeout)
                {
                    kryptonContextMenuItem36.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem36.Checked = false;
                }
                //文字色を確認する
                kryptonRibbonGroupColorButton2.SelectedColor = fd.Color;

            }
        }

        //フォントサイズを上げる
        private void kryptonRibbonGroupClusterButton6_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

            if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 8)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "9";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 9)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "10";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 10)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "11";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 11)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "12";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 12)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "14";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 14)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "16";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 16)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "18";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 18)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "20";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 20)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "22";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 22)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "24";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 24)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "26";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 26)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "28";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 28)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "36";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 36)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "48";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 48)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "72";
            }
        }

        //フォントサイズを下げる
        private void kryptonRibbonGroupClusterButton7_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

            if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 9)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "8";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 10)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "9";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 11)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "10";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 12)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "11";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 14)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "12";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 16)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "14";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 18)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "16";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 20)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "18";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 22)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "20";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 24)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "22";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 26)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "24";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 28)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "26";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 36)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "28";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 48)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "36";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 72)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "48";
            }
        }

        private void kryptonRibbonGroupButton20_Click(object sender, EventArgs e)
        {
            OpenFileDialog od = new OpenFileDialog();
            od.Title = "画像ファイルを選択...";
            od.Filter = "PNGファイル(*.png)|*.png|JPEGファイル(*.jpeg)|*.jpeg|JPGファイル(*.jpg)|*.jpg";
            if (od.ShowDialog() == DialogResult.OK)
            {
                Clipboard.SetImage(Image.FromFile(od.FileName));
                Notepads_kryptonRichTextBox_Notepad.Paste();
            }
        }


        private void kryptonRibbonGroupButton15_Click(object sender, EventArgs e)
        {
            if(kryptonRibbonGroupButton15.Checked == true)
            {
                kryptonRibbonGroup13.Visible = true;
            }
            else
            {
                kryptonRibbonGroup13.Visible= false;
            }
        }

        private void Notepads_kryptonRichTextBox_Notepad_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.LinkText);
        }

        private void kryptonCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (kryptonCheckBox1.Checked == true)
            {
                DateTime date = kryptonDateTimePicker1.Value.Date;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                //下記のように西暦ではなく和暦として表示するように設定する
                culturejp.DateTimeFormat.Calendar = new JapaneseCalendar();
                Sheets_DateLabel.Text = date.ToString("ggy年M月d日", culturejp);
            }
            else
            {
                DateTime date = kryptonDateTimePicker1.Value.Date;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                Sheets_DateLabel.Text = date.ToString("yyyy年M月d日");
            }
        }

        private void kryptonContextMenuItem53_Click(object sender, EventArgs e)
        {
            if (kryptonTextBox11.Focused == true)
            {
                kryptonTextBox11.Focus();
                kryptonTextBox11.SelectAll();
            }
            else if (kryptonTextBox1.Focused == true)
            {
                kryptonTextBox1.Focus();
                kryptonTextBox1.SelectAll();
            }
            else if (kryptonComboBox10.Focused == true)
            {
                kryptonComboBox10.Focus();
                kryptonComboBox10.SelectAll();
            }
            else if (kryptonTextBox2.Focused == true)
            {
                kryptonTextBox2.Focus();
                kryptonTextBox2.SelectAll();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Focus();
                kryptonTextBox3.SelectAll();
            }
            else if (kryptonTextBox4.Focused == true)
            {
                kryptonTextBox4.Focus();
                kryptonTextBox4.SelectAll();
            }
            else if (kryptonTextBox5.Focused == true)
            {
                kryptonTextBox5.Focus();
                kryptonTextBox5.SelectAll();
            }
            else if (kryptonComboBox9.Focused == true)
            {
                kryptonComboBox9.Focus();
                kryptonComboBox9.SelectAll();
            }
            else if (kryptonTextBox6.Focused == true)
            {
                kryptonTextBox6.Focus();
                kryptonTextBox6.SelectAll();
            }
            else if (kryptonTextBox7.Focused == true)
            {
                kryptonTextBox7.Focus();
                kryptonTextBox7.SelectAll();
            }
            else if (kryptonComboBox8.Focused == true)
            {
                kryptonComboBox8.Focus();
                kryptonComboBox8.SelectAll();
            }
            else if (kryptonComboBox6.Focused == true)
            {
                kryptonTextBox6.Focus();
                kryptonTextBox6.SelectAll();
            }
            else if (kryptonTextBox14.Focused == true)
            {
                kryptonTextBox14.Focus();
                kryptonTextBox14.SelectAll();
            }
            else if (kryptonTextBox8.Focused == true)
            {
                kryptonTextBox8.Focus();
                kryptonTextBox8.SelectAll();
            }
            else if (kryptonComboBox7.Focused == true)
            {
                kryptonComboBox7.Focus();
                kryptonComboBox7.SelectAll();
            }
            else if (kryptonTextBox9.Focused == true)
            {
                kryptonTextBox9.Focus();
                kryptonTextBox9.SelectAll();
            }
            else if (kryptonTextBox15.Focused == true)
            {
                kryptonTextBox15.Focus();
                kryptonTextBox15.SelectAll();
            }
            else if (kryptonTextBox10.Focused == true)
            {
                kryptonTextBox10.Focus();
                kryptonTextBox10.SelectAll();
            }
            else if (kryptonComboBox2.Focused == true)
            {
                kryptonComboBox2.Focus();
                kryptonComboBox2.SelectAll();
            }
            else if (kryptonComboBox11.Focused == true)
            {
                kryptonComboBox11.Focus();
                kryptonComboBox11.SelectAll();
            }
            else if (kryptonComboBox3.Focused == true)
            {
                kryptonComboBox3.Focus();
                kryptonComboBox3.SelectAll();
            }
            else if (kryptonComboBox4.Focused == true)
            {
                kryptonComboBox4.Focus();
                kryptonComboBox4.SelectAll();
            }
            else if (kryptonTextBox12.Focused == true)
            {
                kryptonTextBox12.Focus();
                kryptonTextBox12.SelectAll();
            }
            else if (kryptonComboBox5.Focused == true)
            {
                kryptonComboBox5.Focus();
                kryptonComboBox5.SelectAll();
            }
            else if (kryptonTextBox13.Focused == true)
            {
                kryptonTextBox13.Focus();
                kryptonTextBox13.SelectAll();
            }
        }

        //発行元部署名
        private void kryptonContextMenuItem57_Click(object sender, EventArgs e)
        {
            kryptonTextBox11.Focus();
        }

        //発行番号
        private void kryptonContextMenuItem58_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown1.Focus();
        }

        //日付
        private void kryptonContextMenuItem79_Click(object sender, EventArgs e)
        {
            kryptonDateTimePicker1.Focus();
        }

        //組織および会社名
        private void kryptonContextMenuItem59_Click(object sender, EventArgs e)
        {
            kryptonTextBox1.Focus();
        }

        //肩書きと氏名
        private void kryptonContextMenuItem61_Click(object sender, EventArgs e)
        {
            kryptonComboBox10.Focus();
        }

        //組織および会社名
        private void kryptonContextMenuItem62_Click(object sender, EventArgs e)
        {
            kryptonTextBox3.Focus();
        }

        //所在地
        private void kryptonContextMenuItem63_Click(object sender, EventArgs e)
        {
            kryptonTextBox4.Focus();
        }

        //建物名と階数
        private void kryptonContextMenuItem64_Click(object sender, EventArgs e)
        {
            kryptonTextBox5.Focus();
        }

        //肩書きと氏名
        private void kryptonContextMenuItem65_Click(object sender, EventArgs e)
        {
            kryptonComboBox9.Focus();
        }

        //メールアドレス
        private void kryptonContextMenuItem66_Click(object sender, EventArgs e)
        {
            kryptonTextBox7.Focus();
        }

        //電話番号
        private void kryptonContextMenuItem67_Click(object sender, EventArgs e)
        {
            kryptonComboBox6.Focus();
        }

        //Fax番号
        private void kryptonContextMenuItem68_Click(object sender, EventArgs e)
        {
            kryptonComboBox7.Focus();
        }

        //表題名
        private void kryptonContextMenuItem69_Click(object sender, EventArgs e)
        {
            kryptonTextBox10.Focus();
        }

        //月
        private void kryptonContextMenuItem70_Click(object sender, EventArgs e)
        {
            kryptonComboBox1.Focus();
        }

        //頭語
        private void kryptonContextMenuItem71_Click(object sender, EventArgs e)
        {
            kryptonComboBox2.Focus();
        }

        //候文
        private void kryptonContextMenuItem72_Click(object sender, EventArgs e)
        {
            kryptonComboBox11.Focus();
        }

        //感謝のあいさつ1
        private void kryptonContextMenuItem73_Click(object sender, EventArgs e)
        {
            kryptonComboBox3.Focus();
        }

        //感謝のあいさつ2
        private void kryptonContextMenuItem74_Click(object sender, EventArgs e)
        {
            kryptonComboBox4.Focus();
        }

        //結語
        private void kryptonContextMenuItem75_Click(object sender, EventArgs e)
        {
            kryptonComboBox5.Focus();
        }


        //内容文
        private void kryptonContextMenuItem76_Click(object sender, EventArgs e)
        {
            kryptonTextBox12.Focus();
        }

        //記し書き文
        private void kryptonContextMenuItem77_Click(object sender, EventArgs e)
        {
            kryptonTextBox13.Focus();
        }

        //ToDo
        private void kryptonContextMenuItem37_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・ToDo\r\n1.\r\n2.\r\n3.";
            kryptonRibbonGroupButton5.TextLine1 = "・ToDo";
        }

        //やることリスト
        private void kryptonContextMenuItem38_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・やることリスト\r\n1.\r\n2.\r\n3.";
            kryptonRibbonGroupButton5.TextLine1 = "・やることリスト";
        }

        //宛先
        private void kryptonContextMenuItem39_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・宛先";
            kryptonRibbonGroupButton5.TextLine1 = "・宛先";
        }

        //発信者
        private void kryptonContextMenuItem40_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・発信者";
            kryptonRibbonGroupButton5.TextLine1 = "・発信者";
        }

        //表題
        private void kryptonContextMenuItem41_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・表題";
            kryptonRibbonGroupButton5.TextLine1 = "・表題";
        }

        //内容と記し書き
        private void kryptonContextMenuItem42_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・内容\r\n\r\n・記し書き";
            kryptonRibbonGroupButton5.TextLine1 = "・内容と記し書き";
        }

        //概要
        private void kryptonContextMenuItem43_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・概要";
            kryptonRibbonGroupButton5.TextLine1 = "・概要";
        }

        //要点
        private void kryptonContextMenuItem44_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・要点";
            kryptonRibbonGroupButton5.TextLine1 = "・要点";
        }

        //注意
        private void kryptonContextMenuItem45_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・注意";
            kryptonRibbonGroupButton5.TextLine1 = "・注意";
        }

        private void kryptonRibbonGroupButton5_Click(object sender, EventArgs e)
        {
            if(kryptonRibbonGroupButton5.TextLine1 == "・ToDo")
            {
                Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・ToDo\r\n1.\r\n2.\r\n3.";
            }
            else if(kryptonRibbonGroupButton5.TextLine1 == "・やることリスト")
            {
                Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・やることリスト\r\n1.\r\n2.\r\n3.";
            }
            else if(kryptonRibbonGroupButton5.TextLine1 == "・内容と記し書き")
            {
                Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・内容\r\n\r\n・記し書き";
            }
            else
            {
                Notepads_kryptonRichTextBox_Notepad.Text += "\r\n" + kryptonRibbonGroupButton5.TextLine1;
            }

        }

        //最高
        private void kryptonContextMenuItem46_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・最高";
            kryptonRibbonGroupButton8.TextLine1 = "・最高";
        }

        //高
        private void kryptonContextMenuItem47_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・高";
            kryptonRibbonGroupButton8.TextLine1 = "・高";
        }

        //中
        private void kryptonContextMenuItem48_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・中";
            kryptonRibbonGroupButton8.TextLine1 = "・中";
        }

        //小
        private void kryptonContextMenuItem49_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・小";
            kryptonRibbonGroupButton8.TextLine1 = "・小";
        }

        //緊急
        private void kryptonContextMenuItem50_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・緊急";
            kryptonRibbonGroupButton8.TextLine1 = "・緊急";
        }

        //要確認
        private void kryptonContextMenuItem51_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・要確認";
            kryptonRibbonGroupButton8.TextLine1 = "・要確認";
        }

        //状態
        private void kryptonContextMenuItem52_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・状態";
            kryptonRibbonGroupButton8.TextLine1 = "・状態";
        }

        private void kryptonRibbonGroupButton8_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n" + kryptonRibbonGroupButton5.TextLine1;
        }

        private void kryptonContextMenuItem92_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.SelectAll();
        }

        //ショートカット
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Control && e.KeyCode == Keys.F9)
            {
                kryptonNavigator_Workbench.SelectedPage = kryptonPage1;
            }
            else if (e.Control && e.KeyCode == Keys.F11)
            {
                kryptonNavigator_Workbench.SelectedPage = AddressTab;
            }
            else if (e.Control && e.KeyCode == Keys.F12)
            {
                kryptonNavigator_Workbench.SelectedPage = kryptonPage3;
            }

            if(e.Control && e.Shift &&  e.KeyCode == Keys.W)
            {
                kryptonLabel1.Text = "出力中...";
                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                word.Visible = true;
                Document doc = word.Documents.Add();


                //外枠の余白を設定
                doc.PageSetup.TopMargin = Sheets_TopPanel.Height;
                doc.PageSetup.BottomMargin = Sheets_ButtomPanel.Height;
                doc.PageSetup.LeftMargin = Sheets_LeftPanel.Width;
                doc.PageSetup.RightMargin = Sheets_RightPanel.Width;

                foreach (Range range in doc.StoryRanges)
                {
                    range.Font.Size = 10; // フォントサイズを10に設定
                }

                //発行番号
                if (Sheets_NumberLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                    paragraph1.Range.Text = Sheets_NumberLabel.Text;
                    paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph1.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph1.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph1.Range.InsertParagraphAfter();
                }
                //日付
                if (Sheets_DateLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                    paragraph2.Range.Text = Sheets_DateLabel.Text;
                    paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph2.Range.InsertParagraphAfter();
                }
                //相手先会社名
                if (Sheets_AddressCompanyLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph3 = doc.Paragraphs.Add();
                    paragraph3.Range.Text = Sheets_AddressCompanyLabel.Text;
                    paragraph3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph3.Range.InsertParagraphAfter();
                }
                //相手先氏名
                if (Sheets_AddressTitleAndNameLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                    paragraph4.Range.Text = Sheets_AddressTitleAndNameLabel.Text;
                    paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph4.Range.InsertParagraphAfter();
                }
                //発信者会社名
                if (Sheets_CallerCompanyLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph5 = doc.Paragraphs.Add();
                    paragraph5.Range.Text = Sheets_CallerCompanyLabel.Text;
                    paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph5.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph5.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph5.Range.InsertParagraphAfter();
                }
                //発信者所在地
                if (Sheets_CallerLocationLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph6 = doc.Paragraphs.Add();
                    paragraph6.Range.Text = Sheets_CallerLocationLabel.Text;
                    paragraph6.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph6.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph6.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph6.Range.InsertParagraphAfter();
                }
                //発信者建物名と階数
                if (Sheets_BuildingNameLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph7 = doc.Paragraphs.Add();
                    paragraph7.Range.Text = Sheets_BuildingNameLabel.Text;
                    paragraph7.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph7.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph7.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph7.Range.InsertParagraphAfter();
                }
                //発信者氏名
                if (Sheets_CallerTitleAndNameLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                    paragraph8.Range.Text = Sheets_CallerTitleAndNameLabel.Text;
                    paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph8.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph8.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph8.Range.InsertParagraphAfter();
                }
                //メールアドレス
                if (Sheets_CallerMallAddressLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph9 = doc.Paragraphs.Add();
                    paragraph9.Range.Text = "メールアドレス:" + Sheets_CallerMallAddressLabel.Text;
                    paragraph9.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph9.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph9.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph9.Range.InsertParagraphAfter();
                }
                //電話番号
                if (Sheets_CallerTelLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph10 = doc.Paragraphs.Add();
                    paragraph10.Range.Text = "電話番号:" + Sheets_CallerTelLabel.Text;
                    paragraph10.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph10.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph10.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph10.Range.InsertParagraphAfter();
                }
                //Fax番号
                if (Sheets_CallerFaxTelLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph11 = doc.Paragraphs.Add();
                    paragraph11.Range.Text = "Fax番号:" + Sheets_CallerFaxTelLabel.Text;
                    paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph11.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph11.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph11.Range.InsertParagraphAfter();
                }
                //表題
                if (Sheets_TitleButton.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph12 = doc.Paragraphs.Add();
                    if (kryptonRibbonButton_Bold.Checked == true)
                    {
                        paragraph12.Range.Bold = 1;
                    }
                    if (kryptonRibbonButton_Italic.Checked == true)
                    {
                        paragraph12.Range.Italic = 1;
                    }
                    if (kryptonContextMenuItem15.Checked == true)
                    {
                        paragraph12.Range.Underline = WdUnderline.wdUnderlineSingle;
                    }
                    if (kryptonContextMenuItem16.Checked == true)
                    {
                        paragraph12.Range.Font.StrikeThrough = 1;
                    }
                    paragraph12.Range.Font.Name = Sheets_TitleButton.Font.Name;
                    paragraph12.Range.Font.Size = (int)kryptonTextBox10.StateCommon.Content.Font.Size;
                    paragraph12.Range.Text = Sheets_TitleButton.Text;
                    paragraph12.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    SetWordRangeColor(paragraph12.Range, Sheets_TitleButton.ForeColor);
                    paragraph12.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph12.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph12.Range.InsertParagraphAfter();

                }
                //あいさつ文
                if (Sheets_ContentLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph13 = doc.Paragraphs.Add();
                    paragraph13.Range.Bold = 0;
                    paragraph13.Range.Italic = 0;
                    paragraph13.Range.Underline = WdUnderline.wdUnderlineNone;
                    paragraph13.Range.Font.StrikeThrough = 0;
                    paragraph13.Range.Font.Name = "游明朝";
                    paragraph13.Range.Font.Size = 10;
                    paragraph13.Range.Text = Sheets_ContentLabel.Text;
                    paragraph13.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    paragraph13.Range.Font.Color = WdColor.wdColorBlack;
                    paragraph13.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph13.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph13.Range.InsertParagraphAfter();
                }
                //内容
                try
                {
                    int LinesCount = 0;
                    while (true)
                    {
                        Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                        W_Contents.Range.Text = kryptonTextBox12.Lines[LinesCount];
                        W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                        W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                        W_Contents.Range.InsertParagraphAfter();
                        LinesCount = LinesCount + 1;
                        if (LinesCount == kryptonTextBox12.Lines.Length)
                        {
                            break;
                        }
                    }
                }
                catch { }
                //結語
                //kryptonRibbonGroupCheckBoxにチェックがない場合に動作する
                if (kryptonRibbonGroupCheckBox1.Checked != true)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph14 = doc.Paragraphs.Add();
                    paragraph14.Range.Text = Sheet_ConclusionLabel.Text;
                    paragraph14.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph14.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph14.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph14.Range.InsertParagraphAfter();
                }
                kryptonLabel1.Text = "出力完了";
                stausUpdate();
                //記
                if (Sheet_NoteLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph15 = doc.Paragraphs.Add();
                    paragraph15.Range.Text = Sheet_NoteLabel.Text;
                    paragraph15.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    paragraph15.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph15.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph15.Range.InsertParagraphAfter();
                }
                //記し書き
                try
                {
                    int LinesCount2 = 0;
                    while (true)
                    {
                        Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                        W_Contents.Range.Text = kryptonTextBox13.Lines[LinesCount2];
                        W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                        W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                        W_Contents.Range.InsertParagraphAfter();
                        LinesCount2 = LinesCount2 + 1;
                        if (LinesCount2 == kryptonTextBox13.Lines.Length)
                        {
                            break;
                        }
                    }
                }
                catch { }
                //以上
                if (Sheets_EndLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph16 = doc.Paragraphs.Add();
                    paragraph16.Range.Text = Sheets_EndLabel.Text;
                    paragraph16.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph16.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph16.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph16.Range.InsertParagraphAfter();
                }
                GC.Collect();

                doc.PrintPreview();
            }
        }

        private void kryptonNavigator_Workbench_KeyDown(object sender, KeyEventArgs e)
        {

        }


        private void buttonSpecAppMenu2_Click(object sender, EventArgs e)
        {

            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                keboradShortCut.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                keboradShortCut.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                keboradShortCut.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                keboradShortCut.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                keboradShortCut.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                keboradShortCut.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            keboradShortCut.Show();
        }
        public string flName { get; set; }
        public string loaction {  get; set; }

        public string ContactEmailAddress { get; set; }
        public string MailAddress_User {get;set;}
        public string MailAddress_Domain { get;set;}

        public string PhoneNumber1 { get;set;}
        public string PhoneNumber2 { get;set;}
        public string PhoneNumber3 { get;set;}

        public string FaxNumber1 { get;set;}
        public string FaxNumber2 { get;set;}
        public string FacNumber3 { get;set;}

        private void kryptonListBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            kryptonButton6.Show();
            // Outlookアプリケーションのインスタンスを取得
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace outlookNs = outlookApp.GetNamespace("MAPI");

            // 連絡先フォルダを取得
            Microsoft.Office.Interop.Outlook.MAPIFolder contactsFolder = outlookNs.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

            // 検索したい名前
            string targetName = kryptonListBox3.SelectedItem.ToString();

            // 連絡先を検索
            Microsoft.Office.Interop.Outlook.Items contactItems = contactsFolder.Items;
            Microsoft.Office.Interop.Outlook.ContactItem contact = contactItems.Find($"[FullName] = '{targetName}'") as Microsoft.Office.Interop.Outlook.ContactItem;

            if (contact != null)
            {

                loaction = contact.BusinessAddress;

                Address_NameLabel.Text = kryptonListBox3.SelectedItem.ToString();
                kryptonCheckBox10.Text = "所在地:〒" + contact.BusinessAddress;
                kryptonCheckBox11.Text = "メールアドレス:" + contact.Email1Address;
                ContactEmailAddress = contact.Email1Address;
                kryptonCheckBox12.Text = "会社電話番号:" + contact.BusinessTelephoneNumber;
                kryptonCheckBox13.Text = "会社Fax番号:" + contact.BusinessFaxNumber;

                //メールアドレス
                // 変更箇所: kryptonListBox3_SelectedIndexChanged 内の contact.Email1Address を分割してユーザー名とドメイン名を文字列に格納する処理
                // 以下を該当メソッド内の該当行（kryptonRadioButton4.Text = ... の代わり）に置き換えてください。

                // フルアドレスをプロパティに保管（既存プロパティを活用）
                ContactEmailAddress = contact.Email1Address ?? string.Empty;

                // ユーザー名とドメイン名を分割して格納
                MailAddress_User = string.Empty;
                MailAddress_Domain = string.Empty;
                string email = ContactEmailAddress.Trim();

                if (!string.IsNullOrEmpty(email))
                {
                    int at = email.IndexOf('@');
                    if (at > 0 && at < email.Length - 1)
                    {
                        MailAddress_User = email.Substring(0, at);
                        MailAddress_Domain = email.Substring(at + 1);
                    }
                    else
                    {
                        // @ が無い・不正な形式の場合は全体をユーザー部として扱う
                        MailAddress_User = email;
                        MailAddress_Domain = string.Empty;
                    }
                }

                //電話番号
                // 会社電話番号を最大3つのパートに分割して格納
                PhoneNumber1 = PhoneNumber2 = PhoneNumber3 = string.Empty;
                string tel = (contact.BusinessTelephoneNumber ?? string.Empty).Trim();

                if (!string.IsNullOrEmpty(tel))
                {
                    // 数字以外で分割（ハイフンやスペース、括弧などを区切りにする）
                    string[] rawParts = System.Text.RegularExpressions.Regex.Split(tel, @"\D+");
                    System.Collections.Generic.List<string> parts = new System.Collections.Generic.List<string>();
                    foreach (var p in rawParts)
                    {
                        if (!string.IsNullOrEmpty(p)) parts.Add(p);
                    }

                    if (parts.Count >= 3)
                    {
                        PhoneNumber1 = parts[0];
                        PhoneNumber2 = parts[1];
                        PhoneNumber3 = parts[2];
                    }
                    else if (parts.Count == 2)
                    {
                        PhoneNumber1 = parts[0];
                        PhoneNumber2 = parts[1];
                        PhoneNumber3 = string.Empty;
                    }
                    else if (parts.Count == 1)
                    {
                        // 桁数に応じて分割（簡易フォールバック）
                        string digits = parts[0];
                        int len = digits.Length;
                        if (len >= 7)
                        {
                            // 末尾4桁を最後のパートに確保し、先頭を残りで分割
                            int last = 4;
                            int first = Math.Max(2, len - 7); // 最低2桁を先頭に
                            int middle = len - first - last;
                            if (first > 0 && middle >= 0)
                            {
                                PhoneNumber1 = digits.Substring(0, first);
                                if (middle > 0) PhoneNumber2 = digits.Substring(first, middle);
                                PhoneNumber3 = digits.Substring(len - last);
                            }
                            else
                            {
                                PhoneNumber1 = digits;
                            }
                        }
                        else
                        {
                            // 小さい桁数は全体をPhoneNumber1へ
                            PhoneNumber1 = digits;
                        }
                    }

                    //Fax番号
                    ScanFaxNumber();

                }
            }
        }

        public void ScanFaxNumber()
        {
            kryptonButton6.Show();
            // Outlookアプリケーションのインスタンスを取得
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace outlookNs = outlookApp.GetNamespace("MAPI");

            // 連絡先フォルダを取得
            Microsoft.Office.Interop.Outlook.MAPIFolder contactsFolder = outlookNs.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

            // 検索したい名前
            string targetName = kryptonListBox3.SelectedItem.ToString();

            // 連絡先を検索
            Microsoft.Office.Interop.Outlook.Items contactItems = contactsFolder.Items;
            Microsoft.Office.Interop.Outlook.ContactItem contact = contactItems.Find($"[FullName] = '{targetName}'") as Microsoft.Office.Interop.Outlook.ContactItem;

            // 例: kryptonListBox3_SelectedIndexChanged 内で contact.BusinessFaxNumber を分割して格納
            FaxNumber1 = FaxNumber2 = FacNumber3 = string.Empty;
            string fax = (contact.BusinessFaxNumber ?? string.Empty).Trim();

            if (!string.IsNullOrEmpty(fax))
            {
                // 数字以外で分割（ハイフンやスペース、括弧などを区切りにする）
                string[] rawParts = System.Text.RegularExpressions.Regex.Split(fax, @"\D+");
                System.Collections.Generic.List<string> parts = new System.Collections.Generic.List<string>();
                foreach (var p in rawParts)
                {
                    if (!string.IsNullOrEmpty(p)) parts.Add(p);
                }

                if (parts.Count >= 3)
                {
                    FaxNumber1 = parts[0];
                    FaxNumber2 = parts[1];
                    FacNumber3 = parts[2];
                }
                else if (parts.Count == 2)
                {
                    FaxNumber1 = parts[0];
                    FaxNumber2 = parts[1];
                    FacNumber3 = string.Empty;
                }
                else if (parts.Count == 1)
                {
                    // 桁数に応じて分割（簡易フォールバック）
                    string digits = parts[0];
                    int len = digits.Length;
                    if (len >= 7)
                    {
                        int last = 4;
                        int first = Math.Max(2, len - 7);
                        int middle = len - first - last;
                        if (first > 0 && middle >= 0)
                        {
                            FaxNumber1 = digits.Substring(0, first);
                            if (middle > 0) FaxNumber2 = digits.Substring(first, middle);
                            FacNumber3 = digits.Substring(len - last);
                        }
                        else
                        {
                            FaxNumber1 = digits;
                        }
                    }
                    else
                    {
                        FaxNumber1 = digits;
                    }
                }
            }
        }

        private void kryptonButton16_Click(object sender, EventArgs e)
        {
            ContactsAuth();
        }

        private void kryptonButton6_Click(object sender, EventArgs e)
        {
            kryptonNavigator_Workbench.SelectedPage = kryptonPage1;
            //名前
            //組織だった場合
            if(kryptonRadioButton4.Checked == true)
            {
                if (Address_NameLabel.Text != "名前を選択してください。")
                {
                    kryptonTextBox3.Text = Address_NameLabel.Text;
                }
                else if (Address_NameLabel.Text != string.Empty)
                {
                    kryptonTextBox3.Text = Address_NameLabel.Text;
                }
            }
            //名前だった場合
            else if(kryptonRadioButton5.Checked == true)
            {
                if (Address_NameLabel.Text != "名前を選択してください。")
                {
                    kryptonTextBox6.Text = Address_NameLabel.Text;
                }
                else if (Address_NameLabel.Text != string.Empty)
                {
                    kryptonTextBox6.Text = Address_NameLabel.Text;
                }
            }

            //会社場所
            if (kryptonCheckBox10.Text != "所在地:〒")
            {
                if (kryptonCheckBox10.Checked == true)
                {
                    kryptonTextBox4.Text = "〒"+loaction;
                }
            }
            //メールアドレス
            if(kryptonCheckBox11.Text != "メールアドレス:")
            {
                if(kryptonCheckBox11.Checked == true)
                {
                    kryptonTextBox7.Text = MailAddress_User;
                    kryptonComboBox8.Text = MailAddress_Domain;
                }
            }
            //電話番号
            if(kryptonCheckBox12.Text != "会社電話番号:")
            {
                if(kryptonCheckBox12.Checked == true)
                {
                    kryptonComboBox6.Text = PhoneNumber1;
                    kryptonTextBox14.Text = PhoneNumber2;
                    kryptonTextBox8.Text = PhoneNumber3;
                }
            }
            //Fax番号
            if(kryptonCheckBox13.Text != "Fax番号:")
            {
                if(kryptonCheckBox13.Checked == true)
                {
                    kryptonComboBox7.Text = FaxNumber1;
                    kryptonTextBox9.Text = FaxNumber2;
                    kryptonTextBox15.Text = FacNumber3;
                }
            }
        }


        async System.Threading.Tasks.Task ContactsAuth()
        {

            var availableWindowsHello = await UserConsentVerifier.CheckAvailabilityAsync();
            if(availableWindowsHello != UserConsentVerifierAvailability.Available)
            {
                kryptonButton16.Enabled = false;
                kryptonLabel45.Visible = true;
            }
            else
            {
                var result = await UserConsentVerifier.RequestVerificationAsync("Microsoft Outlook の連絡先を取得・使用するには認証してください。");

                if (result == UserConsentVerificationResult.Verified)
                {
                    //認証出来た場合
                    kryptonPanel20.Hide();

                    Address_NameLabel.Show();
                    kryptonPanel19.Show();

                    kryptonRibbonGroupButton_AddContact.Enabled = true;
                    kryptonRibbonGroupButton_RemoveContact.Enabled = true;
                    kryptonRibbonGroupButton_UpdateContacts.Enabled = true;
                    kryptonRibbonGroupButton21.Enabled = true;

                    //連絡先取得処理
                    // Outlookアプリケーションを初期化
                    Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
                    NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

                    // 連絡先フォルダを取得
                    MAPIFolder contactsFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

                    // 連絡先アイテムを取得
                    Items contactItems = contactsFolder.Items;

                    // 連絡先をループで表示
                    foreach (object item in contactItems)
                    {
                        if (item is ContactItem contact)
                        {
                            kryptonListBox3.Items.Add(contact.FullName);
                        }
                    }
                }
                else
                {
                    //認証をキャンセルした場合
                    kryptonPanel20.Show();

                    Address_NameLabel.Hide();
                    kryptonButton6.Hide();
                    kryptonPanel19.Hide();

                    kryptonRibbonGroupButton_AddContact.Enabled = false;
                    kryptonRibbonGroupButton_RemoveContact.Enabled = false;
                    kryptonRibbonGroupButton_UpdateContacts.Enabled = false;
                    kryptonRibbonGroupButton21.Enabled = false;
                }
            }

        }

        private void kryptonCheckBox11_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonGroupButton21_Click(object sender, EventArgs e)
        {
            //認証をキャンセルした場合
            kryptonPanel20.Show();

            Address_NameLabel.Hide();
            kryptonButton6.Hide();
            kryptonPanel19.Hide();

            kryptonRibbonGroupButton_AddContact.Enabled = false;
            kryptonRibbonGroupButton_RemoveContact.Enabled = false;
            kryptonRibbonGroupButton_UpdateContacts.Enabled = false;
            kryptonRibbonGroupButton21.Enabled = false;


            kryptonButton6.Visible = false;
            kryptonListBox3.Items.Clear();

            Address_NameLabel.Text = "名前を選択してください。";
            kryptonCheckBox10.Text = "所在地:";
            kryptonCheckBox11.Text = "メールアドレス:";
            kryptonCheckBox12.Text = "会社電話番号:";
            kryptonCheckBox13.Text = "会社Fax番号:";
        }

        private void kryptonRibbonGroupButton_UpdateContacts_Click(object sender, EventArgs e)
        {
            ContactUpDate();
        }


        public void ContactUpDate()
        {
            kryptonListBox3.Items.Clear();
            //連絡先取得処理
            // Outlookアプリケーションを初期化
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

            // 連絡先フォルダを取得
            MAPIFolder contactsFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

            // 連絡先アイテムを取得
            Items contactItems = contactsFolder.Items;

            // 連絡先をループで表示
            foreach (object item in contactItems)
            {
                if (item is ContactItem contact)
                {
                    kryptonListBox3.Items.Add(contact.FullName);
                }
            }
        }

        private void kryptonRibbonGroupButton_AddContact_Click(object sender, EventArgs e)
        {
            // Outlookアプリケーションのインスタンス取得
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();

            // 新規連絡先アイテムを作成
            Microsoft.Office.Interop.Outlook.ContactItem contact =
                outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olContactItem)
                as Microsoft.Office.Interop.Outlook.ContactItem;

            if (contact != null)
            {
                // 追加画面（Outlookの連絡先フォーム）を表示
                contact.Display(true); // true: モーダル表示, false: 非モーダル
                
            }
            ContactUpDate();
        }

        private void kryptonRibbonGroupButton_RemoveContact_Click(object sender, EventArgs e)
        {
            DeleteOutlookContactByName(flName);
        }

        public void DeleteOutlookContactByName(string fullName)
        {
            if(kryptonListBox3.SelectedItem != null)
            {
                DialogResult result = MessageBox.Show(kryptonListBox3.SelectedItem.ToString() + "の連絡先情報を削除しようとしています。この操作は Microsoft Outlook でも連絡先情報が完全に削除され、元に戻すことはできません。\r\nよろしいですか?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                {
                    if (result == DialogResult.Yes)
                    {
                        fullName = kryptonListBox3.SelectedItem.ToString();
                        // Outlookアプリケーションのインスタンス取得
                        Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
                        Microsoft.Office.Interop.Outlook.NameSpace outlookNs = outlookApp.GetNamespace("MAPI");

                        // 連絡先フォルダを取得
                        Microsoft.Office.Interop.Outlook.MAPIFolder contactsFolder =
                            outlookNs.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

                        // 連絡先アイテムを検索
                        Microsoft.Office.Interop.Outlook.Items contactItems = contactsFolder.Items;
                        Microsoft.Office.Interop.Outlook.ContactItem contact =
                            contactItems.Find($"[FullName] = '{fullName}'") as Microsoft.Office.Interop.Outlook.ContactItem;

                        if (contact != null)
                        {
                            contact.Delete(); // 連絡先を削除
                            MessageBox.Show($"「{fullName}」の連絡先を削除しました。", "削除完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            ContactUpDate();
                        }
                        else
                        {
                            MessageBox.Show($"「{fullName}」の連絡先が見つかりません。", "削除失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            ContactUpDate();
                        }
                    }
                }

            }
            else
            {
                MessageBox.Show("削除する連絡先の名前を選択してください。","ノート",MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void Notepads_kryptonRichTextBox_Notepad_TextChanged_2(object sender, EventArgs e)
        {

        }

        public void AutoSave()
        {
            String str = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)+ @"\DocuQuick";

            if (Directory.Exists(str))
            {
                Notepads_kryptonRichTextBox_Notepad.SaveFile(str+@"\SaveFile.rtf");
            }
            else
            {
                //フォルダを作成してからファイルを保存
                Directory.CreateDirectory(str);
                Notepads_kryptonRichTextBox_Notepad.SaveFile(str + @"\SaveFile.rtf");
            }
        }

        private void kryptonRibbonGroupButton1_NotepadShowExplorer_Click(object sender, EventArgs e)
        {
            String str = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DocuQuick";
            if (Directory.Exists(str))
            {
                System.Diagnostics.Process.Start("explorer.exe", str);
            }
        }

        //印刷
        private void kryptonContextMenuItem29_Click(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "出力中...";
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = false;
            Document doc = word.Documents.Add();


            //外枠の余白を設定
            doc.PageSetup.TopMargin = Sheets_TopPanel.Height;
            doc.PageSetup.BottomMargin = Sheets_ButtomPanel.Height;
            doc.PageSetup.LeftMargin = Sheets_LeftPanel.Width;
            doc.PageSetup.RightMargin = Sheets_RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //発行番号
            if (Sheets_NumberLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                paragraph1.Range.Text = Sheets_NumberLabel.Text;
                paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph1.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph1.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph1.Range.InsertParagraphAfter();
            }
            //日付
            if (Sheets_DateLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                paragraph2.Range.Text = Sheets_DateLabel.Text;
                paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph2.Range.InsertParagraphAfter();
            }
            //相手先会社名
            if (Sheets_AddressCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph3 = doc.Paragraphs.Add();
                paragraph3.Range.Text = Sheets_AddressCompanyLabel.Text;
                paragraph3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph3.Range.InsertParagraphAfter();
            }
            //相手先氏名
            if (Sheets_AddressTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                paragraph4.Range.Text = Sheets_AddressTitleAndNameLabel.Text;
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph4.Range.InsertParagraphAfter();
            }
            //発信者会社名
            if (Sheets_CallerCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph5 = doc.Paragraphs.Add();
                paragraph5.Range.Text = Sheets_CallerCompanyLabel.Text;
                paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph5.Range.InsertParagraphAfter();
            }
            //発信者所在地
            if (Sheets_CallerLocationLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph6 = doc.Paragraphs.Add();
                paragraph6.Range.Text = Sheets_CallerLocationLabel.Text;
                paragraph6.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph6.Range.InsertParagraphAfter();
            }
            //発信者建物名と階数
            if (Sheets_BuildingNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph7 = doc.Paragraphs.Add();
                paragraph7.Range.Text = Sheets_BuildingNameLabel.Text;
                paragraph7.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph7.Range.InsertParagraphAfter();
            }
            //発信者氏名
            if (Sheets_CallerTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                paragraph8.Range.Text = Sheets_CallerTitleAndNameLabel.Text;
                paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph8.Range.InsertParagraphAfter();
            }
            //メールアドレス
            if (Sheets_CallerMallAddressLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph9 = doc.Paragraphs.Add();
                paragraph9.Range.Text = "メールアドレス:" + Sheets_CallerMallAddressLabel.Text;
                paragraph9.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph9.Range.InsertParagraphAfter();
            }
            //電話番号
            if (Sheets_CallerTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph10 = doc.Paragraphs.Add();
                paragraph10.Range.Text = "電話番号:" + Sheets_CallerTelLabel.Text;
                paragraph10.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph10.Range.InsertParagraphAfter();
            }
            //Fax番号
            if (Sheets_CallerFaxTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph11 = doc.Paragraphs.Add();
                paragraph11.Range.Text = "Fax番号:" + Sheets_CallerFaxTelLabel.Text;
                paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph11.Range.InsertParagraphAfter();
            }
            //表題
            if (Sheets_TitleButton.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph12 = doc.Paragraphs.Add();
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    paragraph12.Range.Bold = 1;
                }
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    paragraph12.Range.Italic = 1;
                }
                if (kryptonContextMenuItem15.Checked == true)
                {
                    paragraph12.Range.Underline = WdUnderline.wdUnderlineSingle;
                }
                if (kryptonContextMenuItem16.Checked == true)
                {
                    paragraph12.Range.Font.StrikeThrough = 1;
                }
                paragraph12.Range.Font.Size = (int)kryptonTextBox10.StateCommon.Content.Font.Size;
                paragraph12.Range.Text = Sheets_TitleButton.Text;
                paragraph12.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                SetWordRangeColor(paragraph12.Range, Sheets_TitleButton.ForeColor);
                paragraph12.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph12.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph12.Range.InsertParagraphAfter();

            }
            //あいさつ文
            if (Sheets_ContentLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph13 = doc.Paragraphs.Add();
                paragraph13.Range.Bold = 0;
                paragraph13.Range.Italic = 0;
                paragraph13.Range.Underline = WdUnderline.wdUnderlineNone;
                paragraph13.Range.Font.StrikeThrough = 0;
                paragraph13.Range.Font.Size = 10;
                paragraph13.Range.Text = Sheets_ContentLabel.Text;
                paragraph13.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph13.Range.Font.Color = WdColor.wdColorBlack;
                paragraph13.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph13.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph13.Range.InsertParagraphAfter();
            }
            //内容
            try
            {
                int LinesCount = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox12.Lines[LinesCount];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount = LinesCount + 1;
                    if (LinesCount == kryptonTextBox12.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //結語
            //kryptonRibbonGroupCheckBoxにチェックがない場合に動作する
            if (kryptonRibbonGroupCheckBox1.Checked != true)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph14 = doc.Paragraphs.Add();
                paragraph14.Range.Text = Sheet_ConclusionLabel.Text;
                paragraph14.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph14.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph14.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph14.Range.InsertParagraphAfter();
            }
            kryptonLabel1.Text = "出力完了";
            stausUpdate();
            //記
            if (Sheet_NoteLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph15 = doc.Paragraphs.Add();
                paragraph15.Range.Text = Sheet_NoteLabel.Text;
                paragraph15.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph15.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph15.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph15.Range.InsertParagraphAfter();
            }
            //記し書き
            try
            {
                int LinesCount2 = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox13.Lines[LinesCount2];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount2 = LinesCount2 + 1;
                    if (LinesCount2 == kryptonTextBox13.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //以上
            if (Sheets_EndLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph16 = doc.Paragraphs.Add();
                paragraph16.Range.Text = Sheets_EndLabel.Text;
                paragraph16.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph16.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph16.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph16.Range.InsertParagraphAfter();
            }
            GC.Collect();

            PrintDialog pd = new PrintDialog();
            if(pd.ShowDialog() == DialogResult.OK)
            {
                word.ActivePrinter = pd.PrinterSettings.PrinterName;
                doc.PrintOut();
            }

            try
            {
                doc.Close(false);
                word.Quit();
            }
            catch { }
        }

        private void kryptonButton17_Click(object sender, EventArgs e)
        {
            if(kryptonComboBox18.Text == "発行元部署")
            {
                string str = kryptonTextBox11.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox11.Text = str;
            }
            else if(kryptonComboBox18.Text == "宛先の組織・会社名")
            {
                string str = kryptonTextBox1.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox1.Text = str;
            }
            else if(kryptonComboBox18.Text == "宛先の肩書き")
            {
                string str = kryptonComboBox10.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox10.Text = str;
            }
            else if (kryptonComboBox18.Text == "宛先の氏名")
            {
                string str = kryptonTextBox2.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox2.Text = str;
            }
            else if (kryptonComboBox18.Text == "発信者の会社名")
            {
                string str = kryptonTextBox3.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox3.Text = str;
            }
            else if (kryptonComboBox18.Text == "発信者の所在地")
            {
                string str = kryptonTextBox4.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox4.Text = str;
            }
            else if (kryptonComboBox18.Text == "発信者の建物名")
            {
                string str = kryptonTextBox5.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox5.Text = str;
            }
            else if (kryptonComboBox18.Text == "発信者の肩書き")
            {
                string str = kryptonComboBox9.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonComboBox9.Text = str;
            }
            else if (kryptonComboBox18.Text == "発信者の氏名")
            {
                string str = kryptonTextBox6.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox6.Text = str;
            }
            else if (kryptonComboBox18.Text == "発信者のメールアドレス(ユーザー)")
            {
                string str = kryptonTextBox7.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox7.Text = str;
            }
            else if (kryptonComboBox18.Text == "発信者のメールアドレス(ドメイン)")
            {
                string str = kryptonComboBox8.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonComboBox8.Text = str;
            }
            else if (kryptonComboBox18.Text == "発信者の電話番号(1)")
            {
                string str = kryptonComboBox6.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonComboBox6.Text = str;
            }
            else if (kryptonComboBox18.Text == "発信者の電話番号(2)")
            {
                string str = kryptonTextBox14.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox14.Text = str;
            }
            else if (kryptonComboBox18.Text == "発信者の電話番号(3)")
            {
                string str = kryptonTextBox8.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox8.Text = str;
            }
            else if (kryptonComboBox18.Text == "発信者のFax番号(1)")
            {
                string str = kryptonComboBox7.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonComboBox7.Text = str;
            }
            else if (kryptonComboBox18.Text == "発信者のFax番号(2)")
            {
                string str = kryptonTextBox9.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox9.Text = str;
            }
            else if (kryptonComboBox18.Text == "発信者のFax番号(3)")
            {
                string str = kryptonTextBox15.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox15.Text = str;
            }
            else if (kryptonComboBox18.Text == "表題")
            {
                string str = kryptonTextBox10.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonTextBox10.Text = str;
            }
            else if (kryptonComboBox18.Text == "頭語")
            {
                string str = kryptonComboBox2.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonComboBox2.Text = str;
            }
            else if (kryptonComboBox18.Text == "候文")
            {
                string str = kryptonComboBox11.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonComboBox11.Text = str;
            }
            else if (kryptonComboBox18.Text == "前文")
            {
                string str = kryptonComboBox3.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonComboBox3.Text = str;
            }
            else if (kryptonComboBox18.Text == "感謝のあいさつ")
            {
                string str = kryptonComboBox4.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonComboBox4.Text = str;
            }
            else if (kryptonComboBox18.Text == "結語")
            {
                string str = kryptonComboBox5.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonComboBox5.Text = str;
            }
            else if (kryptonComboBox18.Text == "内容")
            {
                string str = kryptonTextBox12.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonComboBox12.Text = str;
            }
            else if (kryptonComboBox18.Text == "記し書き")
            {
                string str = kryptonTextBox13.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                kryptonComboBox13.Text = str;
            }
        }

        private void kryptonRibbonGroupButton16_Click(object sender, EventArgs e)
        {
            if(kryptonRibbonGroupButton16.Checked == true)
            {
                Transition
                    .With(kryptonPanel21, nameof(Height), 36)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                kryptonRibbonGroupButton16.Checked = true;
            }
            else
            {
                Transition
                    .With(kryptonPanel21, nameof(Height), 0)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                kryptonRibbonGroupButton16.Checked = false;
            }

        }

        private void kryptonRibbonGroupButton22_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Directory.GetCurrentDirectory() + @"\DocuQuick.exe");
        }

        private void kryptonButton18_Click(object sender, EventArgs e)
        {
            kryptonTextBox16.Text = String.Empty;
            kryptonTextBox17.Text = String.Empty;
            kryptonComboBox12.Text = String.Empty;
            kryptonTextBox18.Text = String.Empty;
            kryptonTextBox19.Text = String.Empty;
            kryptonTextBox32.Text = String.Empty;
            kryptonTextBox20.Text = String.Empty;
            kryptonNumericUpDown3.Value = 1;
            kryptonComboBox13.Text =String.Empty;
            kryptonTextBox21.Text = String.Empty;
            kryptonTextBox22.Text = String.Empty;
            kryptonComboBox14.Text =String.Empty;
            kryptonComboBox15.Text =String.Empty;
            kryptonTextBox23.Text = String.Empty;
            kryptonTextBox24.Text = String.Empty;
            kryptonComboBox16.Text =String.Empty;
            kryptonTextBox26.Text = String.Empty;
            kryptonTextBox25.Text = String.Empty;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {

        }

        private void kryptonButton5_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown4.Value = Properties.Settings.Default.Space_Top;
            kryptonNumericUpDown7.Value = Properties.Settings.Default.Space_Buttom;
            kryptonNumericUpDown5.Value = Properties.Settings.Default.Space_Left;
            kryptonNumericUpDown6.Value = Properties.Settings.Default.Space_Right;
        }

        //標準
        private void kryptonContextMenuItem31_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown4.Value = 138;
            kryptonNumericUpDown7.Value = 118;
            kryptonNumericUpDown5.Value = 118;
            kryptonNumericUpDown6.Value = 118;
        }

        //狭い
        private void kryptonContextMenuItem32_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown4.Value = 50;
            kryptonNumericUpDown7.Value = 50;
            kryptonNumericUpDown5.Value = 50;
            kryptonNumericUpDown6.Value = 50;
        }

        //やや狭い
        private void kryptonContextMenuItem33_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown4.Value = 100;
            kryptonNumericUpDown7.Value = 100;
            kryptonNumericUpDown5.Value = 75;
            kryptonNumericUpDown6.Value = 75;
        }

        //広い
        private void kryptonContextMenuItem34_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown4.Value = 100;
            kryptonNumericUpDown7.Value = 100;
            kryptonNumericUpDown5.Value = 200;
            kryptonNumericUpDown6.Value = 200;
        }

        private void kryptonLinkLabel1_LinkClicked(object sender, EventArgs e)
        {
            Properties.Settings.Default.ShowNotepadWarningPanel = false;
            WarningPanel1.Visible = false;
        }

        private void kryptonRibbonGroupButton_TextReset_Click(object sender, EventArgs e)
        {
            if(Properties.Settings.Default.ShowResetDialog == true)
            {

                TextResetDialog textResetDialog = new TextResetDialog();
                //Office2007青色
                if (this.BackColor == Color.FromArgb(191, 219, 255))
                {
                    textResetDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
                }
                //Office2007銀色
                else if (this.BackColor == Color.FromArgb(208, 212, 221))
                {
                    textResetDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
                }
                //Office2007ブラック
                else if (this.BackColor == Color.FromArgb(83, 83, 83))
                {
                    textResetDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
                }
                //Office2010青色
                else if (this.BackColor == Color.FromArgb(187, 206, 230))
                {
                    textResetDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
                }
                //Office2010銀色
                else if (this.BackColor == Color.FromArgb(227, 230, 232))
                {
                    textResetDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
                }
                //Office2010黒色
                else if (this.BackColor == Color.FromArgb(113, 113, 113))
                {
                    textResetDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
                }
                textResetDialog.ShowDialog();
                if (textResetDialog.DialogResult == DialogResult.Yes)
                {
                    SetSheetSpace();
                    SetSheetText();
                    kryptonRibbonButton_Bold.Checked = false;
                    kryptonRibbonButton_Italic.Checked = false;
                    kryptonContextMenuItem15.Checked = false;
                    kryptonContextMenuItem16.Checked = false;

                    kryptonNumericUpDown1.Value = 0;
                    kryptonDateTimePicker1.Value = DateTime.Today;

                    kryptonCheckBox3.Checked = false;
                    kryptonCheckBox2.Checked = false;

                    if(Properties.Settings.Default.IsUseEraName == true)
                    {
                        kryptonCheckBox1.Checked = true;
                    }
                    else
                    {
                        kryptonCheckBox1.Checked = false;
                    }
                }
            }
            else
            {
                SetSheetSpace();
                SetSheetText();
                kryptonRibbonButton_Bold.Checked = false;
                kryptonRibbonButton_Italic.Checked = false;
                kryptonContextMenuItem15.Checked = false;
                kryptonContextMenuItem16.Checked = false;

                kryptonNumericUpDown1.Value = 0;
                kryptonDateTimePicker1.Value = DateTime.Today;

                kryptonCheckBox3.Checked = false;
                kryptonCheckBox2.Checked = false;

                if (Properties.Settings.Default.IsUseEraName == true)
                {
                    kryptonCheckBox1.Checked = true;
                }
                else
                {
                    kryptonCheckBox1.Checked = false;
                }
            }
        }

        //上
        private void kryptonContextMenuItem22_Click(object sender, EventArgs e)
        {
            Sheets_TopPanel.Height = Properties.Settings.Default.Space_Top;

            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = Properties.Settings.Default.Space_Top;
        }

        //下
        private void kryptonContextMenuItem23_Click(object sender, EventArgs e)
        {
            Sheets_ButtomPanel.Height = Properties.Settings.Default.Space_Buttom;

            kryptonRibbonGroupNumericUpDown1.Value = Properties.Settings.Default.Space_Buttom;
        }

        //右
        private void kryptonContextMenuItem24_Click(object sender, EventArgs e)
        {
            Sheets_RightPanel.Width = Properties.Settings.Default.Space_Right;

            kryptonRibbonGroupNumericUpDown2.Value = Properties.Settings.Default.Space_Right;
        }

        //左
        private void kryptonContextMenuItem25_Click(object sender, EventArgs e)
        {
            Sheets_LeftPanel.Width = Properties.Settings.Default.Space_Left;

            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = Properties.Settings.Default.Space_Left;
        }


        //すべて
        private void kryptonContextMenuItem26_Click(object sender, EventArgs e)
        {
            Sheets_TopPanel.Height = Properties.Settings.Default.Space_Top;
            Sheets_ButtomPanel.Height = Properties.Settings.Default.Space_Buttom;
            Sheets_LeftPanel.Width = Properties.Settings.Default.Space_Left;
            Sheets_RightPanel.Width = Properties.Settings.Default.Space_Right;

            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = Properties.Settings.Default.Space_Top;
            kryptonRibbonGroupNumericUpDown1.Value = Properties.Settings.Default.Space_Buttom;
            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = Properties.Settings.Default.Space_Left;
            kryptonRibbonGroupNumericUpDown2.Value = Properties.Settings.Default.Space_Right;
        }

        private void kryptonButton4_Click(object sender, EventArgs e)
        {
            //設定の復元
            SetSettings();
            if (kryptonRibbonGroupButton16.Checked == true)
            {
                Transition
                    .With(kryptonPanel21, nameof(Height), 36)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                kryptonRibbonGroupButton16.Checked = true;
            }
            else
            {
                Transition
                    .With(kryptonPanel21, nameof(Height), 0)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                kryptonRibbonGroupButton16.Checked = false;
            }

            if (kryptonTrackBar1.Value == 0)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }
            else if (kryptonTrackBar1.Value == 10)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = false;
                kryptonLabel42.Enabled = true;
            }
            else
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }


            kryptonPage2.Visible = false;
            kryptonNavigator_Workbench.NavigatorMode = NavigatorMode.BarTabGroup;
            kryptonNavigator_Workbench.SelectedPage = kryptonPage1;

            kryptonRibbon.MinimizedMode = false;
            kryptonRibbon.Enabled = true;

            kryptonLabel7.Enabled = true;
            kryptonCheckButton1.Enabled = true;
            kryptonCheckButton2.Enabled = true;
            kryptonLabel1.Enabled = true;

            if (kryptonTextBox10.Text != string.Empty)
            {
                Sheets_TitleButton.Text = kryptonTextBox10.Text;
                this.Text = kryptonTextBox10.Text + " - DocuQuick Designer";
            }
            else
            {
                this.Text = "無題 - DocuQuick Designer";
            }

            kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

            Sheets_Sheet.Anchor = AnchorStyles.Top;

            // 親コントロールのサイズを取得
            int parentWidth = this.ClientSize.Width;
            int parentHeight = this.ClientSize.Height;

            // パネルのサイズを取得
            int panelWidth = Sheets_Sheet.Width;
            int panelHeight = Sheets_Sheet.Height;

            // パネルの位置を中央に設定
            Sheets_Sheet.Location = new System.Drawing.Point(
                (parentWidth - panelWidth) / 2 - 10,
                90
            );

            Sheets_Sheet.Top = 59;

        }
    }

}
