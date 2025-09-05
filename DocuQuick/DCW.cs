using AeroWizard;
using ComponentFactory.Krypton.Navigator;
using ComponentFactory.Krypton.Ribbon;
using ComponentFactory.Krypton.Toolkit;
using FluentTransitions;
using Krypton.Toolkit;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.UI.Xaml.Documents;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Document_Maker
{
    public partial class DCW : Form
    {

        public DCW()
        {
            InitializeComponent();
        }

        #region TreeNodeの追加
        TreeNode treeNode1 = new TreeNode();
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
        }
        #endregion

        private void DCW_Resize(object sender, EventArgs e)
        {
            if (this.Width >= 541)
            {
                panel1.Top = 143;
            }
            else if (this.Width <= 541)
            {
                panel1.Top = 136;
            }
        }


        private void wizardPage5_Commit(object sender, AeroWizard.WizardPageConfirmEventArgs e)
        {

        }

        private void wizardPage6_Commit(object sender, AeroWizard.WizardPageConfirmEventArgs e)
        {

        }

        private void kryptonLabel24_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonComboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonLabel26_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonLabel25_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonLabel27_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonComboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            button3.Text = kryptonComboBox4.Text + "　" + kryptonComboBox11.Text + kryptonListBox1.SelectedItem + kryptonListBox2.SelectedItem + "\r\n" + kryptonComboBox5.Text;

        }

        private void DCW_Load(object sender, EventArgs e)
        {

            //日付の設定
            if (kryptonCheckBox3.Checked == true)
            {
                DateTime date = kryptonMonthCalendar1.SelectionStart;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                //下記のように西暦ではなく和暦として表示するように設定する
                culturejp.DateTimeFormat.Calendar = new JapaneseCalendar();
                kryptonLabel6.Text = date.ToString("選択中の日付:" + "ggy年M月d日", culturejp);
            }
            else
            {
                DateTime date = kryptonMonthCalendar1.SelectionStart;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                kryptonLabel6.Text = date.ToString("選択中の日付:" + "yyyy年M月d日");
            }

            //kryptonComboBox10の月を教の月に変更する
            DateTime dt = DateTime.Today;
            kryptonComboBox10.Text = dt.Month.ToString();

            kryptonListBox1.SelectedItem = (string)"貴社ますますご盛栄のこととお慶び申し上げます。";
            kryptonListBox2.SelectedItem = (string)"平素は格別のご高配を賜り、厚く御礼申し上げます。";

            //あいさつ文プレビューの適用
            button3.Text = kryptonComboBox4.Text + "　" + kryptonComboBox11.Text + kryptonListBox1.SelectedItem + kryptonListBox2.SelectedItem + "\r\n" + kryptonComboBox5.Text;

            //フォントの設定
            InstalledFontCollection fonts = new InstalledFontCollection();
            FontFamily[] fontFamilies = fonts.Families;

            //TreeNodeの追加
            AddTreeNodes();


            foreach (FontFamily font in fontFamilies)
            {
                kryptonComboBox6.Items.Add(font.Name);
                kryptonComboBox6.AutoCompleteCustomSource.Add(font.Name);


                kryptonComboBox6.Text = button1.Font.Name;
                kryptonComboBox7.Text = button1.Font.Size.ToString();
            }

            wizardControl1.NextButtonText = "次へ";
            wizardControl1.CancelButtonText = "キャンセル";
            wizardControl1.FinishButtonText = "完了";


            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonMonthCalendar1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonMonthCalendar1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonMonthCalendar1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonColorButton1.StateCommon.Content.ShortText.Color1 = Color.Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                this.BackColor = Color.FromArgb(187, 206, 230);
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonMonthCalendar1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue ;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonMonthCalendar1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonMonthCalendar1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;

                kryptonLabel9.StateCommon.ShortText.Color1 = Color.White;
                kryptonLabel29.StateCommon.ShortText.Color1 = Color.White;
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            

        }

        private void wizardControl1_SelectedPageChanged(object sender, EventArgs e)
        {
            if (wizardPage6.Visible == true)
            {
                this.ControlBox = true;
                Transition
                    .With(this, nameof(Width), 819)
                    .CriticalDamp(TimeSpan.FromSeconds(0.6));
            }
            else if(wizardPage8.Visible == true)
            {
                this.ControlBox = true;
                Transition
                    .With(this, nameof(Width), 819)
                    .CriticalDamp(TimeSpan.FromSeconds(0.6));
            }
            else if(wizardPage7.Visible == true)
            {
                this.ControlBox = false;
                wizardPage7.ShowCancel = false;
                Transition
                    .With(this, nameof(Width), 540)
                    .CriticalDamp(TimeSpan.FromSeconds(0.6));
            }
            else
            {
                this.ControlBox = true;
                Transition
                    .With(this, nameof(Width), 540)
                    .CriticalDamp(TimeSpan.FromSeconds(0.6));
            }
        }

        private string PageConfirmationResult;



        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.ToOrCaller = 0;
            if (wizardPage3.Visible == true)
            {
                PageConfirmationResult = wizardPage3.Name.ToString();

            }
            else if (wizardPage4.Visible == true)
            {
                PageConfirmationResult = wizardPage4.Name.ToString();
            }

            AddressWindow addressWindow = new AddressWindow(PageConfirmationResult);
            if (kryptonPalette1.BasePaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue)
            {
                addressWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            else if (kryptonPalette1.BasePaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver)
            {
                addressWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            else if (kryptonPalette1.BasePaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black)
            {
                addressWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            else if (kryptonPalette1.BasePaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue)
            {
                addressWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            else if (kryptonPalette1.BasePaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver)
            {
                addressWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            else if (kryptonPalette1.BasePaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black)
            {
                addressWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
            addressWindow.ShowDialog();
            //情報を受け渡す。
            if (addressWindow.DialogResult == DialogResult.OK)
            {
                if(addressWindow.HumanNameOrCompanyName == 0)
                {
                    kryptonTextBox2.Text = addressWindow.flName;
                }
                else if(addressWindow.HumanNameOrCompanyName == 1)
                {
                    kryptonTextBox3.Text = addressWindow.flName;
                }
            }



        }


        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.ToOrCaller = 1;
            if (wizardPage3.Visible == true)
            {
                PageConfirmationResult = wizardPage3.Name.ToString();
            }
            else if (wizardPage4.Visible == true)
            {
                PageConfirmationResult = wizardPage4.Name.ToString();
            }

            AddressWindow addressWindow = new AddressWindow(PageConfirmationResult);
            if (kryptonPalette1.BasePaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue)
            {
                addressWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            else if (kryptonPalette1.BasePaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver)
            {
                addressWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            else if (kryptonPalette1.BasePaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black)
            {
                addressWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            else if (kryptonPalette1.BasePaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue)
            {
                addressWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            else if (kryptonPalette1.BasePaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver)
            {
                addressWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            else if (kryptonPalette1.BasePaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black)
            {
                addressWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
            addressWindow.ShowDialog();

            //情報を受け渡す。
            if (addressWindow.DialogResult == DialogResult.OK)
            {
                if (addressWindow.HumanNameOrCompanyName == 0)
                {
                    kryptonTextBox4.Text = addressWindow.flName;
                }
                else if (addressWindow.HumanNameOrCompanyName == 1)
                {
                    kryptonTextBox7.Text = addressWindow.flName;
                }
                kryptonTextBox5.Text = addressWindow.loaction;
                kryptonTextBox8.Text = addressWindow.MailAddress_User;
                kryptonComboBox3.Text = addressWindow.MailAddress_Domain;
                kryptonComboBox8.Text = addressWindow.PhoneNumber1;
                kryptonTextBox10.Text = addressWindow.PhoneNumber2;
                kryptonTextBox11.Text = addressWindow.PhoneNumber3;
                kryptonComboBox9.Text = addressWindow.FaxNumber1;
                kryptonTextBox13.Text = addressWindow.FaxNumber2;
                kryptonTextBox12.Text = addressWindow.FacNumber3;
            }

        }


        private void wizardControl1_Cancelling(object sender, CancelEventArgs e)
        {
            IsWizardFinished = false;
        }

        private void kryptonCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (kryptonCheckBox1.Checked == true)
            {
                NoIssueNumber = true;
                kryptonTextBox1.Enabled = false;
                kryptonLabel3.Enabled = false;
                kryptonNumericUpDown1.Enabled = false;
                kryptonLabel4.Enabled = false;
            }
            else
            {
                NoIssueNumber = false;
                kryptonTextBox1.Enabled = true;
                kryptonLabel3.Enabled = true;
                kryptonNumericUpDown1.Enabled = true;
                kryptonLabel4.Enabled = true;
            }
        }

        private void kryptonCheckBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (kryptonCheckBox3.Checked == true)
            {
                DateTime date = kryptonMonthCalendar1.SelectionStart;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                //下記のように西暦ではなく和暦として表示するように設定する
                culturejp.DateTimeFormat.Calendar = new JapaneseCalendar();
                kryptonLabel6.Text = date.ToString("選択中の日付:" + "ggy年M月d日", culturejp);
            }
            else
            {
                DateTime date = kryptonMonthCalendar1.SelectionStart;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                kryptonLabel6.Text = date.ToString("選択中の日付:" + "yyyy年M月d日");
            }
        }

        private void kryptonCheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (kryptonCheckBox2.Checked == true)
            {
                NoDate = true;
                kryptonMonthCalendar1.Enabled = false;
                kryptonLabel6.Enabled = false;
                kryptonCheckBox3.Enabled = false;
            }
            else
            {
                NoDate = false;
                kryptonMonthCalendar1.Enabled = true;
                kryptonLabel6.Enabled = true;
                kryptonCheckBox3.Enabled = true;

                if (kryptonCheckBox3.Checked == true)
                {
                    DateTime date = kryptonMonthCalendar1.SelectionStart;
                    CultureInfo culturejp = new CultureInfo("ja-Jp");
                    //下記のように西暦ではなく和暦として表示するように設定する
                    culturejp.DateTimeFormat.Calendar = new JapaneseCalendar();
                    kryptonLabel6.Text = date.ToString("選択中の日付:" + "ggy年M月d日", culturejp);
                }
                else
                {
                    DateTime date = kryptonMonthCalendar1.SelectionStart;
                    CultureInfo culturejp = new CultureInfo("ja-Jp");
                    kryptonLabel6.Text = date.ToString("選択中の日付:" + "yyyy年M月d日");
                }
            }
        }

        private void kryptonComboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            if (kryptonComboBox1.Text == "お客様各位")
            {
                kryptonTextBox2.Enabled = false;
                kryptonTextBox3.Enabled = false;
            }
            else if (kryptonComboBox1.Text == "従業員各位")
            {
                kryptonTextBox2.Enabled = false;
                kryptonTextBox3.Enabled = false;
            }
            else
            {
                kryptonTextBox2.Enabled = true;
                kryptonTextBox3.Enabled = true;
            }
        }

        private void kryptonComboBox2_SelectedValueChanged(object sender, EventArgs e)
        {

        }

        private void kryptonTextBox15_TextChanged(object sender, EventArgs e)
        {
            button1.Text = kryptonTextBox15.Text;
        }

        private void kryptonComboBox4_SelectedValueChanged(object sender, EventArgs e)
        {
            //一般的
            if (kryptonComboBox4.Text == "拝啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox4.Text == "拝呈")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox4.Text == "啓上")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox4.Text == "敬白")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox4.Text == "拝進")
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
            else if (kryptonComboBox4.Text == "謹啓")
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
            else if (kryptonComboBox4.Text == "謹呈")
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
            else if (kryptonComboBox4.Text == "粛啓")
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
            else if (kryptonComboBox4.Text == "慕啓")
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
            else if (kryptonComboBox4.Text == "謹白")
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
            else if (kryptonComboBox4.Text == "急啓")
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
            else if (kryptonComboBox4.Text == "急呈")
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
            else if (kryptonComboBox4.Text == "急白")
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
            else if (kryptonComboBox4.Text == "前略")
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
            else if (kryptonComboBox4.Text == "冠省")
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
            else if (kryptonComboBox4.Text == "略啓")
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
            else if (kryptonComboBox4.Text == "寸啓")
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
            else if (kryptonComboBox4.Text == "草啓")
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
            else if (kryptonComboBox4.Text == "初めてお手紙を差し上げます")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox4.Text == "突然お手紙を差し上げますご無礼お許しください")
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
            else if (kryptonComboBox4.Text == "拝復")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝答",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox4.Text == "複啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝答",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox4.Text == "謹復")
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
            else if (kryptonComboBox4.Text == "合掌")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "合掌",
                });
                kryptonComboBox5.Text = "合掌";
            }
            else if (kryptonComboBox4.Text == "敬具")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                });
                kryptonComboBox5.Text = "敬具";
            }

            button3.Text = kryptonComboBox4.Text + "　" + kryptonComboBox11.Text + kryptonListBox1.SelectedItem + kryptonListBox2.SelectedItem + "\r\n" + kryptonComboBox5.Text;
        }

        private void kryptonComboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonComboBox10_SelectedValueChanged(object sender, EventArgs e)
        {
            if (kryptonComboBox10.Text == "1")
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
            else if (kryptonComboBox10.Text == "2")
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
            else if (kryptonComboBox10.Text == "3")
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
            else if (kryptonComboBox10.Text == "4")
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
            else if (kryptonComboBox10.Text == "5")
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
            else if (kryptonComboBox10.Text == "6")
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
            else if (kryptonComboBox10.Text == "7")
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
            else if (kryptonComboBox10.Text == "8")
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
            else if (kryptonComboBox10.Text == "9")
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
            else if (kryptonComboBox10.Text == "10")
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
            else if (kryptonComboBox10.Text == "11")
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
            else if (kryptonComboBox10.Text == "12")
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

            button3.Text = kryptonComboBox4.Text + "　" + kryptonComboBox11.Text + kryptonListBox1.SelectedItem + kryptonListBox2.SelectedItem + "\r\n" + kryptonComboBox5.Text;
        }

        private void kryptonListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button3.Text = kryptonComboBox4.Text + "　" + kryptonComboBox11.Text + kryptonListBox1.SelectedItem + kryptonListBox2.SelectedItem + "\r\n" + kryptonComboBox5.Text;
        }

        private void kryptonListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            button3.Text = kryptonComboBox4.Text + "　" + kryptonComboBox11.Text + kryptonListBox1.SelectedItem + kryptonListBox2.SelectedItem + "\r\n" + kryptonComboBox5.Text;
        }

        //フォント
        private void kryptonComboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonComboBox7.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                button1.Font = new System.Drawing.Font(
                    kryptonComboBox6.Text,
                    fontSize,
                    button1.Font.Style
                );
                button1.Font = button1.Font;
            }
        }

        //フォントサイズ
        private void kryptonComboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonComboBox7.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                button1.Font = new System.Drawing.Font(
                    kryptonComboBox6.Text,
                    fontSize,
                    button1.Font.Style
                );
                button1.Font = button1.Font;
            }
        }

        public void FontReset()
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonComboBox7.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                button1.Font = new System.Drawing.Font(
                    kryptonComboBox6.Text,
                    fontSize,
                    FontStyle.Regular
                );
                button1.Font = button1.Font;
            }
        }

        //太字
        private void kryptonCheckButton1_Click(object sender, EventArgs e)
        {

            if (kryptonCheckButton1.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    button1.Font = new System.Drawing.Font(
                        kryptonComboBox6.Text,
                        fontSize,
                        button1.Font.Style | FontStyle.Bold
                    );
                    button1.Font = button1.Font;
                }
            }
            else if (kryptonCheckButton1.Checked == false)
            {
                //太字ボタンをチェックをオフにする
                kryptonCheckButton1.Checked = false;
                //フォントスタイルのみ初期化
                FontReset();
                float fontSize;

                //斜体が有効な場合
                if (kryptonCheckButton2.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        button1.Font = new System.Drawing.Font(
                            button1.Font.Name,
                            fontSize,
                            button1.Font.Style | FontStyle.Italic
                        );
                        button1.Font = button1.Font;
                        kryptonCheckButton2.Checked = true;
                    }
                }

                //下線が有効な場合
                if (kryptonContextMenuItem1.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        button1.Font = new System.Drawing.Font(
                            button1.Font.Name,
                            fontSize,
                            button1.Font.Style | FontStyle.Underline
                        );
                        button1.Font = button1.Font;
                        kryptonContextMenuItem1.Checked = true;
                    }
                }

                //打ち消し線が有効な場合
                if (kryptonContextMenuItem2.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        button1.Font = new System.Drawing.Font(
                            button1.Font.Name,
                            fontSize,
                            button1.Font.Style | FontStyle.Strikeout
                        );
                        button1.Font = button1.Font;
                        kryptonContextMenuItem2.Checked = true;
                    }
                }
            }
        }

        //斜体
        private void kryptonCheckButton2_Click(object sender, EventArgs e)
        {

            if (kryptonCheckButton2.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    button1.Font = new System.Drawing.Font(
                        button1.Font.Name,
                        fontSize,
                        button1.Font.Style | FontStyle.Italic
                    );
                    button1.Font = button1.Font;
                    kryptonCheckButton2.Checked = true;
                }
            }
            else if (kryptonCheckButton2.Checked == false)
            {
                //斜体ボタンをチェックをオフにする
                kryptonCheckButton2.Checked = false;
                //フォントスタイルのみ初期化
                FontReset();
                float fontSize;

                //太字が有効な場合
                if (kryptonCheckButton1.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        button1.Font = new System.Drawing.Font(
                            kryptonComboBox6.Text,
                            fontSize,
                            button1.Font.Style | FontStyle.Bold
                        );
                        button1.Font = button1.Font;
                    }
                }

                //下線が有効な場合
                if (kryptonContextMenuItem1.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        button1.Font = new System.Drawing.Font(
                            button1.Font.Name,
                            fontSize,
                            button1.Font.Style | FontStyle.Underline
                        );
                        button1.Font = button1.Font;
                        kryptonContextMenuItem1.Checked = true;
                    }
                }

                //打ち消し線が有効な場合
                if (kryptonContextMenuItem2.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        button1.Font = new System.Drawing.Font(
                            button1.Font.Name,
                            fontSize,
                            button1.Font.Style | FontStyle.Strikeout
                        );
                        button1.Font = button1.Font;
                        kryptonContextMenuItem2.Checked = true;
                    }
                }
            }
        }

        //下線
        private void kryptonContextMenuItem1_Click(object sender, EventArgs e)
        {
            //下線が有効な場合
            if (kryptonContextMenuItem1.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    button1.Font = new System.Drawing.Font(
                        button1.Font.Name,
                        fontSize,
                        button1.Font.Style | FontStyle.Underline
                    );
                    button1.Font = button1.Font;
                    kryptonContextMenuItem1.Checked = true;
                }
            }
            else if (kryptonContextMenuItem1.Checked == false)
            {
                //斜体ボタンをチェックをオフにする
                kryptonContextMenuItem1.Checked = false;
                //フォントスタイルのみ初期化
                FontReset();
                float fontSize;

                //太字が有効な場合
                if (kryptonCheckButton1.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        button1.Font = new System.Drawing.Font(
                            kryptonComboBox6.Text,
                            fontSize,
                            button1.Font.Style | FontStyle.Bold
                        );
                        button1.Font = button1.Font;
                    }
                }


                //斜体が有効な場合
                if (kryptonCheckButton2.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        button1.Font = new System.Drawing.Font(
                            button1.Font.Name,
                            fontSize,
                            button1.Font.Style | FontStyle.Italic
                        );
                        button1.Font = button1.Font;
                        kryptonCheckButton2.Checked = true;
                    }
                }

                //打ち消し線が有効な場合
                if (kryptonContextMenuItem2.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        button1.Font = new System.Drawing.Font(
                            button1.Font.Name,
                            fontSize,
                            button1.Font.Style | FontStyle.Strikeout
                        );
                        button1.Font = button1.Font;
                        kryptonContextMenuItem2.Checked = true;
                    }
                }
            }
        }

        //打ち消し線
        private void kryptonContextMenuItem2_Click(object sender, EventArgs e)
        {
            //打ち消し線が有効な場合
            if (kryptonContextMenuItem2.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    button1.Font = new System.Drawing.Font(
                        button1.Font.Name,
                        fontSize,
                        button1.Font.Style | FontStyle.Strikeout
                    );
                    button1.Font = button1.Font;
                    kryptonContextMenuItem2.Checked = true;
                }
            }
            else if (kryptonContextMenuItem2.Checked == false)
            {
                //斜体ボタンをチェックをオフにする
                kryptonContextMenuItem1.Checked = false;
                //フォントスタイルのみ初期化
                FontReset();
                float fontSize;

                //太字が有効な場合
                if (kryptonCheckButton1.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        button1.Font = new System.Drawing.Font(
                            kryptonComboBox6.Text,
                            fontSize,
                            button1.Font.Style | FontStyle.Bold
                        );
                        button1.Font = button1.Font;
                    }
                }


                //斜体が有効な場合
                if (kryptonCheckButton2.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        button1.Font = new System.Drawing.Font(
                            button1.Font.Name,
                            fontSize,
                            button1.Font.Style | FontStyle.Italic
                        );
                        button1.Font = button1.Font;
                        kryptonCheckButton2.Checked = true;
                    }
                }

                //下線が有効な場合
                if (kryptonContextMenuItem1.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonComboBox7.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        button1.Font = new System.Drawing.Font(
                            button1.Font.Name,
                            fontSize,
                            button1.Font.Style | FontStyle.Underline
                        );
                        button1.Font = button1.Font;
                        kryptonContextMenuItem1.Checked = true;
                    }
                }

            }
        }

        private void kryptonColorButton1_SelectedColorChanged(object sender, ComponentFactory.Krypton.Toolkit.ColorEventArgs e)
        {
            button1.ForeColor = e.Color;
        }

        private void kryptonColorButton1_Click(object sender, EventArgs e)
        {
            button1.ForeColor = kryptonColorButton1.SelectedColor;
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            FontDialog fd = new FontDialog();
            fd.Font = button1.Font;
            if(fd.ShowDialog() == DialogResult.OK)
            {
                button1.Font = fd.Font;
                kryptonComboBox6.Text = fd.Font.Name;
                kryptonComboBox7.Text = fd.Font.Size.ToString();

                //他のフォントスタイルを確認する
                if (button1.Font.Style == FontStyle.Bold)
                {
                    kryptonCheckButton1.Checked = true;
                }
                else
                {
                    kryptonCheckButton1.Checked = false;
                }

                if (button1.Font.Style == FontStyle.Italic)
                {
                    kryptonCheckButton2.Checked = true;
                }
                else
                {
                    kryptonCheckButton2.Checked = false;
                }

                if (button1.Font.Style == FontStyle.Underline)
                {
                    kryptonContextMenuItem1.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem1.Checked = false;
                }

                if (button1.Font.Style == FontStyle.Strikeout)
                {
                    kryptonContextMenuItem2.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem2.Checked = false;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(kryptonLabel31.Top == 395)
            {
                Transition.
                    With(kryptonLabel31, nameof(Top), 434)
                    .With(kryptonLabel32, nameof(Top), 29)
                    .With(kryptonTextBox9, nameof(Top), 55)
                    .With(kryptonLabel33, nameof(Top), 219)
                    .With(kryptonTextBox14, nameof(Top), 245)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
            }
            else if(kryptonLabel31.Top == 434)
            {

                Transition.
                    With(kryptonLabel31, nameof(Top), 395)
                    .With(kryptonLabel32, nameof(Top), 18)
                    .With(kryptonTextBox9, nameof(Top), 44)
                    .With(kryptonLabel33, nameof(Top), 208)
                    .With(kryptonTextBox14, nameof(Top), 234)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
            }


        }

        private void kryptonTextBox1_TextChanged(object sender, EventArgs e)
        {
            if(kryptonTextBox1.Text == string.Empty)
            {
                kryptonLabel3.Text = "　第";
            }
            else
            {
                kryptonLabel3.Text = "発第";
            }
        }


        //ウィザード終了確認用
        public bool IsWizardFinished { get; set; }
        //文書入力用
        //発行番号
        public string IssueNumber_Publisher { get; set; }
        public int IssueNumber { get; set; }
        public string IssueNumberAll { get; set; }
        public bool NoIssueNumber { get; set; }
        //日付
        public DateTime Date { get; set; }
        public bool UseEraName {  get; set; }
        public bool NoDate { get; set; }
        //発行番号
        //宛先用
        public string AdCompany { get; set; }
        public string AdTitle { get; set; }
        public string AdName { get; set; }
        //発信者用
        public string CaCampany { get; set; }
        public string CaLocation { get; set; }
        public string CaBuildingName { get; set; }
        public int CaFloorNumber { get; set; }
        public string CaTitle { get; set; }
        public string CaName { get; set; }
        public string CaMailAddress { get; set; }
        public string CaMailAddress_Domain { get; set; }
        public string CaPhoneNumber1 { get; set; }
        public string CaPhoneNumber2 { get; set; }
        public string CaPhoneNumber3 { get; set; }
        public string CaFaxNumber1 { get; set; }
        public string CaFaxNumber2 { get; set; }
        public string CaFaxNumber3 { get; set; }
        //表題
        public string title { get; set; }
        public string ftName { get; set; }
        public float ftSize {  get; set; }
        public bool titleBold { get; set; }
        public bool titleItalic { get; set; }
        public bool titleUnderline { get; set; }
        public bool titleStrikeout { get; set; }
        public Color titleColor {  get; set; }
        //あいさつ文
        public string UseSourouBunDate {  get; set; }
        //頭語
        public string acronym { get; set; }
        //候文
        public string souroubun {  get; set; }
        //前文
        public string PreviousText { get; set; }
        //感謝のあいさつ
        public string ThankYouGreeting {  get; set; }
        //結語
        public string Conclusion { get; set; }
        //内容
        public string  Content { get; set; }
        public string Notetaking { get; set; }

        private void wizardControl1_Finished(object sender, EventArgs e)
        {
            IsWizardFinished = true;
            //完了ボタンをクリックしたのみ以下を実行する
            //発行番号
            if (kryptonCheckBox1.Checked == false)
            {

                if (kryptonTextBox1.Text != string.Empty)
                {
                    IssueNumber_Publisher = kryptonTextBox1.Text;
                    IssueNumber = ((int)kryptonNumericUpDown1.Value);
                    IssueNumberAll = kryptonTextBox1.Text + "発第" + kryptonNumericUpDown1.Value + "号";
                    NoIssueNumber = false;
                }
                else
                {
                    IssueNumber_Publisher = string.Empty;
                    IssueNumber = ((int)kryptonNumericUpDown1.Value);
                    IssueNumberAll = "第" + kryptonNumericUpDown1.Value + "号";
                    NoIssueNumber = false;
                }
            }
            else
            {
                NoIssueNumber = true;
            }

            //日付
            if (kryptonCheckBox2.Checked == false)
            {
                NoDate = false;
                // 修正: BoldedDatesはDateTime[]型なので、SelectionStartを使ってDateTime型を取得
                Date = kryptonMonthCalendar1.SelectionStart;
                if(kryptonCheckBox3.Checked == true)
                {
                    UseEraName = true;
                }
                else
                {
                    UseEraName= false;
                }
            }
            else
            {
                NoDate = true;
            }

            //宛先セクション
            //組織および会社名
            AdCompany = kryptonTextBox2.Text;
            //肩書きと氏名
            AdTitle = kryptonComboBox1.Text;
            AdName = kryptonTextBox3.Text;

            //発信者セクション
            //組織および会社名
            CaCampany = kryptonTextBox4.Text;
            //所在地
            CaLocation = kryptonTextBox5.Text;
            //建物名と階数
            CaBuildingName = kryptonTextBox6.Text;
            CaFloorNumber = (int)kryptonNumericUpDown2.Value;
            //肩書きと氏名
            CaTitle = kryptonComboBox2.Text;
            CaName = kryptonTextBox7.Text;
            //メールアドレス
            CaMailAddress = kryptonTextBox8.Text;
            CaMailAddress_Domain = kryptonComboBox3.Text;
            //電話番号
            CaPhoneNumber1 = kryptonComboBox8.Text;
            CaPhoneNumber2 = kryptonTextBox10.Text;
            CaPhoneNumber3 = kryptonTextBox11.Text;
            //Fax番号
            CaFaxNumber1 = kryptonComboBox9.Text;
            CaFaxNumber2 = kryptonTextBox13.Text;
            CaFaxNumber3 = kryptonTextBox12.Text;

            //表題
            title = kryptonTextBox15.Text;

            //フォント名
            ftName = button1.Font.Name;
            //フォントサイズ
            ftSize = button1.Font.Size;
            //フォントスタイルを確認
            if (kryptonCheckButton1.Checked == true)
            {
                titleBold = true;
            }
            else
            {
                titleBold = false;
            }

            if (kryptonCheckButton2.Checked == true)
            {
                titleItalic = true;
            }
            else
            {
                titleItalic = false;
            }

            if (kryptonContextMenuItem1.Checked == true)
            {
                titleUnderline = true;
            }
            else
            {
                titleUnderline = false;
            }

            if (kryptonContextMenuItem2.Checked == true)
            {
                titleStrikeout = true;
            }
            else
            {
                titleStrikeout = false;
            }
            titleColor = button1.ForeColor;

            //あいさつ文
            UseSourouBunDate = kryptonComboBox10.Text;
            //頭語
            acronym = kryptonComboBox4.Text;
            //候文
            souroubun = kryptonComboBox11.Text;
            //前文
            PreviousText = kryptonListBox1.SelectedItem.ToString();
            //感謝のあいさつ
            ThankYouGreeting = kryptonListBox2.SelectedItem.ToString();
            //結語
            Conclusion = kryptonComboBox5.Text;

            //内容
            Content = kryptonTextBox9.Text;
            //記し書き
            Notetaking = kryptonTextBox14.Text;
        }

        private void kryptonCheckBox3_CheckedChanged(object sender, DateRangeEventArgs e)
        {
            if (kryptonCheckBox3.Checked == true)
            {
                DateTime date = kryptonMonthCalendar1.SelectionStart;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                //下記のように西暦ではなく和暦として表示するように設定する
                culturejp.DateTimeFormat.Calendar = new JapaneseCalendar();
                kryptonLabel6.Text = date.ToString("選択中の日付:" + "ggy年M月d日", culturejp);
            }
            else
            {
                DateTime date = kryptonMonthCalendar1.SelectionStart;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                kryptonLabel6.Text = date.ToString("選択中の日付:" + "yyyy年M月d日");
            }
        }


        private void SetWordRangeColor(Range range, Color color)
        {
            // Word の RGB 値は Red + (Green << 8) + (Blue << 16)
            int rgb = color.R | (color.G << 8) | (color.B << 16);
            range.Font.Color = (WdColor)rgb;
        }


        //Word出力処理
        private void kryptonCommandLinkButton1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = true;
            Document doc = word.Documents.Add();


            //外枠の余白を設定
            doc.PageSetup.TopMargin = Properties.Settings.Default.dCW_TopSpace;
            doc.PageSetup.BottomMargin = Properties.Settings.Default.dCW_ButtomSpace;
            doc.PageSetup.LeftMargin = Properties.Settings.Default.dCW_LeftSpace;
            doc.PageSetup.RightMargin = Properties.Settings.Default.dCW_RightSpace;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //発行元部署＋発行番号
            if (kryptonCheckBox1.Checked == false)
            {
                if (kryptonTextBox1.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                    paragraph1.Range.Text = kryptonTextBox1.Text + "発第" + kryptonNumericUpDown1.Value + "号";
                    paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph1.Range.InsertParagraphAfter();
                }
                else
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                    paragraph1.Range.Text = kryptonNumericUpDown1.Value + "号";
                    paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph1.Range.InsertParagraphAfter();
                }
            }

            //日付
            if (kryptonCheckBox2.Checked == false)
            {
                if (kryptonCheckBox3.Checked == true)
                {
                    DateTime date = kryptonMonthCalendar1.SelectionStart;
                    CultureInfo culturejp = new CultureInfo("ja-Jp");
                    //下記のように西暦ではなく和暦として表示するように設定する
                    culturejp.DateTimeFormat.Calendar = new JapaneseCalendar();

                    Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                    paragraph2.Range.Text = date.ToString("ggy年M月d日", culturejp);
                    paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph2.Range.InsertParagraphAfter();
                }
                else
                {
                    DateTime date = kryptonMonthCalendar1.SelectionStart;
                    CultureInfo culturejp = new CultureInfo("ja-Jp");

                    Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                    paragraph2.Range.Text = date.ToString("yyyy年M月d日", culturejp);
                    paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph2.Range.InsertParagraphAfter();
                }
            }

            //宛先
            //組織および会社名
            if (kryptonTextBox2.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph3 = doc.Paragraphs.Add();
                paragraph3.Range.Text = kryptonTextBox2.Text;
                paragraph3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph3.Range.InsertParagraphAfter();
            }

            //肩書と氏名
            if (kryptonComboBox1.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                paragraph4.Range.Text = kryptonComboBox1.Text + "　" + kryptonTextBox3.Text;
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Range.InsertParagraphAfter();
            }
            else
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                paragraph4.Range.Text = kryptonTextBox3.Text;
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Range.InsertParagraphAfter();
            }

            //発信者
            //組織および会社名
            if (kryptonTextBox4.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph5 = doc.Paragraphs.Add();
                paragraph5.Range.Text = kryptonTextBox4.Text;
                paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph5.Range.InsertParagraphAfter();
            }
            //所在地
            if (kryptonTextBox5.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph6 = doc.Paragraphs.Add();
                paragraph6.Range.Text = kryptonTextBox5.Text;
                paragraph6.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph6.Range.InsertParagraphAfter();
            }
            //建物名+階数
            if (kryptonTextBox6.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph7 = doc.Paragraphs.Add();
                if (kryptonNumericUpDown2.Value <= 0)
                {
                    int negativeNumber = (int)kryptonNumericUpDown2.Value;
                    int positiveNumber = Math.Abs(negativeNumber);

                    paragraph7.Range.Text = kryptonTextBox6.Text + "　" + "地下" + positiveNumber + "階";
                }
                else
                {
                    paragraph7.Range.Text = kryptonTextBox6.Text + "　" + kryptonNumericUpDown2.Value + "階";
                }
                paragraph7.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph7.Range.InsertParagraphAfter();
            }
            //肩書きと氏名
            if (kryptonComboBox2.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                paragraph8.Range.Text = kryptonComboBox2.Text + "　" + kryptonTextBox7.Text;
                paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph8.Range.InsertParagraphAfter();
            }
            else
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                paragraph8.Range.Text = kryptonTextBox7.Text;
                paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph8.Range.InsertParagraphAfter();
            }
            //メールアドレス
            if (kryptonTextBox8.Text != string.Empty)
            {
                if (kryptonComboBox3.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph9 = doc.Paragraphs.Add();
                    paragraph9.Range.Text = "メールアドレス:" + kryptonTextBox8.Text + "@" + kryptonComboBox3.Text;
                    paragraph9.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph9.Range.InsertParagraphAfter();
                }
            }
            //電話番号
            if (kryptonComboBox8.Text != string.Empty)
            {
                if (kryptonTextBox10.Text != string.Empty)
                {
                    if (kryptonTextBox11.Text != string.Empty)
                    {
                        Microsoft.Office.Interop.Word.Paragraph paragraph10 = doc.Paragraphs.Add();
                        paragraph10.Range.Text = "電話番号:" + kryptonComboBox8.Text + "-" + kryptonTextBox10.Text + "-" + kryptonTextBox11.Text;
                        paragraph10.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                        paragraph10.Range.InsertParagraphAfter();
                    }
                }
            }
            //Fax番号
            if (kryptonComboBox9.Text != string.Empty)
            {
                if (kryptonTextBox13.Text != string.Empty)
                {
                    if (kryptonTextBox12.Text != string.Empty)
                    {
                        Microsoft.Office.Interop.Word.Paragraph paragraph11 = doc.Paragraphs.Add();
                        paragraph11.Range.Text = "Fax番号:" + kryptonComboBox9.Text + "-" + kryptonTextBox13.Text + "-" + kryptonTextBox12.Text;
                        paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                        paragraph11.Range.InsertParagraphAfter();
                    }
                }
            }

            //表題
            if (kryptonTextBox15.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph12 = doc.Paragraphs.Add();
                if (kryptonCheckButton1.Checked == true)
                {
                    paragraph12.Range.Bold = 1;
                }
                if (kryptonCheckButton2.Checked == true)
                {
                    paragraph12.Range.Italic = 1;
                }
                if (kryptonContextMenuItem1.Checked == true)
                {
                    paragraph12.Range.Underline = WdUnderline.wdUnderlineSingle;
                }
                if (kryptonContextMenuItem2.Checked == true)
                {
                    paragraph12.Range.Font.StrikeThrough = 1;
                }
                paragraph12.Range.Font.Name = button1.Font.Name;
                paragraph12.Range.Font.Size = button1.Font.Size;
                SetWordRangeColor(paragraph12.Range, button1.ForeColor);
                paragraph12.Range.Text = kryptonTextBox15.Text;
                paragraph12.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph12.Range.InsertParagraphAfter();
            }


            //あいさつ文
            Microsoft.Office.Interop.Word.Paragraph paragraph13 = doc.Paragraphs.Add();
            paragraph13.Range.Text = kryptonComboBox4.Text + "　" + kryptonComboBox11.Text + kryptonListBox1.SelectedItem.ToString() + kryptonListBox2.SelectedItem.ToString();
            paragraph13.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            paragraph13.Range.Font.Name = "游明朝";
            paragraph13.Range.Font.Color = WdColor.wdColorBlack;
            paragraph13.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph13.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            paragraph13.Range.InsertParagraphAfter();

            //内容
            try
            {
                int LinesCount = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox9.Lines[LinesCount];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount = LinesCount + 1;
                    if (LinesCount == kryptonTextBox9.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }

            //結語
            Microsoft.Office.Interop.Word.Paragraph paragraph14 = doc.Paragraphs.Add();
            paragraph14.Range.Text = kryptonComboBox5.Text;
            paragraph14.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            paragraph14.Range.InsertParagraphAfter();

            //記
            Microsoft.Office.Interop.Word.Paragraph paragraph15 = doc.Paragraphs.Add();
            paragraph15.Range.Text = "記";
            paragraph15.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph15.Range.InsertParagraphAfter();

            //記し書き
            try
            {
                int LinesCount2 = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox14.Lines[LinesCount2];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount2 = LinesCount2 + 1;
                    if (LinesCount2 == kryptonTextBox14.Lines.Length)
                    {
                        break;
                    }
                }


            }
            catch (Exception ex) { }

            //以上
            Microsoft.Office.Interop.Word.Paragraph paragraph16 = doc.Paragraphs.Add();
            paragraph16.Range.Text = "記";
            paragraph16.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            paragraph16.Range.InsertParagraphAfter();

            System.Windows.Forms.Application.Exit();
            GC.Collect();

        }
    }
}
