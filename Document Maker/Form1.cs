using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Xml.Linq;
using Vanara.Interop.DesktopWindowManager;
using static System.Collections.Specialized.BitVector32;
using Word = Microsoft.Office.Interop.Word;
using Krypton.Toolkit;

namespace Document_Maker
{

    public partial class Form1 : RibbonForm
    {


        public Form1()
        {
            InitializeComponent();
        }
        
        private void ribbonButton26_Click(object sender, EventArgs e)
        {
        }

        private void ribbonButton27_Click(object sender, EventArgs e)
        {
        }

        private void ribbonButton28_Click(object sender, EventArgs e)
        {
            ribbon1.OrbStyle = RibbonOrbStyle.Office_2007;
            ribbonButton28.Checked = true;
            ribbonButton29.Checked = false;
            ribbon1.ThemeColor = RibbonTheme.Normal;
        }

        private void ribbonButton29_Click(object sender, EventArgs e)
        {
            ribbon1.OrbStyle = RibbonOrbStyle.Office_2010;
            ribbonButton28.Checked = false;
            ribbonButton29.Checked = true;
            ribbon1.ThemeColor = RibbonTheme.Blue_2010;
        }

        private void ribbonButton30_Click(object sender, EventArgs e)
        {
            this.Hide();
            this.FormBorderStyle = FormBorderStyle.Sizable;
            ribbon1.BorderMode = RibbonWindowMode.NonClientAreaGlass;
            this.ControlBox = true;
            this.Show();

            ribbonButton77.Checked = false;
            ribbonButton30.Checked = true;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void ribbonUpDown1_UpButtonClicked(object sender, MouseEventArgs e)
        {
            TopPanel.Height += 1;
            ButtomPanel.Height += 1;
            ribbonUpDown1.TextBoxText = TopPanel.Height.ToString(); // Fix: Convert int to string using ToString()
        }

        private void ribbonUpDown1_DownButtonClicked(object sender, MouseEventArgs e)
        {
            TopPanel.Height -= 1;
            ButtomPanel.Height -= 1;
            ribbonUpDown1.TextBoxText = TopPanel.Height.ToString(); // Fix: Convert int to string using ToString()
        }

        private void ribbonUpDown1_TextBoxTextChanged(object sender, EventArgs e)
        {
            // Fix: Convert the string value of TextBoxText to an integer using int.TryParse
            if (int.TryParse(ribbonUpDown1.TextBoxText, out int newHeight))
            {
                TopPanel.Height = newHeight;
                ButtomPanel.Height = newHeight; // Assuming this is the intended behavior
            }
            else if (string.IsNullOrEmpty(ribbonUpDown1.TextBoxText))
            {
                // Handle the case where the TextBoxText is empty
                TopPanel.Height = 0;
                ButtomPanel.Height = 0; // Assuming this is the intended behavior
            }
            else
            {
                // Handle invalid input
                ribbonUpDown1.TextBoxText = RightPanel.Width.ToString();
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            // Set the initial values for the ribbonUpDown controls
            ribbonUpDown1.TextBoxText = TopPanel.Height.ToString();
            ribbonUpDown2.TextBoxText = RightPanel.Width.ToString();
            this.Size = new Size(1231, 700); // Set the initial size of the form

            EditPanel.AutoScroll = true; // Enable auto-scrolling for the EditPanel
            this.Text = textBox10.Text + " - Document Maker";

            WordInstallationValidationWizard wordInstallationValidationWizard = new WordInstallationValidationWizard();
            wordInstallationValidationWizard.ShowDialog(); // Show the Word installation validation wizard


            //Sheet1の内容を設定
            button1.Text = "○○発第○○号";

            string str3 = dateTimePicker1.Value.ToString("yyyy年M月d日");
            button2.Text = str3;
            checkBox1.Checked = false;

            textBox1.Text = "○○株式会社";
            textBox2.Text = "○○部　○○○○様";

            textBox3.Text = "○○○○株式会社";
            textBox4.Text = "東京都渋谷区○○町〇丁目〇番地〇号";
            textBox5.Text = "○○ビル 40階";
            textBox6.Text = "代表取締役　○○○○";
            textBox7.Text = "電話:0000-0000-0000";
            textBox8.Text = "FAX:0000-0000-0000";
            textBox9.Text = "メールアドレス:my@example.com";

            textBox10.Text = "新商品発表会のご案内";

            button3.Text = "拝啓、時下ますますご清栄のこととお慶び申し上げます。平素は並々ならぬお引き立てを賜り、厚く御礼申し上げます。";

            label2.Text = "敬具";



            textBox11.Text = "○○";
            numericUpDown1.Value = 0;
            dateTimePicker1.Value = DateTime.Now;

            textBox12.Text = "○○株式会社";
            textBox13.Text = "○○部　○○○○様";

            textBox14.Text = "○○○○株式会社";
            textBox15.Text = "東京都渋谷区○○町〇丁目〇番地〇号";
            textBox16.Text = "○○ビル 40階";
            textBox17.Text = "代表取締役　○○○○";
            textBox18.Text = "電話:0000-0000-0000";
            textBox19.Text = "FAX:0000-0000-0000";
            textBox21.Text = "メールアドレス:my@example.com";

            textBox20.Text = "新商品発表会のご案内";
        }

        private void ribbonUpDown2_UpButtonClicked(object sender, MouseEventArgs e)
        {
            RightPanel.Width += 1;
            LeftPanel.Width += 1;
            ribbonUpDown2.TextBoxText = RightPanel.Width.ToString(); // Fix: Convert int to string using ToString()
        }

        private void ribbonUpDown2_DownButtonClicked(object sender, MouseEventArgs e)
        {
            RightPanel.Width -= 1;
            LeftPanel.Width -= 1;
            ribbonUpDown2.TextBoxText = RightPanel.Width.ToString(); // Fix: Convert int to string using ToString()
        }

        private void ribbonUpDown2_TextBoxTextChanged(object sender, EventArgs e)
        {
            // Fix: Convert the string value of TextBoxText to an integer using int.TryParse
            if (int.TryParse(ribbonUpDown2.TextBoxText, out int newHeight))
            {
                RightPanel.Width = newHeight;
                LeftPanel.Width = newHeight; // Assuming this is the intended behavior
            }
            else if (string.IsNullOrEmpty(ribbonUpDown2.TextBoxText))
            {
                // Handle the case where the TextBoxText is empty
                RightPanel.Width = 0;
                LeftPanel.Width = 0; // Assuming this is the intended behavior
            }
            else
            {
                // Handle invalid input
                ribbonUpDown2.TextBoxText = RightPanel.Width.ToString();
            }
        }

        private void ribbonButton51_Click(object sender, EventArgs e)
        {
            //Wordインスタンスを生成
            Word.Application word = new Word.Application();
            //Wordウィンドウを表示
            word.Visible = true;
            //新規文書を作成
            Word.Document doc = word.Documents.Add();

            //ページの余白を設定
            doc.PageSetup.TopMargin = TopPanel.Height;    // 上余白 (1インチ = 72ポイント)
            doc.PageSetup.BottomMargin = ButtomPanel.Height; // 下余白
            doc.PageSetup.LeftMargin = LeftPanel.Width;   // 左余白
            doc.PageSetup.RightMargin = RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            if(button1.Text != string.Empty)
            {
                //1.発行番号を書く(paragraph)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph = doc.Content.Paragraphs.Add();
                paragraph.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                                                  //発行番号を書く
                paragraph.Range.Text = button1.Text;
                // 段落を右揃えに設定
                paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 改行を追加
                paragraph.Range.InsertParagraphAfter();
            }

            if (button2.Text != string.Empty)
            {
                //2.発行日を書く(paragraph2)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph2 = doc.Content.Paragraphs.Add();
                paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                                                   //発行番号を書く
                paragraph2.Range.Text = button2.Text;
                // 段落を右揃えに設定
                paragraph2.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 改行を追加
                paragraph2.Range.InsertParagraphAfter();
            }



            //相手先
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox1.Text != string.Empty)
            {
                //3.相手先の会社名を書く(paragraph3)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph3 = doc.Content.Paragraphs.Add();
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph3.Range.Text = textBox1.Text;
                // 段落を右揃えに設定
                paragraph3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph3.Range.InsertParagraphAfter();
            }

            if(textBox2.Text != string.Empty)
            {
                //3.相手先の名前を書く(paragraph4)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph4 = doc.Content.Paragraphs.Add();
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph4.Range.Text = textBox2.Text;
                // 段落を右揃えに設定
                paragraph4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph4.Range.InsertParagraphAfter();
            }


            //発信者
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox3.Text != string.Empty)
            {
                //3.発信者の会社名を書く(paragraph5)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph5 = doc.Content.Paragraphs.Add();
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph5.Range.Text = textBox3.Text;
                // 段落を右揃えに設定
                paragraph5.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph5.Range.InsertParagraphAfter();
            }

            if (textBox4.Text != string.Empty)
            {
                //3.発信者の住所を書く(paragraph6)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph6 = doc.Content.Paragraphs.Add();
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph6.Range.Text = textBox4.Text;
                // 段落を右揃えに設定
                paragraph6.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph6.Range.InsertParagraphAfter();
            }

            if (textBox5.Text != string.Empty)
            {
                //3.発信者の建物名称を書く(paragraph7)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph7 = doc.Content.Paragraphs.Add();
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph7.Range.Text = textBox5.Text;
                // 段落を右揃えに設定
                paragraph7.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph7.Range.InsertParagraphAfter();
            }

            if (textBox6.Text != string.Empty)
            {
                //3.発信者の名前を書く(paragraph8)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph8 = doc.Content.Paragraphs.Add();
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph8.Range.Text = textBox6.Text;
                // 段落を右揃えに設定
                paragraph8.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph8.Range.InsertParagraphAfter();
            }

            if (textBox7.Text != string.Empty)
            {
                //3.発信者の電話番号を書く(paragraph9)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph9 = doc.Content.Paragraphs.Add();
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph9.Range.Text = textBox7.Text;
                // 段落を右揃えに設定
                paragraph9.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph9.Range.InsertParagraphAfter();
            }

            if(textBox8.Text != string.Empty)
            {
                //3.発信者のFaxを書く(paragraph10)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph10 = doc.Content.Paragraphs.Add();
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph10.Range.Text = textBox8.Text;
                // 段落を右揃えに設定
                paragraph10.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph10.Range.InsertParagraphAfter();
            }

            if(textBox9.Text != string.Empty)
            {
                //3.発信者のメールアドレスを書く(paragraph11)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph11 = doc.Content.Paragraphs.Add();
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph11.Range.Text = textBox9.Text;
                // 段落を右揃えに設定
                paragraph11.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph11.Range.InsertParagraphAfter();
            }

            //内容
            //内容を入力するテキストボックスが空でない場合にのみ、内容を書く
            if (textBox10.Text != string.Empty)
            {
                //3.表題を書く(paragraph12)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph12 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph12.Range.Text = textBox10.Text;
                // 段落を右揃えに設定
                paragraph12.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                // 段落の変更を適用
                paragraph12.Range.InsertParagraphAfter();
            }


            if (button3.Text != string.Empty)
            {
                //3.内容を書く(paragraph13)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph13 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph13.Range.Text = button3.Text;
                // 段落を右揃えに設定
                paragraph13.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph13.Range.InsertParagraphAfter();
                paragraph13.Range.InsertParagraphAfter();
            }


            //「結語を入力しない」を選択した場合、敬具を入力しない
            if (ribbonCheckBox1.Checked == false)
            {
                //敬具
                Word.Paragraph paragraph14 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph14.Range.Text = label2.Text;
                // 段落を右揃えに設定
                paragraph14.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph14.Range.InsertParagraphAfter();
            }

            //記・以上
            Word.Paragraph paragraph15 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph15.Range.Text = label3.Text;
            // 段落を右揃えに設定
            paragraph15.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // 段落の変更を適用
            paragraph15.Range.InsertParagraphAfter();
            paragraph15.Range.InsertParagraphAfter();

            Word.Paragraph paragraph16 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph16.Range.Text = label5.Text;
            // 段落を右揃えに設定
            paragraph16.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 段落の変更を適用
            paragraph16.Range.InsertParagraphAfter();

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Wordドキュメント(*.docx)|*.docx";
            saveFileDialog1.Title = "文書ファイルを保存する場所を選択";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog1.FileName; // 選択されたファイルパスを取得
                doc.SaveAs2(filePath, Word.WdSaveFormat.wdFormatDocumentDefault); // ファイルを保存
                MessageBox.Show("指定された場所にファイルが正しく保存されました: " + filePath, "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                WCloseDialogYesNo dialog = new WCloseDialogYesNo();
                dialog.ShowDialog();
                if (dialog.DialogResult == DialogResult.Yes)
                {
                    // ドキュメントを保存せずに閉じる
                    doc.Close(false);
                    //Wordを終了
                    word.Quit();
                    //ガベージコレクション
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                else
                {
                    dialog.Close();
                }
            }
            //ガベージコレクション
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void ribbonButton44_Click(object sender, EventArgs e)
        {
            //Wordインスタンスを生成
            Word.Application word = new Word.Application();
            //Wordウィンドウを表示
            word.Visible = true;
            //新規文書を作成
            Word.Document doc = word.Documents.Add();

            //ページの余白を設定
            doc.PageSetup.TopMargin = TopPanel.Height;    // 上余白 (1インチ = 72ポイント)
            doc.PageSetup.BottomMargin = ButtomPanel.Height; // 下余白
            doc.PageSetup.LeftMargin = LeftPanel.Width;   // 左余白
            doc.PageSetup.RightMargin = RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //1.発行番号を書く(paragraph)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph = doc.Content.Paragraphs.Add();
            paragraph.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph.Range.Text = button1.Text;
            // 段落を右揃えに設定
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph.Range.InsertParagraphAfter();


            //2.発行日を書く(paragraph2)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph2 = doc.Content.Paragraphs.Add();
            paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph2.Range.Text = button2.Text;
            // 段落を右揃えに設定
            paragraph2.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph2.Range.InsertParagraphAfter();

            //相手先
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox1.Text != string.Empty)
            {
                //3.相手先の会社名を書く(paragraph3)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph3 = doc.Content.Paragraphs.Add();
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph3.Range.Text = textBox1.Text;
                // 段落を右揃えに設定
                paragraph3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph3.Range.InsertParagraphAfter();
            }

            if (textBox2.Text != string.Empty)
            {
                //3.相手先の名前を書く(paragraph4)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph4 = doc.Content.Paragraphs.Add();
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph4.Range.Text = textBox2.Text;
                // 段落を右揃えに設定
                paragraph4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph4.Range.InsertParagraphAfter();
            }


            //発信者
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox3.Text != string.Empty)
            {
                //3.発信者の会社名を書く(paragraph5)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph5 = doc.Content.Paragraphs.Add();
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph5.Range.Text = textBox3.Text;
                // 段落を右揃えに設定
                paragraph5.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph5.Range.InsertParagraphAfter();
            }

            if (textBox4.Text != string.Empty)
            {
                //3.発信者の住所を書く(paragraph6)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph6 = doc.Content.Paragraphs.Add();
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph6.Range.Text = textBox4.Text;
                // 段落を右揃えに設定
                paragraph6.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph6.Range.InsertParagraphAfter();
            }

            if (textBox5.Text != string.Empty)
            {
                //3.発信者の建物名称を書く(paragraph7)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph7 = doc.Content.Paragraphs.Add();
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph7.Range.Text = textBox5.Text;
                // 段落を右揃えに設定
                paragraph7.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph7.Range.InsertParagraphAfter();
            }

            if (textBox6.Text != string.Empty)
            {
                //3.発信者の名前を書く(paragraph8)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph8 = doc.Content.Paragraphs.Add();
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph8.Range.Text = textBox6.Text;
                // 段落を右揃えに設定
                paragraph8.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph8.Range.InsertParagraphAfter();
            }

            if (textBox7.Text != string.Empty)
            {
                //3.発信者の電話番号を書く(paragraph9)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph9 = doc.Content.Paragraphs.Add();
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph9.Range.Text = textBox7.Text;
                // 段落を右揃えに設定
                paragraph9.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph9.Range.InsertParagraphAfter();
            }

            if (textBox8.Text != string.Empty)
            {
                //3.発信者のFaxを書く(paragraph10)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph10 = doc.Content.Paragraphs.Add();
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph10.Range.Text = textBox8.Text;
                // 段落を右揃えに設定
                paragraph10.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph10.Range.InsertParagraphAfter();
            }

            if (textBox9.Text != string.Empty)
            {
                //3.発信者のメールアドレスを書く(paragraph11)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph11 = doc.Content.Paragraphs.Add();
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph11.Range.Text = textBox9.Text;
                // 段落を右揃えに設定
                paragraph11.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph11.Range.InsertParagraphAfter();
            }

            //内容
            //内容を入力するテキストボックスが空でない場合にのみ、内容を書く
            if (textBox10.Text != string.Empty)
            {
                //3.表題を書く(paragraph12)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph12 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph12.Range.Text = textBox10.Text;
                // 段落を右揃えに設定
                paragraph12.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                // 段落の変更を適用
                paragraph12.Range.InsertParagraphAfter();
            }


            if (button3.Text != string.Empty)
            {
                //3.内容を書く(paragraph13)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph13 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph13.Range.Text = button3.Text;
                // 段落を右揃えに設定
                paragraph13.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph13.Range.InsertParagraphAfter();
                paragraph13.Range.InsertParagraphAfter();
            }

            //「結語を入力しない」を選択した場合、敬具を入力しない
            if (ribbonCheckBox1.Checked == false)
            {
                //敬具
                Word.Paragraph paragraph14 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph14.Range.Text = label2.Text;
                // 段落を右揃えに設定
                paragraph14.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph14.Range.InsertParagraphAfter();
            }


            //記・以上
            Word.Paragraph paragraph15 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph15.Range.Text = label3.Text;
            // 段落を右揃えに設定
            paragraph15.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // 段落の変更を適用
            paragraph15.Range.InsertParagraphAfter();
            paragraph15.Range.InsertParagraphAfter();

            Word.Paragraph paragraph16 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph16.Range.Text = label5.Text;
            // 段落を右揃えに設定
            paragraph16.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 段落の変更を適用
            paragraph16.Range.InsertParagraphAfter();

        }

        private void ribbonButton21_Click(object sender, EventArgs e)
        {
            RightPanel.Width = 120;
            LeftPanel.Width = 120;
            ribbonUpDown2.TextBoxText = RightPanel.Width.ToString(); // Fix: Convert int to string using ToString()

            TopPanel.Height = 120;
            ButtomPanel.Height = 120;
            ribbonUpDown1.TextBoxText = TopPanel.Height.ToString(); // Fix: Convert int to string using ToString()
        }

        private void ribbonButton39_Click(object sender, EventArgs e)
        {
            RightPanel.Width = 72;
            LeftPanel.Width = 72;
            ribbonUpDown2.TextBoxText = RightPanel.Width.ToString(); // Fix: Convert int to string using ToString()

            TopPanel.Height = 72;
            ButtomPanel.Height = 72;
            ribbonUpDown1.TextBoxText = TopPanel.Height.ToString(); // Fix: Convert int to string using ToString()
        }

        private void ribbonButton40_Click(object sender, EventArgs e)
        {
            RightPanel.Width = 40;
            LeftPanel.Width = 40;
            ribbonUpDown2.TextBoxText = RightPanel.Width.ToString(); // Fix: Convert int to string using ToString()

            TopPanel.Height = 40;
            ButtomPanel.Height = 40;
            ribbonUpDown1.TextBoxText = TopPanel.Height.ToString(); // Fix: Convert int to string using ToString()
        }

        private void ribbonButton41_Click(object sender, EventArgs e)
        {
            RightPanel.Width = 72;
            LeftPanel.Width = 72;
            ribbonUpDown2.TextBoxText = RightPanel.Width.ToString(); // Fix: Convert int to string using ToString()

            TopPanel.Height = 72;
            ButtomPanel.Height = 72;
            ribbonUpDown1.TextBoxText = TopPanel.Height.ToString(); // Fix: Convert int to string using ToString()
        }

        private void ribbonButton42_Click(object sender, EventArgs e)
        {

            TopPanel.Height = 72;
            ButtomPanel.Height = 72;
            ribbonUpDown1.TextBoxText = TopPanel.Height.ToString(); // Fix: Convert int to string using ToString()
        }

        private void ribbonButton43_Click(object sender, EventArgs e)
        {
            RightPanel.Width = 72;
            LeftPanel.Width = 72;
            ribbonUpDown2.TextBoxText = RightPanel.Width.ToString(); // Fix: Convert int to string using ToString()
        }

        private void ribbonButton46_Click(object sender, EventArgs e)
        {
            RightPanel.Width = 72;
            LeftPanel.Width = 72;
            ribbonUpDown2.TextBoxText = RightPanel.Width.ToString(); // Fix: Convert int to string using ToString()

            TopPanel.Height = 72;
            ButtomPanel.Height = 72;
            ribbonUpDown1.TextBoxText = TopPanel.Height.ToString(); // Fix: Convert int to string using ToString()
        }

        private void ribbonOrbOptionButton1_Click(object sender, EventArgs e)
        {
            AppExit appExit = new AppExit();
            appExit.Show();
        }

        private void ribbon1_OrbDropDown_Click(object sender, EventArgs e)
        {
            ribbon1.ShowOrbDropDown();
        }

        private void ribbonButton48_Click(object sender, EventArgs e)
        {
            About about = new About();
            about.ShowDialog();
        }

        private void ribbonButton47_Click(object sender, EventArgs e)
        {
            Third_parties third_Parties = new Third_parties();
            third_Parties.ShowDialog();
        }

        private void ribbonOrbOptionButton1_MouseUp(object sender, MouseEventArgs e)
        {
            AppExit appExit = new AppExit();
            appExit.Show();
        }

        private void ribbonButton18_Click(object sender, EventArgs e)
        {
            //Wordインスタンスを生成
            Word.Application word = new Word.Application();
            //Wordウィンドウを表示
            word.Visible = true;
            //新規文書を作成
            Word.Document doc = word.Documents.Add();

            //ページの余白を設定
            doc.PageSetup.TopMargin = TopPanel.Height;    // 上余白 (1インチ = 72ポイント)
            doc.PageSetup.BottomMargin = ButtomPanel.Height; // 下余白
            doc.PageSetup.LeftMargin = LeftPanel.Width;   // 左余白
            doc.PageSetup.RightMargin = RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //1.発行番号を書く(paragraph)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph = doc.Content.Paragraphs.Add();
            paragraph.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph.Range.Text = button1.Text;
            // 段落を右揃えに設定
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph.Range.InsertParagraphAfter();


            //2.発行日を書く(paragraph2)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph2 = doc.Content.Paragraphs.Add();
            paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph2.Range.Text = button2.Text;
            // 段落を右揃えに設定
            paragraph2.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph2.Range.InsertParagraphAfter();

            //相手先
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox1.Text != string.Empty)
            {
                //3.相手先の会社名を書く(paragraph3)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph3 = doc.Content.Paragraphs.Add();
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph3.Range.Text = textBox1.Text;
                // 段落を右揃えに設定
                paragraph3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph3.Range.InsertParagraphAfter();
            }

            if (textBox2.Text != string.Empty)
            {
                //3.相手先の名前を書く(paragraph4)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph4 = doc.Content.Paragraphs.Add();
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph4.Range.Text = textBox2.Text;
                // 段落を右揃えに設定
                paragraph4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph4.Range.InsertParagraphAfter();
            }


            //発信者
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox3.Text != string.Empty)
            {
                //3.発信者の会社名を書く(paragraph5)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph5 = doc.Content.Paragraphs.Add();
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph5.Range.Text = textBox3.Text;
                // 段落を右揃えに設定
                paragraph5.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph5.Range.InsertParagraphAfter();
            }

            if (textBox4.Text != string.Empty)
            {
                //3.発信者の住所を書く(paragraph6)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph6 = doc.Content.Paragraphs.Add();
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph6.Range.Text = textBox4.Text;
                // 段落を右揃えに設定
                paragraph6.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph6.Range.InsertParagraphAfter();
            }

            if (textBox5.Text != string.Empty)
            {
                //3.発信者の建物名称を書く(paragraph7)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph7 = doc.Content.Paragraphs.Add();
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph7.Range.Text = textBox5.Text;
                // 段落を右揃えに設定
                paragraph7.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph7.Range.InsertParagraphAfter();
            }

            if (textBox6.Text != string.Empty)
            {
                //3.発信者の名前を書く(paragraph8)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph8 = doc.Content.Paragraphs.Add();
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph8.Range.Text = textBox6.Text;
                // 段落を右揃えに設定
                paragraph8.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph8.Range.InsertParagraphAfter();
            }

            if (textBox7.Text != string.Empty)
            {
                //3.発信者の電話番号を書く(paragraph9)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph9 = doc.Content.Paragraphs.Add();
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph9.Range.Text = textBox7.Text;
                // 段落を右揃えに設定
                paragraph9.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph9.Range.InsertParagraphAfter();
            }

            if (textBox8.Text != string.Empty)
            {
                //3.発信者のFaxを書く(paragraph10)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph10 = doc.Content.Paragraphs.Add();
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph10.Range.Text = textBox8.Text;
                // 段落を右揃えに設定
                paragraph10.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph10.Range.InsertParagraphAfter();
            }

            if (textBox9.Text != string.Empty)
            {
                //3.発信者のメールアドレスを書く(paragraph11)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph11 = doc.Content.Paragraphs.Add();
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph11.Range.Text = textBox9.Text;
                // 段落を右揃えに設定
                paragraph11.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph11.Range.InsertParagraphAfter();
            }

            //内容
            //内容を入力するテキストボックスが空でない場合にのみ、内容を書く
            if (textBox10.Text != string.Empty)
            {
                //3.表題を書く(paragraph12)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph12 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph12.Range.Text = textBox10.Text;
                // 段落を右揃えに設定
                paragraph12.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                // 段落の変更を適用
                paragraph12.Range.InsertParagraphAfter();
            }


            if (button3.Text != string.Empty)
            {
                //3.内容を書く(paragraph13)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph13 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph13.Range.Text = button3.Text;
                // 段落を右揃えに設定
                paragraph13.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph13.Range.InsertParagraphAfter();
                paragraph13.Range.InsertParagraphAfter();
            }

            //敬具
            Word.Paragraph paragraph14 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph14.Range.Text = label2.Text;
            // 段落を右揃えに設定
            paragraph14.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 段落の変更を適用
            paragraph14.Range.InsertParagraphAfter();

            //記・以上
            Word.Paragraph paragraph15 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph15.Range.Text = label3.Text;
            // 段落を右揃えに設定
            paragraph15.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // 段落の変更を適用
            paragraph15.Range.InsertParagraphAfter();
            paragraph15.Range.InsertParagraphAfter();

            Word.Paragraph paragraph16 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph16.Range.Text = label5.Text;
            // 段落を右揃えに設定
            paragraph16.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 段落の変更を適用
            paragraph16.Range.InsertParagraphAfter();
        }

        private void ribbonOrbMenuItem2_Click(object sender, EventArgs e)
        {
            //Wordインスタンスを生成
            Word.Application word = new Word.Application();
            //Wordウィンドウを表示
            word.Visible = true;
            //新規文書を作成
            Word.Document doc = word.Documents.Add();

            //ページの余白を設定
            doc.PageSetup.TopMargin = TopPanel.Height;    // 上余白 (1インチ = 72ポイント)
            doc.PageSetup.BottomMargin = ButtomPanel.Height; // 下余白
            doc.PageSetup.LeftMargin = LeftPanel.Width;   // 左余白
            doc.PageSetup.RightMargin = RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //1.発行番号を書く(paragraph)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph = doc.Content.Paragraphs.Add();
            paragraph.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph.Range.Text = button1.Text;
            // 段落を右揃えに設定
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph.Range.InsertParagraphAfter();


            //2.発行日を書く(paragraph2)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph2 = doc.Content.Paragraphs.Add();
            paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph2.Range.Text = button2.Text;
            // 段落を右揃えに設定
            paragraph2.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph2.Range.InsertParagraphAfter();

            //相手先
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox1.Text != string.Empty)
            {
                //3.相手先の会社名を書く(paragraph3)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph3 = doc.Content.Paragraphs.Add();
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph3.Range.Text = textBox1.Text;
                // 段落を右揃えに設定
                paragraph3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph3.Range.InsertParagraphAfter();
            }

            if (textBox2.Text != string.Empty)
            {
                //3.相手先の名前を書く(paragraph4)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph4 = doc.Content.Paragraphs.Add();
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph4.Range.Text = textBox2.Text;
                // 段落を右揃えに設定
                paragraph4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph4.Range.InsertParagraphAfter();
            }


            //発信者
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox3.Text != string.Empty)
            {
                //3.発信者の会社名を書く(paragraph5)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph5 = doc.Content.Paragraphs.Add();
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph5.Range.Text = textBox3.Text;
                // 段落を右揃えに設定
                paragraph5.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph5.Range.InsertParagraphAfter();
            }

            if (textBox4.Text != string.Empty)
            {
                //3.発信者の住所を書く(paragraph6)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph6 = doc.Content.Paragraphs.Add();
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph6.Range.Text = textBox4.Text;
                // 段落を右揃えに設定
                paragraph6.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph6.Range.InsertParagraphAfter();
            }

            if (textBox5.Text != string.Empty)
            {
                //3.発信者の建物名称を書く(paragraph7)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph7 = doc.Content.Paragraphs.Add();
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph7.Range.Text = textBox5.Text;
                // 段落を右揃えに設定
                paragraph7.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph7.Range.InsertParagraphAfter();
            }

            if (textBox6.Text != string.Empty)
            {
                //3.発信者の名前を書く(paragraph8)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph8 = doc.Content.Paragraphs.Add();
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph8.Range.Text = textBox6.Text;
                // 段落を右揃えに設定
                paragraph8.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph8.Range.InsertParagraphAfter();
            }

            if (textBox7.Text != string.Empty)
            {
                //3.発信者の電話番号を書く(paragraph9)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph9 = doc.Content.Paragraphs.Add();
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph9.Range.Text = textBox7.Text;
                // 段落を右揃えに設定
                paragraph9.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph9.Range.InsertParagraphAfter();
            }

            if (textBox8.Text != string.Empty)
            {
                //3.発信者のFaxを書く(paragraph10)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph10 = doc.Content.Paragraphs.Add();
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph10.Range.Text = textBox8.Text;
                // 段落を右揃えに設定
                paragraph10.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph10.Range.InsertParagraphAfter();
            }

            if (textBox9.Text != string.Empty)
            {
                //3.発信者のメールアドレスを書く(paragraph11)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph11 = doc.Content.Paragraphs.Add();
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph11.Range.Text = textBox9.Text;
                // 段落を右揃えに設定
                paragraph11.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph11.Range.InsertParagraphAfter();
            }

            //内容
            //内容を入力するテキストボックスが空でない場合にのみ、内容を書く
            if (textBox10.Text != string.Empty)
            {
                //3.表題を書く(paragraph12)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph12 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph12.Range.Text = textBox10.Text;
                // 段落を右揃えに設定
                paragraph12.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                // 段落の変更を適用
                paragraph12.Range.InsertParagraphAfter();
            }


            if (button3.Text != string.Empty)
            {
                //3.内容を書く(paragraph13)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph13 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph13.Range.Text = button3.Text;
                // 段落を右揃えに設定
                paragraph13.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph13.Range.InsertParagraphAfter();
                paragraph13.Range.InsertParagraphAfter();
            }


            //「結語を入力しない」を選択した場合、敬具を入力しない
            if (ribbonCheckBox1.Checked == false)
            {
                //敬具
                Word.Paragraph paragraph14 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph14.Range.Text = label2.Text;
                // 段落を右揃えに設定
                paragraph14.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph14.Range.InsertParagraphAfter();
            }

            //記・以上
            Word.Paragraph paragraph15 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph15.Range.Text = label3.Text;
            // 段落を右揃えに設定
            paragraph15.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // 段落の変更を適用
            paragraph15.Range.InsertParagraphAfter();
            paragraph15.Range.InsertParagraphAfter();

            Word.Paragraph paragraph16 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph16.Range.Text = label5.Text;
            // 段落を右揃えに設定
            paragraph16.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 段落の変更を適用
            paragraph16.Range.InsertParagraphAfter();
        }

        private void ribbonOrbMenuItem3_Click(object sender, EventArgs e)
        {

            //Wordインスタンスを生成
            Word.Application word = new Word.Application();
            //Wordウィンドウを表示
            word.Visible = true;
            //新規文書を作成
            Word.Document doc = word.Documents.Add();

            //ページの余白を設定
            doc.PageSetup.TopMargin = TopPanel.Height;    // 上余白 (1インチ = 72ポイント)
            doc.PageSetup.BottomMargin = ButtomPanel.Height; // 下余白
            doc.PageSetup.LeftMargin = LeftPanel.Width;   // 左余白
            doc.PageSetup.RightMargin = RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //1.発行番号を書く(paragraph)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph = doc.Content.Paragraphs.Add();
            paragraph.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph.Range.Text = button1.Text;
            // 段落を右揃えに設定
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph.Range.InsertParagraphAfter();


            //2.発行日を書く(paragraph2)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph2 = doc.Content.Paragraphs.Add();
            paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph2.Range.Text = button2.Text;
            // 段落を右揃えに設定
            paragraph2.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph2.Range.InsertParagraphAfter();

            //相手先
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox1.Text != string.Empty)
            {
                //3.相手先の会社名を書く(paragraph3)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph3 = doc.Content.Paragraphs.Add();
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph3.Range.Text = textBox1.Text;
                // 段落を右揃えに設定
                paragraph3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph3.Range.InsertParagraphAfter();
            }

            if (textBox2.Text != string.Empty)
            {
                //3.相手先の名前を書く(paragraph4)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph4 = doc.Content.Paragraphs.Add();
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph4.Range.Text = textBox2.Text;
                // 段落を右揃えに設定
                paragraph4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph4.Range.InsertParagraphAfter();
            }


            //発信者
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox3.Text != string.Empty)
            {
                //3.発信者の会社名を書く(paragraph5)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph5 = doc.Content.Paragraphs.Add();
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph5.Range.Text = textBox3.Text;
                // 段落を右揃えに設定
                paragraph5.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph5.Range.InsertParagraphAfter();
            }

            if (textBox4.Text != string.Empty)
            {
                //3.発信者の住所を書く(paragraph6)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph6 = doc.Content.Paragraphs.Add();
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph6.Range.Text = textBox4.Text;
                // 段落を右揃えに設定
                paragraph6.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph6.Range.InsertParagraphAfter();
            }

            if (textBox5.Text != string.Empty)
            {
                //3.発信者の建物名称を書く(paragraph7)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph7 = doc.Content.Paragraphs.Add();
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph7.Range.Text = textBox5.Text;
                // 段落を右揃えに設定
                paragraph7.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph7.Range.InsertParagraphAfter();
            }

            if (textBox6.Text != string.Empty)
            {
                //3.発信者の名前を書く(paragraph8)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph8 = doc.Content.Paragraphs.Add();
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph8.Range.Text = textBox6.Text;
                // 段落を右揃えに設定
                paragraph8.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph8.Range.InsertParagraphAfter();
            }

            if (textBox7.Text != string.Empty)
            {
                //3.発信者の電話番号を書く(paragraph9)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph9 = doc.Content.Paragraphs.Add();
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph9.Range.Text = textBox7.Text;
                // 段落を右揃えに設定
                paragraph9.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph9.Range.InsertParagraphAfter();
            }

            if (textBox8.Text != string.Empty)
            {
                //3.発信者のFaxを書く(paragraph10)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph10 = doc.Content.Paragraphs.Add();
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph10.Range.Text = textBox8.Text;
                // 段落を右揃えに設定
                paragraph10.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph10.Range.InsertParagraphAfter();
            }

            if (textBox9.Text != string.Empty)
            {
                //3.発信者のメールアドレスを書く(paragraph11)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph11 = doc.Content.Paragraphs.Add();
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph11.Range.Text = textBox9.Text;
                // 段落を右揃えに設定
                paragraph11.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph11.Range.InsertParagraphAfter();
            }

            //内容
            //内容を入力するテキストボックスが空でない場合にのみ、内容を書く
            if (textBox10.Text != string.Empty)
            {
                //3.表題を書く(paragraph12)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph12 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph12.Range.Text = textBox10.Text;
                // 段落を右揃えに設定
                paragraph12.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                // 段落の変更を適用
                paragraph12.Range.InsertParagraphAfter();
            }


            if (button3.Text != string.Empty)
            {
                //3.内容を書く(paragraph13)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph13 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph13.Range.Text = button3.Text;
                // 段落を右揃えに設定
                paragraph13.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph13.Range.InsertParagraphAfter();
                paragraph13.Range.InsertParagraphAfter();
            }


            //「結語を入力しない」を選択した場合、敬具を入力しない
            if (ribbonCheckBox1.Checked == false)
            {
                //敬具
                Word.Paragraph paragraph14 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph14.Range.Text = label2.Text;
                // 段落を右揃えに設定
                paragraph14.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph14.Range.InsertParagraphAfter();
            }

            //記・以上
            Word.Paragraph paragraph15 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph15.Range.Text = label3.Text;
            // 段落を右揃えに設定
            paragraph15.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // 段落の変更を適用
            paragraph15.Range.InsertParagraphAfter();
            paragraph15.Range.InsertParagraphAfter();

            Word.Paragraph paragraph16 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph16.Range.Text = label5.Text;
            // 段落を右揃えに設定
            paragraph16.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 段落の変更を適用
            paragraph16.Range.InsertParagraphAfter();

            //エラーが出やすいのでtry-catchで囲む
            try
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Wordドキュメント(*.docx)|*.docx";
                saveFileDialog1.Title = "文書ファイルを保存する場所を選択";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog1.FileName; // 選択されたファイルパスを取得
                    doc.SaveAs2(filePath, Word.WdSaveFormat.wdFormatDocumentDefault); // ファイルを保存
                    MessageBox.Show("指定された場所にファイルが正しく保存されました: " + filePath, "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    WCloseDialogYesNo dialog = new WCloseDialogYesNo();
                    dialog.ShowDialog();
                    if (dialog.DialogResult == DialogResult.Yes)
                    {
                        // ドキュメントを保存せずに閉じる
                        doc.Close(false);
                        //Wordを終了
                        word.Quit();
                        //ガベージコレクション
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                    else
                    {
                        dialog.Close();
                    }
                }
            }
            catch(Exception)
            {
                //ガベージコレクション
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            //ガベージコレクション
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void ribbonButton22_Click(object sender, EventArgs e)
        {
            //Wordインスタンスを生成
            Word.Application word = new Word.Application();
            //Wordウィンドウを表示
            word.Visible = true;
            //新規文書を作成
            Word.Document doc = word.Documents.Add();

            //ページの余白を設定
            doc.PageSetup.TopMargin = TopPanel.Height;    // 上余白 (1インチ = 72ポイント)
            doc.PageSetup.BottomMargin = ButtomPanel.Height; // 下余白
            doc.PageSetup.LeftMargin = LeftPanel.Width;   // 左余白
            doc.PageSetup.RightMargin = RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //1.発行番号を書く(paragraph)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph = doc.Content.Paragraphs.Add();
            paragraph.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph.Range.Text = button1.Text;
            // 段落を右揃えに設定
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph.Range.InsertParagraphAfter();


            //2.発行日を書く(paragraph2)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph2 = doc.Content.Paragraphs.Add();
            paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph2.Range.Text = button2.Text;
            // 段落を右揃えに設定
            paragraph2.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph2.Range.InsertParagraphAfter();

            //相手先
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox1.Text != string.Empty)
            {
                //3.相手先の会社名を書く(paragraph3)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph3 = doc.Content.Paragraphs.Add();
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph3.Range.Text = textBox1.Text;
                // 段落を右揃えに設定
                paragraph3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph3.Range.InsertParagraphAfter();
            }

            if (textBox2.Text != string.Empty)
            {
                //3.相手先の名前を書く(paragraph4)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph4 = doc.Content.Paragraphs.Add();
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph4.Range.Text = textBox2.Text;
                // 段落を右揃えに設定
                paragraph4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph4.Range.InsertParagraphAfter();
            }


            //発信者
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox3.Text != string.Empty)
            {
                //3.発信者の会社名を書く(paragraph5)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph5 = doc.Content.Paragraphs.Add();
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph5.Range.Text = textBox3.Text;
                // 段落を右揃えに設定
                paragraph5.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph5.Range.InsertParagraphAfter();
            }

            if (textBox4.Text != string.Empty)
            {
                //3.発信者の住所を書く(paragraph6)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph6 = doc.Content.Paragraphs.Add();
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph6.Range.Text = textBox4.Text;
                // 段落を右揃えに設定
                paragraph6.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph6.Range.InsertParagraphAfter();
            }

            if (textBox5.Text != string.Empty)
            {
                //3.発信者の建物名称を書く(paragraph7)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph7 = doc.Content.Paragraphs.Add();
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph7.Range.Text = textBox5.Text;
                // 段落を右揃えに設定
                paragraph7.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph7.Range.InsertParagraphAfter();
            }

            if (textBox6.Text != string.Empty)
            {
                //3.発信者の名前を書く(paragraph8)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph8 = doc.Content.Paragraphs.Add();
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph8.Range.Text = textBox6.Text;
                // 段落を右揃えに設定
                paragraph8.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph8.Range.InsertParagraphAfter();
            }

            if (textBox7.Text != string.Empty)
            {
                //3.発信者の電話番号を書く(paragraph9)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph9 = doc.Content.Paragraphs.Add();
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph9.Range.Text = textBox7.Text;
                // 段落を右揃えに設定
                paragraph9.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph9.Range.InsertParagraphAfter();
            }

            if (textBox8.Text != string.Empty)
            {
                //3.発信者のFaxを書く(paragraph10)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph10 = doc.Content.Paragraphs.Add();
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph10.Range.Text = textBox8.Text;
                // 段落を右揃えに設定
                paragraph10.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph10.Range.InsertParagraphAfter();
            }

            if (textBox9.Text != string.Empty)
            {
                //3.発信者のメールアドレスを書く(paragraph11)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph11 = doc.Content.Paragraphs.Add();
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph11.Range.Text = textBox9.Text;
                // 段落を右揃えに設定
                paragraph11.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph11.Range.InsertParagraphAfter();
            }

            //内容
            //内容を入力するテキストボックスが空でない場合にのみ、内容を書く
            if (textBox10.Text != string.Empty)
            {
                //3.表題を書く(paragraph12)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph12 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph12.Range.Text = textBox10.Text;
                // 段落を右揃えに設定
                paragraph12.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                // 段落の変更を適用
                paragraph12.Range.InsertParagraphAfter();
            }


            if (button3.Text != string.Empty)
            {
                //3.内容を書く(paragraph13)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph13 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph13.Range.Text = button3.Text;
                // 段落を右揃えに設定
                paragraph13.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph13.Range.InsertParagraphAfter();
                paragraph13.Range.InsertParagraphAfter();
            }


            //「結語を入力しない」を選択した場合、敬具を入力しない
            if (ribbonCheckBox1.Checked == false)
            {
                //敬具
                Word.Paragraph paragraph14 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph14.Range.Text = label2.Text;
                // 段落を右揃えに設定
                paragraph14.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph14.Range.InsertParagraphAfter();
            }

            //記・以上
            Word.Paragraph paragraph15 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph15.Range.Text = label3.Text;
            // 段落を右揃えに設定
            paragraph15.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // 段落の変更を適用
            paragraph15.Range.InsertParagraphAfter();
            paragraph15.Range.InsertParagraphAfter();

            Word.Paragraph paragraph16 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph16.Range.Text = label5.Text;
            // 段落を右揃えに設定
            paragraph16.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 段落の変更を適用
            paragraph16.Range.InsertParagraphAfter();


            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Wordドキュメント(*.docx)|*.docx";
            saveFileDialog1.Title = "文書ファイルを保存する場所を選択";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog1.FileName; // 選択されたファイルパスを取得
                doc.SaveAs2(filePath, Word.WdSaveFormat.wdFormatDocumentDefault); // ファイルを保存
                MessageBox.Show("指定された場所にファイルが正しく保存されました: " + filePath, "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                WCloseDialogYesNo dialog = new WCloseDialogYesNo();
                dialog.ShowDialog();
                if (dialog.DialogResult == DialogResult.Yes)
                {
                    // ドキュメントを保存せずに閉じる
                    doc.Close(false);
                    //Wordを終了
                    word.Quit();
                    //ガベージコレクション
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                else
                {
                    dialog.Close();
                }
            }
            //ガベージコレクション
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void ribbonButton24_Click(object sender, EventArgs e)
        {
            //Wordインスタンスを生成
            Word.Application word = new Word.Application();
            //Wordウィンドウを表示
            word.Visible = true;
            //新規文書を作成
            Word.Document doc = word.Documents.Add();

            //ページの余白を設定
            doc.PageSetup.TopMargin = TopPanel.Height;    // 上余白 (1インチ = 72ポイント)
            doc.PageSetup.BottomMargin = ButtomPanel.Height; // 下余白
            doc.PageSetup.LeftMargin = LeftPanel.Width;   // 左余白
            doc.PageSetup.RightMargin = RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //1.発行番号を書く(paragraph)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph = doc.Content.Paragraphs.Add();
            paragraph.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph.Range.Text = button1.Text;
            // 段落を右揃えに設定
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph.Range.InsertParagraphAfter();


            //2.発行日を書く(paragraph2)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph2 = doc.Content.Paragraphs.Add();
            paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph2.Range.Text = button2.Text;
            // 段落を右揃えに設定
            paragraph2.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph2.Range.InsertParagraphAfter();

            //相手先
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox1.Text != string.Empty)
            {
                //3.相手先の会社名を書く(paragraph3)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph3 = doc.Content.Paragraphs.Add();
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph3.Range.Text = textBox1.Text;
                // 段落を右揃えに設定
                paragraph3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph3.Range.InsertParagraphAfter();
            }

            if (textBox2.Text != string.Empty)
            {
                //3.相手先の名前を書く(paragraph4)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph4 = doc.Content.Paragraphs.Add();
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph4.Range.Text = textBox2.Text;
                // 段落を右揃えに設定
                paragraph4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph4.Range.InsertParagraphAfter();
            }


            //発信者
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox3.Text != string.Empty)
            {
                //3.発信者の会社名を書く(paragraph5)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph5 = doc.Content.Paragraphs.Add();
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph5.Range.Text = textBox3.Text;
                // 段落を右揃えに設定
                paragraph5.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph5.Range.InsertParagraphAfter();
            }

            if (textBox4.Text != string.Empty)
            {
                //3.発信者の住所を書く(paragraph6)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph6 = doc.Content.Paragraphs.Add();
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph6.Range.Text = textBox4.Text;
                // 段落を右揃えに設定
                paragraph6.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph6.Range.InsertParagraphAfter();
            }

            if (textBox5.Text != string.Empty)
            {
                //3.発信者の建物名称を書く(paragraph7)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph7 = doc.Content.Paragraphs.Add();
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph7.Range.Text = textBox5.Text;
                // 段落を右揃えに設定
                paragraph7.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph7.Range.InsertParagraphAfter();
            }

            if (textBox6.Text != string.Empty)
            {
                //3.発信者の名前を書く(paragraph8)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph8 = doc.Content.Paragraphs.Add();
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph8.Range.Text = textBox6.Text;
                // 段落を右揃えに設定
                paragraph8.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph8.Range.InsertParagraphAfter();
            }

            if (textBox7.Text != string.Empty)
            {
                //3.発信者の電話番号を書く(paragraph9)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph9 = doc.Content.Paragraphs.Add();
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph9.Range.Text = textBox7.Text;
                // 段落を右揃えに設定
                paragraph9.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph9.Range.InsertParagraphAfter();
            }

            if (textBox8.Text != string.Empty)
            {
                //3.発信者のFaxを書く(paragraph10)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph10 = doc.Content.Paragraphs.Add();
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph10.Range.Text = textBox8.Text;
                // 段落を右揃えに設定
                paragraph10.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph10.Range.InsertParagraphAfter();
            }

            if (textBox9.Text != string.Empty)
            {
                //3.発信者のメールアドレスを書く(paragraph11)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph11 = doc.Content.Paragraphs.Add();
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph11.Range.Text = textBox9.Text;
                // 段落を右揃えに設定
                paragraph11.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph11.Range.InsertParagraphAfter();
            }

            //内容
            //内容を入力するテキストボックスが空でない場合にのみ、内容を書く
            if (textBox10.Text != string.Empty)
            {
                //3.表題を書く(paragraph12)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph12 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph12.Range.Text = textBox10.Text;
                // 段落を右揃えに設定
                paragraph12.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                // 段落の変更を適用
                paragraph12.Range.InsertParagraphAfter();
            }


            if (button3.Text != string.Empty)
            {
                //3.内容を書く(paragraph13)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph13 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph13.Range.Text = button3.Text;
                // 段落を右揃えに設定
                paragraph13.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph13.Range.InsertParagraphAfter();
                paragraph13.Range.InsertParagraphAfter();
            }


            //「結語を入力しない」を選択した場合、敬具を入力しない
            if (ribbonCheckBox1.Checked == false)
            {
                //敬具
                Word.Paragraph paragraph14 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph14.Range.Text = label2.Text;
                // 段落を右揃えに設定
                paragraph14.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph14.Range.InsertParagraphAfter();
            }

            //記・以上
            Word.Paragraph paragraph15 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph15.Range.Text = label3.Text;
            // 段落を右揃えに設定
            paragraph15.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // 段落の変更を適用
            paragraph15.Range.InsertParagraphAfter();
            paragraph15.Range.InsertParagraphAfter();

            Word.Paragraph paragraph16 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph16.Range.Text = label5.Text;
            // 段落を右揃えに設定
            paragraph16.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 段落の変更を適用
            paragraph16.Range.InsertParagraphAfter();
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            this.Text = textBox10.Text + " - Document Maker";
            if(textBox10.Text == string.Empty)
            {
                this.Text = "無題 - Document Maker";
            }
        }

        private void textBox10_ParentChanged(object sender, EventArgs e)
        {
            textBox20.Text = textBox10.Text;
        }

        private void ribbonButton9_Click(object sender, EventArgs e)
        {
            DCW dCW = new DCW();
            dCW.ShowDialog();
        }

        private void ribbonOrbMenuItem1_Click(object sender, EventArgs e)
        {
            DCW dCW = new DCW();
            dCW.ShowDialog();
        }

        private void ribbonOrbMenuItem1_MouseUp(object sender, MouseEventArgs e)
        {
            DCW dCW = new DCW();
            dCW.ShowDialog();
        }

        private void ribbonOrbMenuItem2_MouseUp(object sender, MouseEventArgs e)
        {
            //Wordインスタンスを生成
            Word.Application word = new Word.Application();
            //Wordウィンドウを表示
            word.Visible = true;
            //新規文書を作成
            Word.Document doc = word.Documents.Add();

            //ページの余白を設定
            doc.PageSetup.TopMargin = TopPanel.Height;    // 上余白 (1インチ = 72ポイント)
            doc.PageSetup.BottomMargin = ButtomPanel.Height; // 下余白
            doc.PageSetup.LeftMargin = LeftPanel.Width;   // 左余白
            doc.PageSetup.RightMargin = RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //1.発行番号を書く(paragraph)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph = doc.Content.Paragraphs.Add();
            paragraph.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph.Range.Text = button1.Text;
            // 段落を右揃えに設定
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph.Range.InsertParagraphAfter();


            //2.発行日を書く(paragraph2)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph2 = doc.Content.Paragraphs.Add();
            paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph2.Range.Text = button2.Text;
            // 段落を右揃えに設定
            paragraph2.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph2.Range.InsertParagraphAfter();

            //相手先
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox1.Text != string.Empty)
            {
                //3.相手先の会社名を書く(paragraph3)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph3 = doc.Content.Paragraphs.Add();
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph3.Range.Text = textBox1.Text;
                // 段落を右揃えに設定
                paragraph3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph3.Range.InsertParagraphAfter();
            }

            if (textBox2.Text != string.Empty)
            {
                //3.相手先の名前を書く(paragraph4)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph4 = doc.Content.Paragraphs.Add();
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph4.Range.Text = textBox2.Text;
                // 段落を右揃えに設定
                paragraph4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph4.Range.InsertParagraphAfter();
            }


            //発信者
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox3.Text != string.Empty)
            {
                //3.発信者の会社名を書く(paragraph5)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph5 = doc.Content.Paragraphs.Add();
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph5.Range.Text = textBox3.Text;
                // 段落を右揃えに設定
                paragraph5.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph5.Range.InsertParagraphAfter();
            }

            if (textBox4.Text != string.Empty)
            {
                //3.発信者の住所を書く(paragraph6)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph6 = doc.Content.Paragraphs.Add();
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph6.Range.Text = textBox4.Text;
                // 段落を右揃えに設定
                paragraph6.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph6.Range.InsertParagraphAfter();
            }

            if (textBox5.Text != string.Empty)
            {
                //3.発信者の建物名称を書く(paragraph7)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph7 = doc.Content.Paragraphs.Add();
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph7.Range.Text = textBox5.Text;
                // 段落を右揃えに設定
                paragraph7.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph7.Range.InsertParagraphAfter();
            }

            if (textBox6.Text != string.Empty)
            {
                //3.発信者の名前を書く(paragraph8)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph8 = doc.Content.Paragraphs.Add();
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph8.Range.Text = textBox6.Text;
                // 段落を右揃えに設定
                paragraph8.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph8.Range.InsertParagraphAfter();
            }

            if (textBox7.Text != string.Empty)
            {
                //3.発信者の電話番号を書く(paragraph9)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph9 = doc.Content.Paragraphs.Add();
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph9.Range.Text = textBox7.Text;
                // 段落を右揃えに設定
                paragraph9.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph9.Range.InsertParagraphAfter();
            }

            if (textBox8.Text != string.Empty)
            {
                //3.発信者のFaxを書く(paragraph10)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph10 = doc.Content.Paragraphs.Add();
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph10.Range.Text = textBox8.Text;
                // 段落を右揃えに設定
                paragraph10.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph10.Range.InsertParagraphAfter();
            }

            if (textBox9.Text != string.Empty)
            {
                //3.発信者のメールアドレスを書く(paragraph11)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph11 = doc.Content.Paragraphs.Add();
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph11.Range.Text = textBox9.Text;
                // 段落を右揃えに設定
                paragraph11.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph11.Range.InsertParagraphAfter();
            }

            //内容
            //内容を入力するテキストボックスが空でない場合にのみ、内容を書く
            if (textBox10.Text != string.Empty)
            {
                //3.表題を書く(paragraph12)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph12 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph12.Range.Text = textBox10.Text;
                // 段落を右揃えに設定
                paragraph12.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                // 段落の変更を適用
                paragraph12.Range.InsertParagraphAfter();
            }


            if (button3.Text != string.Empty)
            {
                //3.内容を書く(paragraph13)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph13 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph13.Range.Text = button3.Text;
                // 段落を右揃えに設定
                paragraph13.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph13.Range.InsertParagraphAfter();
                paragraph13.Range.InsertParagraphAfter();
            }

            //敬具
            Word.Paragraph paragraph14 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph14.Range.Text = label2.Text;
            // 段落を右揃えに設定
            paragraph14.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 段落の変更を適用
            paragraph14.Range.InsertParagraphAfter();

            //記・以上
            Word.Paragraph paragraph15 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph15.Range.Text = label3.Text;
            // 段落を右揃えに設定
            paragraph15.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // 段落の変更を適用
            paragraph15.Range.InsertParagraphAfter();
            paragraph15.Range.InsertParagraphAfter();

            Word.Paragraph paragraph16 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph16.Range.Text = label5.Text;
            // 段落を右揃えに設定
            paragraph16.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 段落の変更を適用
            paragraph16.Range.InsertParagraphAfter();
        }

        private void ribbonOrbMenuItem3_MouseUp(object sender, MouseEventArgs e)
        {
            //Wordインスタンスを生成
            Word.Application word = new Word.Application();
            //Wordウィンドウを表示
            word.Visible = true;
            //新規文書を作成
            Word.Document doc = word.Documents.Add();

            //ページの余白を設定
            doc.PageSetup.TopMargin = TopPanel.Height;    // 上余白 (1インチ = 72ポイント)
            doc.PageSetup.BottomMargin = ButtomPanel.Height; // 下余白
            doc.PageSetup.LeftMargin = LeftPanel.Width;   // 左余白
            doc.PageSetup.RightMargin = RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //1.発行番号を書く(paragraph)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph = doc.Content.Paragraphs.Add();
            paragraph.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph.Range.Text = button1.Text;
            // 段落を右揃えに設定
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph.Range.InsertParagraphAfter();


            //2.発行日を書く(paragraph2)
            // ドキュメントに段落を追加
            Word.Paragraph paragraph2 = doc.Content.Paragraphs.Add();
            paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
            paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
            //発行番号を書く
            paragraph2.Range.Text = button2.Text;
            // 段落を右揃えに設定
            paragraph2.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 改行を追加
            paragraph2.Range.InsertParagraphAfter();

            //相手先
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox1.Text != string.Empty)
            {
                //3.相手先の会社名を書く(paragraph3)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph3 = doc.Content.Paragraphs.Add();
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph3.Range.Text = textBox1.Text;
                // 段落を右揃えに設定
                paragraph3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph3.Range.InsertParagraphAfter();
            }

            if (textBox2.Text != string.Empty)
            {
                //3.相手先の名前を書く(paragraph4)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph4 = doc.Content.Paragraphs.Add();
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                //発行番号を書く
                paragraph4.Range.Text = textBox2.Text;
                // 段落を右揃えに設定
                paragraph4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph4.Range.InsertParagraphAfter();
            }


            //発信者
            //内容を入力するテキストボックスが空でない場合にのみ、相手先の会社名を書く
            if (textBox3.Text != string.Empty)
            {
                //3.発信者の会社名を書く(paragraph5)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph5 = doc.Content.Paragraphs.Add();
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph5.Range.Text = textBox3.Text;
                // 段落を右揃えに設定
                paragraph5.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph5.Range.InsertParagraphAfter();
            }

            if (textBox4.Text != string.Empty)
            {
                //3.発信者の住所を書く(paragraph6)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph6 = doc.Content.Paragraphs.Add();
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph6.Range.Text = textBox4.Text;
                // 段落を右揃えに設定
                paragraph6.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph6.Range.InsertParagraphAfter();
            }

            if (textBox5.Text != string.Empty)
            {
                //3.発信者の建物名称を書く(paragraph7)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph7 = doc.Content.Paragraphs.Add();
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph7.Range.Text = textBox5.Text;
                // 段落を右揃えに設定
                paragraph7.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph7.Range.InsertParagraphAfter();
            }

            if (textBox6.Text != string.Empty)
            {
                //3.発信者の名前を書く(paragraph8)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph8 = doc.Content.Paragraphs.Add();
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph8.Range.Text = textBox6.Text;
                // 段落を右揃えに設定
                paragraph8.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph8.Range.InsertParagraphAfter();
            }

            if (textBox7.Text != string.Empty)
            {
                //3.発信者の電話番号を書く(paragraph9)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph9 = doc.Content.Paragraphs.Add();
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph9.Range.Text = textBox7.Text;
                // 段落を右揃えに設定
                paragraph9.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph9.Range.InsertParagraphAfter();
            }

            if (textBox8.Text != string.Empty)
            {
                //3.発信者のFaxを書く(paragraph10)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph10 = doc.Content.Paragraphs.Add();
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph10.Range.Text = textBox8.Text;
                // 段落を右揃えに設定
                paragraph10.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph10.Range.InsertParagraphAfter();
            }

            if (textBox9.Text != string.Empty)
            {
                //3.発信者のメールアドレスを書く(paragraph11)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph11 = doc.Content.Paragraphs.Add();
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔
                //発行番号を書く
                paragraph11.Range.Text = textBox9.Text;
                // 段落を右揃えに設定
                paragraph11.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                // 段落の変更を適用
                paragraph11.Range.InsertParagraphAfter();
            }

            //内容
            //内容を入力するテキストボックスが空でない場合にのみ、内容を書く
            if (textBox10.Text != string.Empty)
            {
                //3.表題を書く(paragraph12)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph12 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph12.Range.Text = textBox10.Text;
                // 段落を右揃えに設定
                paragraph12.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                // 段落の変更を適用
                paragraph12.Range.InsertParagraphAfter();
            }


            if (button3.Text != string.Empty)
            {
                //3.内容を書く(paragraph13)
                // ドキュメントに段落を追加
                Word.Paragraph paragraph13 = doc.Content.Paragraphs.Add();
                //発行番号を書く
                paragraph13.Range.Text = button3.Text;
                // 段落を右揃えに設定
                paragraph13.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // 段落の変更を適用
                paragraph13.Range.InsertParagraphAfter();
                paragraph13.Range.InsertParagraphAfter();
            }

            //敬具
            Word.Paragraph paragraph14 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph14.Range.Text = label2.Text;
            // 段落を右揃えに設定
            paragraph14.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 段落の変更を適用
            paragraph14.Range.InsertParagraphAfter();

            //記・以上
            Word.Paragraph paragraph15 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph15.Range.Text = label3.Text;
            // 段落を右揃えに設定
            paragraph15.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // 段落の変更を適用
            paragraph15.Range.InsertParagraphAfter();
            paragraph15.Range.InsertParagraphAfter();

            Word.Paragraph paragraph16 = doc.Content.Paragraphs.Add();
            //発行番号を書く
            paragraph16.Range.Text = label5.Text;
            // 段落を右揃えに設定
            paragraph16.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            // 段落の変更を適用
            paragraph16.Range.InsertParagraphAfter();

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Wordドキュメント(*.docx)|*.docx";
            saveFileDialog1.Title = "文書ファイルを保存する場所を選択";

            try
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog1.FileName; // 選択されたファイルパスを取得
                    doc.SaveAs2(filePath, Word.WdSaveFormat.wdFormatDocumentDefault); // ファイルを保存
                    MessageBox.Show("指定された場所にファイルが正しく保存されました: " + filePath, "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    WCloseDialogYesNo dialog = new WCloseDialogYesNo();
                    dialog.ShowDialog();
                    if (dialog.DialogResult == DialogResult.Yes)
                    {
                        // ドキュメントを保存せずに閉じる
                        doc.Close(false);
                        //Wordを終了
                        word.Quit();
                        //ガベージコレクション
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                    else
                    {
                        dialog.Close();
                    }
                }
                //ガベージコレクション
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch(Exception)
            {
                //ガベージコレクション
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            //ガベージコレクション
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void panel26_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            WelcomePanel.Visible = false;

            ribbonButton10.Checked = true;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = false;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "発行番号";
            panel23.Visible = true;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = false;
            panel27.Visible = false;
            panel28.Visible = false;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            WelcomePanel.Visible = false;

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = true;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = false;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "日付";
            panel23.Visible = false;
            panel24.Visible = true;
            panel25.Visible = false;
            panel26.Visible = false;
            panel27.Visible = false;
            panel28.Visible = false;
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            WelcomePanel.Visible = false;

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = true;
            ribbonButton13.Checked = false;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "相手先";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = true;
            panel26.Visible = false;
            panel27.Visible = false;
            panel28.Visible = false;
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            WelcomePanel.Visible = false;

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = true;
            ribbonButton13.Checked = false;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "相手先";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = true;
            panel26.Visible = false;
            panel27.Visible = false;
            panel28.Visible = false;
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            WelcomePanel.Visible = false;

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = true;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "発信者";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = true;
            panel27.Visible = false;
            panel28.Visible = false;
        }

        private void textBox4_Click(object sender, EventArgs e)
        {
            WelcomePanel.Visible = false;

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = true;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "発信者";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = true;
            panel27.Visible = false;
            panel28.Visible = false;
        }

        private void textBox5_Click(object sender, EventArgs e)
        {
            WelcomePanel.Visible = false;

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = true;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "発信者";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = true;
            panel27.Visible = false;
            panel28.Visible = false;
        }

        private void textBox6_Click(object sender, EventArgs e)
        {
            WelcomePanel.Visible = false;

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = true;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "発信者";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = true;
            panel27.Visible = false;
            panel28.Visible = false;
        }

        private void textBox7_Click(object sender, EventArgs e)
        {
            WelcomePanel.Visible = false;

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = true;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "発信者";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = true;
            panel27.Visible = false;
            panel28.Visible = false;
        }

        private void textBox8_Click(object sender, EventArgs e)
        {
            WelcomePanel.Visible = false;

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = true;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "発信者";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = true;
            panel27.Visible = false;
            panel28.Visible = false;
        }

        private void textBox9_Click(object sender, EventArgs e)
        {
            WelcomePanel.Visible = false;

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = true;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "発信者";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = true;
            panel27.Visible = false;
            panel28.Visible = false;
        }
        private void textBox10_Click(object sender, EventArgs e)
        {
            WelcomePanel.Visible = false;

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = false;
            ribbonButton14.Checked = true;
            ribbonButton15.Checked = false;

            label6.Text = "表題";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = false;
            panel27.Visible = true;
            panel28.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            WelcomePanel.Visible = false;

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = false;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = true;

            label6.Text = "内容";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = false;
            panel27.Visible = false;
            panel28.Visible = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

            timer.Interval = 10;

            button5.Visible = true;
            timer.Start();
            timer.Tick += (s, args) =>
            {

                if (panel21.Width <= 290)
                {
                    panel21.Width = panel21.Width - 2;
                    if(panel21.Width <= 270)
                    {
                        panel21.Width = panel21.Width - 4;
                        if (panel21.Width <= 250)
                        {
                            panel21.Width = panel21.Width - 6;
                            if (panel21.Width <= 230)
                            {
                                panel21.Width = panel21.Width - 8;
                                if (panel21.Width <= 210)
                                {
                                    panel21.Width = panel21.Width - 10;
                                }
                            }
                        }
                    }
                }
                if (panel21.Width <= 0)
                {
                    timer.Stop(); // タイマーを停止
                    panel21.Width = 0; // パネルの幅を0に設定
                }
            };
        }

        private void button5_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

            timer.Interval = 1;

            timer.Start();
            timer.Tick += (s, args) =>
            {

                if (panel21.Width >= 0)
                {
                    panel21.Width = panel21.Width  + 10;
                    if (panel21.Width <= 210)
                    {
                        panel21.Width = panel21.Width + 10;
                        if (panel21.Width <= 230)
                        {
                            panel21.Width = panel21.Width + 8;
                            if (panel21.Width <= 250)
                            {
                                panel21.Width = panel21.Width + 6;
                                if (panel21.Width <= 270)
                                {
                                    panel21.Width = panel21.Width + 1;
                                }
                            }
                        }
                    }
                }
                if (panel21.Width == 290)
                {
                    timer.Stop(); // タイマーを停止
                    panel21.Width = 290; // パネルの幅を0に設定
                }
                button5.Visible = false;
            };
        }

        private void button5_MouseEnter(object sender, EventArgs e)
        {
            button5.FlatAppearance.BorderSize = 1;
        }

        private void button5_MouseLeave(object sender, EventArgs e)
        {
            button5.FlatAppearance.BorderSize = 0;
        }

        //リボンのボタンを押したときの処理
        private void ribbonButton10_Click(object sender, EventArgs e)
        {
            if (panel21.Width == 0)
            {
                // ウェルカムパネルを非表示にする
                WelcomePanel.Visible = false;

                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

                timer.Interval = 1;

                timer.Start();
                timer.Tick += (s, args) =>
                {

                    if (panel21.Width >= 0)
                    {
                        panel21.Width = panel21.Width + 10;
                        if (panel21.Width <= 210)
                        {
                            panel21.Width = panel21.Width + 10;
                            if (panel21.Width <= 230)
                            {
                                panel21.Width = panel21.Width + 8;
                                if (panel21.Width <= 250)
                                {
                                    panel21.Width = panel21.Width + 6;
                                    if (panel21.Width <= 270)
                                    {
                                        panel21.Width = panel21.Width + 1;
                                    }
                                }
                            }
                        }
                    }
                    if (panel21.Width == 290)
                    {
                        timer.Stop(); // タイマーを停止
                        panel21.Width = 290; // パネルの幅を0に設定
                    }
                    button5.Visible = false;
                };

            }

            ribbonButton10.Checked = true;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = false;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "発行番号";
            panel23.Visible = true;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = false;
            panel27.Visible = false;
            panel28.Visible = false;
        }

        private void ribbonButton11_Click(object sender, EventArgs e)
        {
            // ウェルカムパネルを非表示にする
            WelcomePanel.Visible = false;

            if (panel21.Width == 0)
            {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

                timer.Interval = 1;

                timer.Start();
                timer.Tick += (s, args) =>
                {

                    if (panel21.Width >= 0)
                    {
                        panel21.Width = panel21.Width + 10;
                        if (panel21.Width <= 210)
                        {
                            panel21.Width = panel21.Width + 10;
                            if (panel21.Width <= 230)
                            {
                                panel21.Width = panel21.Width + 8;
                                if (panel21.Width <= 250)
                                {
                                    panel21.Width = panel21.Width + 6;
                                    if (panel21.Width <= 270)
                                    {
                                        panel21.Width = panel21.Width + 1;
                                    }
                                }
                            }
                        }
                    }
                    if (panel21.Width == 290)
                    {
                        timer.Stop(); // タイマーを停止
                        panel21.Width = 290; // パネルの幅を0に設定
                    }
                    button5.Visible = false;
                };

            }

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = true;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = false;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "日付";
            panel23.Visible = false;
            panel24.Visible = true;
            panel25.Visible = false;
            panel26.Visible = false;
            panel27.Visible = false;
            panel28.Visible = false;
        }

        private void ribbonButton12_Click(object sender, EventArgs e)
        {
            // ウェルカムパネルを非表示にする
            WelcomePanel.Visible = false;

            if (panel21.Width == 0)
            {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

                timer.Interval = 1;

                timer.Start();
                timer.Tick += (s, args) =>
                {

                    if (panel21.Width >= 0)
                    {
                        panel21.Width = panel21.Width + 10;
                        if (panel21.Width <= 210)
                        {
                            panel21.Width = panel21.Width + 10;
                            if (panel21.Width <= 230)
                            {
                                panel21.Width = panel21.Width + 8;
                                if (panel21.Width <= 250)
                                {
                                    panel21.Width = panel21.Width + 6;
                                    if (panel21.Width <= 270)
                                    {
                                        panel21.Width = panel21.Width + 1;
                                    }
                                }
                            }
                        }
                    }
                    if (panel21.Width == 290)
                    {
                        timer.Stop(); // タイマーを停止
                        panel21.Width = 290; // パネルの幅を0に設定
                    }
                    button5.Visible = false;
                };

            }

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = true;
            ribbonButton13.Checked = false;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "相手先";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = true;
            panel26.Visible = false;
            panel27.Visible = false;
            panel28.Visible = false;
        }

        private void ribbonButton13_Click(object sender, EventArgs e)
        {
            // ウェルカムパネルを非表示にする
            WelcomePanel.Visible = false;

            if (panel21.Width == 0)
            {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

                timer.Interval = 1;

                timer.Start();
                timer.Tick += (s, args) =>
                {

                    if (panel21.Width >= 0)
                    {
                        panel21.Width = panel21.Width + 10;
                        if (panel21.Width <= 210)
                        {
                            panel21.Width = panel21.Width + 10;
                            if (panel21.Width <= 230)
                            {
                                panel21.Width = panel21.Width + 8;
                                if (panel21.Width <= 250)
                                {
                                    panel21.Width = panel21.Width + 6;
                                    if (panel21.Width <= 270)
                                    {
                                        panel21.Width = panel21.Width + 1;
                                    }
                                }
                            }
                        }
                    }
                    if (panel21.Width == 290)
                    {
                        timer.Stop(); // タイマーを停止
                        panel21.Width = 290; // パネルの幅を0に設定
                    }
                    button5.Visible = false;
                };

            }

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = true;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = false;

            label6.Text = "発信者";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = true;
            panel27.Visible = false;
            panel28.Visible = false;
        }

        private void ribbonButton14_Click(object sender, EventArgs e)
        {
            // ウェルカムパネルを非表示にする
            WelcomePanel.Visible = false;

            if (panel21.Width == 0)
            {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

                timer.Interval = 1;

                timer.Start();
                timer.Tick += (s, args) =>
                {

                    if (panel21.Width >= 0)
                    {
                        panel21.Width = panel21.Width + 10;
                        if (panel21.Width <= 210)
                        {
                            panel21.Width = panel21.Width + 10;
                            if (panel21.Width <= 230)
                            {
                                panel21.Width = panel21.Width + 8;
                                if (panel21.Width <= 250)
                                {
                                    panel21.Width = panel21.Width + 6;
                                    if (panel21.Width <= 270)
                                    {
                                        panel21.Width = panel21.Width + 1;
                                    }
                                }
                            }
                        }
                    }
                    if (panel21.Width == 290)
                    {
                        timer.Stop(); // タイマーを停止
                        panel21.Width = 290; // パネルの幅を0に設定
                    }
                    button5.Visible = false;
                };

            }

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = false;
            ribbonButton14.Checked = true;
            ribbonButton15.Checked = false;

            label6.Text = "表題";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = false;
            panel27.Visible = true;
            panel28.Visible = false;
        }

        private void ribbonButton15_Click(object sender, EventArgs e)
        {
            // ウェルカムパネルを非表示にする
            WelcomePanel.Visible = false;

            if (panel21.Width == 0)
            {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

                timer.Interval = 1;

                timer.Start();
                timer.Tick += (s, args) =>
                {

                    if (panel21.Width >= 0)
                    {
                        panel21.Width = panel21.Width + 10;
                        if (panel21.Width <= 210)
                        {
                            panel21.Width = panel21.Width + 10;
                            if (panel21.Width <= 230)
                            {
                                panel21.Width = panel21.Width + 8;
                                if (panel21.Width <= 250)
                                {
                                    panel21.Width = panel21.Width + 6;
                                    if (panel21.Width <= 270)
                                    {
                                        panel21.Width = panel21.Width + 1;
                                    }
                                }
                            }
                        }
                    }
                    if (panel21.Width == 290)
                    {
                        timer.Stop(); // タイマーを停止
                        panel21.Width = 290; // パネルの幅を0に設定
                    }
                    button5.Visible = false;
                };

            }

            ribbonButton10.Checked = false;
            ribbonButton11.Checked = false;
            ribbonButton12.Checked = false;
            ribbonButton13.Checked = false;
            ribbonButton14.Checked = false;
            ribbonButton15.Checked = true;

            label6.Text = "内容";
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = false;
            panel26.Visible = false;
            panel27.Visible = false;
            panel28.Visible = true;
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            if (textBox11.Text == string.Empty)
            {
                string str1 = numericUpDown1.Value.ToString();
                button1.Text = "第"+str1+"号";
            }
            else
            {
                label7.ForeColor = Color.FromArgb(21, 66, 139);
                string str2 = textBox11.Text + "発第" + numericUpDown1.Value.ToString()+"号";
                button1.Text = str2;
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (textBox11.Text == string.Empty)
            {
                string str1 = numericUpDown1.Value.ToString();
                button1.Text = "第"+str1+"号";
            }
            else
            {
                string str2 = textBox11.Text + "発第" + numericUpDown1.Value.ToString() + "号";
                button1.Text = str2;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                panel2.Hide();
                DayPanel.Height = 87 - 28;
                button1.Text = string.Empty;
                textBox11.Enabled = false;
                label7.Enabled = false;
                numericUpDown1.Enabled = false;
                label8.Enabled = false;

                if (panel4.Visible == true)
                {
                    DayPanel.Height = 87 - 28;
                }
                else
                {
                    DayPanel.Height = 0;
                }
            }
            else
            {
                panel2.Show();
                DayPanel.Height = 57 + 28;
                textBox11.Enabled = true;
                label7.Enabled = true;
                numericUpDown1.Enabled = true;
                label8.Enabled = true;

                if (textBox11.Text == string.Empty)
                {
                    string str1 = numericUpDown1.Value.ToString();
                    button1.Text = "第"+str1+"号";
                }
                else
                {
                    string str2 = textBox11.Text + "発第" + numericUpDown1.Value.ToString() + "号";
                    button1.Text = str2;
                }

                if (panel4.Visible == true)
                {
                    DayPanel.Height = 57 + 28;
                }
                else
                {
                    DayPanel.Height = 57;
                }
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                var cultureJp = new CultureInfo("ja-jp", false);
                cultureJp.DateTimeFormat.Calendar = new JapaneseCalendar();
                string str3 = dateTimePicker1.Value.ToString("ggy年M月d日", cultureJp);
                button2.Text = str3;
            }
            else
            {
                string str3 = dateTimePicker1.Value.ToString("yyyy年M月d日");
                button2.Text = str3;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox3_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                var cultureJp = new CultureInfo("ja-jp", false);
                cultureJp.DateTimeFormat.Calendar = new JapaneseCalendar();
                string str3 = dateTimePicker1.Value.ToString("ggy年M月d日", cultureJp);
                button2.Text = str3;
            }
            else
            {
                string str3 = dateTimePicker1.Value.ToString("yyyy年M月d日");
                button2.Text = str3;
            }
        }

        private void checkBox2_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                dateTimePicker1.Enabled = false;
                checkBox3.Enabled = false;

                button2.Text = string.Empty;
                panel4.Visible = false;
                if (panel2.Visible == true)
                {
                    DayPanel.Height = 87 - 28;
                }
                else
                {
                    DayPanel.Height = 0;
                }
            }
            else
            {
                dateTimePicker1.Enabled = true;
                checkBox3.Enabled = true;

                if (checkBox3.Checked == true)
                {
                    var cultureJp = new CultureInfo("ja-jp", false);
                    cultureJp.DateTimeFormat.Calendar = new JapaneseCalendar();
                    string str3 = dateTimePicker1.Value.ToString("ggy年M月d日", cultureJp);
                    button2.Text = str3;
                }
                else
                {
                    string str3 = dateTimePicker1.Value.ToString("yyyy年M月d日");
                    button2.Text = str3;
                }

                panel4.Visible = true;
                if (panel2.Visible == true)
                {
                    DayPanel.Height = 57 + 28;
                }
                else
                {
                    DayPanel.Height = 57;
                }
            }
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = textBox12.Text;
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            textBox2.Text = textBox13.Text;
        }

        private void ribbonButton49_Click(object sender, EventArgs e)
        {
            ribbonButton17.Checked = false;
            ribbonButton49.Checked = true;
            textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            textBox3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            textBox4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            textBox5.BorderStyle = System.Windows.Forms.BorderStyle.None;
            textBox6.BorderStyle = System.Windows.Forms.BorderStyle.None;
            textBox7.BorderStyle = System.Windows.Forms.BorderStyle.None;
            textBox8.BorderStyle = System.Windows.Forms.BorderStyle.None;
            textBox9.BorderStyle = System.Windows.Forms.BorderStyle.None;
            textBox10.BorderStyle = System.Windows.Forms.BorderStyle.None;

        }

        private void ribbonButton17_Click(object sender, EventArgs e)
        {
            ribbonButton17.Checked = true;
            ribbonButton49.Checked = false;
            textBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            textBox2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            textBox3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            textBox4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            textBox5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            textBox6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            textBox7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            textBox8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            textBox9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            textBox10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
        }

        private void ribbonButton50_Click(object sender, EventArgs e)
        {
            if(ribbonButton50.Checked == true)
            {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

                timer.Interval = 10;

                button5.Visible = false;
                timer.Start();
                timer.Tick += (s, args) =>
                {

                    if (panel21.Width <= 290)
                    {
                        panel21.Width = panel21.Width - 2;
                        if (panel21.Width <= 270)
                        {
                            panel21.Width = panel21.Width - 4;
                            if (panel21.Width <= 250)
                            {
                                panel21.Width = panel21.Width - 6;
                                if (panel21.Width <= 230)
                                {
                                    panel21.Width = panel21.Width - 8;
                                    if (panel21.Width <= 210)
                                    {
                                        panel21.Width = panel21.Width - 10;
                                    }
                                }
                            }
                        }
                    }
                    if (panel21.Width <= 0)
                    {
                        timer.Stop(); // タイマーを停止
                        panel21.Width = 0; // パネルの幅を0に設定
                    }
                };

                ribbonPanel1.Enabled = false;
                ribbonPanel2.Enabled = false;
                ribbonPanel3.Enabled = false;
                ribbonPanel7.Enabled = false;
                ribbonPanel11.Enabled = false;

                //無効化
                ribbonButton17.Checked = false;
                ribbonButton17.Enabled = false;

                ribbonButton49.Checked = true;

                ribbonButton17.Checked = false;
                ribbonButton49.Checked = true;
                textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox3.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox4.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox5.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox6.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox7.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox8.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox9.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox10.BorderStyle = System.Windows.Forms.BorderStyle.None;

                //Sheet1内のテキストボックスの入力をすべて無効化
                textBox1.ReadOnly = true;
                textBox2.ReadOnly = true;
                textBox3.ReadOnly = true;
                textBox4.ReadOnly = true;
                textBox5.ReadOnly = true;
                textBox6.ReadOnly = true;
                textBox7.ReadOnly = true;
                textBox8.ReadOnly = true;
                textBox9.ReadOnly = true;
                textBox10.ReadOnly = true;

                //Orbitボタンを無効化
                ribbonOrbMenuItem1.Enabled = false;

                //クイックアクセスツールバーのボタンを無効化
                ribbonButton22.Enabled = false;
            }
            else
            {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

                timer.Interval = 1;

                timer.Start();
                timer.Tick += (s, args) =>
                {

                    if (panel21.Width >= 0)
                    {
                        panel21.Width = panel21.Width + 10;
                        if (panel21.Width <= 210)
                        {
                            panel21.Width = panel21.Width + 10;
                            if (panel21.Width <= 230)
                            {
                                panel21.Width = panel21.Width + 8;
                                if (panel21.Width <= 250)
                                {
                                    panel21.Width = panel21.Width + 6;
                                    if (panel21.Width <= 270)
                                    {
                                        panel21.Width = panel21.Width + 1;
                                    }
                                }
                            }
                        }
                    }
                    if (panel21.Width <= 290)
                    {
                        timer.Stop(); // タイマーを停止
                        panel21.Width = 290; // パネルの幅を0に設定
                    }
                    button5.Visible = false;
                };

                ribbonPanel1.Enabled = true;
                ribbonPanel2.Enabled = true;
                ribbonPanel3.Enabled = true;
                ribbonPanel7.Enabled = true;
                ribbonPanel11.Enabled = true;

                //有効化
                ribbonButton17.Checked = true;
                ribbonButton17.Enabled = true;

                ribbonButton49.Checked = false;

                ribbonButton17.Checked = true;
                ribbonButton49.Checked = false;
                textBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
                textBox2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
                textBox3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
                textBox4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
                textBox5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
                textBox6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
                textBox7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
                textBox8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
                textBox9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
                textBox10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;

                //Sheet1内のテキストボックスの入力をすべて有効化
                textBox1.ReadOnly = false;
                textBox2.ReadOnly = false;
                textBox3.ReadOnly = false;
                textBox4.ReadOnly = false;
                textBox5.ReadOnly = false;
                textBox6.ReadOnly = false;
                textBox7.ReadOnly = false;
                textBox8.ReadOnly = false;
                textBox9.ReadOnly = false;
                textBox10.ReadOnly = false;

                //Orbitボタンを有効化
                ribbonOrbMenuItem1.Enabled = true;

                //クイックアクセスツールバーのボタンを有効化
                ribbonButton22.Enabled = true;
            }
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            textBox3.Text = textBox14.Text;
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            textBox4.Text = textBox15.Text;
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            textBox16.Text = textBox5.Text;
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            textBox5.Text = textBox16.Text;
        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {
            textBox6.Text = textBox17.Text;
        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {
            textBox7.Text = textBox18.Text;
        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {
            textBox8.Text = textBox19.Text;
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            textBox9.Text = textBox21.Text;
        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            textBox10.Text = textBox20.Text;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox12.Text = textBox1.Text;
        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            textBox13.Text = textBox2.Text;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox14.Text = textBox3.Text;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            textBox15.Text = textBox4.Text;
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            textBox17.Text = textBox6.Text;
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            textBox18.Text = textBox7.Text;
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            textBox19.Text = textBox8.Text;
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            textBox21.Text = textBox9.Text;
        }

        private void ribbonButton65_Click(object sender, EventArgs e)
        {
            DialogYsNo dialogYesNO = new DialogYsNo();
            dialogYesNO.ShowDialog(this);
            if (dialogYesNO.DialogResult == DialogResult.Yes)
            {
                button1.Text = "○○発第○○号"; 

                string str3 = dateTimePicker1.Value.ToString("yyyy年M月d日");
                button2.Text = str3;
                checkBox1.Checked = false;

                textBox1.Text = "○○株式会社";
                textBox2.Text = "○○部　○○○○様";

                textBox3.Text = "○○○○株式会社";
                textBox4.Text = "東京都渋谷区○○町〇丁目〇番地〇号";
                textBox5.Text = "○○ビル 40階";
                textBox6.Text = "代表取締役　○○○○";
                textBox7.Text = "電話:0000-0000-0000";
                textBox8.Text = "FAX:0000-0000-0000";
                textBox9.Text = "メールアドレス:my@example.com";

                textBox10.Text = "新商品発表会のご案内";

                button3.Text = "拝啓、時下ますますご清栄のこととお慶び申し上げます。平素は並々ならぬお引き立てを賜り、厚く御礼申し上げます。";

                label2.Text = "敬具";



                textBox11.Text = "○○";
                numericUpDown1.Value = 0;
                dateTimePicker1.Value = DateTime.Now;

                textBox12.Text = "○○株式会社";
                textBox13.Text = "○○部　○○○○様";

                textBox14.Text = "○○○○株式会社";
                textBox15.Text = "東京都渋谷区○○町〇丁目〇番地〇号";
                textBox16.Text = "○○ビル 40階";
                textBox17.Text = "代表取締役　○○○○";
                textBox18.Text = "電話:0000-0000-0000";
                textBox19.Text = "FAX:0000-0000-0000";
                textBox21.Text = "メールアドレス:my@example.com";

                textBox20.Text = "新商品発表会のご案内";


            }
            else if (dialogYesNO.DialogResult == DialogResult.No)
            {
                //何もしない
            }

        }

        private void ribbonButton52_Click(object sender, EventArgs e)
        {
            
            textBox12.Text = textBox1.Text;
            textBox13.Text = textBox2.Text;

            textBox14.Text = textBox3.Text;
            textBox15.Text = textBox4.Text;
            textBox16.Text = textBox5.Text;
            textBox17.Text = textBox6.Text;
            textBox18.Text = textBox7.Text;
            textBox19.Text = textBox8.Text;
            textBox21.Text = textBox9.Text;

            textBox20.Text = textBox10.Text;
        }

        private void ribbonOrbMenuItem4_Click(object sender, EventArgs e)
        {
            Settings settings = new Settings();
            settings.ShowDialog();
        }

        private void ribbonOrbMenuItem4_MouseUp(object sender, MouseEventArgs e)
        {
            Settings settings = new Settings();
            settings.ShowDialog();
        }

        private void ribbonButton77_Click(object sender, EventArgs e)
        {
            this.Hide();

            this.ControlBox = false; // コントロールボックスを非表示にする
            this.FormBorderStyle = FormBorderStyle.Sizable;
            // Ribbonのカスタム描画モードを設定
            ribbon1.BorderMode = RibbonWindowMode.NonClientAreaCustomDrawn;

            this.Show();
            ribbonButton77.Checked = true;
            ribbonButton30.Checked = false;


        }

        private void ribbon1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_MouseEnter(object sender, EventArgs e)
        {

        }

        private void ribbon1_MouseEnter(object sender, EventArgs e)
        {

        }

        private void ribbon1_MouseLeave(object sender, EventArgs e)
        {

        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
        }

        private void Form1_ResizeEnd(object sender, EventArgs e)
        {

        }

        private void ribbon1_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void ribbon1_MouseUp(object sender, MouseEventArgs e)
        {

        }

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void Form1_MouseMove(object sender, MouseEventArgs e)
        {

        }

        private void ribbonColorChooser2_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser2.Color;
            textBox10.ForeColor = ribbonColorChooser2.Color;
        }

        private void ribbonColorChooser3_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser3.Color;
            textBox10.ForeColor = ribbonColorChooser3.Color;
        }

        private void ribbonColorChooser4_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser4.Color;
            textBox10.ForeColor = ribbonColorChooser4.Color;
        }

        private void ribbonColorChooser5_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser5.Color;
            textBox10.ForeColor = ribbonColorChooser5.Color;
        }

        private void ribbonColorChooser6_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser6.Color;
            textBox10.ForeColor = ribbonColorChooser6.Color;
        }

        private void ribbonColorChooser7_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser7.Color;
            textBox10.ForeColor = ribbonColorChooser7.Color;
        }

        private void ribbonColorChooser8_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser8.Color;
            textBox10.ForeColor = ribbonColorChooser8.Color;
        }

        private void ribbonColorChooser9_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser9.Color;
            textBox10.ForeColor = ribbonColorChooser9.Color;
        }

        private void ribbonColorChooser10_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser10.Color;
            textBox10.ForeColor = ribbonColorChooser10.Color;
        }

        private void ribbonColorChooser11_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser11.Color;
            textBox10.ForeColor = ribbonColorChooser11.Color;
        }

        private void ribbonColorChooser12_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser12.Color;
            textBox10.ForeColor = ribbonColorChooser12.Color;
        }

        private void ribbonColorChooser13_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser13.Color;
            textBox10.ForeColor = ribbonColorChooser13.Color;
        }

        private void ribbonColorChooser14_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser14.Color;
            textBox10.ForeColor = ribbonColorChooser14.Color;
        }

        private void ribbonColorChooser15_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser15.Color;
            textBox10.ForeColor = ribbonColorChooser15.Color;
        }

        private void ribbonColorChooser16_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser16.Color;
            textBox10.ForeColor = ribbonColorChooser16.Color;
        }

        private void ribbonColorChooser17_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser17.Color;
            textBox10.ForeColor = ribbonColorChooser17.Color;
        }

        private void ribbonColorChooser18_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser18.Color;
            textBox10.ForeColor = ribbonColorChooser18.Color;
        }

        private void ribbonColorChooser19_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser19.Color;
            textBox10.ForeColor = ribbonColorChooser19.Color;
        }

        private void ribbonColorChooser20_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser20.Color;
            textBox10.ForeColor = ribbonColorChooser20.Color;
        }

        private void ribbonColorChooser21_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser21.Color;
            textBox10.ForeColor = ribbonColorChooser21.Color;
        }

        private void ribbonColorChooser22_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser22.Color;
            textBox10.ForeColor = ribbonColorChooser22.Color;
        }

        private void ribbonColorChooser23_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser23.Color;
            textBox10.ForeColor = ribbonColorChooser23.Color;
        }

        private void ribbonColorChooser24_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser24.Color;
            textBox10.ForeColor = ribbonColorChooser24.Color;
        }

        private void ribbonColorChooser25_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser25.Color;
            textBox10.ForeColor = ribbonColorChooser25.Color;
        }

        private void ribbonColorChooser26_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser26.Color;
            textBox10.ForeColor = ribbonColorChooser26.Color;
        }

        private void ribbonColorChooser27_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser27.Color;
            textBox10.ForeColor = ribbonColorChooser27.Color;
        }

        private void ribbonColorChooser28_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser28.Color;
            textBox10.ForeColor = ribbonColorChooser28.Color;
        }

        private void ribbonColorChooser29_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser29.Color;
            textBox10.ForeColor = ribbonColorChooser29.Color;
        }

        private void ribbonColorChooser30_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser30.Color;
            textBox10.ForeColor = ribbonColorChooser30.Color;
        }

        private void ribbonColorChooser31_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser31.Color;
            textBox10.ForeColor = ribbonColorChooser31.Color;
        }

        private void ribbonColorChooser32_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser32.Color;
            textBox10.ForeColor = ribbonColorChooser32.Color;
        }

        private void ribbonColorChooser33_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser33.Color;
            textBox10.ForeColor = ribbonColorChooser33.Color;
        }

        private void ribbonColorChooser34_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser34.Color;
            textBox10.ForeColor = ribbonColorChooser34.Color;
        }

        private void ribbonColorChooser35_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser35.Color;
            textBox10.ForeColor = ribbonColorChooser35.Color;
        }

        private void ribbonColorChooser36_Click(object sender, EventArgs e)
        {
            ribbonColorChooser1.Color = ribbonColorChooser36.Color;
            textBox10.ForeColor = ribbonColorChooser36.Color;
        }

        private void ribbonButton74_Click(object sender, EventArgs e)
        {
            KryptonColorDialog cd = new KryptonColorDialog();
            cd.Color = ribbonColorChooser1.Color;
            cd.AllowFullOpen = true;
            if (cd.ShowDialog() ==DialogResult.OK)
            {
               ribbonColorChooser1.Color = cd.Color;
                textBox10.ForeColor = cd.Color;
            }
        }

        private void ribbonColorChooser1_Click(object sender, EventArgs e)
        {
            textBox10.ForeColor = ribbonColorChooser1.Color;
        }

        //太字ボタンを押したときの処理
        private void ribbonButton5_Click(object sender, EventArgs e)
        {
            if(ribbonButton5.Checked == true)
            {
                textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Bold);
            }
            else if(ribbonButton5.Checked == false)
            {
                textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Regular);
                //下線が選択されている場合
                if (ribbonButton8.Checked == true)
                {
                    textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Underline);
                }
                //斜体が選択されている場合
                if (ribbonButton7.Checked == true)
                {
                    textBox10.Font = new System.Drawing.Font(textBox10.Font,FontStyle.Italic);
                    //斜体が選択されていて、下線が選択されている場合
                    if (ribbonButton8.Checked == true)
                    {
                        textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Underline);
                    }
                }
                //斜体が選択されていない場合
                else if (ribbonButton7.Checked == true)
                {
                    textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Regular);
                    //斜体が選択されておらず、下線が選択されている場合
                    if (ribbonButton8.Checked == true)
                    {
                        textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Underline);
                    }
                }
            }
        }


        //斜体ボタンを押したときの処理
        private void ribbonButton7_Click(object sender, EventArgs e)
        {
            if (ribbonButton7.Checked == true)
            {
                textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Italic);
            }
            else if (ribbonButton7.Checked == false)
            {
                textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Regular);
                //太字が選択されている場合
                if (ribbonButton5.Checked == true)
                {
                    textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Bold);
                    //太字が選択されていて、下線が選択されている場合
                    if (ribbonButton8.Checked == true)
                    {
                        textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Underline);
                    }
                }
                //太字が選択されていない場合
                else if (ribbonButton5.Checked == false)
                {
                    textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Regular);
                    //太字が選択されておらず、下線が選択されている場合
                    if (ribbonButton8.Checked == true)
                    {
                        textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Underline);
                    }
                }
            }
        }

        //下線(取り消し線)ボタンを押したときの処理
        private void ribbonButton8_Click(object sender, EventArgs e)
        {
            if (ribbonButton8.Checked == true)
            {
                textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Underline);
            }
            else if (ribbonButton8.Checked == false)
            {
                ribbonButton75.Checked = false;
                textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Regular);
                //太字が選択されている場合
                if (ribbonButton5.Checked == true)
                {
                    textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Bold);
                    //太字が選択されていて、斜体が選択されている場合
                    if (ribbonButton7.Checked == true)
                    {
                        textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Italic);
                    }
                }
                //太字が選択されていない場合
                else if (ribbonButton5.Checked == false)
                {
                    textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Regular);
                    //太字が選択されておらず、斜体が選択されている場合
                    if (ribbonButton7.Checked == true)
                    {
                        textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Italic);
                    }
                }
            }
        }

        private void ribbonButton75_Click(object sender, EventArgs e)
        {
            ribbonButton8.Checked = true;
            if (ribbonButton8.Checked == true)
            {
                textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Underline);
            }
            else if (ribbonButton8.Checked == false)
            {
                ribbonButton75.Checked = false;
                textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Regular);
                //太字が選択されている場合
                if (ribbonButton5.Checked == true)
                {
                    textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Bold);
                    //太字が選択されていて、斜体が選択されている場合
                    if (ribbonButton7.Checked == true)
                    {
                        textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Italic);

                    }
                }
                //太字が選択されていない場合
                else if (ribbonButton5.Checked == false)
                {
                    textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Regular);
                    //太字が選択されておらず、斜体が選択されている場合
                    if (ribbonButton7.Checked == true)
                    {
                        textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Italic);
                    }
                }
            }
        }

        private void ribbonButton76_Click(object sender, EventArgs e)
        {
            ribbonButton8.Checked = true;
            if (ribbonButton8.Checked == true)
            {

                textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Strikeout);
            }
            else if (ribbonButton8.Checked == false)
            {
                ribbonButton75.Checked = false;
                textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Regular);
                //太字が選択されている場合
                if (ribbonButton5.Checked == true)
                {
                    textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Bold);
                    //太字が選択されていて、斜体が選択されている場合
                    if (ribbonButton7.Checked == true)
                    {
                        textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Italic);
                    }
                }
                //太字が選択されていない場合
                else if (ribbonButton5.Checked == false)
                {
                    textBox10.Font = new System.Drawing.Font(textBox10.Font, FontStyle.Regular);
                    //太字が選択されておらず、斜体が選択されている場合
                    if (ribbonButton7.Checked == true)
                    {
                        textBox10.Font = new System.Drawing.Font(textBox10.Font, textBox10.Font.Style | FontStyle.Italic);
                    }
                }
            }
        }

        private void ribbonButton53_Click(object sender, EventArgs e)
        {
            float fontSize = 5f;
            textBox10.Font = new System.Drawing.Font(textBox10.Font.OriginalFontName, fontSize);

        }

        private void ribbonButton54_Click(object sender, EventArgs e)
        {
            float fontSize = 10f;
            textBox10.Font = new System.Drawing.Font(textBox10.Font.OriginalFontName, fontSize);
        }

        private void ribbonButton55_Click(object sender, EventArgs e)
        {
            float fontSize = 15f;
            textBox10.Font = new System.Drawing.Font(textBox10.Font.OriginalFontName, fontSize);
        }

        private void ribbonButton56_Click(object sender, EventArgs e)
        {
            float fontSize = 20f;
            textBox10.Font = new System.Drawing.Font(textBox10.Font.OriginalFontName, fontSize);
        }

        private void ribbonButton57_Click(object sender, EventArgs e)
        {
            float fontSize = 25f;
            textBox10.Font = new System.Drawing.Font(textBox10.Font.OriginalFontName, fontSize);
        }

        private void ribbonButton58_Click(object sender, EventArgs e)
        {
            float fontSize = 30f;
            textBox10.Font = new System.Drawing.Font(textBox10.Font.OriginalFontName, fontSize);
        }

        private void ribbonButton59_Click(object sender, EventArgs e)
        {
            float fontSize = 35f;
            textBox10.Font = new System.Drawing.Font(textBox10.Font.OriginalFontName, fontSize);
        }

        private void ribbonButton60_Click(object sender, EventArgs e)
        {
            float fontSize = 40f;
            textBox10.Font = new System.Drawing.Font(textBox10.Font.OriginalFontName, fontSize);
        }

        private void ribbonButton61_Click(object sender, EventArgs e)
        {
            float fontSize = 45f;
            textBox10.Font = new System.Drawing.Font(textBox10.Font.OriginalFontName, fontSize);
        }

        private void ribbonButton62_Click(object sender, EventArgs e)
        {
            float fontSize = 50f;
            textBox10.Font = new System.Drawing.Font(textBox10.Font.OriginalFontName, fontSize);
        }

        private void ribbonComboBox2_TextBoxTextChanged(object sender, EventArgs e)
        {

        }

        private void ribbonButton66_Click(object sender, EventArgs e)
        {
            float fontSize =   12f;
            textBox10.Font = new System.Drawing.Font("游明朝", fontSize);
        }

        private void ribbonButton67_Click(object sender, EventArgs e)
        {
            float fontSize = 12f;
            textBox10.Font = new System.Drawing.Font("BIZ UDP明朝 Medium", fontSize);
        }

        private void ribbonButton68_Click(object sender, EventArgs e)
        {
            float fontSize = 12f;
            textBox10.Font = new System.Drawing.Font("HGP教科書体", fontSize);
        }

        private void ribbonButton69_Click(object sender, EventArgs e)
        {
            float fontSize = 12f;
            textBox10.Font = new System.Drawing.Font("HGS教科書体", fontSize);
        }

        private void ribbonButton70_Click(object sender, EventArgs e)
        {
            float fontSize = 12f;
            textBox10.Font = new System.Drawing.Font("ＭＳ 明朝", fontSize);
        }

        private void ribbonComboBox1_TextBoxTextChanged(object sender, EventArgs e)
        {
        }

        ClipBoradWindow clipBoradWindow = new ClipBoradWindow();
        private void ribbonButton73_Click(object sender, EventArgs e)
        {
            clipBoradWindow.Show();
            ribbonButton73.Enabled = false; // クリップボードウィンドウを開いたら無効化

        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            if (clipBoradWindow.Visible)
            {
                ribbonButton73.Enabled = false;
            }
            else
            {
                ribbonButton73.Enabled = true; // アプリケーションがアクティブでない場合は有効化
            }
        }

        private void ribbonButton1_Click(object sender, EventArgs e)
        {
            Clipboard.GetText();
            if(textBox1.Focused == true)
            {
                textBox1.Paste();
            }
            if (textBox2.Focused == true)
            {
                textBox2.Paste();
            }
            if (textBox3.Focused == true)
            {
                textBox3.Paste();
            }
            if (textBox4.Focused == true)
            {
                textBox4.Paste();
            }
            if (textBox5.Focused == true)
            {
                textBox5.Paste();
            }
            if (textBox6.Focused == true)
            {
                textBox6.Paste();
            }
            if (textBox7.Focused == true)
            {
                textBox7.Paste();
            }
            if (textBox8.Focused == true)
            {
                textBox8.Paste();
            }
            if (textBox9.Focused == true)
            {
                textBox9.Paste();
            }
            if (textBox10.Focused == true)
            {
                textBox10.Paste();
            }
            if(textBox11.Focused == true)
            {
                textBox11.Paste();
            }
            if (textBox12.Focused == true)
            {
                textBox12.Paste();
            }
            if (textBox13.Focused == true)
            {
                textBox13.Paste();
            }
            if (textBox14.Focused == true)
            {
                textBox14.Paste();
            }
            if (textBox15.Focused == true)
            {
                textBox15.Paste();
            }
            if (textBox16.Focused == true)
            {
                textBox16.Paste();
            }
            if (textBox17.Focused == true)
            {
                textBox17.Paste();
            }
            if (textBox18.Focused == true)
            {
                textBox18.Paste();
            }
            if (textBox19.Focused == true)
            {
                textBox19.Paste();
            }
            if (textBox20.Focused == true)
            {
                textBox20.Paste();
            }
            if (textBox21.Focused == true)
            {
                textBox21.Paste();
            }
        }

        private void ribbonButton2_Click(object sender, EventArgs e)
        {
            if(textBox1.Focused == true)
            {
                textBox1.Cut();
            }
            if (textBox2.Focused == true)
            {
                textBox2.Cut();
            }
            if (textBox3.Focused == true)
            {
                textBox3.Cut();
            }
            if (textBox4.Focused == true)
            {
                textBox4.Cut();
            }
            if (textBox5.Focused == true)
            {
                textBox5.Cut();
            }
            if (textBox6.Focused == true)
            {
                textBox6.Cut();
            }
            if (textBox7.Focused == true)
            {
                textBox7.Cut();
            }
            if (textBox8.Focused == true)
            {
                textBox8.Cut();
            }
            if (textBox9.Focused == true)
            {
                textBox9.Cut();
            }
            if (textBox10.Focused == true)
            {
                textBox10.Cut();
            }
            if (textBox11.Focused == true)
            {
                textBox11.Cut();
            }
            if (textBox12.Focused == true)
            {
                textBox12.Cut();
            }
            if (textBox13.Focused == true)
            {
                textBox13.Cut();
            }
            if (textBox14.Focused == true)
            {
                textBox14.Cut();
            }
            if (textBox15.Focused == true)
            {
                textBox15.Cut();
            }
            if (textBox16.Focused == true)
            {
                textBox16.Cut();
            }
            if (textBox17.Focused == true)
            {
                textBox17.Cut();
            }
            if (textBox18.Focused == true)
            {
                textBox18.Cut();
            }
            if (textBox19.Focused == true)
            {
                textBox19.Cut();
            }
            if (textBox20.Focused == true)
            {
                textBox20.Cut();
            }
            if (textBox21.Focused == true)
            {
                textBox21.Cut();
            }
        }

        private void ribbonButton3_Click(object sender, EventArgs e)
        {
            if (textBox1.Focused == true)
            {
                textBox1.Copy();
            }
            if (textBox2.Focused == true)
            {
                textBox2.Copy();
            }
            if (textBox3.Focused == true)
            {
                textBox3.Copy();
            }
            if (textBox4.Focused == true)
            {
                textBox4.Copy();
            }
            if (textBox5.Focused == true)
            {
                textBox5.Copy();
            }
            if (textBox6.Focused == true)
            {
                textBox6.Copy();
            }
            if (textBox7.Focused == true)
            {
                textBox7.Copy();
            }
            if (textBox8.Focused == true)
            {
                textBox8.Copy();
            }
            if (textBox9.Focused == true)
            {
                textBox9.Copy();
            }
            if (textBox10.Focused == true)
            {
                textBox10.Copy();
            }
            if (textBox11.Focused == true)
            {
                textBox11.Copy();
            }
            if (textBox12.Focused == true)
            {
                textBox12.Copy();
            }
            if (textBox13.Focused == true)
            {
                textBox13.Copy();
            }
            if (textBox14.Focused == true)
            {
                textBox14.Copy();
            }
            if (textBox15.Focused == true)
            {
                textBox15.Copy();
            }
            if (textBox16.Focused == true)
            {
                textBox16.Copy();
            }
            if (textBox17.Focused == true)
            {
                textBox17.Copy();
            }
            if (textBox18.Focused == true)
            {
                textBox18.Copy();
            }
            if (textBox19.Focused == true)
            {
                textBox19.Copy();
            }
            if (textBox20.Focused == true)
            {
                textBox20.Copy();
            }
            if (textBox21.Focused == true)
            {
                textBox21.Copy();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //描画先とするImageオブジェクトを作成する
            Bitmap canvas = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            //ImageオブジェクトのGraphicsオブジェクトを作成する
            Graphics g = Graphics.FromImage(canvas);

            //縦に白から黒へのグラデーションのブラシを作成
            //g.VisibleClipBoundsは表示クリッピング領域に外接する四角形
            LinearGradientBrush gb = new LinearGradientBrush(
                    g.VisibleClipBounds,
                    SystemColors.GradientInactiveCaption,
                    SystemColors.GradientActiveCaption,
                    LinearGradientMode.Vertical);

            //四角を描く
            g.FillRectangle(gb, g.VisibleClipBounds);

            //リソースを解放する
            gb.Dispose();
            g.Dispose();

            //PictureBox1に表示する
            pictureBox1.Image = canvas;

        }

        private void kryptonContextMenu1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void button1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                kryptonContextMenu1.Show(Cursor);
            }
        }

        private void button2_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                kryptonContextMenu2.Show(Cursor);
            }
        }

        private void button3_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                kryptonContextMenu3.Show(Cursor);
            }
        }
    }
}
