using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using static System.Collections.Specialized.BitVector32;
using Word = Microsoft.Office.Interop.Word;

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
        }

        private void ribbonButton29_Click(object sender, EventArgs e)
        {
            ribbon1.OrbStyle = RibbonOrbStyle.Office_2010;
            ribbonButton28.Checked = false;
            ribbonButton29.Checked = true;
        }

        private void ribbonButton30_Click(object sender, EventArgs e)
        {

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
            this.Size = new Size(935, 576); // Set the initial size of the form
            Sheet1.Location = new System.Drawing.Point(9, 56); // Set the initial location of the sheet
            EditPanel.AutoScroll = true; // Enable auto-scrolling for the EditPanel
            this.Text = textBox10.Text + " - Document Maker";
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
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog1.FileName; // 選択されたファイルパスを取得
                doc.SaveAs2(filePath, Word.WdSaveFormat.wdFormatDocumentDefault); // ファイルを保存
                MessageBox.Show("指定された場所にファイルが正しく保存されました: " + filePath, "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                DialogResult result = MessageBox.Show("文書作成ソフトが起動しています。保存せずに文書作成ソフトを閉じますか?", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    // ドキュメントを保存せずに閉じる
                    doc.Close(false);
                    //Wordを終了
                    word.Quit();
                    //ガベージコレクション
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
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
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog1.FileName; // 選択されたファイルパスを取得
                doc.SaveAs2(filePath, Word.WdSaveFormat.wdFormatDocumentDefault); // ファイルを保存
                MessageBox.Show("指定された場所にファイルが正しく保存されました: " + filePath, "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                DialogResult result = MessageBox.Show("文書作成ソフトが起動しています。保存せずに文書作成ソフトを閉じますか?", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    // ドキュメントを保存せずに閉じる
                    doc.Close(false);
                    //Wordを終了
                    word.Quit();
                    //ガベージコレクション
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
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
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog1.FileName; // 選択されたファイルパスを取得
                doc.SaveAs2(filePath, Word.WdSaveFormat.wdFormatDocumentDefault); // ファイルを保存
                MessageBox.Show("指定された場所にファイルが正しく保存されました: " + filePath, "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                DialogResult result = MessageBox.Show("文書作成ソフトが起動しています。保存せずに文書作成ソフトを閉じますか?", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    // ドキュメントを保存せずに閉じる
                    doc.Close(false);
                    //Wordを終了
                    word.Quit();
                    //ガベージコレクション
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
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

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            this.Text = textBox10.Text + " - Document Maker";
            if(textBox10.Text == string.Empty)
            {
                this.Text = "無題 - Document Maker";
            }
        }

        private void textBox10_Click(object sender, EventArgs e)
        {

        }

        private void textBox10_ParentChanged(object sender, EventArgs e)
        {

        }
    }
}
