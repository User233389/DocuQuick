using Microsoft.Win32;
using System;
using System.Threading;
using System.Windows.Forms;

namespace Document_Maker
{
    public partial class WordInstallationValidationWizard : Form
    {
        public WordInstallationValidationWizard()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }


        private bool IsWordInstalled()
        {
            const string wordRegistryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE";

            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(wordRegistryKey))
            {
                return key != null;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Loading ld = new Loading();
            ld.Show();
            Microsoft.Office.Interop.Word.Application wordApp = null;

            try
            {
                // Microsoft Wordのインストールを確認する
                if (IsWordInstalled())
                {
                    // Wordを起動する
                    wordApp = new Microsoft.Office.Interop.Word.Application();
                    wordApp.Visible = true;

                    pictureBox1.Show();
                    pictureBox4.Hide();

                    pictureBox2.Show();
                    pictureBox5.Hide();

                    Thread.Sleep(2000); // 2秒待機

                    // Wordを終了する
                    try
                    {
                        wordApp.Quit();
                        pictureBox3.Show();
                        pictureBox6.Hide();
                        label3.Text = "Microsoft Wordが正常にインストールされ起動・終了しました。\nWordは問題なく使用できます。";
                        ld.Close();
                    }
                    catch (System.Runtime.InteropServices.COMException comEx)
                    {
                        pictureBox3.Hide();
                        pictureBox6.Show();
                        label3.Text = "Microsoft Wordの終了に失敗しました。\nエラー内容: " + comEx.Message;
                        ld.Close();
                    }
                }
                else
                {
                    pictureBox1.Hide();
                    pictureBox4.Show();

                    pictureBox2.Hide();
                    pictureBox5.Show();

                    pictureBox3.Hide();
                    pictureBox6.Show();
                    label3.Text = "Microsoft Wordがインストールされておらず\n起動できませんでした。";
                    ld.Close();
                }
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                pictureBox1.Show();
                pictureBox4.Hide();

                pictureBox2.Hide();
                pictureBox5.Show();

                pictureBox3.Hide();
                pictureBox6.Hide();
                label3.Text = "Microsoft Wordの起動に失敗しました。\nエラー内容: " + comEx.Message;
                ld.Close();
            }
            catch (Exception ex)
            {
                label3.Text = "予期しないエラーが発生しました。\nエラー内容: " + ex.Message;
                ld.Close();
            }
            finally
            {
                // リソース解放
                if (wordApp != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                    wordApp = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
