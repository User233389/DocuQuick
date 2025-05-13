using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Document_Maker
{
    public partial class DCW : Form
    {
        public DCW()
        {
            InitializeComponent();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                numericUpDown1.Enabled = false;
                label3.Enabled = false;
                label6.Enabled = false;
                label7.Enabled = false;
            }
            else
            {
                numericUpDown1.Enabled = true;
                label3.Enabled = true;
                label6.Enabled = true;
                label7.Enabled = true;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                monthCalendar1.Enabled = false;
                label4.Enabled = false;
            }
            else
            {

                monthCalendar1.Enabled = true;
                label4.Enabled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            FontDialog fontDialog = new FontDialog();

            // Set the initial font and color
            fontDialog.ShowColor = true;
            fontDialog.Font = label18.Font;
            fontDialog.Site = label18.Site;
            fontDialog.Color = label18.ForeColor;
            // Set the font dialog options
            if (fontDialog.ShowDialog() == DialogResult.OK)
            {
                label18.Font = fontDialog.Font;
                label18.ForeColor = fontDialog.Color;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.Color = label18.ForeColor;
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                label18.ForeColor = colorDialog.Color;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            label18.Font = new System.Drawing.Font("游明朝", 11);
            label18.ForeColor = System.Drawing.Color.Black;
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            if (textBox10.Text == string.Empty)
            {
                label18.Text = "書類送付のお知らせ(例)";
            }
            else
            {
                label18.Text = textBox10.Text;
            }
        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            //一般的な頭語の結語を追加
            if ((string)comboBox1.SelectedItem == "拝啓（一般的）")
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
            }
            else if ((string)comboBox1.SelectedItem == "拝呈（一般的）")
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
            }
            else if ((string)comboBox1.SelectedItem == "啓上（一般的）")
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
            }
            else if ((string)comboBox1.SelectedItem == "啓白（一般的）")
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
            }
            else if ((string)comboBox1.SelectedItem == "拝進（一般的）")
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
            }
            //丁寧さの頭語の結語を追加
            else if ((string)comboBox1.SelectedItem == "謹啓（丁寧さ）")
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
            }
            else if ((string)comboBox1.SelectedItem == "謹呈（丁寧さ）")
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
            }
            else if ((string)comboBox1.SelectedItem == "粛啓（丁寧さ）")
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
            }
            else if ((string)comboBox1.SelectedItem == "恭啓（丁寧さ）")
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
            }
            else if ((string)comboBox1.SelectedItem == "謹白（丁寧さ）")
            {
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
            }

        }

        private void commandLink1_Click(object sender, EventArgs e)
        {

        }
    }
}   
