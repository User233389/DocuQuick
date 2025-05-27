using Microsoft.Win32;
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

namespace Document_Maker
{
    public partial class Splash : Form
    {
        public Splash()
        {
            InitializeComponent();
            SetLabelTextColorByTitleBar();

        }


        // DWM APIを使用してタイトルバーの色を取得するためのP/Invoke宣言
        [DllImport("dwmapi.dll", EntryPoint = "DwmGetColorizationColor")]
        private static extern int DwmGetColorizationColor(out uint pcrColorization, out bool pfOpaqueBlend);

        private Color GetTitleBarColor()
        {
            DwmGetColorizationColor(out uint colorizationColor, out bool opaqueBlend);
            byte a = (byte)((colorizationColor >> 24) & 0xFF);
            byte r = (byte)((colorizationColor >> 16) & 0xFF);
            byte g = (byte)((colorizationColor >> 8) & 0xFF);
            byte b = (byte)(colorizationColor & 0xFF);
            return Color.FromArgb(a, r, g, b);
        }


        private void SetLabelTextColorByTitleBar()
        {
            Color titleBarColor = GetTitleBarColor();

            // 明るさを計算（輝度の簡易計算）
            double brightness = (0.299 * titleBarColor.R + 0.587 * titleBarColor.G + 0.114 * titleBarColor.B);

            // しきい値を基に文字色を決定（128は中間）
            Color textColor = brightness > 128 ? Color.Black : Color.White;

            // ラベルの背景色をタイトルバーの色に設定
            // ラベルの文字色を変更（複数ある場合はループ）
            themedLabel1.ForeColor = textColor;
            themedLabel2.ForeColor = textColor; // 他のラベルも同様に設定
        }

        //タイトルバーにアクセントカラーが適用されているかを確認するメソッド
        private static bool IsAccentColorOnTitleBar()
        {
            const string keyPath = @"Software\Microsoft\Windows\DWM";
            const string valueName = "ColorPrevalence";

            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath))
            {
                if (key != null)
                {
                    object value = key.GetValue(valueName);
                    if (value != null && value is int intValue)
                    {
                        return intValue == 1;
                    }
                }
            }

            return false;
        }

        private void Splash_Shown(object sender, EventArgs e)
        {
            if (IsAccentColorOnTitleBar() == false)
            // タイトルバーにアクセントカラーが適用されていない場合、ラベルの文字色を黒に設定
            {
                themedLabel1.ForeColor = Color.Black;
                themedLabel2.ForeColor = Color.Black; // 他のラベルも同様に設定
            }

            //処理の開始
            //Refreshによりコントロールを適切に描画する
            this.Refresh(); // フォームを更新して描画を反映
            progressBar1.Value = 10;
            Thread.Sleep(500); // 0.5秒待機
            progressBar1.Value = 30;
            Thread.Sleep(500); // 0.5秒待機
            progressBar1.Value = 50;
            Thread.Sleep(500); // 0.5秒待機
            progressBar1.Value = 70;
            Thread.Sleep(500); // 0.5秒待機
            progressBar1.Value = 90;
            Thread.Sleep(500); // 0.5秒待機
            progressBar1.Value = 100;
            this.Close(); // スプラッシュスクリーンを閉じる
        }
    }
}
