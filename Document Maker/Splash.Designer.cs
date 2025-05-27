namespace Document_Maker
{
    partial class Splash
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.glassExtenderProvider1 = new Vanara.Interop.DesktopWindowManager.GlassExtenderProvider();
            this.themedLabel1 = new AeroWizard.ThemedLabel();
            this.themedLabel2 = new AeroWizard.ThemedLabel();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // themedLabel1
            // 
            this.themedLabel1.Font = new System.Drawing.Font("Segoe UI", 27.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.themedLabel1.ForeColor = System.Drawing.Color.DarkBlue;
            this.themedLabel1.Location = new System.Drawing.Point(29, 32);
            this.themedLabel1.Name = "themedLabel1";
            this.themedLabel1.Size = new System.Drawing.Size(427, 46);
            this.themedLabel1.TabIndex = 1;
            this.themedLabel1.Text = "Document Maker";
            // 
            // themedLabel2
            // 
            this.themedLabel2.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.themedLabel2.Location = new System.Drawing.Point(29, 90);
            this.themedLabel2.Name = "themedLabel2";
            this.themedLabel2.Size = new System.Drawing.Size(427, 22);
            this.themedLabel2.TabIndex = 2;
            this.themedLabel2.Text = "設定情報を読み込み中...";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 229);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(461, 23);
            this.progressBar1.TabIndex = 0;
            // 
            // Splash
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ClientSize = new System.Drawing.Size(485, 266);
            this.ControlBox = false;
            this.Controls.Add(this.themedLabel2);
            this.Controls.Add(this.themedLabel1);
            this.Controls.Add(this.progressBar1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.glassExtenderProvider1.SetGlassMargins(this, new System.Windows.Forms.Padding(0, 215, 0, 0));
            this.Name = "Splash";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Document Maker 1.0.2";
            this.Shown += new System.EventHandler(this.Splash_Shown);
            this.ResumeLayout(false);

        }

        #endregion
        private Vanara.Interop.DesktopWindowManager.GlassExtenderProvider glassExtenderProvider1;
        private AeroWizard.ThemedLabel themedLabel1;
        private AeroWizard.ThemedLabel themedLabel2;
        private System.Windows.Forms.ProgressBar progressBar1;
    }
}