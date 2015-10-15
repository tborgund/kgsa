namespace KGSA
{
    partial class FormSplash
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormSplash));
            this.labelVersion = new System.Windows.Forms.Label();
            this.labelProgress = new System.Windows.Forms.Label();
            this.labelBuildDate = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelVersion
            // 
            this.labelVersion.AutoSize = true;
            this.labelVersion.BackColor = System.Drawing.Color.Transparent;
            this.labelVersion.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelVersion.ForeColor = System.Drawing.Color.White;
            this.labelVersion.Location = new System.Drawing.Point(38, 37);
            this.labelVersion.Name = "labelVersion";
            this.labelVersion.Size = new System.Drawing.Size(15, 20);
            this.labelVersion.TabIndex = 0;
            this.labelVersion.Text = "-";
            // 
            // labelProgress
            // 
            this.labelProgress.AutoSize = true;
            this.labelProgress.BackColor = System.Drawing.Color.Transparent;
            this.labelProgress.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelProgress.ForeColor = System.Drawing.Color.White;
            this.labelProgress.Location = new System.Drawing.Point(39, 140);
            this.labelProgress.Name = "labelProgress";
            this.labelProgress.Size = new System.Drawing.Size(59, 16);
            this.labelProgress.TabIndex = 2;
            this.labelProgress.Text = "Laster..";
            // 
            // labelBuildDate
            // 
            this.labelBuildDate.AutoSize = true;
            this.labelBuildDate.BackColor = System.Drawing.Color.Transparent;
            this.labelBuildDate.ForeColor = System.Drawing.Color.White;
            this.labelBuildDate.Location = new System.Drawing.Point(133, 42);
            this.labelBuildDate.Name = "labelBuildDate";
            this.labelBuildDate.Size = new System.Drawing.Size(10, 13);
            this.labelBuildDate.TabIndex = 3;
            this.labelBuildDate.Text = "-";
            // 
            // FormSplash
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::KGSA.Properties.Resources.splash;
            this.ClientSize = new System.Drawing.Size(477, 180);
            this.Controls.Add(this.labelBuildDate);
            this.Controls.Add(this.labelProgress);
            this.Controls.Add(this.labelVersion);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormSplash";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "KGSA - Laster..";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelVersion;
        private System.Windows.Forms.Label labelProgress;
        private System.Windows.Forms.Label labelBuildDate;

    }
}