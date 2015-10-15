namespace KGSA
{
    partial class FormMacro
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMacro));
            this.label1 = new System.Windows.Forms.Label();
            this.labelLog = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.panel1 = new System.Windows.Forms.Panel();
            this.buttonStartStop = new System.Windows.Forms.Button();
            this.panelWarning = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel1.SuspendLayout();
            this.panelWarning.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(90, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "Status:";
            // 
            // labelLog
            // 
            this.labelLog.AutoSize = true;
            this.labelLog.Location = new System.Drawing.Point(136, 35);
            this.labelLog.Name = "labelLog";
            this.labelLog.Size = new System.Drawing.Size(41, 13);
            this.labelLog.TabIndex = 12;
            this.labelLog.Text = "Venter.";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(93, 9);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(349, 23);
            this.progressBar.TabIndex = 13;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.buttonStartStop);
            this.panel1.Controls.Add(this.progressBar);
            this.panel1.Controls.Add(this.labelLog);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 312);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(454, 54);
            this.panel1.TabIndex = 14;
            // 
            // buttonStartStop
            // 
            this.buttonStartStop.Location = new System.Drawing.Point(12, 9);
            this.buttonStartStop.Name = "buttonStartStop";
            this.buttonStartStop.Size = new System.Drawing.Size(75, 23);
            this.buttonStartStop.TabIndex = 14;
            this.buttonStartStop.Text = "Stop";
            this.buttonStartStop.UseVisualStyleBackColor = true;
            this.buttonStartStop.Click += new System.EventHandler(this.buttonStartStop_Click);
            // 
            // panelWarning
            // 
            this.panelWarning.Controls.Add(this.pictureBox1);
            this.panelWarning.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelWarning.Location = new System.Drawing.Point(0, 0);
            this.panelWarning.Name = "panelWarning";
            this.panelWarning.Size = new System.Drawing.Size(454, 312);
            this.panelWarning.TabIndex = 16;
            this.panelWarning.Visible = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::KGSA.Properties.Resources.macro_img_sign;
            this.pictureBox1.Location = new System.Drawing.Point(82, 10);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(296, 296);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // FormMacro
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(454, 366);
            this.ControlBox = false;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panelWarning);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormMacro";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "KGSA - Makro";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormMacro_FormClosing);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panelWarning.ResumeLayout(false);
            this.panelWarning.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelLog;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button buttonStartStop;
        private System.Windows.Forms.Panel panelWarning;
        private System.Windows.Forms.PictureBox pictureBox1;

    }
}