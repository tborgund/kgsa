using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    public partial class FormProcessing : Form
    {
        private const int CS_DROPSHADOW = 0x00020000;
        delegate void SetProgressBarValueCallback(int b);
        delegate void SetProgressTextCallback(string s);
        delegate void SetCancelButtonCallback(bool value);
        delegate void SetProgressStyleCallback(ProgressBarStyle style);
        delegate void SetUpdateFormCallback();
        private Form main;
        private BackgroundWorker bgWorker;
        private Timer timer;
        public bool userPushedCancelButton = false;

        public FormProcessing(FormMain form)
        {
            InitializeComponent();
            this.main = form;
            this.Owner = main;
            timer = new Timer();
            timer.Interval = 1500;
            timer.Tick += new EventHandler(timer_Close);
        }

        private void timer_Close(object sender, EventArgs e)
        {
            timer.Stop();
            this.Hide();
            ResetVariables();
        }

        private void ResetVariables()
        {
            SetProgressBarStyle(ProgressBarStyle.Marquee);
            textBoxProgressText.Text = "Arbeider..";
            buttonCancel.Text = "Avbryt";
        }

        public bool SupportsCancelation
        {
            get { return this.buttonCancel.Enabled; }
            set { SetCancelButton(value); }
        }

        private void SetCancelButton(bool value)
        {
            if (buttonCancel.InvokeRequired)
            {
                SetCancelButtonCallback d = new SetCancelButtonCallback(SetCancelButton);
                this.Invoke(d, new object[] { value });
            }
            else
            {
                buttonCancel.Enabled = value;
            }
        }

        private void SetFormPosition()
        {
            this.Location = new Point(
                (main.Location.X + main.Width / 2) - (this.Width / 2),
                (main.Location.Y + main.Height / 3) - (this.Height / 2));
        }

        public bool SetVisible
        {
            get { return this.Visible; }
            set { if (value) { MakeVisible(); } else { this.Hide(); } }
        }

        private void MakeVisible()
        {
            userPushedCancelButton = false;
            timer.Stop();
            SetCancelButton(false);
            SetProgressBarStyle(ProgressBarStyle.Marquee);
            SetProgressBarValue(0);
            SetFormPosition();
            ProgressText("Arbeider..");

            if (this.InvokeRequired)
            {
                SetUpdateFormCallback d = new SetUpdateFormCallback(MakeVisible);
                this.Invoke(d, new object[] { });
            }
            else
            {
                this.Show();
                this.Update();
            }
        }

        public void HideDelayed()
        {
            timer.Start();
            buttonCancel.Enabled = false;
        }

        public string SetText
        {
            set { ProgressText(value); }
        }

        public BackgroundWorker SetBackgroundWorker
        {
            set { this.bgWorker = value; SupportCancel(bgWorker); }
        }

        private void ProgressText(string s)
        {
            if (textBoxProgressText.InvokeRequired)
            {
                SetProgressTextCallback d = new SetProgressTextCallback(ProgressText);
                this.Invoke(d, new object[] { s });
            }
            else
            {
                textBoxProgressText.Text = s;
            }
        }

        public int SetValue
        {
            set { SetProgressBarValue(value); }
        }

        private void SetProgressBarValue(int b)
        {
            if (progressBar.InvokeRequired)
            {
                SetProgressBarValueCallback d = new SetProgressBarValueCallback(SetProgressBarValue);
                this.Invoke(d, new object[] { b });
            }
            else
            {
                if (progressBar.Style != ProgressBarStyle.Continuous)
                    progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = b;
            }
        }

        public ProgressBarStyle SetProgressStyle
        {
            set { SetProgressBarStyle(value); }
        }

        private void SetProgressBarStyle(ProgressBarStyle style)
        {
            if (progressBar.InvokeRequired)
            {
                SetProgressStyleCallback d = new SetProgressStyleCallback(SetProgressBarStyle);
                this.Invoke(d, new object[] { style });
            }
            else
                progressBar.Style = style;
        }


        /// <summary>
        /// Aktiverer avbryt knapp hvis BackgroundWorker støttet kanselering.
        /// </summary>
        /// <param name="bg">Background worker som skal sjekkes</param>
        /// <returns>Returnerer sant hvis BackgroundWorker støtter kanselering</returns>
        private bool SupportCancel(BackgroundWorker bg)
        {
            this.bgWorker = bg;
            if (bgWorker != null)
            {
                if (bgWorker.WorkerSupportsCancellation)
                {
                    SetProgressbarButton(true);
                    return true;
                }
                else
                {
                    SetProgressbarButton(false);
                    return false;
                }
            }
            SetProgressbarButton(false);
            buttonCancel.Enabled = false;
            return false;
        }

        private void SetProgressbarButton(bool value)
        {
            if (buttonCancel.InvokeRequired)
            {
                SetCancelButtonCallback d = new SetCancelButtonCallback(SetProgressbarButton);
                this.Invoke(d, new object[] { value });
            }
            else
            {
                buttonCancel.Enabled = value;
            }
        }

        protected override void OnLostFocus(EventArgs e)
        {
            if (this.Visible)
            {
                base.OnLostFocus(e);
                this.Focus();
            }
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= CS_DROPSHADOW;
                return cp;
            }
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            userPushedCancelButton = true;
            if (bgWorker == null)
            {
                Logg.Log("Kan ikke avbryte prosessen!", Color.Red);
                return;
            }
            textBoxProgressText.Text = "Avbryter..";
            buttonCancel.Text = "Avbryter..";
            bgWorker.CancelAsync();
            while(bgWorker.IsBusy)
            {
                Application.DoEvents();
                System.Threading.Thread.Sleep(50);
            }
            this.Hide();
            ResetVariables();
        }
    }
}
