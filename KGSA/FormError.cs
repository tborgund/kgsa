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
    public partial class FormError : Form
    {
        public FormError(string tittel, Exception ex, string detaljer = "")
        {

            if (ex == null)
            {
                ex = new Exception("Exception object was null");
            }
            Log.n("Uhåndtert unntak oppstod! Unntak: " + ex.Message, Color.Red);
            Log.d("Unntak beskjed: " + ex.Message);
            Log.d("Unntak: " + ex.ToString());
            Log.d("KGSA versjon: " + FormMain.version);
            Log.d("OS versjon: " + Environment.OSVersion.Version.ToString());
            Log.d("Tid og Dato: " + DateTime.Now.ToShortTimeString() + " " + DateTime.Now.ToShortDateString());
            InitializeComponent();
            this.Text = "KGSA (" + FormMain.version + ") - Kritisk feil";

            labelErrorTitle.Text = tittel;
            if (!String.IsNullOrEmpty(detaljer))
                textBoxErrorMessage.Text = detaljer + Environment.NewLine;
            textBoxErrorMessage.Text += "Unntak beskjed: " + ex.Message;
            textBoxErrorMessage.Text += Environment.NewLine + "Unntak: " + ex.ToString();
            textBoxErrorMessage.Text += Environment.NewLine + "KGSA versjon: " + FormMain.version;
            textBoxErrorMessage.Text += Environment.NewLine + "OS versjon: " + Environment.OSVersion.Version.ToString();
            textBoxErrorMessage.Text += Environment.NewLine + "Tid og Dato: " + DateTime.Now.ToShortTimeString() + " " + DateTime.Now.ToShortDateString();

            try
            {
                textBoxErrorMessage.Text += Environment.NewLine + "settingsPath: " + FormMain.settingsPath;
                textBoxErrorMessage.Text += Environment.NewLine + "settingsFile: " + FormMain.settingsFile;
                textBoxErrorMessage.Text += Environment.NewLine + "settingsTemp: " + FormMain.settingsTemp;
                textBoxErrorMessage.Text += Environment.NewLine + "Gjeldene brukers temp mappe: " + System.IO.Path.GetTempPath();
            }
            catch
            {
                textBoxErrorMessage.Text += Environment.NewLine + "Obs! Kunne ikke hente frem miljø variabler.";
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
