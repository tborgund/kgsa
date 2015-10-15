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
            Logg.Log("Uhåndtert unntak oppstod! Unntak: " + ex.Message, Color.Red);
            Logg.Debug("Unntak beskjed: " + ex.Message);
            Logg.Debug("Unntak: " + ex.ToString());
            Logg.Debug("KGSA versjon: " + FormMain.version);
            Logg.Debug("OS versjon: " + Environment.OSVersion.Version.ToString());
            Logg.Debug("Tid og Dato: " + DateTime.Now.ToShortTimeString() + " " + DateTime.Now.ToShortDateString());
            InitializeComponent();
            this.Text = "KGSA (" + FormMain.version + ") - Kritisk feil";

            labelErrorTitle.Text = tittel;
            if (detaljer != "")
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
