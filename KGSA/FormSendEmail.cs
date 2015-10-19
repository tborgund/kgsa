using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace KGSA
{
    public partial class FormSendEmail : Form
    {
        FormMain main;
        public FormSendEmail(FormMain form)
        {
            this.main = form;
            InitializeComponent();
            comboBoxGruppe.SelectedIndex = 1;
            comboBoxType.SelectedIndex = 0;
            textBoxTitle.Text = main.appConfig.epostEmne;
            textBoxContent.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Er du sikker?", "KGSA - Varsel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FormEmailAddressbook emailForm = new FormEmailAddressbook();

            emailForm.ShowDialog();
        }
    }
}
