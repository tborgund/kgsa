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
    public partial class FormMsgAutoranking : Form
    {
        AppSettings appConfig;
        public FormMsgAutoranking(AppSettings settings)
        {
            this.appConfig = settings;
            InitializeComponent();
            textBox1.Text = appConfig.epostNesteMelding.Replace("\n", Environment.NewLine);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            appConfig.epostNesteMelding = textBox1.Text.Replace(Environment.NewLine, "\n");        
        }
    }
}
