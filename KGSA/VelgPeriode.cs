using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    public partial class VelgPeriode : Form
    {
        FormMain main;
        public DateTime dtFra;
        public DateTime dtTil;
        public DateTime dtCurrent;
        public static CultureInfo norway = new CultureInfo("nb-NO");

        public VelgPeriode(FormMain form)
        {
            this.main = form;
            InitializeComponent();
        }

        private void VelgPeriode_Shown(object sender, EventArgs e)
        {
            dateTimePicker1.MinDate = main.appConfig.dbFrom;
            dateTimePicker1.MaxDate = main.appConfig.dbTo;
            dateTimePicker1.Value = FormMain.datoPeriodeFra;

            dateTimePicker2.MinDate = main.appConfig.dbFrom;
            dateTimePicker2.MaxDate = main.appConfig.dbTo;
            dateTimePicker2.Value = FormMain.datoPeriodeTil;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dtFra = dateTimePicker1.Value;
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            dtTil = dateTimePicker2.Value;
        }
    }
}
