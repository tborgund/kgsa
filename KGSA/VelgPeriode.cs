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
        public DateTime dtFra;
        public DateTime dtTil;
        public DateTime dtCurrent;
        public static CultureInfo norway = new CultureInfo("nb-NO");

        public VelgPeriode()
        {
            InitializeComponent();
        }

        private void VelgPeriode_Shown(object sender, EventArgs e)
        {
            dateTimePicker1.MinDate = FormMain.dbFraDT;
            dateTimePicker1.MaxDate = FormMain.dbTilDT;
            dateTimePicker1.Value = FormMain.datoPeriodeFra;

            dateTimePicker2.MinDate = FormMain.dbFraDT;
            dateTimePicker2.MaxDate = FormMain.dbTilDT;
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
