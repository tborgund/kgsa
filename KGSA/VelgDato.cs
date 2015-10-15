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
    public partial class VelgDato : Form
    {
        private AppSettings appConfig;
        public VelgDato(AppSettings app)
        {
            this.appConfig = app;
            InitializeComponent();
        }

        private void VelgDato_Shown(object sender, EventArgs e)
        {
            dateTimePicker.MinDate = appConfig.dbStoreFrom;
            dateTimePicker.MaxDate = appConfig.dbStoreTo;
            if (appConfig.dbStoreViewpoint > appConfig.dbStoreFrom && appConfig.dbStoreViewpoint < appConfig.dbStoreTo)
                dateTimePicker.Value = appConfig.dbStoreViewpoint;
            else
                dateTimePicker.Value = appConfig.dbStoreFrom;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dateTimePicker.Value.Date >= appConfig.dbStoreFrom && dateTimePicker.Value <= appConfig.dbStoreTo.Date)
                appConfig.dbStoreViewpoint = dateTimePicker.Value;
        }
    }
}
