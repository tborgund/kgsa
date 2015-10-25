using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    public partial class FormKrav : Form
    {
        FormMain main;
        private DataSet dsSk;
        private SqlCeDataAdapter daSk;
        private SqlCeConnection conSk = new SqlCeConnection(FormMain.SqlConStr);

        public FormKrav(FormMain form)
        {
            this.main = form;
            InitializeComponent();

            InitSelgerkoder();
            importSettings();
        }

        private void InitSelgerkoder()
        {
            dsSk = new DataSet();
            conSk.Open();
            daSk = new SqlCeDataAdapter("Select * from tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'", conSk);
            var cmdBldr = new SqlCeCommandBuilder(daSk);
            daSk.Fill(dsSk, "tblSelgerkoder");
            bindingSourceSk.DataSource = dsSk;
            bindingSourceSk.DataMember = "tblSelgerkoder";
            dataGridViewSk.DataSource = bindingSourceSk;
            bindingNavigatorSk.BindingSource = bindingSourceSk;

            toolStripComboBoxSkFilter.SelectedIndex = 0;


            conSk.Close();
        }

        private void toolStripComboBoxSkFilter_DropDownClosed(object sender, EventArgs e)
        {
            OppdaterSelgerkoder();
        }

        private void OppdaterSelgerkoder()
        {
            if (dataGridViewSk.CurrentRow == null)
                InitSelgerkoder();

            dsSk.Clear();
            conSk.Open();
            if (toolStripComboBoxSkFilter.SelectedIndex > 0)
                daSk = new SqlCeDataAdapter("Select * from tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "' AND Kategori = '" + toolStripComboBoxSkFilter.SelectedItem.ToString() + "'", conSk);
            else
                daSk = new SqlCeDataAdapter("Select * from tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'", conSk);
            var cmdBldr = new SqlCeCommandBuilder(daSk);
            daSk.Fill(dsSk, "tblSelgerkoder");
            conSk.Close();
        }

        private void dataGridViewSk_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Log.Status("");
            dataGridViewSk.Rows[e.RowIndex].ErrorText = String.Empty;
        }

        private void dataGridViewSk_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                if (String.IsNullOrEmpty(e.FormattedValue.ToString()))
                    return;
                if (dataGridViewSk.Rows[e.RowIndex].IsNewRow || e.ColumnIndex < 4)
                    return;

                var value = e.FormattedValue.ToString();
                int number;
                bool result = Int32.TryParse(value, out number);
                if (result)
                {
                    // Alt er OK!
                    Log.d("Endret krav til " + number);
                }
                else
                {
                    // Format feil!
                    if (value == null) value = "";
                    dataGridViewSk.Rows[e.RowIndex].ErrorText = "Feil format i krav. Må være hele ikke-negative tall."; //error massage
                    Log.n("Feil format i krav. Må være hele ikke-negative tall.", Color.Red);
                    e.Cancel = true;
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                dataGridViewSk[e.ColumnIndex, e.RowIndex].Value = "0";
            }
        }

        private void dataGridViewSk_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            try
            {
                if (dataGridViewSk.CurrentRow != null)
                    for (int i = 0; i < dsSk.Tables["tblSelgerkoder"].Rows.Count; i++)
                        dsSk.Tables["tblSelgerkoder"].Rows[i]["Avdeling"] = main.appConfig.Avdeling;
            }
            catch (DeletedRowInaccessibleException ex)
            {
                Log.n("Unntak oppstod under oppdatering av selgerliste. Exception: " + ex, Color.Red);
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Unntak oppstod under oppdatering av selgerliste", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void LagreSelgerkoder()
        {
            try
            {
                if (dataGridViewSk.CurrentRow != null)
                    dataGridViewSk.CurrentRow.DataGridView.EndEdit();
                dataGridViewSk.EndEdit();
                bindingSourceSk.EndEdit();

                conSk.Open();
                daSk.Update(dsSk, "tblSelgerkoder");
                conSk.Close();
                OppdaterSelgerkoder();
                Log.n("Selgerkoder Lagret.", Color.Green);
                main.salesCodes.Update();
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil oppstod under lagring av selgerkoder.", ex, "Sjekk om alle påkrevde felt er utfylt.\n\nException: ");
                errorMsg.ShowDialog(this);
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            LagreSelgerkoder();
        }

        private void importSettings()
        {
            // OVERSIKT - Krav
            checkOversiktVisKrav.Checked = main.appConfig.oversiktKravVis;
            checkKravFinans.Checked = main.appConfig.oversiktKravFinans;
            checkKravMod.Checked = main.appConfig.oversiktKravMod;
            checkKravStrom.Checked = main.appConfig.oversiktKravStrom;
            checkKravRtgsa.Checked = main.appConfig.oversiktKravRtgsa;
            checkKravAntallFinans.Checked = main.appConfig.oversiktKravFinansAntall;
            checkKravAntallMod.Checked = main.appConfig.oversiktKravModAntall;
            checkKravAntallStrom.Checked = main.appConfig.oversiktKravStromAntall;
            checkKravAntallRtgsa.Checked = main.appConfig.oversiktKravRtgsaAntall;
            checkKravMTD.Checked = main.appConfig.oversiktKravMtd;
            checkKravMtdShowTarget.Checked = main.appConfig.oversiktKravMtdShowTarget;
        }

        private void exportSettings()
        {
            // OVERSIKT - Krav
            main.appConfig.oversiktKravVis = checkOversiktVisKrav.Checked;
            main.appConfig.oversiktKravFinansAntall = checkKravAntallFinans.Checked;
            main.appConfig.oversiktKravModAntall = checkKravAntallMod.Checked;
            main.appConfig.oversiktKravStromAntall = checkKravAntallStrom.Checked;
            main.appConfig.oversiktKravRtgsaAntall = checkKravAntallRtgsa.Checked;
            main.appConfig.oversiktKravFinans = checkKravFinans.Checked;
            main.appConfig.oversiktKravMod = checkKravMod.Checked;
            main.appConfig.oversiktKravStrom = checkKravStrom.Checked;
            main.appConfig.oversiktKravRtgsa = checkKravRtgsa.Checked;
            main.appConfig.oversiktKravMtd = checkKravMTD.Checked;
            main.appConfig.oversiktKravMtdShowTarget = checkKravMtdShowTarget.Checked;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            exportSettings();
            LagreSelgerkoder();
        }
    }
}
