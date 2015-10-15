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
    public partial class FormBudgetTimer : Form
    {
        AppSettings appConfig;
        DataSet ds;
        SqlCeDataAdapter da;
        SqlCeConnection con = new SqlCeConnection(FormMain.SqlConStr);
        BudgetInfo budgetInfo;
        int budgetId;

        public FormBudgetTimer(AppSettings app, BudgetInfo info, int id)
        {
            this.appConfig = app;
            this.budgetInfo = info;
            this.budgetId = id;
            InitializeComponent();
            AddColumns();


            RefreshDataGrid();
        }

        private bool SaveDb()
        {
            try
            {
                if (dataGridView1.CurrentRow != null)
                    dataGridView1.CurrentRow.DataGridView.EndEdit();
                dataGridView1.EndEdit();
                bindingSource1.EndEdit();

                con.Open();
                da.Update(ds, "tblBudgetTimer");
                con.Close();

                RefreshDataGrid();
                return true;
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil oppstod under lagring av timer.", ex, "Sjekk om alle påkrevde felt er utfylt.\n\nException: ");
                errorMsg.ShowDialog(this);
            }
            return false;
        }

        private void AddColumns()
        {
            var id = new DataGridViewTextBoxColumn();
            id.Visible = false;
            id.DataPropertyName = "Id";
            dataGridView1.Columns.Add(id);

            var budgetid = new DataGridViewTextBoxColumn();
            budgetid.Visible = false;
            budgetid.DataPropertyName = "BudgetId";
            dataGridView1.Columns.Add(budgetid);

            var selgerkode = new DataGridViewTextBoxColumn();
            selgerkode.HeaderText = "Selgerkode";
            selgerkode.Name = "Selgerkode";
            selgerkode.MinimumWidth = 75;
            selgerkode.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            selgerkode.FillWeight = 100;
            selgerkode.DataPropertyName = "Selgerkode";
            selgerkode.ReadOnly = true;
            dataGridView1.Columns.Add(selgerkode);

            for (int i = 1; i <= 31; i++)
            {
                var dag = new DataGridViewTextBoxColumn();
                dag.HeaderText = "D" + i;
                dag.Name = "D" + i;
                dag.Width = 30;
                dag.DataPropertyName = i.ToString();
                dag.MaxInputLength = 4;
                dag.ValueType = typeof(Decimal);
                dag.ToolTipText = FormMain.GetFirstDayOfMonth(budgetInfo.date).AddDays(i - 1).ToString("dddd d. MMMM", FormMain.norway);
                dataGridView1.Columns.Add(dag);

            }
        }

        private void RefreshDataGrid()
        {
            ds = new DataSet();
            con.Open();
            da = new SqlCeDataAdapter("Select * from tblBudgetTimer WHERE BudgetId = " + budgetId, con);
            var cmdBldr = new SqlCeCommandBuilder(da);
            da.Fill(ds, "tblBudgetTimer");
            bindingSource1.DataSource = ds;
            bindingSource1.DataMember = "tblBudgetTimer";
            dataGridView1.DataSource = bindingSource1;
            bindingNavigator1.BindingSource = bindingSource1;
            con.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveDb();

            if (MessageBox.Show("Vil du oppdatere timeantall?", "Oppdater timeantall", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes)
                this.DialogResult = System.Windows.Forms.DialogResult.OK;

            this.Close();
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            Logg.Log("Format feil oppdaget i time tabellen!", Color.Red);

        }
    }
}
