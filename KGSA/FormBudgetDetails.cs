using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    public partial class FormBudgetDetails : Form
    {
        AppSettings appConfig;
        DataSet ds;
        SqlCeDataAdapter da;
        SqlCeConnection con = new SqlCeConnection(FormMain.SqlConStr);
        BudgetObj budget;
        BudgetInfo budgetInfo;
        int budgetId;
        public FormBudgetDetails(AppSettings app, BudgetObj bud, int id)
        {
            this.appConfig = app;
            this.budget = bud;
            this.budgetId = id;
            InitializeComponent();

            budget.UpdateBudgetSelgerkoder(budgetId);
            budgetInfo = budget.GetBudgetInfo(budgetId);


            SetUI();

            RefreshDataGridView();
        }

        private void SetUI()
        {
            this.Text = "Budsjett: " + budgetInfo.kategori + " selgerkoder";

            dateTime_date.Format = DateTimePickerFormat.Custom;
            dateTime_date.CustomFormat = "MMMM yyyy";
            dateTime_date.Value = budgetInfo.date;
            textBox_date.Text = budgetInfo.date.ToString("MMMM yyyy");
            textBox_dager.Text = budgetInfo.dager.ToString();
            textBox_omsetning.Text = budgetInfo.omsetning.ToString();
            textBox_inntjening.Text = budgetInfo.inntjening.ToString();
            textBox_margin.Text = budgetInfo.margin.ToString();
            textBox_ta.Text = budgetInfo.ta.ToString();
            textBox_strom.Text = budgetInfo.strom.ToString();
            textBox_finans.Text = budgetInfo.finans.ToString();
            textBox_rtgsa.Text = budgetInfo.rtgsa.ToString();
            textBox_acc.Text = budgetInfo.acc.ToString();
            textBox_vinn.Text = budgetInfo.vinn.ToString();

            comboBox_Acc.Text = budgetInfo.TypeToString(budgetInfo.acc_type);
            comboBox_Finans.Text = budgetInfo.TypeToString(budgetInfo.finans_type);
            comboBox_Rtgsa.Text = budgetInfo.TypeToString(budgetInfo.rtgsa_type);
            comboBox_Strom.Text = budgetInfo.TypeToString(budgetInfo.strom_type);
            comboBox_TA.Text = budgetInfo.TypeToString(budgetInfo.ta_type);
            comboBox_Vinn.Text = budgetInfo.TypeToString(budgetInfo.vinn_type);
        }

        private void RefreshDataGridView()
        {
            ds = new DataSet();
            con.Open();
            da = new SqlCeDataAdapter("Select * from tblBudgetSelger WHERE BudgetId =" + budgetId, con);
            var cmdBldr = new SqlCeCommandBuilder(da);
            da.Fill(ds, "tblBudgetSelger");
            bindingSource1.DataSource = ds;
            bindingSource1.DataMember = "tblBudgetSelger";
            dataGridView1.DataSource = bindingSource1;
            bindingNavigator1.BindingSource = bindingSource1;
            con.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ToggleUiForEditing();
        }

        private void ValidateTextInputAsNumeric(TextBox control, bool decimals = false)
        {
            decimal d = 0;
            string str = control.Text;

            char[] arr = str.ToCharArray();

            arr = Array.FindAll<char>(arr, (c => (char.IsDigit(c) || c == '-' || c == ',')));
            str = new string(arr);

            if (!Decimal.TryParse(str, NumberStyles.Any, FormMain.norway, out d))
                control.Text = "";
            if (d != 0)
            {
                if (decimals)
                    control.Text = d.ToString("#,##0.00");
                else
                    control.Text = d.ToString("#,##0");
            }
            else
                control.Text = "";
        }

        private void ToggleUiForEditing()
        {
            bool boo = !textBox_omsetning.ReadOnly;

            dateTime_date.Visible = !boo;
            textBox_date.Visible = boo;

            textBox_date.ReadOnly = boo;
            textBox_omsetning.ReadOnly = boo;
            textBox_inntjening.ReadOnly = boo;
            textBox_margin.ReadOnly = boo;
            textBox_ta.ReadOnly = boo;
            textBox_strom.ReadOnly = boo;
            textBox_finans.ReadOnly = boo;
            textBox_rtgsa.ReadOnly = boo;
            textBox_acc.ReadOnly = boo;
            textBox_vinn.ReadOnly = boo;

            comboBox_Acc.Enabled = !boo;
            comboBox_TA.Enabled = !boo;
            comboBox_Strom.Enabled = !boo;
            comboBox_Rtgsa.Enabled = !boo;
            comboBox_Finans.Enabled = !boo;
            comboBox_Vinn.Enabled = !boo;

            if (!boo)
                SetUI();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SaveDb();
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
                da.Update(ds, "tblBudgetSelger");
                con.Close();

                RefreshDataGridView();
                return true;
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil oppstod under lagring av budsjett.", ex, "Sjekk om alle påkrevde felt er utfylt.\n\nException: ");
                errorMsg.ShowDialog(this);
            }
            return false;
        }

        private void buttonLagre_Click(object sender, EventArgs e)
        {
            SaveBudget();
            ToggleUiForEditing();
        }

        private void SaveBudget()
        {
            BudgetCategory kategori = budgetInfo.kategori;
            DateTime date = FormMain.GetLastDayOfMonth(dateTime_date.Value);
            int dager = Convert.ToInt32(textBox_dager.Text);
            decimal omsetning = Convert.ToDecimal(textBox_omsetning.Text);
            decimal inntjening = Convert.ToDecimal(textBox_inntjening.Text);
            decimal budget_margin = Convert.ToDecimal(textBox_margin.Text);
            decimal ta = 0, strom = 0, rtgsa = 0, finans = 0, acc = 0, vinn = 0;
            string ta_type = "", strom_type = "", rtgsa_type = "", finans_type = "", acc_type = "", vinn_type = "";

            ValidateTextInputAsNumeric(textBox_ta, true);
            ValidateTextInputAsNumeric(textBox_strom, true);
            ValidateTextInputAsNumeric(textBox_finans, true);
            ValidateTextInputAsNumeric(textBox_rtgsa, true);
            ValidateTextInputAsNumeric(textBox_acc, true);
            ValidateTextInputAsNumeric(textBox_margin, true);
            ValidateTextInputAsNumeric(textBox_omsetning, false);
            ValidateTextInputAsNumeric(textBox_inntjening, false);
            ValidateTextInputAsNumeric(textBox_vinn, true);

            if (comboBox_TA.SelectedIndex >= 0 && textBox_ta.Text.Length > 0)
            {
                ta = Convert.ToDecimal(textBox_ta.Text);
                ta_type = comboBox_TA.GetItemText(comboBox_TA.SelectedItem);
            }
            if (comboBox_Strom.SelectedIndex >= 0 && textBox_strom.Text.Length > 0)
            {
                strom = Convert.ToDecimal(textBox_strom.Text);
                strom_type = comboBox_Strom.GetItemText(comboBox_Strom.SelectedItem);
            }
            if (comboBox_Finans.SelectedIndex >= 0 && textBox_finans.Text.Length > 0)
            {
                finans = Convert.ToDecimal(textBox_finans.Text);
                finans_type = comboBox_Finans.GetItemText(comboBox_Finans.SelectedItem);
            }
            if (comboBox_Rtgsa.SelectedIndex >= 0 && textBox_rtgsa.Text.Length > 0)
            {
                rtgsa = Convert.ToDecimal(textBox_rtgsa.Text);
                rtgsa_type = comboBox_Rtgsa.GetItemText(comboBox_Rtgsa.SelectedItem);
            }
            if (comboBox_Acc.SelectedIndex >= 0 && textBox_acc.Text.Length > 0)
            {
                acc = Convert.ToDecimal(textBox_acc.Text);
                acc_type = comboBox_Acc.GetItemText(comboBox_Acc.SelectedItem);
            }
            if (comboBox_Vinn.SelectedIndex >= 0 && textBox_vinn.Text.Length > 0)
            {
                vinn = Convert.ToDecimal(textBox_vinn.Text);
                vinn_type = comboBox_Vinn.GetItemText(comboBox_Vinn.SelectedItem);
            }

            if (budget.AddBudget(appConfig.Avdeling, kategori, date, dager, omsetning, inntjening, budget_margin, ta, ta_type, strom, strom_type, finans, finans_type, rtgsa, rtgsa_type, acc, acc_type, vinn, vinn_type, true, budgetInfo.budget_id) == -1)
            {
                MessageBox.Show("Budsjett ble ikke lagret!\nSe logg for detaljer.");
                return;
            }
        }

        private void OpenBudgetTimer()
        {
            if (budgetId >= 0)
            {
                var form = new FormBudgetTimer(appConfig, budgetInfo, budgetId);
                if (form.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
                {
                    budget.SumWorkHoursAndDays(budgetInfo);
                }

                RefreshDataGridView();
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            OpenBudgetTimer();
        }
    }
}
