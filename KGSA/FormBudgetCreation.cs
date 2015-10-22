using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace KGSA
{
    public partial class FormBudgetCreation : Form
    {
        DataSet ds;
        SqlCeDataAdapter da;
        SqlCeConnection con = new SqlCeConnection(FormMain.SqlConStr);
        BudgetObj budget;
        FormMain main;
        string[] tips = new string[2] { "Tips: Margin, Hitrate, SoM og SoB\nskrives med to desimaltegn.\nEks: 0,10 (= 10 %)", "Tips: La feltet være tomt hvis det ikke er ønskelig\n å ta det med i budsjettet." };

        public FormBudgetCreation(FormMain form)
        {
            this.main = form;
            InitializeComponent();
            ImportSettings();
            dateTimePicker_Date.Format = DateTimePickerFormat.Custom;
            dateTimePicker_Date.CustomFormat = "MMMM yyyy";
            dateTimePicker_Date.Value = FormMain.GetLastDayOfMonth(main.appConfig.dbTo);
            comboBox_Acc.SelectedIndex = 0;
            comboBox_Finans.SelectedIndex = 0;
            comboBox_Kategori.SelectedIndex = 0;
            comboBox_Rtgsa.SelectedIndex = 0;
            comboBox_Strom.SelectedIndex = 0;
            comboBox_TA.SelectedIndex = 0;
            comboBox_Vinn.SelectedIndex = 0;
            budget = new BudgetObj(main);
            UpdateDb();
            if (dataGridView1.RowCount != 0)
                tabControl1.SelectedTab = tabPageSettings;

            Random rnd = new Random();
            labelTips.Text = tips[rnd.Next(tips.Length)];
        }

        public void ImportSettings()
        {
            try
            {
                if (main.appConfig.budgetChartPostWidth >= 10 && main.appConfig.budgetChartPostWidth <= 100)
                    numericBudgetChartPostWidth.Value = main.appConfig.budgetChartPostWidth;
                else
                    numericBudgetChartPostWidth.Value = 60;

                if (main.appConfig.budgetChartMinPosts >= 0 && main.appConfig.budgetChartMinPosts <= 50)
                    numericBudgetChartMin.Value = main.appConfig.budgetChartMinPosts;
                else
                    numericBudgetChartMin.Value = 0;

                checkBoxBudgetShowEfficiency.Checked = main.appConfig.budgetChartShowEfficiency;
                checkBoxBudgetShowQuality.Checked = main.appConfig.budgetChartShowQuality;
                checkBoxBudgetIsolateCrossSalesrep.Checked = main.appConfig.budgetIsolateCrossSalesRep;
                checkBoxBudgetInclAllSalesRepUnderCross.Checked = main.appConfig.budgetInclAllSalesRepUnderCross;

                checkBox_ShowMda.Checked = main.appConfig.budgetShowMda;
                checkBox_ShowAv.Checked = main.appConfig.budgetShowAudioVideo;
                checkBox_ShowSda.Checked = main.appConfig.budgetShowSda;
                checkBox_ShowTele.Checked = main.appConfig.budgetShowTele;
                checkBox_ShowData.Checked = main.appConfig.budgetShowData;
                checkBox_ShowCross.Checked = main.appConfig.budgetShowCross;
                checkBox_ShowKasse.Checked = main.appConfig.budgetShowKasse;
                checkBox_ShowAftersales.Checked = main.appConfig.budgetShowAftersales;
                checkBox_ShowMdaSda.Checked = main.appConfig.budgetShowMdasda;
                checkBox_ShowButikk.Checked = main.appConfig.budgetShowButikk;

                // Dagsbudsjett
                checkSettingsDailyInclude.Checked = main.appConfig.dailyBudgetIncludeInQuickRanking;
                checkSettingsDailyAuto.Checked = main.appConfig.dailyBudgetQuickRankingAutoUpdate;
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public void ExportSettings()
        {
            if (numericBudgetChartPostWidth.Value >= 10 && numericBudgetChartPostWidth.Value <= 100)
                main.appConfig.budgetChartPostWidth = numericBudgetChartPostWidth.Value;
            else
                main.appConfig.budgetChartPostWidth = 60;

            if (numericBudgetChartMin.Value >= 0 && numericBudgetChartMin.Value <= 50)
                main.appConfig.budgetChartMinPosts = (int)numericBudgetChartMin.Value;
            else
                main.appConfig.budgetChartMinPosts = 0;

            main.appConfig.budgetChartShowEfficiency = checkBoxBudgetShowEfficiency.Checked;
            main.appConfig.budgetChartShowQuality = checkBoxBudgetShowQuality.Checked;
            main.appConfig.budgetIsolateCrossSalesRep = checkBoxBudgetIsolateCrossSalesrep.Checked;
            main.appConfig.budgetInclAllSalesRepUnderCross = checkBoxBudgetInclAllSalesRepUnderCross.Checked;

            main.appConfig.budgetShowMda = checkBox_ShowMda.Checked;
            main.appConfig.budgetShowAudioVideo = checkBox_ShowAv.Checked;
            main.appConfig.budgetShowSda = checkBox_ShowSda.Checked;
            main.appConfig.budgetShowTele = checkBox_ShowTele.Checked;
            main.appConfig.budgetShowData = checkBox_ShowData.Checked;
            main.appConfig.budgetShowCross = checkBox_ShowCross.Checked;
            main.appConfig.budgetShowKasse = checkBox_ShowKasse.Checked;
            main.appConfig.budgetShowAftersales = checkBox_ShowAftersales.Checked;
            main.appConfig.budgetShowMdasda = checkBox_ShowMdaSda.Checked;
            main.appConfig.budgetShowButikk = checkBox_ShowButikk.Checked;

            // Dagsbudsjett
            main.appConfig.dailyBudgetIncludeInQuickRanking = checkSettingsDailyInclude.Checked;
            main.appConfig.dailyBudgetQuickRankingAutoUpdate = checkSettingsDailyAuto.Checked;

            main.SaveSettings();
        }

        private void UpdateDb()
        {
            ds = new DataSet();
            con.Open();

            da = new SqlCeDataAdapter("Select Id, Kategori, Date, Omsetning, Inntjening, Margin, Updated from tblBudget WHERE Avdeling = '" + main.appConfig.Avdeling + "'", con);
            var cmdBldr = new SqlCeCommandBuilder(da);
            da.Fill(ds, "tblBudget");
            bindingSource1.DataSource = ds;
            bindingSource1.DataMember = "tblBudget";
            dataGridView1.DataSource = bindingSource1;
            bindingNavigator1.BindingSource = bindingSource1;
            con.Close();

            if (dataGridView1.Rows.Count > 0)
                toolStripButtonOpen.Enabled = true;
            else
                toolStripButtonOpen.Enabled = false;

            ColorRows();
        }

        private void buttonAddBudget_Click(object sender, EventArgs e)
        {
            if (!CheckInputBoxes()) // Sjekk om alt er OK
            {
                MessageBox.Show("Mangler felt!");
                return;
            }

            AddBudget();


            UpdateDb();
        }

        private void SaveDatagrid()
        {
            try
            {
                if (dataGridView1.CurrentRow != null)
                    dataGridView1.CurrentRow.DataGridView.EndEdit();
                dataGridView1.EndEdit();
                bindingSource1.EndEdit();

                con.Open();
                da.Update(ds, "tblBudget");
                con.Close();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void AddBudget()
        {
            try
            {
                BudgetCategory kategori = BudgetCategoryClass.NameToType(comboBox_Kategori.GetItemText(comboBox_Kategori.SelectedItem));
                DateTime date = FormMain.GetLastDayOfMonth(dateTimePicker_Date.Value);

                int dager = 0;
                int.TryParse(textBox_Dager.Text, out dager);

                decimal omsetning = 0;
                decimal.TryParse(textBox_Omsetning.Text, out omsetning);

                decimal inntjening = 0;
                decimal.TryParse(textBox_Inntjening.Text, out inntjening);

                decimal budget_margin = 0;
                decimal.TryParse(textBox_Margin.Text, out budget_margin);

                decimal ta = 0, strom = 0, rtgsa = 0, finans = 0, acc = 0, vinn = 0;
                string ta_type = "", strom_type = "", rtgsa_type = "", finans_type = "", acc_type = "", vinn_type = "";

                if (comboBox_TA.SelectedIndex >= 0 && textBox_TA.Text.Length > 0)
                {
                    decimal.TryParse(textBox_TA.Text, out ta);
                    ta_type = comboBox_TA.GetItemText(comboBox_TA.SelectedItem);
                }
                if (comboBox_Strom.SelectedIndex >= 0 && textBox_Strom.Text.Length > 0)
                {
                    decimal.TryParse(textBox_Strom.Text, out strom);
                    strom_type = comboBox_Strom.GetItemText(comboBox_Strom.SelectedItem);
                }
                if (comboBox_Finans.SelectedIndex >= 0 && textBox_Finans.Text.Length > 0)
                {
                    decimal.TryParse(textBox_Finans.Text, out finans);
                    finans_type = comboBox_Finans.GetItemText(comboBox_Finans.SelectedItem);
                }
                if (comboBox_Rtgsa.SelectedIndex >= 0 && textBox_Rtgsa.Text.Length > 0)
                {
                    decimal.TryParse(textBox_Rtgsa.Text, out rtgsa);
                    rtgsa_type = comboBox_Rtgsa.GetItemText(comboBox_Rtgsa.SelectedItem);
                }
                if (comboBox_Acc.SelectedIndex >= 0 && textBox_Acc.Text.Length > 0)
                {
                    decimal.TryParse(textBox_Acc.Text, out acc);
                    acc_type = comboBox_Acc.GetItemText(comboBox_Acc.SelectedItem);
                }
                if (comboBox_Vinn.SelectedIndex >= 0 && textBox_Vinn.Text.Length > 0)
                {
                    decimal.TryParse(textBox_Vinn.Text, out vinn);
                    vinn_type = comboBox_Vinn.GetItemText(comboBox_Vinn.SelectedItem);
                }

                int id = budget.AddBudget(main.appConfig.Avdeling, kategori, date, dager, omsetning, inntjening, budget_margin, ta, ta_type, strom, strom_type, finans, finans_type, rtgsa, rtgsa_type, acc, acc_type, vinn, vinn_type);
                if (id == -1)
                {
                    MessageBox.Show("Budsjett ble ikke lagret!\nSe logg for detaljer.");
                    return;
                }
                else
                {
                    Logg.Log("Budsjett opprettet under avdeling " + kategori, Color.Green);
                    var form = new FormBudgetDetails(main.appConfig, budget, id);
                    form.ShowDialog(this);
                }

                Nullstill();
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
                MessageBox.Show("Noe var feil med de utfylte tallene. Sjekk over igjen før du forsøker å opprette nytt budsjett.", "KGSA - Feil", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private bool CheckInputBoxes()
        {
            if (comboBox_Kategori.SelectedIndex < 0)
            {
                Logg.Log("Avdeling ikke valgt!", Color.Red);
                return false;
            }

            try
            {
                RemoveSpaces(textBox_Acc);
                RemoveSpaces(textBox_Dager);
                RemoveSpaces(textBox_Finans);
                RemoveSpaces(textBox_Inntjening);
                RemoveSpaces(textBox_Margin);
                RemoveSpaces(textBox_Omsetning);
                RemoveSpaces(textBox_Rtgsa);
                RemoveSpaces(textBox_Strom);
                RemoveSpaces(textBox_TA);
                RemoveSpaces(textBox_Vinn);
            }
            catch
            {
                Logg.Log("Format feil. Sjekk alle bokser for feil.", Color.Red);
                return false;
            }

            return true;
        }

        private void Nullstill()
        {
            comboBox_Kategori.SelectedIndex = 0;
            dateTimePicker_Date.Value = DateTime.Now;
            textBox_Omsetning.Text = "";
            textBox_Inntjening.Text = "";
            textBox_Margin.Text = "";
            textBox_TA.Text = "";
            textBox_Strom.Text = "";
            textBox_Rtgsa.Text = "";
            textBox_Finans.Text = "";
            textBox_Acc.Text = "";
            textBox_Vinn.Text = "";
            comboBox_Acc.SelectedIndex = 0;
            comboBox_Finans.SelectedIndex = 0;
            comboBox_Kategori.SelectedIndex = 0;
            comboBox_Rtgsa.SelectedIndex = 0;
            comboBox_Strom.SelectedIndex = 0;
            comboBox_TA.SelectedIndex = 0;
            comboBox_Vinn.SelectedIndex = 0;
        }

        private void buttonNullstill_Click(object sender, EventArgs e)
        {
            Nullstill();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            SaveDatagrid();
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            OpenBudgetDetails();
        }

        private void OpenBudgetDetails()
        {
            int b = GetCurrentlySelectedBudgetId();

            if (b >= 0)
            {
                var form = new FormBudgetDetails(main.appConfig, budget, b);
                form.ShowDialog(this);

                UpdateDb();
            }

        }

        private int GetCurrentlySelectedBudgetId()
        {
            if (dataGridView1 != null)
                if (dataGridView1.SelectedRows != null)
                    if (dataGridView1.SelectedRows.Count > 0)
                        return Convert.ToInt32(dataGridView1.SelectedRows[0].Cells["Id"].Value.ToString());
                        
            return -1;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            OpenBudgetDetails();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            OpenBudgetDetails();
        }

        private void dateTimePicker_Date_ValueChanged(object sender, EventArgs e)
        {
            textBox_Dager.Text = GetNumberOfOpenDaysInMonth(FormMain.GetFirstDayOfMonth(dateTimePicker_Date.Value), FormMain.GetLastDayOfMonth(dateTimePicker_Date.Value)).ToString();
        }

        public int GetNumberOfOpenDaysInMonth(DateTime from, DateTime to)
        {
            int days = 0;
            foreach (DateTime day in EachDay(from, to))
            {
                if (day.DayOfWeek != DayOfWeek.Sunday && !(
                    day.Date == new DateTime(from.Year, 1, 1) ||
                    day.Date == new DateTime(from.Year, 5, 1) ||
                    day.Date == new DateTime(from.Year, 5, 17) ||
                    day.Date == new DateTime(from.Year, 12, 25) ||
                    day.Date == new DateTime(from.Year, 12, 26)))
                    days++;
            }
            return days;
        }

        public IEnumerable<DateTime> EachDay(DateTime from, DateTime thru)
        {
            for (var day = from.Date; day.Date <= thru.Date; day = day.AddDays(1))
                yield return day;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            ExportSettings();
        }

        private void FormatTextKeyEnter(TextBox box, KeyEventArgs e, BudgetValueType type)
        {
            if (e.KeyCode == Keys.Enter)
            {
                FormatTextInput(box, type);

                if (box.Text.Length > 0)
                {
                    box.SelectionStart = box.Text.Length;
                    box.SelectionLength = 0;
                }

                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void FormatTextInput(TextBox box, BudgetValueType type)
        {
            decimal value = 0;
            if (!decimal.TryParse(box.Text, out value))
            {
                box.Text = "";
                return;
            }

            if (type == BudgetValueType.Omsetning || type == BudgetValueType.Inntjening)
                box.Text = Math.Round(value, 0).ToString("#,##0");
            else if (type == BudgetValueType.SoM || type == BudgetValueType.SoB)
                box.Text = Math.Round(value, 2).ToString("#,##0.00");
            else
                box.Text = value.ToString();
        }

        private void RemoveSpaces(TextBox control)
        {
            decimal d = 0;
            string str = control.Text;

            char[] arr = str.ToCharArray();

            arr = Array.FindAll<char>(arr, (c => (char.IsDigit(c) || c == '-' || c == ',')));
            str = new string(arr);

            if (!Decimal.TryParse(str, NumberStyles.Any, FormMain.norway, out d))
                control.Text = "";
            if (d != 0)
                control.Text = d.ToString();
            else
                control.Text = "";
        }

        private void textBox_Omsetning_Leave(object sender, EventArgs e)
        {
            FormatTextInput((TextBox)sender, BudgetValueType.Omsetning);
        }

        private void textBox_Omsetning_Enter(object sender, EventArgs e)
        {
            RemoveSpaces((TextBox)sender);
        }

        private void textBox_Omsetning_KeyDown(object sender, KeyEventArgs e)
        {
            FormatTextKeyEnter((TextBox)sender, e, BudgetValueType.Omsetning);
        }

        private void textBox_Inntjening_Enter(object sender, EventArgs e)
        {
            RemoveSpaces((TextBox)sender);
        }

        private void textBox_Inntjening_KeyDown(object sender, KeyEventArgs e)
        {
            FormatTextKeyEnter((TextBox)sender, e, BudgetValueType.Inntjening);
        }

        private void textBox_Inntjening_Leave(object sender, EventArgs e)
        {
            FormatTextInput((TextBox)sender, BudgetValueType.Inntjening);
        }

        private void textBox_Margin_Leave(object sender, EventArgs e)
        {
            FormatTextInput((TextBox)sender, BudgetValueType.SoM);
        }

        private void textBox_Margin_KeyDown(object sender, KeyEventArgs e)
        {
            FormatTextKeyEnter((TextBox)sender, e, BudgetValueType.SoM);
        }

        private void textBox_Margin_Enter(object sender, EventArgs e)
        {
            RemoveSpaces((TextBox)sender);
        }

        private void dataGridView1_Paint(object sender, PaintEventArgs e)
        {
            ColorRows();
        }

        private void ColorRows()
        {
            if (dataGridView1.Rows.Count > 0)
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (!(Convert.ToDateTime(row.Cells[3].Value).Month == main.appConfig.dbTo.Month && Convert.ToDateTime(row.Cells[3].Value).Year == main.appConfig.dbTo.Year))
                    {
                        row.DefaultCellStyle.ForeColor = Color.Gray;
                        this.Update();
                    }
                }
            }
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            DeleteSelectedBudget();
        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            DeleteSelectedBudget();
        }

        private void DeleteSelectedBudget()
        {
            int id = GetCurrentlySelectedBudgetId();
            if (id != -1)
                if (MessageBox.Show("Er du sikker på at du vil slette markert budsjett?", "Sletting av budsjett", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
                {
                    budget.DeleteBudget(id);
                    Logg.Log("Budsjett med id " + id + " slettet.");
                    UpdateDb();
                }
        }
    }
}