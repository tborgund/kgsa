using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    public partial class FormVinnprodukt : Form
    {
        private readonly Timer timerMsgClear = new Timer();
        FormMain main;
        DataSet ds;
        SqlCeDataAdapter da;
        SqlCeConnection con = new SqlCeConnection(FormMain.SqlConStr);

        public FormVinnprodukt(FormMain form)
        {
            this.main = form;
            InitializeComponent();
            UpdateGrid();
            toolStripComboBoxFilterKat.SelectedIndex = 0;
            timerMsgClear.Tick += timer;
            Logg.Debug("Vinnprodukter åpnet.");
        }

        public void UpdateGrid()
        {
            ds = new DataSet();
            con.Open();
            if (toolStripComboBoxFilterKat.Text != "Alle" && !String.IsNullOrEmpty(toolStripComboBoxFilterKat.Text))
                da = new SqlCeDataAdapter("SELECT * FROM tblVinnprodukt WHERE Avdeling = " + main.appConfig.Avdeling + " AND Kategori = '" + toolStripComboBoxFilterKat.Text + "'", con);
            else
                da = new SqlCeDataAdapter("SELECT * FROM tblVinnprodukt WHERE Avdeling = " + main.appConfig.Avdeling, con);
            var cmdBldr = new SqlCeCommandBuilder(da);
            da.Fill(ds, "tblVinnprodukt");
            bindingSource1.DataSource = ds;
            bindingSource1.DataMember = "tblVinnprodukt";
            dataGridView1.DataSource = bindingSource1;
            bindingNavigator1.BindingSource = bindingSource1;

            con.Close();

        }

        private bool SaveGrid()
        {
            try
            {
                if (dataGridView1.Rows.Count == 0)
                    return true;

                if (dataGridView1.CurrentRow != null)
                    dataGridView1.CurrentRow.DataGridView.EndEdit();

                dataGridView1.EndEdit();
                bindingSource1.EndEdit();

                con.Open();
                da.Update(ds, "tblVinnprodukt");
                con.Close();
                return true;
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil oppstod under lagring av tabel.", ex, "Sjekk om alle påkrevde felt er utfylt.\n\nException: ");
                errorMsg.ShowDialog(this);
            }
            return false;
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            SendMessage("Feil i cellformatering, sjekk om alle felt er fylt ut riktig.", Color.Red);
            e.Cancel = true;
        }

        private void dataGridView1_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                if (e.Row.IsNewRow)
                {
                    //e.Row.Cells[6].Value = DateTime.Now;

                    e.Row.Cells[7].Value = DateTime.Now.AddYears(-1);
                    e.Row.Cells[6].Value = DateTime.Now.AddYears(1);
                    e.Row.Cells[5].Value = DateTime.Now;
                    e.Row.Cells["Avdeling"].Value = main.appConfig.Avdeling;
                    e.Row.Cells["Kategori"].Value = "Alle";
                }
            }
            catch { }
        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                string cellValue = e.FormattedValue.ToString();

                if (dataGridView1.Rows[e.RowIndex].IsNewRow)
                    return;

                if (e.ColumnIndex == 4 && cellValue.Contains(" "))
                {
                    e.Cancel = true;
                    dataGridView1.Rows[e.RowIndex].ErrorText = "Mellomrom er ikke lov!";
                    return;
                }

                if (e.ColumnIndex == 0)
                    dataGridView1[e.ColumnIndex, e.RowIndex].Value = cellValue.ToUpper();

                if (String.IsNullOrEmpty(cellValue) && (e.ColumnIndex == 4))
                {
                    dataGridView1.Rows[e.RowIndex].ErrorText = "Feltet kan ikke være tomt";
                    SendMessage("Feltet kan ikke være tomt! Eller har du ikke oppdatert databasen?", Color.Red);
                    e.Cancel = true;
                }
            }
            catch
            {
                dataGridView1.Rows[e.RowIndex].ErrorText = "Format feil!";
                e.Cancel = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void AddVarekoder()
        {
            try
            {
                DateTime date = main.appConfig.dbTo;
                if (date.Day < 15)
                    date = date.AddMonths(-1);

                SendMessage("Legger til..");

                DataTable table = main.database.CallMonthTable(date.AddMonths(-1), main.appConfig.Avdeling);

                var rows = table.Select("[Varekode] LIKE 'ELSTROM*' OR [Varegruppe] % 100 = 83 OR [Varegruppe] = 961");
                DataTable searchTable = rows.Any() ? rows.CopyToDataTable() : table.Clone();

                DataView view = new DataView(searchTable);
                DataTable distinctValues = view.ToTable(true, "Varekode");


                List<string> productsList = new List<string>() { };
                foreach (DataRow row in distinctValues.Rows)
                    productsList.Add(row["Varekode"].ToString());

                productsList.Sort();

                foreach (string varekode in productsList)
                {
                    if (main.vinnprodukt.aborted)
                        break;
                    main.vinnprodukt.Add(varekode, "Alle", 1, DateTime.Now.AddYears(1), DateTime.Now.AddYears(-1));
                }

                main.vinnprodukt.Update();

                if (main.vinnprodukt.aborted)
                {
                    SendMessage("Import avbrutt.", Color.Red);
                    main.vinnprodukt.aborted = false;
                }
                else
                    SendMessage("Finans/Strøm/TA lagt til Vinnprodukt listen.");
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void AddRtgsaVarekoder()
        {
            try
            {
                List<VarekodeList> varekoderAlle = main.appConfig.varekoder.ToList();

                SendMessage("Legger til..");

                foreach (VarekodeList varekode in varekoderAlle)
                {
                    if (main.vinnprodukt.aborted)
                        break;
                    main.vinnprodukt.Add(varekode.kode, "Alle", 1, DateTime.Now.AddYears(1), DateTime.Now.AddYears(-1));
                }

                main.vinnprodukt.Update();

                if (main.vinnprodukt.aborted)
                {
                    SendMessage("Import avbrutt.", Color.Red);
                    main.vinnprodukt.aborted = false;
                }
                else
                    SendMessage("RTG/SA lagt til Vinnprodukt listen.");
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Er du sikker på at du vil tømme listen?", "Tømme Vinnprodukter", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
            {
                main.vinnprodukt.ClearAll();
                UpdateGrid();
                SendMessage("Listen tømt.");
            }
        }

        private void toolStripComboBoxFilterKat_DropDownClosed(object sender, EventArgs e)
        {
            UpdateGrid();
        }

        private void finansTAStrømToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddVarekoder();
            UpdateGrid();
        }

        private void rTGSAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddRtgsaVarekoder();
            UpdateGrid();
        }

        private void fraCSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VelgCSV();
            UpdateGrid();
        }

        private void VelgCSV()
        {
            var fdlg = new OpenFileDialog();
            fdlg.Title = "Velg CVS-fil";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "All files (*.*)|*.*|CVS filer (*.csv)|*.csv";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            fdlg.Multiselect = true;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                SendMessage("Importerer CSV, vent litt..");
                this.Update();
                bool complete = main.vinnprodukt.ImportFromCsv(fdlg.FileName);
                UpdateGrid();
                if (complete)
                    SendMessage("Fullført", Color.Green);
                else
                    SendMessage("Avbrutt");
            }

            fdlg.Dispose();
        }

        public void SendMessage(string str, Color? c = null)
        {
            try
            {
                str = str.Trim();
                str = str.Replace("\n", " ");
                str = str.Replace("\r", String.Empty);
                str = str.Replace("\t", String.Empty);
                errorMessage.ForeColor = c.HasValue ? c.Value : Color.Black;
                if (str.Length <= 60)
                    errorMessage.Text = str;
                else
                    errorMessage.Text = str.Substring(0, 60) + Environment.NewLine + str.Substring(60, str.Length - 60);

                ClearMessageTimer();

                if (!String.IsNullOrEmpty(str))
                    Logg.Log(str, c, true);
            }
            catch
            {
                errorMessage.ForeColor = Color.Red;
                errorMessage.Text = "Feil i meldingssystem!" + Environment.NewLine + "Siste feilmelding: " + str;
            }
        }

        private void timer(object sender, EventArgs e)
        {
            ClearMessageTimerStop();
        }

        private void ClearMessageTimer()
        {
            timerMsgClear.Stop();
            timerMsgClear.Interval = 10 * 1000; // 10 sek
            timerMsgClear.Enabled = true;
            timerMsgClear.Start();
        }

        private void ClearMessageTimerStop()
        {
            if (timerMsgClear.Enabled)
            {
                SendMessage("");
                timerMsgClear.Stop();
                timerMsgClear.Enabled = false;
            }
        }

        private void FormVinnprodukt_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.DialogResult == System.Windows.Forms.DialogResult.OK)
                if (!SaveGrid())
                    e.Cancel = true;
        }

        private void slettAlleVarekoderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Er du sikker på at du vil tømme listen?", "Tømme Vinnprodukter", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
            {
                main.vinnprodukt.ClearAll();
                UpdateGrid();
                SendMessage("Listen tømt.");
            }
        }

        private void importerFraCSVToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            VelgCSV();
        }

        private void eksporterTilCSVToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            LagreCSV();
            UpdateGrid();
        }

        private void LagreCSV()
        {
            try
            {
                var fdlg = new SaveFileDialog();
                fdlg.Title = "Lagre Vinnprodukt CSV som..";
                fdlg.InitialDirectory = @"c:\";
                fdlg.FileName = @"kgsa vinnprodukter.csv";
                fdlg.Filter = "All files (*.*)|*.*|CVS filer (*.csv)|*.csv";
                fdlg.FilterIndex = 2;
                fdlg.RestoreDirectory = true;

                if (fdlg.ShowDialog() == DialogResult.OK)
                {
                    SendMessage("Eksporterer CSV..");
                    if (main.vinnprodukt.ExportToCsv(fdlg.FileName))
                        SendMessage("Fullført", Color.Green);
                    else
                        SendMessage("Avbrutt");
                }

                fdlg.Dispose();
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex, true);
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            SaveGrid();
        }

    }
}
