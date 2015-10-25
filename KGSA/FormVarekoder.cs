using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    public partial class FormVarekoder : Form
    {
        private readonly Timer timerMsgClear = new Timer();
        FormMain main;
        public FormVarekoder(FormMain form)
        {
            this.main = form;
            InitializeComponent();
            BindVarekoder();
            timerMsgClear.Tick += timer;
            Log.d("Varekoder åpnet.");
        }

        private void FilterVarekoder()
        {
            if (!String.IsNullOrEmpty(toolStripComboBoxFilter.Text))
            {
                if (toolStripComboBoxFilter.Text == "Alle")
                    bindingSource1.DataSource = main.appConfig.varekoder;
                else
                    bindingSource1.DataSource = main.appConfig.varekoder.Where(item => item.kategori == toolStripComboBoxFilter.Text);
            }
            else
                bindingSource1.DataSource = main.appConfig.varekoder;
        }

        private void BindVarekoder()
        {
            try
            {
                if (main.appConfig.varekoder == null)
                {
                    main.appConfig.varekoder = new List<VarekodeList> { };
                    main.appConfig.varekoder.Clear();
                    main.appConfig.ResetVarekoder(main.appConfig.varekoder);
                }

                bindingSource1.DataSource = main.appConfig.varekoder;
                dataGridView1.DataSource = bindingSource1;
                bindingNavigator1.BindingSource = bindingSource1;
            }
            catch
            {
                Log.n("Feil oppstod ved binding av varekoder til tabel.", Color.Red);
                SendMessage("Feil oppstod ved binding av varekoder til tabel.", Color.Red);
            }
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
                    Log.n(str, c, true);
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
            timerMsgClear.Interval = 10 * 1000; // 10 sek
            timerMsgClear.Enabled = true;
            timerMsgClear.Start();
        }

        private void ClearMessageTimerStop()
        {
            SendMessage("");
            timerMsgClear.Stop();
        }

        private void toolStripButtonTilStandard_Click(object sender, EventArgs e)
        {
            EraseVarekoder();
        }

        private void EraseVarekoder()
        {
            try
            {
                if (Log.Alert("Er du sikker på at du vil slette alle eksisterende\nvarekoder og sette inn standard?", "KGSA - Informasjon", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == System.Windows.Forms.DialogResult.Yes)
                {
                    if (dataGridView1.CurrentRow != null)
                        dataGridView1.CurrentRow.DataGridView.EndEdit();
                    dataGridView1.EndEdit();
                    bindingSource1.EndEdit();

                    main.appConfig.varekoder.Clear();
                    main.appConfig.ResetVarekoder(main.appConfig.varekoder);
                    bindingSource1.ResetBindings(false);
                    dataGridView1.Refresh();

                    SendMessage("Varekoder resatt.", Color.Blue);
                }
            }
            catch (Exception ex)
            {
                SendMessage("Feil oppstod ved resetting av varekoder.", Color.Red);
                Log.Unhandled(ex);
            }
        }

        private void toolStripButtonLagre_Click(object sender, EventArgs e)
        {
            SaveGrid();
        }

        private bool SaveGrid()
        {
            try
            {
                dataGridView1.CurrentRow.DataGridView.EndEdit();

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1[0, i].Value != null)
                    {
                        string varekode = dataGridView1[0, i].Value.ToString();
                        if (varekode.Length != 0)
                        {
                            if (dataGridView1[6, i].Value == null)
                                dataGridView1[6, i].Value = varekode.ToUpper();
                            else
                            {
                                string alias = dataGridView1[6, i].Value.ToString();
                                if (alias.Length == 0)
                                    dataGridView1[6, i].Value = varekode.ToUpper();
                            }
                        }
                    }
                }

                dataGridView1.EndEdit();
                bindingSource1.EndEdit();
                SendMessage("Varekoder lagret", Color.Green);
                return true;
            }
            catch (Exception ex)
            {
                SendMessage("Feil oppstod ved lagring av varekoder.", Color.Red);
                Log.d("Feil oppstod ved lagring av varekoder.", ex);
            }
            return false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void FormVarekoder_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!SaveGrid() && this.DialogResult == System.Windows.Forms.DialogResult.OK)
                e.Cancel = true;
        }

        private void toolStripComboBoxFilter_DropDownClosed(object sender, EventArgs e)
        {
            FilterVarekoder();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            SendMessage("");
        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                string cellValue = e.FormattedValue.ToString();

                if (dataGridView1.Rows[e.RowIndex].IsNewRow)
                    return;

                if ((e.ColumnIndex == 0 || e.ColumnIndex == 6) && cellValue.Contains(" "))
                {
                    e.Cancel = true;
                    SendMessage("Mellomrom er ikke lov!", Color.Red);
                    return;
                }

                if (e.ColumnIndex == 0)
                    dataGridView1[e.ColumnIndex, e.RowIndex].Value = cellValue.ToUpper();
                else if (e.ColumnIndex == 6 && cellValue.Length == 0)
                {
                    string varekode = dataGridView1[0, e.RowIndex].Value.ToString();
                    dataGridView1[e.ColumnIndex, e.RowIndex].Value = varekode.ToUpper();
                }

            }
            catch
            {
                SendMessage("Feil format i celle!", Color.Red);
                e.Cancel = true;
            }
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.ColumnIndex == 1 || e.ColumnIndex == 2)
            {
                SendMessage("Ugyldige tegn.", Color.Red);
                e.ThrowException = false;
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Varekode: Legg inn varekoden som vist i Elguide.\n\nKategori: Bestemmer hvilken ranking varekoden gjelder for." +
                "\n\nSelger prov: Hva selgeren får i provisjon per salg.\n\nTekniker prov: Hva teknikerern får i provisjon per salg." +
                "\n\nSalgspris: Legg inn salgsprisen som oppgitt i Elguide. Er viktig for å vise rabatteringer i rankingen." +
                "\n\nSynlig: Bestemmer om varekoden skal være synlig i rankingen. Hvis den skjules beregnes bare omsetning og inntjeningen, ikke antall." +
                "\n\nHitrate: Bestemmer om varekoden skal taes med i hitrate beregningen for hovedproduktet. Må også være synlig for å telle med." +
                "\n\nAlias: Vises istedet for varekoden. Varekoder med like Alias grupperes sammen.", "Informasjon", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
