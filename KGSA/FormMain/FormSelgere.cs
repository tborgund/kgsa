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
    partial class FormMain
    {
        private void AddSelgerkoderAuto()
        {
            try
            {
                if (MessageBox.Show("KGSA vil forsøke å legge selgerkoder til i riktig avdeling utifra salgsmønster.\nDette vil overskrive eksisterende selgerkoder!\n\nEr du sikker?", "KGSA - Informasjon", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2) != System.Windows.Forms.DialogResult.Yes)
                    return;

                processing.SetVisible = true;
                Log.Status("Prosesserer..");
                toolStripComboBoxSkFilter.SelectedIndex = 0;
                salesCodes.Clear();
                OppdaterSelgerkoder();
                this.Update();

                if (appConfig.importSetting.StartsWith("Full"))
                {
                    processing.SetValue = 5;
                    Log.Status("Legger til [Kasse] selgere..");
                    processing.SetText = "Legger til [Kasse] selgere..";
                    using (DataTable dt = database.GetSqlDataTable("SELECT TOP 3 Selgerkode, SUM(Antall) AS Antall FROM tblSalg WHERE (Avdeling = '"
                        + appConfig.Avdeling.ToString() + "') AND (Dato >= '" + appConfig.dbTo.AddMonths(-1).ToString("yyy-MM-dd") + "') AND (Dato <= '"
                        + appConfig.dbTo.ToString("yyy-MM-dd") + "') GROUP BY Selgerkode ORDER BY Antall DESC"))
                    {
                        List<string> selgere = new List<string>();
                        if (dt.Rows.Count > 0)
                        {
                            for (int i = 0; i < dt.Rows.Count; i++)
                                selgere.Add(dt.Rows[i][0].ToString().Trim());

                            for (int i = 0; i < selgere.Count; i++)
                                if (selgere[i] != "INT")
                                    salesCodes.Add(selgere[i], "Kasse", "Selger");
                        }
                    }
                }
                processing.SetValue = 15;
                var dataselgere = "";
                Log.Status("Legger til [Data] selgere..");
                processing.SetText = "Legger til [Data] selgere..";
                using (DataTable dt = database.GetSqlDataTable("SELECT TOP 12 Selgerkode, SUM(Antall) AS Antall FROM tblSalg "
                    + " WHERE (Varegruppe = '534' OR Varegruppe = '533') AND (Avdeling = '" + appConfig.Avdeling.ToString()
                    + "') AND (Dato >= '" + appConfig.dbTo.AddMonths(-1).ToString("yyy-MM-dd") + "') AND (Dato <= '"
                    + appConfig.dbTo.ToString("yyy-MM-dd") + "') GROUP BY Selgerkode ORDER BY Antall DESC"))
                {
                    List<string> selgere = new List<string>();
                    if (dt.Rows.Count > 0)
                    {
                        decimal most = (int)dt.Rows[0][1];
                        decimal divide = 6;
                        decimal compare = Math.Round(most / divide, 0);
                        for (int i = 0; i < dt.Rows.Count; i++)
                            if ((int)dt.Rows[i][1] >= compare)
                                selgere.Add(dt.Rows[i][0].ToString().Trim());

                        for (int i = 0; i < selgere.Count; i++)
                            if (selgere[i] != "INT")
                            {
                                salesCodes.Add(selgere[i], "Data", "Selger");
                                dataselgere += selgere[i];
                            }
                    }
                }
                processing.SetValue = 30;
                Log.Status("Legger til [AudioVideo] selgere..");
                processing.SetText = "Legger til [AudioVideo] selgere..";
                using (DataTable dt = database.GetSqlDataTable("SELECT TOP 12 Selgerkode, SUM(Antall) AS Antall FROM tblSalg "
                    + " WHERE (Varegruppe = '224') AND (Avdeling = '" + appConfig.Avdeling.ToString() + "') AND (Dato >= '"
                    + appConfig.dbTo.AddMonths(-1).ToString("yyy-MM-dd") + "') AND (Dato <= '" + appConfig.dbTo.ToString("yyy-MM-dd") + "') GROUP BY Selgerkode ORDER BY Antall DESC"))
                {
                    List<string> selgere = new List<string>();
                    if (dt.Rows.Count > 0)
                    {
                        decimal most = (int)dt.Rows[0][1];
                        decimal divide = 6;
                        decimal compare = Math.Round(most / divide, 0);
                        for (int i = 0; i < dt.Rows.Count; i++)
                            if ((int)dt.Rows[i][1] >= compare)
                                selgere.Add(dt.Rows[i][0].ToString().Trim());

                        for (int i = 0; i < selgere.Count; i++)
                            if (selgere[i] != "INT")
                                salesCodes.Add(selgere[i], "AudioVideo", "Selger");
                    }
                }
                processing.SetValue = 45;
                Log.Status("Legger til [Tele] selgere..");
                processing.SetText = "Legger til [Tele] selgere..";
                using (DataTable dt = database.GetSqlDataTable("SELECT TOP 12 Selgerkode, SUM(Antall) AS Antall "
                    + " FROM tblSalg WHERE (Varegruppe = '431') AND (Avdeling = '" + appConfig.Avdeling.ToString()
                    + "') AND (Dato >= '" + appConfig.dbTo.AddMonths(-1).ToString("yyy-MM-dd") + "') AND (Dato <= '"
                    + appConfig.dbTo.ToString("yyy-MM-dd") + "') GROUP BY Selgerkode ORDER BY Antall DESC"))
                {
                    List<string> selgere = new List<string>();
                    if (dt.Rows.Count > 0)
                    {
                        decimal most = (int)dt.Rows[0][1];
                        decimal divide = 6;
                        decimal compare = Math.Round(most / divide, 0);
                        for (int i = 0; i < dt.Rows.Count; i++)
                            if ((int)dt.Rows[i][1] >= compare)
                                selgere.Add(dt.Rows[i][0].ToString().Trim());

                        for (int i = 0; i < selgere.Count; i++)
                            if (selgere[i] != "INT")
                                salesCodes.Add(selgere[i], "Tele", "Selger");
                    }
                }
                processing.SetValue = 60;
                Log.Status("Legger til [Data] selgere..");
                processing.SetText = "Legger til [Data] selgere..";
                using (DataTable dt = database.GetSqlDataTable("SELECT TOP 12 Selgerkode, SUM(Antall) AS Antall FROM tblSalg WHERE "
                    + " (Varegruppe = '580') AND (Avdeling = '" + appConfig.Avdeling.ToString() + "') AND (Dato >= '"
                    + appConfig.dbTo.AddMonths(-1).ToString("yyy-MM-dd") + "') AND (Dato <= '"
                    + appConfig.dbTo.ToString("yyy-MM-dd") + "') GROUP BY Selgerkode ORDER BY Antall DESC"))
                {
                    List<string> selgere = new List<string>();
                    if (dt.Rows.Count > 0)
                    {
                        decimal most = (int)dt.Rows[0][1];
                        decimal divide = 6;
                        decimal compare = Math.Round(most / divide, 0);
                        for (int i = 0; i < dt.Rows.Count; i++)
                            if ((int)dt.Rows[i][1] >= compare)
                                selgere.Add(dt.Rows[i][0].ToString().Trim());

                        //for (int i = 0; i < selgere.Count; i++) // Legg til første match av potensielle teknikere som "TeknikerAlle"
                        //{
                        //    if (!dataselgere.Contains(selgere[i]))
                        //    {
                        //        sKoder.Add(selgere[i], "Teknikere", "TeknikerAlle");
                        //        dataselgere += selgere[i];
                        //        break;
                        //    }
                        //}

                        for (int i = 0; i < selgere.Count; i++) // Legg til resten av potensielle teknikere
                            if (!dataselgere.Contains(selgere[i]))
                                salesCodes.Add(selgere[i], "Teknikere", "Selger");
                    }
                }
                processing.SetValue = 75;
                if (appConfig.importSetting.StartsWith("Full"))
                {
                    Log.Status("Legger til [SDA] og [MDA] selgere..");
                    processing.SetText = "Legger til [SDA] og [MDA] selgere..";
                    using (DataTable dtMda = database.GetSqlDataTable("SELECT TOP 20 Selgerkode, SUM(Antall) AS Antall "
                        + " FROM tblSalg WHERE (Varegruppe LIKE '3%') AND (Avdeling = '" + appConfig.Avdeling.ToString()
                        + "') AND (Dato >= '" + appConfig.dbTo.AddMonths(-1).ToString("yyy-MM-dd") + "') AND (Dato <= '"
                        + appConfig.dbTo.ToString("yyy-MM-dd") + "') GROUP BY Selgerkode ORDER BY Antall DESC"))
                    {
                        DataTable dtSda = database.GetSqlDataTable("SELECT TOP 20 Selgerkode, SUM(Antall) AS Antall "
                            + " FROM tblSalg WHERE (Varegruppe LIKE '1%') AND (Avdeling = '" + appConfig.Avdeling.ToString()
                            + "') AND (Dato >= '" + appConfig.dbTo.AddMonths(-1).ToString("yyy-MM-dd") + "') AND (Dato <= '"
                            + appConfig.dbTo.ToString("yyy-MM-dd") + "') GROUP BY Selgerkode ORDER BY Antall DESC");
                        List<string> selgereMda = new List<string>();
                        List<string> selgereSda = new List<string>();
                        if (dtMda.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtSda.Rows.Count; i++)
                            {
                                for (int d = 0; d < dtMda.Rows.Count; d++)
                                    if ((string)dtSda.Rows[i][0] == (string)dtMda.Rows[d][0])
                                    {
                                        if ((int)dtSda.Rows[i][1] >= (int)dtMda.Rows[d][1])
                                            selgereMda.Add(dtSda.Rows[i][0].ToString().Trim());
                                        else
                                            selgereSda.Add(dtMda.Rows[d][0].ToString().Trim());
                                    }
                            }

                            for (int i = 0; i < selgereMda.Count; i++)
                                if (selgereMda[i] != "INT")
                                    salesCodes.Add(selgereMda[i], "MDA", "Selger");

                            for (int i = 0; i < selgereSda.Count; i++)
                                if (selgereSda[i] != "INT")
                                    salesCodes.Add(selgereSda[i], "SDA", "Selger");
                        }
                    }
                }
                processing.SetValue = 99;
                ClearHash(); // med fra om oppdatering av ranking
                OppdaterSelgerkoder();
                Log.n("Fullført automatisk utfylling av selgerkoder. Kontroller selgerkode listen!", Color.Green);
                processing.SetText = "Ferdig!";
                processing.HideDelayed();
                this.Activate();
                if (appConfig.importSetting.StartsWith("Full"))
                    MessageBox.Show("Automatisk utfylling av selgerkoder fullført.\n\nSelgere er blitt tilknyttet " +
                        "en kategori basert på en analyse av transaksjoner den siste måneden.\nKategorien [Aftersales] må " +
                        "velges manuelt, samt [Kasse], [MDA] og [SDA] er vanskelig å finne automatisk.\n" +
                        "Sjekk over selgerkodene og manuelt legg de til i riktig avdeling.",
                        "KGSA - Informasjon", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    MessageBox.Show("Automatisk utfylling av selgerkoder fullført.\n\nSelgere er blitt tilknyttet " +
                        "en kategori basert på en analyse av transaksjoner den siste måneden.\nKategoriene [Aftersales], [MDA], [SDA] og [Kasse] må " +
                        "velges manuelt.\n" +
                        "Sjekk over selgerkodene og legg de til i riktig avdeling.",
                        "KGSA - Informasjon", MessageBoxButtons.OK, MessageBoxIcon.Information);

                LagreSelgerkoder();
            }
            catch (Exception ex)
            {
                processing.SetVisible = false;
                FormError errorMsg = new FormError("Feil oppstod under prosessering av selgerkoder.", ex);
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

                var msg = salesCodes.DeleteDuplicates();
                if (!String.IsNullOrEmpty(msg))
                {
                    Log.n(msg, Color.Red);
                    MessageBox.Show(msg, "KGSA - Advarsel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    daSk.Update(dsSk, "tblSelgerkoder");
                    OppdaterSelgerkoder();
                    return;
                }
                daSk.Update(dsSk, "tblSelgerkoder");
                OppdaterSelgerkoder();
                Log.n("Selgerkoder Lagret.", Color.Green);
                salesCodes.Update();
                ClearHash();
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil oppstod under lagring av selgerkoder.", ex,  "Sjekk om alle påkrevde felt er utfylt.\n\nException: ");
                errorMsg.ShowDialog(this);
            }
        }

        DataSet dsSk;
        SqlCeDataAdapter daSk;
        private void InitSelgerkoder()
        {
            SelgerkoderPopulateSelgere();

            dsSk = new DataSet();
            daSk = new SqlCeDataAdapter("Select * from tblSelgerkoder WHERE Avdeling = '" + appConfig.Avdeling + "'", connection);
            var cmdBldr = new SqlCeCommandBuilder(daSk);
            daSk.Fill(dsSk, "tblSelgerkoder");
            bindingSourceSk.DataSource = dsSk;
            bindingSourceSk.DataMember = "tblSelgerkoder";
            dataGridViewSk.DataSource = bindingSourceSk;
            bindingNavigatorSk.BindingSource = bindingSourceSk;

            toolStripComboBoxSkFilter.SelectedIndex = 0;
            comboBoxKategorier.SelectedIndex = 0;
        }


        private void SelgerkoderPopulateSelgere()
        {
            try
            {
                listBoxSk.Items.Clear();

                if (selgerkodeList.Count == 0)
                {
                    UpdateSelgerkoderUI();
                }
                else
                {
                    listBoxSk.Items.AddRange(selgerkodeList.ToArray());
                }
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil oppstod under henting av selgerkoder", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void OppdaterSelgerkoder()
        {
            if (dataGridViewSk.CurrentRow == null)
                InitSelgerkoder();

            dsSk.Clear();
            if (toolStripComboBoxSkFilter.SelectedIndex > 0)
                daSk = new SqlCeDataAdapter("Select * from tblSelgerkoder WHERE Avdeling = '" + appConfig.Avdeling + "' AND Kategori = '" + toolStripComboBoxSkFilter.SelectedItem.ToString() + "'", connection);
            else
                daSk = new SqlCeDataAdapter("Select * from tblSelgerkoder WHERE Avdeling = '" + appConfig.Avdeling + "'", connection);
            var cmdBldr = new SqlCeCommandBuilder(daSk);
            daSk.Fill(dsSk, "tblSelgerkoder");
        }

        private ContextMenuStrip listboxContextMenu;

        private void SelgerkodeClick(object sender, EventArgs e)
        {
            try
            {
                var kat = sender.ToString();
                if (tabControlMain.SelectedTab == tabPageTrans)
                {
                    var cellvalue = lastRightClickValue;
                    lastRightClickValue = "";
                    if (!salesCodes.Add(cellvalue.Trim(), kat, "Selger"))
                        Log.n("Kunne ikke legge til selger '" + cellvalue + "' for selgeren finnes allerede.", Color.Red);
                    else
                    {
                        Log.n("Selger '" + cellvalue.Trim() + "' lagt til avdeling '" + kat + "' med provisjon Selger.", Color.Green);
                    }
                }
                else
                {
                    foreach (string element in listBoxSk.SelectedItems)
                    {
                        if (!salesCodes.Add(element.Trim(), kat, "Selger"))
                            Log.n("Kunne ikke legge til selger '" + element + "' for selgeren finnes allerede.", Color.Red);
                        else
                            Log.n("Selger '" + element.Trim() + "' lagt til avdeling '" + kat + "' med provisjon Selger.", Color.Green);
                    }
                    listBoxSk.ClearSelected();
                    LagreSelgerkoder();
                }
                OppdaterSelgerkoder();
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Unntak oppstod under lagring av ny selgerkode.", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void listboxContextMenu_Opening( object sender, CancelEventArgs e )
        {
            // Tøm meny
            listboxContextMenu.Items.Clear();
            // Sjekk om vi har noe valgt..
            var selected = listBoxSk.SelectedItems.Count;
            if (selected > 0)
            {
                ToolStripMenuItem menu = new ToolStripMenuItem();
                menu.Enabled = false;
                if (selected == 1)
                    menu.Text = listBoxSk.SelectedItem.ToString();
                else
                    menu.Text = selected + " valgt";
                listboxContextMenu.Items.Add(menu);

                var menuItem = new ToolStripMenuItem("Legg til..");
                listboxContextMenu.Items.Add(menuItem);
                menuItem.DropDownItems.Add("SDA", null, this.SelgerkodeClick);
                menuItem.DropDownItems.Add("AudioVideo", null, this.SelgerkodeClick);
                menuItem.DropDownItems.Add("MDA", null, this.SelgerkodeClick);
                menuItem.DropDownItems.Add("Tele", null, this.SelgerkodeClick);
                menuItem.DropDownItems.Add("Data", null, this.SelgerkodeClick);
                menuItem.DropDownItems.Add("Teknikere", null, this.SelgerkodeClick);
                menuItem.DropDownItems.Add("Kasse", null, this.SelgerkodeClick);
                menuItem.DropDownItems.Add("Aftersales", null, this.SelgerkodeClick);
                menuItem.DropDownItems.Add("Cross", null, this.SelgerkodeClick);
            }
        }
    }
}
