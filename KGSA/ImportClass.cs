using FileHelpers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace KGSA
{
    public partial class ImportClass : Form
    {
        delegate void SetTextCallback(string text);
        private const int CS_DROPSHADOW = 0x00020000;
        AppSettings appConfig = new AppSettings();
        public string message = "";
        private SqlCeConnection con;
        private DataSet ds;
        private Timer tm;

        public ImportClass()
        {
            LoadSettings();
            InitializeComponent();
            bgWorkerImport.RunWorkerAsync();
            tm = new Timer();
            tm.Interval = 2 * 1000;
            tm.Tick += new EventHandler(tm_Tick);
        }

        private void tm_Tick(object sender, EventArgs e)
        {
            tm.Stop();
            if (!message.ToLower().Contains("feil"))
                this.DialogResult = DialogResult.OK;
            this.Close();
        }

        public void LoadSettings()
        {
            if (!File.Exists(Form1.settingsFile))
            {
                return;
            }

            XmlSerializer mySerializer = new XmlSerializer(typeof(AppSettings));
            using (StreamReader myXmlReader = new StreamReader(Form1.settingsFile))
            {
                try
                {
                    appConfig = (AppSettings)mySerializer.Deserialize(myXmlReader);
                }
                catch
                {
                    Message("Feil ved lasting av program innstillinger.");
                }
            }
        }

        public void Message(string str)
        {
            try
            {

                str = str.Trim();
                str = str.Replace("\n", " ");
                str = str.Replace("\r", String.Empty);
                str = str.Replace("\t", String.Empty);

                if (this.labelLog.InvokeRequired)
                {
                    SetTextCallback d = new SetTextCallback(Message);
                    this.Invoke(d, new object[] { str });
                }
                else
                {
                    this.labelLog.Text = str;
                }
                message = str;
            }
            catch
            {
                message = str;
            }
        }

        public void ImportAllCsv()
        {
            string fileToImport = appConfig.csvFile;
            if (Form1.browseCSV != "")
            {
                fileToImport = Form1.browseCSV;
                Form1.browseCSV = "";
            }
            else
            {
                try
                {
                    if (!File.Exists(appConfig.csvFile))
                        Message("Import feil: Fant ikke CSV fil eller ingen tilgang.");
                }
                catch
                {
                    Message("Import feil: Fant ikke CSV fil eller ingen tilgang.");
                }
            }
            Message("Importerer fra " + fileToImport + "..");
            try
            {
                try
                {
                    var engine = new FileHelperEngine(typeof(csvImport));
                    var resCSV = engine.ReadFile(fileToImport) as csvImport[];

                    try
                    {
                        if (resCSV.Length > 0)
                        {
                            DateTime dtFirst = DateTime.MaxValue;
                            DateTime dtLast = DateTime.MinValue;
                            int valider = 0;
                            var norway = new CultureInfo("nb-NO");

                            // sjekk for først og siste dato, samt hvilken kategorier den inneholder
                            for (int i = 0; i < resCSV.Length; i++)
                            {
                                DateTime dtTemp = Convert.ToDateTime(resCSV[i].Dato.ToString());
                                if (DateTime.Compare(dtTemp, dtFirst) < 0)
                                {
                                    // Datoen er eldre
                                    dtFirst = dtTemp;
                                }
                                if (DateTime.Compare(dtTemp, dtLast) > 0)
                                {
                                    // Datoen er nyere
                                    dtLast = dtTemp;
                                }
                                if (resCSV[i].Kgm.StartsWith("2") || resCSV[i].Kgm.StartsWith("4") || resCSV[i].Kgm.StartsWith("5"))
                                {
                                    valider++;
                                }
                            }
                            Message("CSV inneholder " + valider + " gyldige transaksjoner, solgt mellom " + dtFirst.ToString("dddd d. MMMM yyyy", norway) + " og " + dtLast.ToString("dddd d. MMMM yyyy", norway) + ".");

                            string strFirst = dtFirst.ToString("yyy-MM-dd");
                            string strLast = dtLast.ToString("yyy-MM-dd");

                            if (valider > 0)
                            {
                                con = new SqlCeConnection(Form1.SqlConStr);
                                con.Open();
                                var command = new SqlCeCommand("DELETE FROM tblSalg WHERE (Dato >= '" + strFirst + "') AND (Dato <= '" + strLast + "')", con);
                                var result = command.ExecuteNonQuery();
                                if (appConfig.debug)
                                    bgWorkerImport.ReportProgress(0, "Slettet " + result + " transaksjoner.");

                                command = new SqlCeCommand("SELECT * FROM tblSalg;", con);
                                var da = new SqlCeDataAdapter(command);

                                ds = new DataSet();
                                da.Fill(ds, "tblSalg");
                                var ca = new SqlCeCommandBuilder(da);

                                Message("Prosesserer " + resCSV.Length.ToString("#,##0") + " transaksjoner..");

                                for (int i = 0; i < resCSV.Length; i++)
                                {
                                    // Ta bare med antall ulik 0
                                    if (resCSV[i].Antall != 0 && !resCSV[i].Kgm.Contains("99999") && appConfig.importAll
                                        ||
                                        (resCSV[i].Antall != 0
                                        &&
                                        (resCSV[i].Kgm.StartsWith("531") || resCSV[i].Kgm.StartsWith("533") ||
                                            resCSV[i].Kgm.StartsWith("580") || resCSV[i].Kgm.StartsWith("534") ||
                                            resCSV[i].Kgm.StartsWith("280") || resCSV[i].Kgm.StartsWith("224") ||
                                            resCSV[i].Kgm.StartsWith("480") || resCSV[i].Kgm.StartsWith("431"))
                                        && !appConfig.importAll)
                                        )
                                    {
                                        // dRow[?]
                                        // 0 = SalgID
                                        // 1 = Selgerkode
                                        // 2 = Varegruppe
                                        // 3 = Varekode
                                        // 4 = Dato
                                        // 5 = Antall
                                        DataRow dRow = ds.Tables["tblSalg"].NewRow();
                                        dRow[1] = resCSV[i].Sk; // Selgerkode
                                        if (!appConfig.importAll)
                                        {
                                            if (resCSV[i].Kgm.StartsWith("531")) // desktops
                                                dRow[2] = 531;
                                            if (resCSV[i].Kgm.StartsWith("533")) // laptops
                                                dRow[2] = 533;
                                            if (resCSV[i].Kgm.StartsWith("534")) // nettbrett
                                                dRow[2] = 534;
                                            if (resCSV[i].Kgm.StartsWith("580")) // tjenester
                                                dRow[2] = 580;
                                            if (resCSV[i].Kgm.StartsWith("280")) // nettbrett
                                                dRow[2] = 280;
                                            if (resCSV[i].Kgm.StartsWith("224")) // tjenester
                                                dRow[2] = 224;
                                            if (resCSV[i].Kgm.StartsWith("480")) // nettbrett
                                                dRow[2] = 480;
                                            if (resCSV[i].Kgm.StartsWith("431")) // tjenester
                                                dRow[2] = 431;
                                        }
                                        else
                                        {
                                            dRow[2] = resCSV[i].Kgm.Substring(0, 3); // Ta med alle kategorier.
                                        }
                                        dRow[3] = resCSV[i].Varenummer;
                                        string varDato = resCSV[i].Dato.ToString();
                                        dRow[4] = Convert.ToDateTime(varDato);
                                        dRow[5] = resCSV[i].Antall;
                                        dRow[6] = resCSV[i].Btokr;
                                        dRow[7] = resCSV[i].Avd;
                                        dRow[8] = resCSV[i].Salgspris;
                                        dRow[9] = resCSV[i].BilagsNr;
                                        ds.Tables["tblSalg"].Rows.Add(dRow);
                                    }
                                }

                                // Send data til SQL server og avslutt forbindelsen
                                da.Update(ds, "tblSalg");
                                con.Close();
                                Message("Importering fullført!");
                            }
                            else
                            {
                                Message("Import feil: Fant ingen gyldige transaksjoner.");
                            }
                        }
                        else
                        {
                            Message("Import feil: Ingen transaksjoner funnet! Kontroller om eksportering er korrekt eller sjekk innstillinger.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Message("Import feil: Unntak ved prosessering av transaksjoner. Feilmelding: " + ex.ToString());
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Feil ved lesing av CSV.\n" + ex.ToString(), "KGSA - Importering avbrutt", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Message("Import feil: Unntak ved lesing av CSV: " + ex.ToString());
                }
            }
            catch (Exception ex)
            {
                Message("Import feil: Ukjent feil ved imprtering. (" + fileToImport + ") Feilmelding: " + ex.ToString());
            }
        }

        protected override CreateParams CreateParams
        {
            get
            {
                // add the drop shadow flag for automatically drawing
                // a drop shadow around the form
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= CS_DROPSHADOW;
                return cp;
            }
        }

        private void bgWorkerImport_DoWork(object sender, DoWorkEventArgs e)
        {
            ImportAllCsv();
        }

        private void bgWorkerImport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string str = e.UserState.ToString();
            Message(str);
        }

        private void bgWorkerImport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            tm.Start();
        }

    }
}
