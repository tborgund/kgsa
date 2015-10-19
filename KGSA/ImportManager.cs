using ErikEJ.SqlCe;
using FileHelpers;
using System;
using System.Collections;
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
    /// <summary>
    /// A class used for generic importing methods
    /// </summary>
    public class ImportManager
    {
        FormMain main;
        private static string[] strGroupImportNormal = new string[8] { "224", "280", "431", "480", "531", "533", "534", "580" };
        //private static List<int> intGroupImportNormal = new List<int> { 224, 280, 431, 480, 531, 533, 534, 580 };
        //private DateTime importTimerLastValue = DateTime.MinValue;
        //private DateTime importTimerLastValueChanged = DateTime.MinValue;
        private List<string> csvFilesToImport = new List<string>();
        private static decimal numberScale = 145.7M;
        public decimal csvSizeGuess = 0;
        string dateFormat = "dd/MM/yyyy HH:mm:ss";

        public int importReadErrors = 0;
        public int importCount = 0;

        public int returnCode = -1;

        public ImportManager(FormMain form, List<string> csvFiles)
        {
            this.main = form;
            this.csvFilesToImport = csvFiles;
        }

        /// <summary>
        /// Metode for imprtering av Elguide transaksjoner fra Program 137
        /// </summary>
        /// <param name="bw">Backgroundworker</param>
        /// <param name="unattended">Vil ikke bruke dialogbokser ved unattended kjøring</param>
        /// <returns>Success or not</returns>
        public bool DoImportTransactions(BackgroundWorker bw, bool unattended)
        {
            if (bw == null)
                throw new ArgumentNullException("Backgroundworker er null!");
            if (csvFilesToImport.Count == 0)
                csvFilesToImport.Add(main.appConfig.csvElguideExportFolder + @"\irank.csv");

            this.returnCode = 0;
            this.importReadErrors = 0;
            //this.importTimerLastValue = DateTime.MinValue;
            //this.importTimerLastValueChanged = DateTime.MinValue;
            csvImport[] csvTransAll = new csvImport[] { };

            main.processing.SetValue = 0;

            try
            {
                // Les inn alle CSV filene til csvTransAll..
                foreach (String file in csvFilesToImport)
                {
                    if (bw.CancellationPending) {
                        this.returnCode = 2;
                        return true;
                    }

                    if (!File.Exists(file))
                    {
                        Logg.Log("Finner ikke fil " + file, Color.Red);
                        if (unattended) {
                            if (MessageBox.Show("Finner ikke fil " + file + "\n\nKontroller import lokasjon!",
                                "KGSA - Finner ikke fil", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                                == System.Windows.Forms.DialogResult.Cancel)
                                continue;
                        }
                        else
                            this.returnCode = 9;
                    }

                    FileInfo f = new FileInfo(file);
                    decimal s1 = f.Length;
                    this.csvSizeGuess = Math.Round(s1 / numberScale, 2);

                    Logg.Log("Leser CSV fra Elguide '" + file + "'..");
                    var csvTrans = ReadCsv(file, bw, unattended);
                    if (csvTrans == null)
                    {
                        this.returnCode = 9;
                        return false;
                    }

                    int oldLength = csvTransAll.Length;
                    Array.Resize<csvImport>(ref csvTransAll, oldLength + csvTrans.Length);
                    Array.Copy(csvTrans, 0, csvTransAll, oldLength, csvTrans.Length);

                    Logg.Debug("Lest ferdig '" + file + "' - Totalt: " + csvTrans.Length);
                }

                if (bw.CancellationPending) {
                    this.returnCode = 2;
                    return false;
                }

                if (csvTransAll.Length < 25) {
                    Logg.Log("Import: Ingen eller for få transaksjoner funnet!"
                        + "Kontroller om eksportering er korrekt eller sjekk innstillinger.", Color.Red);
                    {
                        this.returnCode = 4;
                        return false;
                    }
                }

                int csvLengde = csvTransAll.Length;

                Logg.Log("Prosesserer " + csvLengde.ToString("#,##0") + " transaksjoner..");
                main.processing.SetText = "Prosesserer " + csvLengde.ToString("#,##0") + " transaksjoner..";
                main.processing.SetProgressStyle = ProgressBarStyle.Marquee;

                DateTime dtFirst = DateTime.MaxValue;
                DateTime dtLast = DateTime.MinValue;
                int valider = 0;

                List<string> avd = new List<string>();

                // sjekk for først og siste dato, samt hvilke avdelinger den inneholder
                string csvAvdelinger = "";
                for (int i = 0; i < csvLengde; i++)
                {
                    if (bw.CancellationPending)
                    {
                        this.returnCode = 2;
                        return false;
                    }

                    if (!csvAvdelinger.Contains(csvTransAll[i].Avd))
                        csvAvdelinger += csvTransAll[i].Avd.ToString() + ";";

                    DateTime dtTemp = DateTime.ParseExact(csvTransAll[i].Dato.ToString(), dateFormat, FormMain.norway);

                    //DateTime dtTemp = Convert.ToDateTime(csvTransAll[i].Dato.ToString());
                    if (DateTime.Compare(dtTemp, dtFirst) < 0)
                        dtFirst = dtTemp;

                    if (DateTime.Compare(dtTemp, dtLast) > 0)
                        dtLast = dtTemp;

                    if (valider < 25) {
                        if (csvTransAll[i].Kgm.StartsWith("2") || csvTransAll[i].Kgm.StartsWith("4") || csvTransAll[i].Kgm.StartsWith("5"))
                            valider++;
                    }
                }
                string[] arrayAvdelinger = csvAvdelinger.Split(';');
                Logg.Debug("Avdelinger funnet i fil (funnetAvdelinger): Antall: (" + arrayAvdelinger.Length + ") Innhold: " + csvAvdelinger);

                string sqlRemoveAvd = "";
                if (main.appConfig.importSetting.Equals("FullFavoritt")) // vi skal bare importere favoritt avdelinger
                    foreach (string arrayAvd in arrayAvdelinger)
                        foreach (string avdel in FormMain.Favoritter)
                            if (arrayAvd.Equals(avdel))  // Avdelingen finnes i csv OG som favoritt avdeling..
                                sqlRemoveAvd += " Avdeling = '" + arrayAvd + "' OR "; // legger til avdeling for sletting før import

                if (main.appConfig.importSetting.Equals("Full")) // vi skal importere ALLE avdelinger i CSV, derfor må vi slette alle avdelinger før import
                    foreach (string arrayAvd in arrayAvdelinger)
                        sqlRemoveAvd += " Avdeling = '" + arrayAvd + "' OR "; // legger til avdeling for sletting før import

                if (sqlRemoveAvd.Length > 3) // Fjen sist "OR" fra sql string for å gjøre den gyldig
                    sqlRemoveAvd = sqlRemoveAvd.Remove(sqlRemoveAvd.Length - 3);

                Logg.Log("Import: CSV inneholder " + csvLengde + " transaksjoner, solgt mellom "
                    + dtFirst.ToString("dddd d. MMMM yyyy", FormMain.norway) + " og "
                    + dtLast.ToString("dddd d. MMMM yyyy", FormMain.norway) + ".", Color.Black, true);

                string strSqlDateFirst = dtFirst.ToString("yyy-MM-dd");
                string strSqlDateLast = dtLast.ToString("yyy-MM-dd");

                if (valider < 25) {
                    Logg.Log("Import: Bare et begrenset antall transaksjoner var gyldig ("
                        + valider + "), kan ikke fortsett. Eksporter transaksjoner på nytt!", Color.Red);
                    {
                        this.returnCode = 1;
                        return false;
                    }
                }

                bool addPrize = false;
                Dictionary<string, decimal> dictPrizes = new Dictionary<string, decimal>();
                if (main.appConfig.dbObsoleteUpdated != FormMain.rangeMin)
                    addPrize = true;

                var table = new DataTable();
                table.Columns.Add("Selgerkode", typeof(string));
                table.Columns.Add("Varegruppe", typeof(Int16));
                table.Columns.Add("Varekode", typeof(string));
                table.Columns.Add("Dato", typeof(DateTime));
                table.Columns.Add("Antall", typeof(int));
                table.Columns.Add("Btokr", typeof(decimal));
                table.Columns.Add("Avdeling", typeof(Int16));
                table.Columns.Add("Salgspris", typeof(decimal));
                table.Columns.Add("Bilagsnr", typeof(int));
                table.Columns.Add("Mva", typeof(decimal));

                main.processing.SetText = "Prosesserer " + csvLengde.ToString("#,##0") + " transaksjoner.. (Konverterer)";

                for (int i = 0; i < csvLengde; i++)
                {
                    if (bw.CancellationPending)
                    {
                        this.returnCode = 2;
                        return false;
                    }

                    if (ImportThisLine(csvTransAll[i]))
                    {
                        int varegruppe = 0, avdeling = 0, bilagsnr = 0;
                        int.TryParse(csvTransAll[i].Kgm.Substring(0, 3), out varegruppe);
                        int.TryParse(csvTransAll[i].Avd, out avdeling);
                        int.TryParse(csvTransAll[i].BilagsNr, out bilagsnr);

                        DataRow dtRow = table.NewRow();
                        dtRow[0] = csvTransAll[i].Sk; // Selgerkode
                        dtRow[1] = varegruppe; // varegruppe
                        dtRow[2] = csvTransAll[i].Varenummer; // varekode
                        dtRow[3] = csvTransAll[i].Dato;
                        dtRow[4] = csvTransAll[i].Antall;
                        dtRow[5] = csvTransAll[i].Btokr;
                        dtRow[6] = avdeling;
                        dtRow[7] = csvTransAll[i].Salgspris;
                        dtRow[8] = bilagsnr;
                        if (csvTransAll[i].Mva == 0)
                            dtRow[9] = 1;
                        else
                            dtRow[9] = (csvTransAll[i].Mva / 100) + 1;

                        // Lagre salgsprisen..
                        if (addPrize)
                            if (csvTransAll[i].Rab == 0 && csvTransAll[i].Antall > 0
                                && csvTransAll[i].Dato.AddDays(main.appConfig.storeMaxAgePrizes) > main.appConfig.dbObsoleteUpdated)
                                if (!dictPrizes.ContainsKey(csvTransAll[i].Varenummer))
                                    dictPrizes.Add(csvTransAll[i].Varenummer, csvTransAll[i].Salgspris);

                        // sjekk lengde på fields før vi legger dem til..
                        if (dtRow[0].ToString().Length < 11 && dtRow[2].ToString().Length < 25)
                            table.Rows.Add(dtRow);
                        else
                        {
                            importReadErrors++;
                            Logg.Log("Import: Format feil ved linje " + i
                                + ": Felt var for lange for databasen, hopper over.", Color.Orange);
                            Logg.Debug("Linje: " + dtRow[0] + ";" + dtRow[1] + ";" + dtRow[2] + ";"
                                + dtRow[3] + ";" + dtRow[4] + ";" + dtRow[5] + ";" + dtRow[6] + ";"
                                + dtRow[7] + ";" + dtRow[8] + ";" + dtRow[9]);
                        }
                    }
                }

                if (addPrize && dictPrizes.Count > 0)
                {
                    Logg.Log("Lager: Oppdaterer vareinfo med priser fra " + dictPrizes.Count + " forskjellige varer..", null, true);
                    UpdatePrizes(dictPrizes);
                    Logg.Log("Lager: Ferdig med oppdatering av priser", null, true);
                }

                // På ditte tidspunktet kan ikke prosessen avbrytes..
                main.processing.SetText = "Sletter gamle transaksjoner..";
                main.processing.SupportsCancelation = false;

                if (table.Rows.Count >= 25)
                {
                    Logg.Log("Import: Klar til å lagre transaksjoner. Sletter overlappende transaksjoner nå..");
                    try
                    {
                        // Slett overlappende transaksjoner vi har i databasen fra før..
                        string sqlRemove = "DELETE FROM tblSalg WHERE (Dato >= '" + strSqlDateFirst
                            + "') AND (Dato <= '" + strSqlDateLast + "')";
                        if (!main.appConfig.importSetting.Equals("Normal"))
                            sqlRemove += " AND (" + sqlRemoveAvd + ")";

                        SqlCeCommand command = new SqlCeCommand(sqlRemove, main.connection);
                        var result = command.ExecuteNonQuery();
                        Logg.Debug("Import: Slettet " + result + " transaksjoner fra databasen.");

                    }
                    catch (Exception ex)
                    {
                        Logg.Unhandled(ex);
                    }
                }

                Logg.Log("Import: Lagrer " + table.Rows.Count.ToString("#,##0") + " transaksjoner..");
                main.processing.SetText = "Lagrer " + table.Rows.Count.ToString("#,##0") + " transaksjoner.. (kan ta litt tid)";

                TimeWatch timewatch = new TimeWatch();
                timewatch.Start();
                // Send data til SQL server og avslutt forbindelsen
                main.database.DoBulkCopy(table, "tblSalg");
                Logg.Log("Import: Lagring av " + table.Rows.Count.ToString("#,##0") + " transaksjoner tok "
                    + timewatch.Stop() + " sekunder.", Color.Black, true);
                main.processing.SupportsCancelation = true; 

                // nuller ut database periode fra-til, tvinger oppdatering ved reload senere
                main.appConfig.dbFrom = DateTime.MinValue;
                main.appConfig.dbTo = DateTime.MinValue;

                Logg.Log("Import: Importering ferdig (Tid: " + timewatch.Stop() + ")");
                main.processing.SetText = "Fullført importering av transaksjoner! Vent litt..";
                main.processing.SetProgressStyle = ProgressBarStyle.Continuous;
                main.processing.SetValue = 90;

                this.importCount = table.Rows.Count;
                this.returnCode = 0;
                return true;
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Sorry! Kritisk uhåndtert feil oppstod under importering.",
                    ex, "CSV filer: " + String.Join("\n", csvFilesToImport.ToArray()));
                errorMsg.ShowDialog(main);
            }
            this.returnCode = -1;
            return false;
        }

        private void UpdatePrizes(Dictionary<string, decimal> dictPrizes)
        {
            try
            {
                foreach (KeyValuePair<string, decimal> entry in dictPrizes)
                {
                    using (SqlCeCommand cmd = new SqlCeCommand("UPDATE tblVareinfo SET Salgspris=@Salgspris "
                        + "WHERE Varekode=@Varekode", main.connection))
                    {
                        cmd.Parameters.Add("@Salgspris", SqlDbType.Decimal);
                        cmd.Parameters["@Salgspris"].Value = entry.Value;
                        cmd.Parameters.AddWithValue("@Varekode", entry.Key);

                        cmd.CommandType = System.Data.CommandType.Text;
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        } 

        public bool RemoveExtraDelimiter(string line, List<string> list)
        {
            try
            {
                string[] col = line.Split(';');
                Logg.Debug("Orginal: " + line);

                decimal t = 0;
                string combined = "";
                if (!decimal.TryParse(col[3], out t)) // er tekst i kollonnen etter
                {
                    for (int b = 0; b < col.Length; b++)
                    {
                        if (b == 5)
                        {
                            combined += col[b] + col[b + 1] + ";"; // slå sammen kolonne 5 og 6
                            b += 1; // hopp over neste kolonne, vi har allerede tatt med den
                        }
                        else
                            combined += col[b] + ";";
                    }
                }

                Logg.Debug("Rettet: " + combined);

                list.Add(combined);
                return true;
            }
            catch
            {
                Logg.Debug("Inport: Unntak ved delimiter fix.");
            }

            return false;
        }

        public string GetTempFilename(string ext)
        {
            string name = Path.GetRandomFileName();
            name = Path.ChangeExtension(name, ext);
            return "infotopdf_" + name;
        }

        private csvImport[] ReadCsv(string file, BackgroundWorker bw, bool unattended)
        {
            try
            {
                FileHelperEngine engine = new FileHelperEngine(typeof(csvImport));
                engine.ErrorManager.ErrorMode = ErrorMode.SaveAndContinue;

                engine.SetProgressHandler(new ProgressChangeHandler(ReadProgressCSV));
                var resCSV = engine.ReadFile(file) as csvImport[];

                if (engine.ErrorManager.HasErrors)
                {
                    List<string> filtLines = new List<string> () {};
                    foreach (ErrorInfo err in engine.ErrorManager.Errors)
                    {
                        if (RemoveExtraDelimiter(err.RecordString, filtLines))
                        {
                            Logg.Debug("Import: Gjenopprettet linje " + err.LineNumber + " med feil format.");
                        }
                        else
                        {
                            importReadErrors++;
                            Logg.Log("Import: Format feil ved linje " + err.LineNumber + ": " + err.RecordString, Color.Red);
                            Logg.Debug("Import: Feilmelding: " + err.ExceptionInfo.ToString());
                            if (importReadErrors > 100)
                            {
                                Logg.Log("CSV er skadet eller inneholder ikke riktige ranking transaksjoner.", Color.Red);
                                if (!unattended)
                                {
                                    FormError errorMsg = new FormError("CSV er skadet eller inneholder ikke riktige ranking transaksjoner.",
                                        err.ExceptionInfo, "Transaksjoner som ikke kunne leses:\n" + filtLines.Count);
                                    errorMsg.ShowDialog();
                                }
                                return null;
                            }
                        }
                    }

                    if (filtLines.Count > 0)
                    {
                        Logg.Debug("Import: Forsøker å legge til " + filtLines.Count + " linjer med feil på nytt etter fix...");
                        Logg.Debug("Import: resCSV length før: " + resCSV.Length);
                        string tmpFile = GetTempFilename(".txt");
                        File.WriteAllLines(tmpFile, filtLines.ToArray(), Encoding.Unicode);
                        var resCSVfixed = engine.ReadFile(tmpFile) as csvImport[];


                        int oldLength = resCSV.Length;
                        Array.Resize<csvImport>(ref resCSV, oldLength + resCSVfixed.Length);
                        Array.Copy(resCSVfixed, 0, resCSV, oldLength, resCSVfixed.Length);


                        Logg.Debug("Import: resCSV length etter: " + resCSV.Length);
                    }
                }
                Logg.Debug("Import: ReadCsv ferdig '" + file + "' - Totalt: " + resCSV.Length);

                return resCSV;
            }
            catch (IOException ex)
            {
                Logg.Log("CSV var låst for lesing. Forleng ventetid i makro hvis overføringen ikke ble ferdig i tide.", Color.Red);
                Logg.Debug("CSV var låst.", ex);
                return null;
            }
            catch (Exception ex)
            {
                throw new Exception("Kritisk feil under lesing av CSV!", ex);
            }
        }

        private void ReadProgressCSV(ProgressEventArgs e)
        {
            if (e.ProgressCurrent % 813 == 0)
            {
                decimal value = (e.ProgressCurrent / (csvSizeGuess + 1)) * 100;
                if (value >= 0 && value < 100)
                    main.processing.SetValue = (int)value;
                main.processing.SetText = "Leser CSV: " + e.ProgressCurrent.ToString("#,##0") + "..";

            }
        }

        private bool ImportThisLine(csvImport line)
        {
            if (main.appConfig.importSetting.Equals("Full")) // Importer alle transaksjoner uten unntak.
                return true;
            if (main.appConfig.importSetting.Equals("FullFavoritt") && FormMain.Favoritter.Contains(line.Avd)) // Importer bare favoritt avdelinger
                return true;
            if (main.appConfig.importSetting.Equals("Normal") && strGroupImportNormal.Contains(line.Kgm.Substring(0, 3)))
                return true;
            return false;
        }
    }
}
