using FileHelpers;
using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace KGSA
{
    public class Obsolete
    {
        FormMain main;
        //DataTable tableHistory;
        public DateTime compareDate { get; set; }
        //BackgroundWorker bwCheckLAtestDate = new BackgroundWorker();
        public void Load(FormMain form)
        {
            this.main = form;
            CheckFromToRecord();
        }

        private void CheckFromToRecord()
        {
            try
            {
                main.appConfig.dbStoreFrom = GetFirstDate();
                main.appConfig.dbStoreTo = GetLatestDate();

                //if (main.appConfig.dbObsoleteUpdated == FormMain.rangeMin || main.appConfig.dbObsoleteUpdated < main.appConfig.dbStoreFrom)
                //    main.appConfig.dbObsoleteUpdated = main.appConfig.dbStoreTo;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        public DataTable GetWeekly(DateTime date)
        {
            try
            {
                string dateStr = date.ToString("yyy-MM-dd");

                DataTable table = main.database.tableWeekly.GetWeeklyTable(main.appConfig.Avdeling, date);
                if (table == null || table.Rows.Count == 0)
                    return new DataTable();

                string sqlProductInfo = "SELECT tblWeekly.ProductCode, MAX(tblVareinfo.Kategori) AS Kategori, MAX(tblVareinfo.Varetekst) AS Varetekst, " +
                    " MAX(tblVareinfo.MerkeNavn) AS MerkeNavn FROM tblWeekly, tblVareinfo " +
                    " WHERE tblVareinfo.Varekode = tblWeekly.ProductCode AND " +
                    " CONVERT(NVARCHAR(10),tblWeekly.Date,121) = CONVERT(NVARCHAR(10),'" + dateStr + "',121) " +
                    " AND tblWeekly.Avdeling = " + main.appConfig.Avdeling + " GROUP BY tblWeekly.ProductCode ";

                DataTable tableInfo = main.database.GetSqlDataTable(sqlProductInfo);
                if (tableInfo == null || tableInfo.Rows.Count == 0)
                {
                    Log.n("Fant ikke vare informasjon for noen av varekodene", Color.Red);
                    tableInfo = new DataTable();
                }

                table.Columns.Add("Kategori", typeof(int));
                table.Columns.Add("Varetekst", typeof(string));
                table.Columns.Add("MerkeNavn", typeof(string));

                foreach (DataRow tblRow in table.Rows)
                {
                    foreach (DataRow rowInfo in tableInfo.Rows)
                    {
                        if (tblRow["ProductCode"].Equals(rowInfo["ProductCode"]))
                        {
                            tblRow["Kategori"] = rowInfo["Kategori"];
                            tblRow["Varetekst"] = rowInfo["Varetekst"];
                            tblRow["MerkeNavn"] = rowInfo["MerkeNavn"];
                        }
                    }
                }

                return table;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
            return null;
        }

        public DataTable GetPopularPrisguideProducts(DateTime date)
        {
            try
            {
                DataTable table = main.database.tablePrisguide.GetPrisguideTable(main.appConfig.Avdeling, date);
                if (table == null || table.Rows.Count == 0)
                    return new DataTable();

                string sqlInfo = "SELECT tblVareinfo.Varekode, tblVareinfo.Varetekst, tblVareinfo.Kategori, tblVareinfo.MerkeNavn " +
                    " FROM tblVareinfo, tblPrisguide WHERE tblVareinfo.Varekode = tblPrisguide.ProductCode " +
                    " AND CONVERT(NVARCHAR(10),tblPrisguide.Date,121) = CONVERT(NVARCHAR(10),'" + date.ToString("yyyy-MM-dd") + "',121) " +
                    " AND tblPrisguide.Avdeling = " + main.appConfig.Avdeling;

                DataTable tableInfo = main.database.GetSqlDataTable(sqlInfo);
                if (tableInfo == null || tableInfo.Rows.Count == 0)
                {
                    Log.n("Ingen match mellom prisguide varer og vareinfo!", Color.Red);
                    tableInfo = new DataTable();
                }

                table.Columns.Add("Kategori", typeof(int));
                table.Columns.Add("Varetekst", typeof(string));
                table.Columns.Add("MerkeNavn", typeof(string));

                foreach (DataRow tblRow in table.Rows)
                {
                    foreach (DataRow rowInfo in tableInfo.Rows)
                    {
                        if (tblRow["ProductCode"].Equals(rowInfo["Varekode"]))
                        {
                            tblRow["Kategori"] = rowInfo["Kategori"];
                            tblRow["Varetekst"] = rowInfo["Varetekst"];
                            tblRow["MerkeNavn"] = rowInfo["MerkeNavn"];
                        }
                    }
                }

                return table;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
            return null;
        }

        public DateTime GetLatestDate()
        {
            try
            {
                using (DataTable dt = main.database.GetSqlDataTable("SELECT TOP(1) Dato FROM tblHistory WHERE Avdeling = " + main.appConfig.Avdeling + " AND Kategori = 'TOTALT' ORDER BY Dato DESC"))
                {
                    if (dt != null)
                    {
                        if (dt.Rows.Count > 0)
                        {
                            DateTime result = Convert.ToDateTime(dt.Rows[0][0]);
                            if (result > FormMain.rangeMin)
                                return result;
                        }
                    }
                    return FormMain.rangeMin;
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return FormMain.rangeMin;
            }
        }

        public DateTime GetFirstDate()
        {
            try
            {
                using (DataTable dt = main.database.GetSqlDataTable("SELECT TOP(1) Dato FROM tblHistory WHERE Avdeling = " + main.appConfig.Avdeling + " AND Kategori = 'TOTALT' ORDER BY Dato ASC"))
                {
                    if (dt != null)
                    {
                        if (dt.Rows.Count > 0)
                        {
                            DateTime result = Convert.ToDateTime(dt.Rows[0][0]);
                            if (result > FormMain.rangeMin)
                                return result;
                        }
                    }
                    return FormMain.rangeMin;
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return FormMain.rangeMin;
            }
        }

        private void ReadProgressCSV(ProgressEventArgs e)
        {
            if (e.ProgressCurrent % 831 == 0)
                main.processing.SetText = "Leser CSV: " + e.ProgressCurrent.ToString("#,##0") + "..";
        }

        public bool ClearDatabase()
        {
            try
            {
                DialogResult msgbox = MessageBox.Show("Sikker på at du vil slette lager databasen?",
                    "KGSA - VIKTIG",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button2);

                if (msgbox == DialogResult.Yes)
                {
                    Log.n("Lager: Nullstiller databasen..");

                    main.database.tableUkurans.Reset();
                    main.database.tableVareinfo.Reset();

                    main.appConfig.dbObsoleteUpdated = FormMain.rangeMin;
                    main.appConfig.dbStoreFrom = FormMain.rangeMin;
                    main.appConfig.dbStoreTo = FormMain.rangeMin;

                    Log.n("Lager: Nullstilling fullført.");
                    return true;
                }
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil oppstod ved sletting av lager databasen.", ex);
                errorMsg.ShowDialog();
            }
            return false;
        }

        public bool SaveHistory(DataTable dt, DateTime date, bool lagerTo = false)
        {
            try
            {
                int lager = (lagerTo) ? main.appConfig.Avdeling + 1000 : main.appConfig.Avdeling;
                if (dt == null)
                    return false;

                if (dt.Rows.Count == 0)
                {
                    Log.d("Lager historikk: Ingen data å lagre.");
                    return false;
                }

                string sqlDelCommand = "DELETE FROM tblHistory WHERE Avdeling = " + lager + " AND Dato = '" + date.ToString("yyy-MM-dd") + "'";
                var command = new SqlCeCommand(sqlDelCommand, main.connection); 
                var result = command.ExecuteNonQuery();
                Log.d("tblHistory: Slettet " + result + " oppføringer.");

                Log.n("Lagrer historisk data for lager..", null, true);

                string sql = "INSERT INTO tblHistory (Avdeling, Dato, Kategori, Lagerantall, Lagerverdi, Ukuransantall, Ukuransverdi, Ukuransprosent) " +
                    "VALUES (@Avdeling, @Dato, @Kategori, @Lagerantall, @Lagerverdi, @Ukuransantall, @Ukuransverdi, @Ukuransprosent)";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    using (SqlCeCommand cmd = new SqlCeCommand(sql, main.connection))
                    {
                        cmd.Parameters.AddWithValue("@Avdeling", dt.Rows[i]["Avdeling"]);
                        cmd.Parameters.Add("@Dato", SqlDbType.DateTime).Value = date.Date;
                        cmd.Parameters.AddWithValue("@Kategori", dt.Rows[i]["Kategori"].ToString());
                        cmd.Parameters.AddWithValue("@Lagerantall", (int)dt.Rows[i]["Lagerantall"]);
                        cmd.Parameters.AddWithValue("@Lagerverdi", dt.Rows[i]["Lagerverdi"]);
                        cmd.Parameters.AddWithValue("@Ukuransantall", (int)dt.Rows[i]["Ukuransantall"]);
                        cmd.Parameters.AddWithValue("@Ukuransverdi", dt.Rows[i]["Ukuransverdi"]);
                        cmd.Parameters.AddWithValue("@Ukuransprosent", dt.Rows[i]["Ukuransprosent"]);
                        cmd.CommandType = System.Data.CommandType.Text;
                        cmd.ExecuteNonQuery();
                    }
                }

                Log.n("Fullført lagring av historisk lager data.", null, true);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return false;
            }
            return false;
        }

        public DateTime FindNearestDate(DateTime date, int minDays = 0)
        {
            try
            {
                // Finn nærmeste dato
                string sql = "SELECT * FROM tblHistory WHERE Avdeling = " + main.appConfig.Avdeling;
                DataTable dtHistory = main.database.GetSqlDataTable(sql);

                int nearestDays = 999;
                DateTime nearestDate = DateTime.Now;
                for (int i = 0; i < dtHistory.Rows.Count; i++)
                {
                    DateTime searchDate = (DateTime)dtHistory.Rows[i]["Dato"];
                    int days = Math.Abs((date - searchDate).Days);
                    if (days < nearestDays && minDays <= days)
                    {
                        nearestDays = days;
                        nearestDate = searchDate;
                    }
                }
                if (nearestDate != DateTime.Now)
                    return nearestDate;
                else
                    return FormMain.rangeMin;
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
                return FormMain.rangeMin;
            }
        }
        public DataTable GetHistory(DateTime dato, bool lagerTo = false)
        {
            try
            {
                int lager = (lagerTo) ? main.appConfig.Avdeling + 1000 : main.appConfig.Avdeling;
                string sql = "SELECT Kategori, Lagerantall, Lagerverdi, Ukuransantall, Ukuransverdi, Ukuransprosent FROM tblHistory WHERE Avdeling = " + lager + " AND Dato = '" + dato.ToString("yyy-MM-dd") + "'";

                DataTable dtResult = main.database.GetSqlDataTable(sql);

                return dtResult;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }

            return null;
        }

        public string Decompress(string filename)
        {
            try
            {
                if (File.Exists(filename) && filename.EndsWith(".zip"))
                {
                    Log.n("Pakker ut " + filename + "..", null, true);

                    string zipToUnpack = filename;
                    string unpackDirectory = System.IO.Path.GetTempPath();
                    using (ZipFile zip1 = ZipFile.Read(zipToUnpack))
                    {
                        foreach (ZipEntry e in zip1)
                        {
                            e.Extract(unpackDirectory, ExtractExistingFileAction.OverwriteSilently);
                        }
                    }

                    Log.n("Utpakket.", null, true);

                    if (File.Exists(System.IO.Path.GetTempPath() + @"\wobsolete.csv"))
                        return System.IO.Path.GetTempPath() + @"\wobsolete.csv";
                    else
                    {
                        Log.n("Feil i utpakking av Zip fil! (" + filename + ")", Color.Red);
                    } 
                }
                return "";
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
                return "";
            }
        }

        public DataTable MakeTableObsolete(bool lagerTo = false)
        {
            try
            {
                int lager = (lagerTo) ? main.appConfig.Avdeling + 1000 : main.appConfig.Avdeling;

                string sql = "SELECT tblUkurans.Avdeling, tblUkurans.Varekode, tblVareinfo.Varegruppe, "
                    + "tblUkurans.Antall, tblUkurans.Kost, tblUkurans.Dato, "
                    + "tblUkurans.UkuransVerdi, tblUkurans.UkuransProsent FROM tblUkurans "
                    + "inner join tblVareinfo on tblUkurans.Varekode = tblVareinfo.Varekode WHERE tblUkurans.Avdeling = " + lager;

                DataTable dtWork = main.database.tableHistory.GetDataTable();
                DataTable sqlce = main.database.GetSqlDataTable(sql);

                if (sqlce.Rows.Count == 0)
                    return dtWork;

                for (int d = 1; d <= 9; d++)
                {
                    if (d == 7)
                        d = 9;

                    string sKat = "";
                    if (d == 1)
                        sKat = "MDA";
                    if (d == 2)
                        sKat = "AudioVideo";
                    if (d == 3)
                        sKat = "SDA";
                    if (d == 4)
                        sKat = "Telecom";
                    if (d == 5)
                        sKat = "Computing";
                    if (d == 6)
                        sKat = "Kitchen";
                    if (d == 9)
                        sKat = "Other";

                    object y;
                    int antall = 0, ukuransAntall = 0;
                    decimal verdi = 0, ukuransVerdi = 0;
                    y = sqlce.Compute("Sum(Antall)", "[Varegruppe] >= '" + (d * 100) + "' AND [Varegruppe] < '" + ((d + 1) * 100) + "'");
                    if (!DBNull.Value.Equals(y))
                        antall = Convert.ToInt32(y);

                    y = sqlce.Compute("Sum(Kost)", "[Varegruppe] >= '" + (d * 100) + "' AND [Varegruppe] < '" + ((d + 1) * 100) + "'");
                    if (!DBNull.Value.Equals(y))
                        verdi = Convert.ToDecimal(y);

                    y = sqlce.Compute("Sum(Antall)", "[Varegruppe] >= '" + (d * 100) + "' AND [Varegruppe] < '" + ((d + 1) * 100) + "' AND [Ukuransverdi] > 0");
                    if (!DBNull.Value.Equals(y))
                        ukuransAntall = Convert.ToInt32(y);

                    y = sqlce.Compute("Sum(Ukuransverdi)", "[Varegruppe] >= '" + (d * 100) + "' AND [Varegruppe] < '" + ((d + 1) * 100) + "' AND [Ukuransverdi] > 0");
                    if (!DBNull.Value.Equals(y))
                        ukuransVerdi = Convert.ToDecimal(y);

                    if (antall != 0)
                    {
                        DataRow dtRow = dtWork.NewRow();
                        dtRow["Avdeling"] = lager;
                        dtRow["Kategori"] = sKat;
                        dtRow["Lagerantall"] = antall;
                        dtRow["Lagerverdi"] = verdi;
                        dtRow["Ukuransantall"] = ukuransAntall;
                        dtRow["Ukuransverdi"] = ukuransVerdi;
                        if (verdi != 0)
                            dtRow["Ukuransprosent"] = Math.Round(ukuransVerdi / verdi * 100, 2);
                        else
                            dtRow["Ukuransprosent"] = 0;
                        dtWork.Rows.Add(dtRow);
                    }
                }

                object r;
                int antallTot = 0, antallUkuransTot = 0;
                decimal verdiTot = 0, verdiUkuransTot = 0;

                r = sqlce.Compute("Sum(Antall)", null);
                if (!DBNull.Value.Equals(r))
                    antallTot = Convert.ToInt32(r);

                r = sqlce.Compute("Sum(Kost)", null);
                if (!DBNull.Value.Equals(r))
                    verdiTot = Convert.ToDecimal(r);

                r = sqlce.Compute("Sum(Antall)", "[Ukuransverdi] > 0");
                if (!DBNull.Value.Equals(r))
                    antallUkuransTot = Convert.ToInt32(r);

                r = sqlce.Compute("Sum(Ukuransverdi)", "[Ukuransverdi] > 0");
                if (!DBNull.Value.Equals(r))
                    verdiUkuransTot = Convert.ToInt32(r);

                // T O T A L T
                DataRow dtRowTot = dtWork.NewRow();
                dtRowTot["Avdeling"] = lager;
                dtRowTot["Kategori"] = "TOTALT";
                dtRowTot["Lagerverdi"] = verdiTot;
                dtRowTot["Lagerantall"] = antallTot;
                dtRowTot["Ukuransantall"] = antallUkuransTot;
                dtRowTot["Ukuransverdi"] = verdiUkuransTot;
                if (verdiTot != 0)
                    dtRowTot["Ukuransprosent"] = Math.Round(verdiUkuransTot / verdiTot * 100, 2);
                else
                    dtRowTot["Ukuransprosent"] = 0;
                dtWork.Rows.Add(dtRowTot);

                return dtWork;
            }
            catch (Exception ex)
            {
                Log.d("Feil oppstod i MakeTableObsolete", ex);
                return null;
            }
        }

        public bool Import(string filename, BackgroundWorker bw = null)
        {
            try
            {
                if (!File.Exists(filename))
                {
                    Log.n("Vare import: Fant ikke fil eller ble nektet tilgag. (" + filename + ")", Color.Red);
                    return false;
                }

                int importReadErrors = 0;
                var engine = new FileHelperEngine(typeof(csvObsolete));
                engine.ErrorManager.ErrorMode = ErrorMode.SaveAndContinue;

                main.processing.SetProgressStyle = ProgressBarStyle.Marquee;
                engine.SetProgressHandler(new ProgressChangeHandler(ReadProgressCSV));
                var resCSV = engine.ReadFile(filename) as csvObsolete[];

                if (engine.ErrorManager.HasErrors)
                    foreach (ErrorInfo err in engine.ErrorManager.Errors)
                    {
                        importReadErrors++;
                        Log.n("Import: Klarte ikke lese linje " + err.LineNumber + ": " + err.RecordString, Color.Red);
                        Log.d("Exception: " + err.ExceptionInfo.ToString());

                        if (importReadErrors > 100)
                        {
                            Log.n("Feil: CSV er ikke en obsolete eksportering eller filen er skadet. (" + filename + ")", Color.Red);
                            return false;
                        }

                    }

                main.processing.SetProgressStyle = ProgressBarStyle.Continuous;
                int count = resCSV.Length;
                string dateFormat = "dd/MM/yyyy HH:mm:ss";

                DateTime dtLast = DateTime.MinValue;
                for (int i = 0; i < count; i++)
                {
                    if (bw != null)
                        if (bw.CancellationPending)
                            return false;

                    DateTime dtTemp = DateTime.ParseExact(resCSV[i].DatoInnLager.ToString(), dateFormat, FormMain.norway);
                    if (DateTime.Compare(dtTemp, dtLast) > 0)
                        dtLast = dtTemp;
                }

                if (count > 0)
                {

                    Log.n("Prosesserer " + count.ToString("#,##0") + " vare oppføringer.. (" + filename + ")");

                    var tableUkurans = main.database.tableUkurans.GetDataTable();

                    DataTable dtVareinfo = main.database.GetSqlDataTable("SELECT Varekode FROM tblVareinfo");

                    var tableVareinfo = main.database.tableVareinfo.GetDataTable();

                    int lager2 = main.appConfig.Avdeling + 1000;
                    for (int i = 0; i < count; i++)
                    {
                        if (bw != null)
                        {
                            if (i % 321 == 0)
                            {
                                bw.ReportProgress(i, new StatusProgress(count, null, 0, 100));
                                main.processing.SetText = "Lagrer vare oppføringer..: " + i.ToString("#,##0") + " / " + count.ToString("#,##0");
                            }
                            if (bw.CancellationPending)
                            {
                                Log.n("Lager importering avbrutt av bruker!", Color.Red);
                                return false;
                            }
                        }

                        var varekode = resCSV[i].Varekode;
                        if (varekode.Length > 0)
                            varekode = varekode.Substring(1, varekode.Length - 1);

                            if (!dtVareinfo.AsEnumerable().Any(row => varekode == row.Field<String>("Varekode")) && !tableVareinfo.AsEnumerable().Any(row => varekode == row.Field<String>("Varekode")))
                            {
                                DataRow dtRow = tableVareinfo.NewRow();
                                dtRow["Varekode"] = varekode;
                                dtRow["Varetekst"] = resCSV[i].VareTekst;
                                dtRow["Kategori"] = (int)resCSV[i].Kat;
                                dtRow["KategoriNavn"] = resCSV[i].KatNavn;
                                dtRow["Varegruppe"] = (int)resCSV[i].Grp;
                                dtRow["VaregruppeNavn"] = resCSV[i].GrpNavn;
                                dtRow["Modgruppe"] = (int)resCSV[i].Mod;
                                dtRow["ModgruppeNavn"] = resCSV[i].ModNavn;
                                dtRow["Merke"] = (int)resCSV[i].Merke;
                                dtRow["MerkeNavn"] = resCSV[i].MerkeNavn;
                                dtRow["Dato"] = DateTime.Now;
                                tableVareinfo.Rows.Add(dtRow);
                            }


                        if (resCSV[i].Avd == main.appConfig.Avdeling || resCSV[i].Avd == lager2)
                        {
                            DataRow dtRow = tableUkurans.NewRow();
                            dtRow["Avdeling"] = Convert.ToInt32(resCSV[i].Avd);
                            dtRow["Varekode"] = varekode;
                            dtRow["Antall"] = Convert.ToInt32(resCSV[i].AntallLager);
                            dtRow["Kost"] = Convert.ToDecimal(resCSV[i].KostVerdiLager);
                            dtRow["Dato"] = resCSV[i].DatoInnLager;
                            dtRow["UkuransVerdi"] = Convert.ToDecimal(resCSV[i].UkuransVerdi);
                            var s = resCSV[i].UkuransProsent;
                            decimal d = 0;
                            if (s.Length > 0)
                            {
                                s = s.Substring(0, s.Length - 1);
                                if (s.StartsWith(","))
                                    s = "0" + s;
                                d = Convert.ToDecimal(s);
                                dtRow["UkuransProsent"] = d;
                            }
                            else
                                dtRow["UkuransProsent"] = 0;

                            tableUkurans.Rows.Add(dtRow);
                        }
                    }

                    main.processing.SetText = "Fullfører vare import..";
                    main.database.tableUkurans.Reset();
                    main.database.DoBulkCopy(tableUkurans, "tblUkurans");

                    if (tableVareinfo.Rows.Count > 0)
                        main.database.DoBulkCopy(tableVareinfo, "tblVareinfo");

                    main.appConfig.dbObsoleteUpdated = dtLast;
                    main.appConfig.dbStoreFrom = FormMain.rangeMin;
                    main.appConfig.dbStoreTo = FormMain.rangeMin;

                    Log.d("Forbereder lagring av historikk for lager..");
                    SaveHistory(MakeTableObsolete(false), dtLast, false);
                    SaveHistory(MakeTableObsolete(true), dtLast, true);
                    Log.d("Lagring av historikk fullført.");

                }
                return true;
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil ved importering av lager", ex);
                errorMsg.ShowDialog();
                return false;
            }
        }

    }
}
