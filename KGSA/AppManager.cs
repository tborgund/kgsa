using FileHelpers;
using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Data.SqlServerCe;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace KGSA
{
    public class AppManager
    {
        #region variables
        FormMain main;
        public static string sqlTblProductData = "DROP TABLE if exists [tblProductData]; " +
            "CREATE TABLE if not exists [tblProductData] ( " +
            "  [Id] bigint NOT NULL " +
            ", [Barcode] nchar(13) NOT NULL " +
            ", [Varekode] nvarchar(25) NULL " +
            ", [Varetekst] nvarchar(50) NOT NULL " +
            ", [Kategori] int NULL " +
            ", [KategoriNavn] nvarchar(15) NULL " +
            ", [Varegruppe] int NULL " +
            ", [VaregruppeNavn] nvarchar(25) NULL " +
            ", [Modgruppe] int NULL " +
            ", [ModgruppeNavn] nvarchar(25) NULL " +
            ", [Merke] int NULL " +
            ", [MerkeNavn] nvarchar(25) NULL " +
            ", [DatoOpprettet] bigint NULL " +
            ", [Salgspris] money NULL " +
            ", CONSTRAINT [sqlite_master_PK_tblProductData] PRIMARY KEY ([Id]) " +
            ");";

        public static string sqlTblInventory = "DROP TABLE if exists [tblInventory]; " +
            "CREATE TABLE if not exists [tblInventory] ( " +
            "  [Id] bigint NOT NULL " +
            ", [Avdeling] int NOT NULL " +
            ", [Varekode] nvarchar(25) NULL " +
            ", [Varetekst] nvarchar(50) NOT NULL " +
            ", [Kategori] int NULL " +
            ", [KategoriNavn] nvarchar(15) NULL " +
            ", [Varegruppe] int NULL " +
            ", [VaregruppeNavn] nvarchar(25) NULL " +
            ", [Modgruppe] int NULL " +
            ", [ModgruppeNavn] nvarchar(25) NULL " +
            ", [Merke] int NULL " +
            ", [MerkeNavn] nvarchar(25) NULL " +
            ", [Antall] int NOT NULL " +
            ", [Kost] money NULL " +
            ", [Dato] bigint NULL " +
            ", [UkuransMnd] int NULL " +
            ", [UkuransVerdi] money NULL " +
            ", [UkuransProsent] money NULL " +
            ", CONSTRAINT [sqlite_master_PK_tblInventory] PRIMARY KEY ([Id]) " +
            ");";

        public static string sqlTblInfo = "DROP TABLE if exists [tblInfo]; " +
            "CREATE TABLE if not exists [tblInfo] (" +
            "[DatoOppdatert] bigint NOT NULL " +
            ", [DatoEksportert] bigint NOT NULL " +
            ", [Versjon] nvarchar(25) NOT NULL " +
            ", [Avdeling] int NOT NULL " +
            ", [AvdelingNavn] nvarchar(25) NOT NULL " +
            ", [Type] nvarchar(25) NOT NULL " +
            ");";

        string sqlTblEanDrop = "DROP TABLE [tblEan];";
        string sqlTblEanCreate = "CREATE TABLE [tblEan] ( " +
                        "[Id] int IDENTITY (1,1) NOT NULL " +
                        ", [Barcode] nchar(13) NOT NULL " +
                        ", [Varekode] nvarchar(25) NOT NULL " +
                        ", [Varetekst] nvarchar(50) NOT NULL " +
                        ");";
        string sqlTblEanAlter = "ALTER TABLE [tblEan] ADD CONSTRAINT [PK_tblEan] PRIMARY KEY ([Id]);";

        BackgroundWorker worker;

        #endregion
        public AppManager(FormMain form)
        {
            this.main = form;
        }

        public void UpdateAllAsync(BackgroundWorker bw)
        {
            this.worker = bw;
            worker.DoWork += new DoWorkEventHandler(bwUpdateAll_DoWork);
            worker.ProgressChanged += new ProgressChangedEventHandler(main.bwProgressReport_ProgressChanged);
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwUpdateAll_Completed);

            main.processing.SetVisible = true;
            main.processing.SetText = "Forbereder..";
            main.processing.SetProgressStyle = ProgressBarStyle.Marquee;
            worker.RunWorkerAsync();
            main.processing.SetBackgroundWorker = worker;
        }

        internal void bwUpdateAll_DoWork(object sender, DoWorkEventArgs e)
        {
            FormMain.appManagerIsBusy = true;
            main.ProgressStart();

            bool invUpdated = ImportAndConvertInventory(worker);
            if (invUpdated && File.Exists(FormMain.settingsPath + @"\" + BluetoothServer.inventoryFilename))
            {
                main.appConfig.blueInventoryReady = true;
                Logg.Log("Varebeholdnings databasen er klar for App.", Color.Green);
            }
            else
                Logg.Log("Varebeholdnings databasen ble ikke oppdatert. Se logg for detaljer", Color.Red);

            bool dataUpdated = ExportProductDatabase(worker);
            if (dataUpdated && File.Exists(FormMain.settingsPath + @"\" + BluetoothServer.dataFilename))
            {
                main.appConfig.blueProductLastDate = main.appConfig.dbTo;
                main.appConfig.blueProductReady = true;
                Logg.Log("Produkt databasen er klar for App.", Color.Green);
            }
            else
                Logg.Log("Produkt databasen ble ikke oppdatert. Se logg for detaljer", Color.Red);

            if (!invUpdated || !dataUpdated)
                e.Result = false;
            else
                e.Result = true;
        }

        internal void bwUpdateAll_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            FormMain.appManagerIsBusy = false;
            main.ProgressStop();
            main.processing.SetProgressStyle = ProgressBarStyle.Continuous;
            main.processing.SetValue = 100;

            if ((bool)e.Result && e.Error == null && !e.Cancelled)
            {
                main.processing.SetText = "Fullført eksportering av produkt-data og varebeholdning";
                main.processing.HideDelayed();
            }
            else if (e.Cancelled)
            {
                main.processing.SetText = "Avbrutt etter ønske";
                main.processing.HideDelayed();
                Logg.Log("Prosessen ble avbrutt etter ønske.", Color.Red);
            }
            else
            {
                main.processing.SetText = "Prosessen ble avbrutt";
                main.processing.HideDelayed();
                Logg.Log("Prosessen ble fullført, men med feil. Se logg for detaljer.", Color.Red);
            }
        }

        #region ExportInventory

        private void ReadProgressCSV(FileHelpers.ProgressEventArgs e)
        {
            if (e.ProgressCurrent % 831 == 0)
                main.processing.SetText = "Leser CSV: " + e.ProgressCurrent.ToString("#,##0") + "..";
        }

        public bool ImportAndConvertInventory(BackgroundWorker bw)
        {
            DateTime exported = FormMain.rangeMin;
            String wobsoleteFile = main.appConfig.csvElguideExportFolder + @"\wobsolete.zip";
            try
            {
                if (File.Exists(wobsoleteFile))
                {
                    exported = File.GetLastWriteTime(wobsoleteFile);
                    Logg.Debug(wobsoleteFile + " eksportert: " + exported.ToShortDateString());
                }
                else
                {
                    Logg.Log("Import: Fant ikke fil eller ble nektet tilgang. (" + wobsoleteFile + ")", Color.Red);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return false;
            }

            string wobsoleteDecompressed = Decompress(wobsoleteFile);
            if (String.IsNullOrEmpty(wobsoleteDecompressed))
            {
                Logg.Log("Import: Feil under utpakking av CSV", Color.Red);
                return false;
            }

            string database = FormMain.settingsPath + @"\" + BluetoothServer.inventoryFilename;
            if (File.Exists(database))
            {
                try
                {
                    File.Delete(database);
                    SQLiteConnection.CreateFile(database);
                }
                catch (Exception ex)
                {
                    Logg.Unhandled(ex);
                    Logg.Log("Import: Eksport destinasjonen " + database + " eksisterer og er låst.", Color.Red);
                    return false;
                }
            }

            SQLiteConnection con = new SQLiteConnection("Data Source=" + database + ";Version=3;");
            try
            {
                int importReadErrors = 0;
                var engine = new FileHelperEngine(typeof(csvObsolete));
                engine.ErrorManager.ErrorMode = ErrorMode.SaveAndContinue;

                main.processing.SetProgressStyle = ProgressBarStyle.Marquee;
                engine.SetProgressHandler(new ProgressChangeHandler(ReadProgressCSV));
                var resCSV = engine.ReadFile(wobsoleteDecompressed) as csvObsolete[];

                if (engine.ErrorManager.HasErrors)
                    foreach (ErrorInfo err in engine.ErrorManager.Errors)
                    {
                        if (bw != null)
                            if (bw.CancellationPending)
                                return false;

                        importReadErrors++;
                        Logg.Log("Import: Klarte ikke lese linje " + err.LineNumber + ": " + err.RecordString, Color.Red);
                        Logg.Debug("Exception: " + err.ExceptionInfo.ToString());

                        if (importReadErrors > 100)
                        {
                            Logg.Log("Feil: CSV er ikke en obsolete eksportering eller filen er skadet. (" + wobsoleteDecompressed + ")", Color.Red);
                            return false;
                        }

                    }

                int count = resCSV.Length;
                if (count == 0)
                {
                    Logg.Log("CSV filer var tom!", Color.Red);
                    return false;
                }

                main.processing.SetProgressStyle = ProgressBarStyle.Continuous;
                Logg.Log("Prosesserer " + count.ToString("#,##0") + " vare oppføringer.. (" + wobsoleteDecompressed + ")");

                DateTime lastDate = FormMain.rangeMin;
                using (var conSqlite = new SQLiteConnection(con))
                {
                    conSqlite.Open();
                    using (var cmdCreateTable = new SQLiteCommand(sqlTblInventory, conSqlite))
                    {
                        cmdCreateTable.ExecuteNonQuery();
                    }
                    using (var cmd = new SQLiteCommand(conSqlite))
                    {
                        using (var transaction = conSqlite.BeginTransaction())
                        {
                            int lager2 = main.appConfig.Avdeling + 1000;
                            for (var i = 0; i < count; i++)
                            {
                                if (bw != null)
                                {
                                    if (i % 321 == 0)
                                    {
                                        bw.ReportProgress(i, new StatusProgress(count, null, 0, 100));
                                        main.processing.SetText = "Lagrer vare oppføringer..: "
                                            + i.ToString("#,##0") + " / " + count.ToString("#,##0");
                                    }
                                }

                                long TICKS_AT_EPOCH = 621355968000000000L;
                                long TICKS_PER_MILLISECOND = 10000;

                                if (resCSV[i].Avd == main.appConfig.Avdeling || resCSV[i].Avd == lager2)
                                {
                                    if (resCSV[i].DatoInnLager > lastDate)
                                        lastDate = resCSV[i].DatoInnLager;

                                    cmd.CommandText = "INSERT INTO tblInventory (Id, Avdeling, Varekode, Varetekst, Kategori, KategoriNavn, Varegruppe, VaregruppeNavn, Modgruppe, ModgruppeNavn, Merke, MerkeNavn, Antall, Kost, Dato, UkuransMnd, UkuransVerdi, UkuransProsent) VALUES (" + i + ", "
                                        + resCSV[i].Avd + ", '"
                                        + resCSV[i].Varekode.Replace("'", "") + "', '"
                                        + resCSV[i].VareTekst.Replace("'", "''") + "', "
                                        + resCSV[i].Kat + ", '"
                                        + resCSV[i].KatNavn + "', "
                                        + resCSV[i].Grp + ", '"
                                        + resCSV[i].GrpNavn + "', "
                                        + resCSV[i].Mod + ", '"
                                        + resCSV[i].ModNavn + "', "
                                        + resCSV[i].Merke + ", '"
                                        + resCSV[i].MerkeNavn + "', "
                                        + resCSV[i].AntallLager + ", "
                                        + resCSV[i].KostVerdiLager.ToString().Replace(",", ".") + ", "
                                        + ((resCSV[i].DatoInnLager.ToUniversalTime().Ticks - TICKS_AT_EPOCH) / TICKS_PER_MILLISECOND).ToString() + ", "
                                        + resCSV[i].MndUkurans + ", "
                                        + resCSV[i].UkuransVerdi.Replace(" ", "").Replace(",", ".").Replace("%", "") + ", "
                                        + resCSV[i].UkuransProsent.Replace(" ", "").Replace(",", ".").Replace("%", "") + ");";

                                    cmd.ExecuteNonQuery();
                                }
                            }
                            transaction.Commit();
                        }

                    }

                    main.appConfig.blueInventoryExportDate = exported;
                    main.appConfig.blueInventoryLastDate = lastDate;

                    CreateInfoTable(conSqlite, "Inventory", main.appConfig.blueInventoryLastDate, main.appConfig.blueInventoryExportDate);
                    Logg.Debug("Opprettet inventory info tabell");

                    conSqlite.Close();
                }

                Logg.Debug(wobsoleteFile + " siste dato: " + lastDate.ToShortDateString() + " og eksportert: " + exported.ToShortDateString());

                return true;
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil ved konvertering av wobsolete", ex);
                errorMsg.ShowDialog();
                return false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        #endregion

        #region ExportProductData

        public bool ExportProductDatabase(BackgroundWorker bw)
        {
            DateTime exported = FormMain.rangeMin;
            String csvPath = main.appConfig.csvElguideExportFolder + @"\irank.csv";
            try
            {
                if (File.Exists(csvPath))
                {
                    exported = File.GetLastWriteTime(csvPath);
                }
                else
                {
                    Logg.Log("Import: Fant ikke fil eller ble nektet tilgang. (" + csvPath + ")", Color.Red);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return false;
            }

            string database = FormMain.settingsPath + @"\" + BluetoothServer.dataFilename;
            if (File.Exists(database))
            {
                try
                {
                    File.Delete(database);
                    SQLiteConnection.CreateFile(database);
                }
                catch (Exception ex)
                {
                    Logg.Unhandled(ex);
                    Logg.Log("Eksport: Databasen er låst og kan ikke slettes/fornyes: " + database + ". Exception: " + ex.Message);
                    return false;
                }
            }

            SQLiteConnection sqliteCon = new SQLiteConnection("Data Source=" + database + ";Version=3;");
            try
            {
                Logg.Debug("Henter varedata fra databasen..");
                main.processing.SetText = "Henter varedata..";
                DataTable table = main.database.GetSqlDataTable("select * from tblVareinfo");
                main.processing.SetText = "Henter EAN koder..";
                DataTable tableEan = main.database.GetSqlDataTable("select Barcode, Varekode, Varetekst from tblEan");

                if (table == null)
                    throw new Exception("Intern feil i databasen. tblVareinfo var null!");
                if (tableEan == null)
                    throw new Exception("Intern feil i databasen. tblEan var null!");

                if (tableEan.Rows.Count == 0)
                {
                    MessageBox.Show("Databasen inneholder ingen EAN koder. Eksporter EAN fra Elguide først.", "Mangler EAN", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Logg.Log("EAN tabellen er tom. Importer EAN koder fra Elguide først.", Color.Red);
                    return false;
                }

                if (table.Rows.Count == 0)
                {
                    Logg.Log("Fant ingen varer i databasen. Har du importert lagervarer fra Elguide?", Color.Red);
                    return false;
                }

                DataColumn column = new DataColumn();
                column.ColumnName = "Barcode";
                column.DataType = typeof(string);
                table.Columns.Add(column);

                table.Columns["Varekode"].AllowDBNull = true;
                table.Columns["Kategori"].AllowDBNull = true;
                table.Columns["KategoriNavn"].AllowDBNull = true;
                table.Columns["Varegruppe"].AllowDBNull = true;
                table.Columns["VaregruppeNavn"].AllowDBNull = true;
                table.Columns["Modgruppe"].AllowDBNull = true;
                table.Columns["ModgruppeNavn"].AllowDBNull = true;
                table.Columns["Merke"].AllowDBNull = true;
                table.Columns["MerkeNavn"].AllowDBNull = true;
                table.Columns["Dato"].AllowDBNull = true;
                table.Columns["Salgspris"].AllowDBNull = true;

                int countvare = table.Rows.Count;
                Logg.Debug("Henter ut EAN koder.. (" + countvare + ")");
                try
                {
                    for (int i = 0; i < countvare; i++)
                    {
                        if (bw.CancellationPending)
                            return false;

                        var filter = tableEan.Select("Varekode = '" + table.Rows[i]["Varekode"] + "'");
                        if (filter.Length == 1)
                            table.Rows[i]["Barcode"] = filter[0]["Barcode"].ToString();
                    }
                }
                catch (Exception ex)
                {
                    Logg.Log("Feil ved EAN kode setting.. " + ex.Message);
                }

                DataTable tableEanOnly = tableEan.Select("Varekode = ''").CopyToDataTable();

                Logg.Debug("Setter EAN kode til varedata.. (Antall: " + tableEanOnly.Rows.Count + ")");

                int countEanOnly = tableEanOnly.Rows.Count;
                for (int i = 0; i < countEanOnly; i++)
                {
                    if (bw.CancellationPending)
                        return false;

                    DataRow dtRow = table.NewRow();
                    dtRow["Barcode"] = tableEanOnly.Rows[i]["Barcode"];
                    dtRow["Varetekst"] = tableEanOnly.Rows[i]["Varetekst"];
                    dtRow["Kategori"] = 0;
                    dtRow["Varegruppe"] = 0;
                    dtRow["Modgruppe"] = 0;
                    dtRow["Merke"] = 0;
                    dtRow["Dato"] = FormMain.rangeMin;
                    table.Rows.Add(dtRow);
                }

                main.processing.SetText = "Behandler " + string.Format("{0:n0}", table.Rows.Count) + " varekoder..";
                Logg.Log("Behandler " + string.Format("{0:n0}", table.Rows.Count) + " varekoder..");

                DateTime lastDate = FormMain.rangeMin;
                using (var con = new SQLiteConnection(sqliteCon))
                {
                    con.Open();
                    using (var command = new SQLiteCommand(sqlTblProductData, con))
                    {
                        command.ExecuteNonQuery();
                    }
                    using (var cmd = new SQLiteCommand(con))
                    {
                        Logg.Debug("Flytter resultat over til Sqlite databasen..");
                        using (var transaction = con.BeginTransaction())
                        {
                            int count = table.Rows.Count;
                            for (var i = 0; i < count; i++)
                            {
                                if (count % 33 == 0)
                                    main.processing.SetText = "Skriver " + i + " av " + count + " varekoder..";

                                decimal dec = 0M;
                                Decimal.TryParse(table.Rows[i]["Salgspris"].ToString(), out dec);

                                DateTime date = Convert.ToDateTime(table.Rows[i]["Dato"]);
                                if (date > lastDate)
                                    lastDate = date;

                                cmd.CommandText =
                                    "INSERT INTO tblProductData (Id, Barcode, Varekode, Varetekst, Kategori, KategoriNavn, "
                                    + "Varegruppe, VaregruppeNavn, Modgruppe, ModgruppeNavn, Merke, MerkeNavn, "
                                    + " DatoOpprettet, Salgspris) VALUES (" + i + ", '"
                                + table.Rows[i]["Barcode"] + "', '" + table.Rows[i]["Varekode"] + "', '"
                                + table.Rows[i]["Varetekst"].ToString().Replace("'", "''") + "', " + table.Rows[i]["Kategori"] + ", '"
                                + table.Rows[i]["KategoriNavn"] + "', " + table.Rows[i]["Varegruppe"] + ", '"
                                + table.Rows[i]["VaregruppeNavn"] + "', " + table.Rows[i]["Modgruppe"] + ", '"
                                + table.Rows[i]["ModgruppeNavn"] + "', " + table.Rows[i]["Merke"] + ", '"
                                + table.Rows[i]["MerkeNavn"] + "', " + FromDateTimeToInteger(date) + ", "
                                + dec.ToString().Replace(",",".") + ");";
                                cmd.ExecuteNonQuery();
                            }
                            transaction.Commit();
                        }
                        Logg.Debug("Ferdig med å flytte resultat over til databasen.");
                    }

                    main.appConfig.blueProductExportDate = exported;
                    main.appConfig.blueProductLastDate = lastDate;

                    CreateInfoTable(con, "ProductData", main.appConfig.blueProductLastDate, main.appConfig.blueInventoryExportDate);
                    Logg.Debug("Opprettet produktdata info tabell");

                    con.Close();
                }

                Logg.Debug(csvPath + " siste dato: " + lastDate.ToShortDateString() + " og eksportert: " + exported.ToShortDateString());

                return true;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                Logg.Log("Feil ved eksport av produtdata: " + ex.Message);
                return false;
            }
            finally
            {
                if (sqliteCon != null)
                {
                    sqliteCon.Close();
                    sqliteCon.Dispose();
                }
            }
        }

        #endregion

        #region importEAN

        private int ean_current_count = 0;
        private int ean_new_count = 0;
        private int ean_nomatch_count = 0;
        private int ean_match_count = 0;
        private int ean_totalt_count = 0;

        public void ImportEan(BackgroundWorker bw)
        {
            this.worker = bw;
            // Browse etter *.dat..
            try
            {
                var fdlg = new OpenFileDialog();
                fdlg.Title = "Velg strekkode fil (DAT) laget av Elguide";
                fdlg.InitialDirectory = @"c:\";
                fdlg.Filter = "All files (*.*)|*.*|DAT filer (*.dat)|*.dat";
                fdlg.FilterIndex = 2;
                fdlg.RestoreDirectory = true;
                fdlg.Multiselect = true;

                if (fdlg.ShowDialog() == DialogResult.OK)
                {
                    List<string> eanfiles = new List<string>() { };
                    eanfiles.AddRange(fdlg.FileNames);
                    main.processing.SetVisible = true;
                    main.processing.SetText = "Starter import av EAN datafil..";
                    main.processing.SetProgressStyle = ProgressBarStyle.Marquee;

                    worker.DoWork += new DoWorkEventHandler(bwEanImport_DoWork);
                    worker.ProgressChanged += new ProgressChangedEventHandler(main.bwProgressReport_ProgressChanged);
                    worker.WorkerReportsProgress = true;
                    worker.WorkerSupportsCancellation = true;
                    worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwEanImport_Completed);
                    worker.RunWorkerAsync(eanfiles);
                    main.processing.SetBackgroundWorker = worker;
                }
                fdlg.Dispose();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void bwEanImport_DoWork(object sender, DoWorkEventArgs e)
        {
            FormMain.appManagerIsBusy = true;
            main.ProgressStart();
            this.ean_current_count = 0;
            this.ean_new_count = 0;
            this.ean_nomatch_count = 0;
            this.ean_match_count = 0;
            this.ean_totalt_count = 0;
            
            List<string> eanfiles = (List<string>)e.Argument;
            if (eanfiles != null)
                e.Result = ImportEanWorkerBrowse(worker, eanfiles);
            else
                throw new ArgumentNullException("Background worker argument was null!");
        }

        private bool ImportEanWorkerBrowse(BackgroundWorker bw, List<string> eanfiles)
        {
            try
            {
                DataTable tableFile = new DataTable();
                tableFile.Columns.Add("Barcode", typeof(string));
                tableFile.Columns.Add("Varetekst", typeof(string));

                foreach (string file in eanfiles)
                {
                    if (bw != null)
                        if (bw.CancellationPending)
                            return false;

                    if (!File.Exists(file) || String.IsNullOrEmpty(file))
                        throw new IOException("Fant ikke fil eller ble nektet tilgang: " + file);

                    Logg.Log("Leser " + file + "..");
                    string line;
                    System.IO.StreamReader fileFile = new System.IO.StreamReader(file, System.Text.Encoding.GetEncoding(865));
                    while ((line = fileFile.ReadLine()) != null)
                    {
                        if (bw != null)
                            if (bw.CancellationPending)
                                return false;

                        string lineStr = line.Trim();
                        if (lineStr.Length > 13)
                        {
                            if (Regex.IsMatch(lineStr.Substring(0, 13), @"^\d+$"))
                            {
                                string tekst = lineStr.Substring(13, lineStr.Length - 13);

                                DataRow dtRow = tableFile.NewRow();
                                dtRow["Barcode"] = lineStr.Substring(0, 13);
                                if (tekst.Length > 29)
                                    dtRow["Varetekst"] = lineStr.Substring(13, 29).Trim();
                                else
                                    dtRow["Varetekst"] = lineStr.Substring(13, lineStr.Length - 13).Trim();
                                tableFile.Rows.Add(dtRow);
                            }
                        }
                    }
                    fileFile.Close();
                }

                Logg.Log("Henter ut eksisterende EAN koder fra databasen..");

                if (!main.connection.TableExists("tblEan"))
                {
                    var cmdCreate = new SqlCeCommand(sqlTblEanCreate, main.connection);
                    cmdCreate.ExecuteNonQuery();
                    var cmdAlter = new SqlCeCommand(sqlTblEanAlter, main.connection);
                    cmdAlter.ExecuteNonQuery();
                }
                DataTable dtCurrent = main.database.GetSqlDataTable("SELECT Barcode, Varetekst FROM tblEan");

                var tableNew = new DataTable();
                tableNew.Columns.Add("Barcode", typeof(string));
                tableNew.Columns.Add("Varetekst", typeof(string));
                this.ean_current_count = dtCurrent.Rows.Count;
                main.processing.SetText = "Prosesserer innhold..";

                if (ean_current_count > 0)
                {
                    Logg.Log("EAN koder i tblEan: " + ean_current_count, null, true);
                    Logg.Log("Filtrerer ut EAN koder vi har fra før..");

                    for (int i = 0; i < tableFile.Rows.Count; i++)
                    {
                        if (bw != null)
                            if (bw.CancellationPending)
                                return false;

                        DataRow[] filteredRows = dtCurrent.Select("Barcode = '" + tableFile.Rows[i]["Barcode"] + "'");
                        if (filteredRows.Length == 0)
                        {
                            DataRow dtRow = tableNew.NewRow();
                            dtRow["Barcode"] = tableFile.Rows[i]["Barcode"];
                            dtRow["Varetekst"] = tableFile.Rows[i]["Varetekst"] ;
                            tableNew.Rows.Add(dtRow);
                        }
                    }
                }
                else
                    tableNew = tableFile.Copy();

                this.ean_new_count = tableNew.Rows.Count;

                if (ean_new_count == 0)
                {
                    Logg.Log("Fant ingen nye EAN koder.");
                    return true;
                }

                Logg.Log("Fant " + string.Format("{0:n0}", tableNew.Rows.Count) + " nye EAN koder.", null, true);

                return MatchEanToDatabase(tableNew, bw);
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return false;
            }
        }

        private bool MatchEanToDatabase(DataTable dtTblCurrent, BackgroundWorker bw)
        {
            try
            {
                DataView view = dtTblCurrent.DefaultView;
                view.Sort = "Varetekst ASC";
                dtTblCurrent = view.ToTable();

                main.processing.SetText = "Behandler EAN databasen..";
                Logg.Log("Henter varekoder fra databasen og matcher varetekst.. (kan ta en stund!)");

                DataTable dtVarekoder = main.database.GetSqlDataTable("SELECT Varekode, SUBSTRING(Varetekst, 1, 29) AS Varetekst FROM tblVareinfo");

                var dtTblMatch = new DataTable();
                dtTblMatch.Columns.Add("Barcode", typeof(string));
                dtTblMatch.Columns.Add("Varekode", typeof(string));
                dtTblMatch.Columns.Add("Varetekst", typeof(string));

                for (int i = 0; i < dtVarekoder.Rows.Count; i++)
                {
                    if (bw != null)
                        if (bw.CancellationPending)
                            return false;

                    DataRow[] result = dtTblCurrent.Select("Varetekst = '" + dtVarekoder.Rows[i]["Varetekst"].ToString().Replace("'", "''") + "'");
                    if (result.Count() > 0)
                    {
                        DataRow dtRow = dtTblMatch.NewRow();
                        dtRow["Barcode"] = result[0]["Barcode"];
                        dtRow["Varekode"] = dtVarekoder.Rows[i]["Varekode"];
                        //dtRow["Varetekst"] = dtVarekoder.Rows[i]["Varetekst"];
                        dtRow["Varetekst"] = "";

                        dtTblMatch.Rows.Add(dtRow);
                    }
                    //else Logg.Debug("Fant ikke match på tekst '" + dtVarekoder.Rows[i]["Varetekst"].ToString() + "'");
                }

                Logg.Log("Varekoder i databasen: " + dtVarekoder.Rows.Count, null, true);
                Logg.Log("Match i EAN databasen: " + dtTblMatch.Rows.Count, null, true);
                this.ean_nomatch_count = (dtVarekoder.Rows.Count - dtTblMatch.Rows.Count);
                this.ean_match_count = dtTblMatch.Rows.Count;
                Logg.Log("EAN med manglende varekode: " + this.ean_nomatch_count, null, true);

                Logg.Log("Oppdaterer databasen..");
                Logg.Log("Legger til strekkoder uten varekode..", null, true);

                DataTable dtFinal = dtTblMatch.Copy();

                for (int i = 0; i < dtTblCurrent.Rows.Count; i++)
                {
                    if (bw != null)
                        if (bw.CancellationPending)
                            return false;

                    DataRow[] filteredRows = dtTblMatch.Select("Barcode = '" + dtTblCurrent.Rows[i]["Barcode"] + "'");
                    if (filteredRows.Length == 0)
                    {
                        DataRow dtRow = dtFinal.NewRow();
                        dtRow["Barcode"] = dtTblCurrent.Rows[i]["Barcode"];
                        dtRow["Varekode"] = "";
                        dtRow["Varetekst"] = dtTblCurrent.Rows[i]["Varetekst"];
                        dtFinal.Rows.Add(dtRow);
                    }
                }
                this.ean_totalt_count = dtFinal.Rows.Count;
                Logg.Log("Lagt til " + (dtFinal.Rows.Count - dtTblMatch.Rows.Count) + " strekkoder uten varkode.");

                main.database.DoBulkCopy(dtFinal, "tblEan");

                Logg.Log("Fullført EAN importering!", Color.Green);
                main.processing.SetText = "Fullført importering av EAN datafil";
                return true;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return false;
        }

        private void bwEanImport_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            FormMain.appManagerIsBusy = false;
            main.ProgressStop();
            if (!e.Cancelled && e.Error == null && (bool)e.Result)
            {
                this.ean_totalt_count += this.ean_current_count;
                main.processing.SetText = "Ferdig!";
                main.processing.HideDelayed();
                Logg.Alert("Nye EAN koder med varekode: " + string.Format("{0:n0}", this.ean_match_count) + "\nNye EAN koder UTEN varekode: "
                    + string.Format("{0:n0}", this.ean_nomatch_count) + "\nNye EAN koder totalt: " + string.Format("{0:n0}",
                    (this.ean_match_count + this.ean_nomatch_count)) + "\n\nTotalt EAN koder i database (Før): "
                    + string.Format("{0:n0}", this.ean_current_count) + "\nTotalt EAN koder i databasen (Etter): "
                    + string.Format("{0:n0}", this.ean_totalt_count), "Import rapport", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (e.Cancelled)
            {
                Logg.Log("Prosessen ble avbrutt.", Color.Red);
                main.processing.SetText = "Avbrutt";
                main.processing.HideDelayed();
            }
            else
            {
                Logg.Log("Prosessen bwEanImport ble fullørt men med feil. Se logg for detaljer.", Color.Red);
                main.processing.SetVisible = false;
            }

        }

        #endregion

        private void CreateInfoTable(SQLiteConnection con, string type, DateTime dateUpdated, DateTime dateExported)
        {
            if (con.State == ConnectionState.Open)
            {
                var cmdCreateTable = new SQLiteCommand(sqlTblInfo, con);
                cmdCreateTable.ExecuteNonQuery();

                using (var cmd = new SQLiteCommand(con))
                {
                    cmd.CommandText = "INSERT INTO tblInfo (DatoOppdatert, DatoEksportert, Versjon, "
                    + "Avdeling, AvdelingNavn, Type) VALUES ("
                    + FromDateTimeToInteger(dateUpdated) + ", "
                    + FromDateTimeToInteger(dateExported) + ", '"
                    + FormMain.version + "', "
                    + main.appConfig.Avdeling + ", '"
                    + main.avdeling.Get(main.appConfig.Avdeling) + "', '"
                    + type + "');";
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void ResetEanDb()
        {
            try
            {
                if (main.connection.TableExists("tblEan"))
                {
                    var cmdDrop = new SqlCeCommand(sqlTblEanDrop, main.connection);
                    cmdDrop.ExecuteNonQuery();
                }

                var cmdCreate = new SqlCeCommand(sqlTblEanCreate, main.connection);
                cmdCreate.ExecuteNonQuery();
                var cmdAlter = new SqlCeCommand(sqlTblEanAlter, main.connection);
                cmdAlter.ExecuteNonQuery();

                Logg.Log("tblEan reset.", Color.Green);
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public static long FromDateTimeToInteger(DateTime date)
        {
            long time = date.Ticks - new DateTime(1970, 1, 1).Ticks;
            time /= TimeSpan.TicksPerSecond;
            return time * 1000;
        }

        public string Decompress(string filename)
        {
            try
            {
                if (File.Exists(filename) && filename.EndsWith(".zip"))
                {
                    Logg.Log("Pakker ut " + filename + "..", null, true);

                    string zipToUnpack = filename;
                    string unpackDirectory = System.IO.Path.GetTempPath();
                    using (ZipFile zip1 = ZipFile.Read(zipToUnpack))
                    {
                        foreach (ZipEntry e in zip1)
                        {
                            e.Extract(unpackDirectory, ExtractExistingFileAction.OverwriteSilently);
                        }
                    }

                    Logg.Log("Utpakket.", null, true);

                    if (File.Exists(System.IO.Path.GetTempPath() + @"\wobsolete.csv"))
                        return System.IO.Path.GetTempPath() + @"\wobsolete.csv";
                    else
                    {
                        Logg.Log("Feil i utpakking av Zip fil! (" + filename + ")", Color.Red);
                    }
                }
                return "";
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }
    }
}

