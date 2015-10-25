using FileHelpers;
using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
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

        #endregion
        public AppManager(FormMain form)
        {
            main = form;
        }

        public void UpdateAllAsync()
        {
            main.worker = new BackgroundWorker();
            main.worker.DoWork += new DoWorkEventHandler(worker_DoWork);
            main.worker.ProgressChanged += new ProgressChangedEventHandler(main.bwProgressReport_ProgressChanged);
            main.worker.WorkerReportsProgress = true;
            main.worker.WorkerSupportsCancellation = true;
            main.worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(worker_Completed);

            main.processing.SetVisible = true;
            main.processing.SetText = "Forbereder..";
            main.processing.SetProgressStyle = ProgressBarStyle.Marquee;
            main.worker.RunWorkerAsync();
            main.processing.SetBackgroundWorker = main.worker;
        }

        public bool UpdateAll(BackgroundWorker bw)
        {
            try
            {
                if (main.IsBusy())
                    return false;

                FormMain.appManagerIsBusy = true;
                main.ProgressStart();
                if (UpdateInventoryDatabase(bw) && UpdateProductDatabase(bw))
                {
                    FormMain.appManagerIsBusy = false;
                    main.ProgressStop();
                    main.appConfig.blueServerDatabaseUpdated = DateTime.Now;
                    main.SaveSettings();
                    return true;
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.e("Uventet feil oppstod under oppdatering av App-databasene. Se logg for detaljer.");
            }
            FormMain.appManagerIsBusy = false;
            main.ProgressStop();
            return false;
        }

        internal void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            FormMain.appManagerIsBusy = true;
            main.ProgressStart();

            bool invUpdated = UpdateInventoryDatabase(main.worker);
            if (invUpdated && File.Exists(FormMain.settingsPath + @"\" + BluetoothServer.inventoryFilename))
                Log.n("Varebeholdnings databasen er klar for App.", Color.Green);

            bool dataUpdated = UpdateProductDatabase(main.worker);
            if (dataUpdated && File.Exists(FormMain.settingsPath + @"\" + BluetoothServer.dataFilename))
                Log.n("Produkt databasen er klar for App.", Color.Green);

            e.Result = invUpdated && dataUpdated;

            if (main.worker != null && main.worker.CancellationPending)
                e.Cancel = true;
        }

        internal void worker_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            FormMain.appManagerIsBusy = false;
            main.ProgressStop();
            main.processing.SetProgressStyle = ProgressBarStyle.Continuous;
            main.processing.SetValue = 100;

            if (!e.Cancelled && e.Error == null && (bool)e.Result)
            {
                main.appConfig.blueServerDatabaseUpdated = DateTime.Now;
                main.SaveSettings();
                main.processing.SetText = "Ferdig";
                main.processing.HideDelayed();
            }
            else if (e.Cancelled)
            {
                Log.e("Prosessen ble avbrutt.");
                main.processing.SetText = "Avbrutt";
                main.processing.HideDelayed();
            }
            else
            {
                main.processing.SetText = "Prosessen ble avbrutt";
                main.processing.HideDelayed();
                Log.e("Prosessen ble fullført, men med feil. Se logg for detaljer.");
            }
        }

        #region Update inventory database

        public bool UpdateInventoryDatabase(BackgroundWorker bw)
        {
            try
            {
                main.processing.SetProgressStyle = ProgressBarStyle.Marquee;
                main.processing.SetText = "Henter produkt-data..";
                Log.n("Henter produkt-data..");
                DataTable table = main.database.tableUkurans.GetInventory(main.appConfig.Avdeling, bw);
                if (table == null)
                    throw new NullReferenceException("Tabellen tableUkurans returnerte NULL");

                int count = table.Rows.Count;
                if (count == 0)
                {
                    Log.e("Databasen inneholder ingen lagervarer. Importer lagervarer først");
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
                        Log.Unhandled(ex);
                        Log.n("Import: Eksport destinasjonen " + database + " eksisterer og er låst.", Color.Red);
                        return false;
                    }
                }

                Log.n("Prosesserer " + string.Format("{0:n0}", count) + " vareoppføringer..");
                main.processing.SetProgressStyle = ProgressBarStyle.Continuous;
                main.processing.SetText = "Prosesserer " + string.Format("{0:n0}", count) + " vareoppføringer..";

                long TICKS_AT_EPOCH = 621355968000000000L;
                long TICKS_PER_MILLISECOND = 10000;

                DateTime lastDate = FormMain.rangeMin;

                SQLiteConnection con = new SQLiteConnection("Data Source=" + database + ";Version=3;");
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
                            for (var i = 0; i < count; i++)
                            {
                                if (bw.CancellationPending)
                                    return false;

                                if (i % 321 == 0 && bw.WorkerReportsProgress)
                                {
                                    bw.ReportProgress(i, new StatusProgress(count, null, 0, 100));
                                    main.processing.SetText = "Lagrer vareoppføringer..: "
                                        + i.ToString("#,##0") + " / " + count.ToString("#,##0");
                                }

                                DateTime date = Convert.ToDateTime(table.Rows[i][TableUkurans.KEY_DATO]);
                                if (date > lastDate)
                                    lastDate = date;

                                cmd.CommandText = "INSERT INTO tblInventory (Id, Avdeling, Varekode, Varetekst, Kategori, KategoriNavn,"
                                    + " Varegruppe, VaregruppeNavn, Modgruppe, ModgruppeNavn, Merke, MerkeNavn, Antall, Kost, Dato, UkuransMnd, UkuransVerdi, UkuransProsent) VALUES (" + i + ", "
                                    + table.Rows[i][TableUkurans.KEY_AVDELING] + ", '"
                                    + table.Rows[i][TableVareinfo.KEY_VAREKODE] + "', '"
                                    + table.Rows[i][TableVareinfo.KEY_TEKST].ToString().Replace("'", "''") + "', "
                                    + table.Rows[i][TableVareinfo.KEY_KAT] + ", '"
                                    + table.Rows[i][TableVareinfo.KEY_KATNAVN] + "', "
                                    + table.Rows[i][TableVareinfo.KEY_GRUPPE] + ", '"
                                    + table.Rows[i][TableVareinfo.KEY_GRUPPENAVN] + "', "
                                    + table.Rows[i][TableVareinfo.KEY_MODGRUPPE] + ", '"
                                    + table.Rows[i][TableVareinfo.KEY_MODGRUPPENAVN] + "', "
                                    + table.Rows[i][TableVareinfo.KEY_MERKE] + ", '"
                                    + table.Rows[i][TableVareinfo.KEY_MERKENAVN] + "', "
                                    + table.Rows[i][TableUkurans.KEY_ANTALL] + ", "
                                    + table.Rows[i][TableUkurans.KEY_KOST].ToString().Replace(" ", "").Replace(",", ".") + ", "
                                    + ((date.ToUniversalTime().Ticks - TICKS_AT_EPOCH) / TICKS_PER_MILLISECOND).ToString() + ", "
                                    + "0" + ", "
                                    + table.Rows[i][TableUkurans.KEY_UKURANS].ToString().Replace(" ", "").Replace(",", ".") + ", "
                                    + table.Rows[i][TableUkurans.KEY_UKURANSPROSENT].ToString().Replace(" ", "").Replace(",", ".") + ");";

                                cmd.ExecuteNonQuery();

                            }
                            transaction.Commit();
                        }

                    }

                    main.appConfig.blueInventoryExportDate = lastDate;
                    main.appConfig.blueInventoryLastDate = lastDate;

                    CreateInfoTable(conSqlite, "Inventory", main.appConfig.blueInventoryLastDate, main.appConfig.blueInventoryExportDate);
                    Log.d("Opprettet inventory info tabell");

                    conSqlite.Close();
                }

                Log.d("Varebeholdning siste dato: " + lastDate.ToShortDateString() + " og eksportert: " + lastDate.ToShortDateString());

                return true;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.e("Uventet feil oppstod under oppdatering av App-databasen Inventory. Se logg for detaljer");
            }
            return false;
        }

        #endregion

        #region Update Product data

        public bool UpdateProductDatabase(BackgroundWorker bw)
        {
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
                    Log.Unhandled(ex);
                    Log.n("Eksport: Databasen er låst og kan ikke slettes/fornyes: " + database + ". Exception: " + ex.Message);
                    return false;
                }
            }

            SQLiteConnection sqliteCon = new SQLiteConnection("Data Source=" + database + ";Version=3;");
            try
            {
                Log.d("Henter varedata fra databasen..");
                main.processing.SetText = "Henter varedata..";
                DataTable table = main.database.tableVareinfo.GetAllProducts();
                if (table == null)
                    throw new Exception("Intern feil i databasen. tblVareinfo var null!");

                if (table.Rows.Count == 0)
                {
                    Log.e("Fant ingen varer i databasen. Har du importert lagervarer fra Elguide?");
                    return false;
                }

                main.processing.SetText = "Henter EAN koder..";
                DataTable tableEan = main.database.tableEan.GetAllRows();
                if (tableEan == null)
                    throw new Exception("Intern feil i databasen. tblEan var null!");

                if (tableEan.Rows.Count == 0)
                {
                    Log.e("Fant ingen EAN koder i databasen. Har du importert EAN koder fra Elguide?");
                    return false;
                }
                table.Columns.Add(new DataColumn("Barcode", typeof(string)));

                int countvare = table.Rows.Count;
                Log.d("Henter ut EAN koder.. (" + countvare + ")");
                try
                {
                    for (int i = 0; i < countvare; i++)
                    {
                        if (bw != null && bw.CancellationPending)
                            return false;

                        var filter = tableEan.Select(TableEan.KEY_PRODUCT_CODE + " = '" + table.Rows[i]["Varekode"] + "'");
                        if (filter.Length == 1)
                            table.Rows[i]["Barcode"] = filter[0][TableEan.KEY_BARCODE].ToString();
                    }
                }
                catch (Exception ex)
                {
                    Log.Unhandled(ex);
                    Log.e("Kritisk feil ved prosessering av EAN tabellen");
                    return false;
                }

                DataTable tableEanOnly = tableEan.Select(TableEan.KEY_PRODUCT_CODE + " = ''").CopyToDataTable();

                Log.d("Setter EAN kode til varedata.. (Antall: " + tableEanOnly.Rows.Count + ")");

                int countEanOnly = tableEanOnly.Rows.Count;
                for (int i = 0; i < countEanOnly; i++)
                {
                    if (bw != null && bw.CancellationPending)
                        return false;

                    DataRow dtRow = table.NewRow();
                    dtRow["Barcode"] = tableEanOnly.Rows[i][TableEan.KEY_BARCODE];
                    dtRow["Varetekst"] = tableEanOnly.Rows[i][TableEan.KEY_PRODUCT_TEXT];
                    dtRow["Kategori"] = 0;
                    dtRow["Varegruppe"] = 0;
                    dtRow["Modgruppe"] = 0;
                    dtRow["Merke"] = 0;
                    dtRow["Dato"] = FormMain.rangeMin;
                    table.Rows.Add(dtRow);
                }

                main.processing.SetText = "Behandler " + string.Format("{0:n0}", table.Rows.Count) + " varekoder..";
                Log.n("Behandler " + string.Format("{0:n0}", table.Rows.Count) + " varekoder..");

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
                        Log.d("Saving produkt-data to sqlite database " + database + "..");
                        using (var transaction = con.BeginTransaction())
                        {
                            int count = table.Rows.Count;
                            for (var i = 0; i < count; i++)
                            {
                                if (bw != null && bw.CancellationPending)
                                    break;

                                if (i % 83 == 0)
                                {
                                    bw.ReportProgress(i, new StatusProgress(count, null, 0, 100));
                                    main.processing.SetText = "Lagrer produktdata..: "
                                        + string.Format("{0:n0}", i) + " / " + string.Format("{0:n0}", count);
                                }

                                decimal dec = 0M;
                                decimal.TryParse(table.Rows[i]["Salgspris"].ToString(), out dec);

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
                        Log.d("Ferdig med å flytte resultat over til databasen.");
                    }

                    main.appConfig.blueProductExportDate = lastDate;
                    main.appConfig.blueProductLastDate = lastDate;

                    CreateInfoTable(con, "ProductData", main.appConfig.blueProductLastDate, main.appConfig.blueInventoryExportDate);
                    Log.d("Opprettet produktdata info tabell");

                    con.Close();
                }

                Log.d("Produkt-data siste dato: " + lastDate.ToShortDateString() + " og eksportert: " + lastDate.ToShortDateString());

                return true;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.n("Feil ved eksport av produtdata: " + ex.Message);
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

        public void ImportEan()
        {
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

                    main.worker = new BackgroundWorker();
                    main.worker.DoWork += new DoWorkEventHandler(bwEanImport_DoWork);
                    main.worker.ProgressChanged += new ProgressChangedEventHandler(main.bwProgressReport_ProgressChanged);
                    main.worker.WorkerReportsProgress = true;
                    main.worker.WorkerSupportsCancellation = true;
                    main.worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwEanImport_Completed);
                    main.worker.RunWorkerAsync(eanfiles);
                    main.processing.SetBackgroundWorker = main.worker;
                }
                fdlg.Dispose();
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private void bwEanImport_DoWork(object sender, DoWorkEventArgs e)
        {
            FormMain.appManagerIsBusy = true;
            main.ProgressStart();

            List<string> eanfiles = (List<string>)e.Argument;
            if (eanfiles != null)
                e.Result = ImportEanCodes(eanfiles);
            else
                throw new ArgumentNullException("No files to import was selected = Returned NULL");

            if (main.worker.CancellationPending)
                e.Cancel = true;
        }

        private bool ImportEanCodes(List<string> eanfiles)
        {
            try
            {
                DataTable tableFile = new DataTable();
                tableFile.Columns.Add(TableEan.KEY_BARCODE, typeof(string));
                tableFile.Columns.Add(TableEan.KEY_PRODUCT_TEXT, typeof(string));

                foreach (string file in eanfiles)
                {
                    if (main.worker != null && main.worker.CancellationPending)
                        return false;

                    if (!File.Exists(file) || string.IsNullOrEmpty(file))
                        throw new IOException("Fant ikke fil eller ble nektet tilgang: " + file);

                    Log.n("Leser bærbar telle-penn fil \"" + file + "\"..");

                    using (StreamReader fileReader = new StreamReader(file, System.Text.Encoding.GetEncoding(865)))
                    {
                        string fileLine;
                        while ((fileLine = fileReader.ReadLine()) != null)
                        {
                            if (main.worker != null && main.worker.CancellationPending)
                                return false;

                            string line = fileLine.Trim();
                            if (line.Length > 13)
                            {
                                if (Regex.IsMatch(line.Substring(0, 13), @"^\d+$"))
                                {
                                    string tekst = line.Substring(13, line.Length - 13);

                                    DataRow dtRow = tableFile.NewRow();
                                    dtRow[0] = line.Substring(0, 13);
                                    if (tekst.Length > 29)
                                        dtRow[1] = line.Substring(13, 29).Trim();
                                    else
                                        dtRow[1] = line.Substring(13, line.Length - 13).Trim();
                                    tableFile.Rows.Add(dtRow);
                                }
                            }
                        }
                        fileReader.Close();
                    }
                }

                int countFile = tableFile.Rows.Count;
                if (countFile == 0)
                {
                    Log.e("Fant ingen EAN koder fra fil(er)");
                    return false;
                }
                Log.d(string.Format("{0:n0}", countFile) + " rows read from file(s)");

                Log.n("Henter EAN koder fra databasen..");
                main.processing.SetText = "Henter EAN koder fra databasen..";

                DataTable tableEan = main.database.tableEan.GetAllRows();
                if (tableEan == null)
                    throw new NullReferenceException("DataTable tableEan retrieved from database returned null");

                int countEan = tableEan.Rows.Count;
                Log.d(string.Format("{0:n0}", countEan) + " EAN codes read from database");

                Log.n("Behandler EAN databasen..");
                main.processing.SetText = "Behandler EAN databasen..";

                TimeWatch time = new TimeWatch();
                time.Start();
                for (int i = 0; i < countFile; i++)
                {
                    if (main.worker.CancellationPending)
                        return false;

                    if (i % 134 == 0)
                    {
                        main.worker.ReportProgress(i, new StatusProgress(countFile, null, 0, 30));
                        main.processing.SetText = "Leser EAN koder..: "
                            + string.Format("{0:n0}", i) + " / " + string.Format("{0:n0}", countFile);
                    }

                    DataRow[] result = tableEan.Select(TableEan.KEY_BARCODE + " = '" + tableFile.Rows[i][TableEan.KEY_BARCODE] + "'");
                    if (result.Count() == 0)
                    {
                        DataRow dtRow = tableEan.NewRow();
                        dtRow[TableEan.INDEX_BARCODE] = tableFile.Rows[i][0];
                        dtRow[TableEan.INDEX_PRODUCT_TEXT] = tableFile.Rows[i][1];
                        dtRow[TableEan.INDEX_PRODUCT_CODE] = "";
                        tableEan.Rows.Add(dtRow);
                    }
                    else if (result.Count() >= 1)
                    {
                        for (int d = 0; d < result.Count(); d++)
                        {
                            if (!result[d][TableEan.INDEX_PRODUCT_TEXT].Equals(tableFile.Rows[i][1]))
                            {
                                Log.d("Updating known EAN (" + result[d][TableEan.INDEX_BARCODE] + ") with new data..");
                                result[0][TableEan.INDEX_BARCODE] = tableFile.Rows[i][0];
                                result[0][TableEan.INDEX_PRODUCT_TEXT] = tableFile.Rows[i][1];
                                result[0][TableEan.INDEX_PRODUCT_CODE] = "";
                                result[0].EndEdit();
                                tableEan.AcceptChanges();
                            }
                        }
                    }
                    else
                        Log.e("Error while searching for EAN (" + tableFile.Rows[i][TableEan.KEY_BARCODE] + ") - This should not happen!");
                }
                int countAfter = tableEan.Rows.Count;
                int countNew = countAfter - countEan;

                Log.d("Searched " + string.Format("{0:n0}", countFile) + " EAN codes in " + time.show + " seconds");
                if (countNew == 0)
                {
                    Log.n("Fullført EAN importering: Ingen nye EAN koder funnet", Color.Blue);
                    return true;
                }
                else
                {
                    Log.n("Fant " + string.Format("{0:n0}", countNew) + " EAN koder. Starter søk i databasen etter varekoder..");
                    return MatchEanToDatabase(tableEan);
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.e("Kritisk feil i ImportEanCodes: Importering avbrutt. Se logg for detaljer.");
            }
            return false;
        }

        private bool MatchEanToDatabase(DataTable tableEan)
        {
            try
            {
                DataView view = tableEan.DefaultView;
                view.Sort = TableEan.KEY_PRODUCT_TEXT + " ASC";
                tableEan = view.ToTable();

                List<string> notFoundList = new List<string> { };
                int countEans = tableEan.Rows.Count;
                int countProductCodes = 0;
                int countFound = 0;
                int countNotFound = 0;
                string sql = "SELECT " + TableVareinfo.KEY_VAREKODE + ", SUBSTRING(" + TableVareinfo.KEY_TEKST
                    + ", 1, 29) AS " + TableVareinfo.KEY_TEKST + " FROM " + TableVareinfo.TABLE_NAME;

                main.processing.SetText = "Henter varekoder fra databasen..";

                using (DataTable tableProducts = main.database.GetSqlDataTable(sql))
                {
                    if (tableProducts == null)
                        throw new NullReferenceException("DataTable tableInfo retrieved from database returned NULL");

                    countProductCodes = tableProducts.Rows.Count;
                    if (countProductCodes == 0)
                    {
                        Log.e("Ingen varer funnet i databasen. Importer lager først");
                        return false;
                    }
                    Log.d("Started matching new EAN codes to product database..");

                    TimeWatch time = new TimeWatch();
                    time.Start();

                    for (int i = 0; i < countEans; i++)
                    {
                        if (main.worker.CancellationPending)
                            return false;

                        if (i % 114 == 0)
                        {
                            main.worker.ReportProgress(i, new StatusProgress(countProductCodes, null, 30, 60));
                            main.processing.SetText = "Søker i produkt databasen..: "
                                + string.Format("{0:n0}", i) + " / " + string.Format("{0:n0}", countEans);
                        }

                        DataRow[] result = tableProducts.Select("Varetekst = '" + tableEan.Rows[i][TableEan.INDEX_PRODUCT_TEXT].ToString().Replace("'", "''") + "'");
                        if (result.Count() > 0)
                        {
                            tableEan.Rows[i][TableEan.INDEX_PRODUCT_CODE] = result[0][0];
                            countFound++;
                        }
                        else
                        {
                            if (notFoundList.Count < 25)
                                notFoundList.Add(countNotFound + ";" + tableEan.Rows[i][TableEan.INDEX_BARCODE] + ";" + tableEan.Rows[i][TableEan.INDEX_PRODUCT_TEXT]);
                            countNotFound++;
                        }
                    }
                    Log.d("Matching search took " + time.show + " seconds");

                    if (main.appConfig.debug)
                    {
                        Log.d(string.Format("{0:n0}", countFound) + " new rows with matching ProductCodes");
                        Log.d(string.Format("{0:n0}", countNotFound) + " rows with no ProductCodes");
                        Log.d("-------------- NOT FOUND ------------------ START");
                        foreach (string line in notFoundList)
                            Log.d(line);
                        Log.d("-------------- NOT FOUND ------------------ END");
                    }
                }

                if (countFound > 0)
                {
                    Log.n("Lagrer endringer i EAN databsen.. (" + string.Format("{0:n0}", countEans) + " oppføringer)");
                    main.processing.SupportsCancelation = false;
                    main.processing.SetText = "Lagrer endringer i EAN databasen.. (" + string.Format("{0:n0}", countEans) + " oppføringer)";
                    main.database.tableEan.Reset();
                    main.database.DoBulkCopy(tableEan, TableEan.TABLE_NAME);

                    Log.n("Fullført EAN importering: EAN databasen oppdatert med " + string.Format("{0:n0}", countFound) + " nye varekoder", Color.Green);
                }
                else
                    Log.n("Fullført EAN importering: Ingen nye EAN koder funnet etter å ha søkt igjennom " + string.Format("{0:n0}", countEans) + " koder", Color.Blue);

                return true;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.e("Kritisk feil i MatchEanToDatabase: Importering avbrutt. Se logg for detaljer.");
            }
            return false;
        }

        private void bwEanImport_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            FormMain.appManagerIsBusy = false;
            main.ProgressStop();
            main.processing.SetProgressStyle = ProgressBarStyle.Continuous;
            main.processing.SetValue = 100;

            if (!e.Cancelled && e.Error == null && (bool)e.Result)
            {
                main.processing.SetText = "Ferdig!";
                main.processing.HideDelayed();
            }
            else if (e.Cancelled)
            {
                Log.e("Prosessen ble avbrutt");
                main.processing.SetText = "Avbrutt";
                main.processing.HideDelayed();
            }
            else
            {
                Log.e("EAN importering ble fullørt men med feil. Se logg for detaljer.");
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

        public static long FromDateTimeToInteger(DateTime date)
        {
            long time = date.Ticks - new DateTime(1970, 1, 1).Ticks;
            time /= TimeSpan.TicksPerSecond;
            return time * 1000;
        }
    }
}

