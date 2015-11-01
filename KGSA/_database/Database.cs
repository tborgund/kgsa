using ErikEJ.SqlCe;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    public class Database
    {
        FormMain main;
        public TableBudget tableBudget;
        public TableBudgetSelger tableBudgetSelger;
        public TableBudgetTimer tableBudgetTimer;
        public TableEan tableEan;
        public TableEmail tableEmail;
        public TableHistory tableHistory;
        public TableSalg tableSalg;
        public TableSelgerkoder tableSelgerkoder;
        public TableService tableService;
        public TableServiceHistory tableServiceHistory;
        public TableServiceLogg tableServiceLogg;
        public TableUkurans tableUkurans;
        public TableVareinfo tableVareinfo;
        public TableVinnprodukt tableVinnprodukt;
        public TablePrisguide tablePrisguide;
        public TableWeekly tableWeekly;
        public TableDailyBudget tableDailyBudget;
        public DataSet sqlceDatabase;
        private DataSet sqlCache = new DataSet();

        public Database(FormMain form)
        {
            main = form;
            try
            {
                if (!File.Exists(FormMain.fileDatabase))
                {
                    SqlCeEngine engine = new SqlCeEngine(FormMain.SqlConStr);
                    engine.CreateDatabase();
                }
            }
            catch (Exception ex)
            {
                Log.ErrorDialog(ex, "Feil oppstod ved oppretting av ny database. Anbefaler omstart av programmet og/eller re-installasjon.", "KGSA Databasen");
                return;
            }

            try
            {
                OpenConnection();

                this.tableBudget = new TableBudget(main);
                this.tableBudgetSelger = new TableBudgetSelger(main);
                this.tableBudgetTimer = new TableBudgetTimer(main);
                this.tableEan = new TableEan(main);
                this.tableEmail = new TableEmail(main);
                this.tableHistory = new TableHistory(main);
                this.tableSalg = new TableSalg(main);
                this.tableSelgerkoder = new TableSelgerkoder(main);
                this.tableService = new TableService(main);
                this.tableServiceHistory = new TableServiceHistory(main);
                this.tableServiceLogg = new TableServiceLogg(main);
                this.tableUkurans = new TableUkurans(main);
                this.tableVareinfo = new TableVareinfo(main);
                this.tableVinnprodukt = new TableVinnprodukt(main);
                this.tablePrisguide = new TablePrisguide(main);
                this.tableWeekly = new TableWeekly(main);
                this.tableDailyBudget = new TableDailyBudget(main);

                VerifyDatabase();
            }
            catch (SqlCeException ex)
            {
                FormError errorMsg = new FormError("SQLCE uhåndtert unntak oppstod ved Database()", ex);
                errorMsg.ShowDialog();
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Generelt uhåndtert unntak oppstod ved Database()", ex);
                errorMsg.ShowDialog();
            }
        }

        public void CloseConnection()
        {
            if (main.connection != null)
                if (main.connection.State != ConnectionState.Closed)
                    main.connection.Close();
        }

        public void OpenConnection()
        {
            if (main.connection.State != ConnectionState.Open)
                main.connection.Open();
        }

        public void VerifyDatabase()
        {
            tableBudget.Create();
            tableBudgetSelger.Create();
            tableBudgetTimer.Create();
            tableEan.Create();
            tableEmail.Create();
            tableHistory.Create();
            tableSalg.Create();
            tableSelgerkoder.Create();
            tableService.Create();
            tableServiceHistory.Create();
            tableServiceLogg.Create();
            tableUkurans.Create();
            tableVareinfo.Create();
            tableVinnprodukt.Create();
            tablePrisguide.Create();
            tableWeekly.Create();
            tableDailyBudget.Create();

            CheckOldDatabaseElements(); // Legacy database sjekk
        }

        public void DoBulkCopy(DataTable table, string tableName)
        {
            using (SqlCeBulkCopy bc = new SqlCeBulkCopy(main.connection))
            {
                bc.DestinationTableName = tableName;
                bc.WriteToServer(table);
            }
        }

        static String BytesToString(long byteCount)
        {
            string[] suf = { " B", " KB", " MB", " GB", " TB", " PB", " EB" }; //Longs run out around EB
            if (byteCount == 0)
                return "0" + suf[0];
            long bytes = Math.Abs(byteCount);
            int place = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
            double num = Math.Round(bytes / Math.Pow(1024, place), 1);
            return (Math.Sign(byteCount) * num).ToString() + suf[place];
        }

        public void PrintStatus()
        {
            try
            {
                Log.n("Database: Henter frem database status.. vent litt");

                string size = BytesToString(new FileInfo(FormMain.fileDatabase).Length);

                Log.n("Database: Filnavn: " + FormMain.fileDatabase + " Størrelse: " + size + " Versjon: " + FormMain.databaseVersion, Color.Blue, true);

                using (SqlCeCommand command = new SqlCeCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES", main.connection))
                {
                    SqlCeDataReader reader = command.ExecuteReader();
                    DataTable table = new DataTable();
                    table.Load(reader);

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        using (var cmdScalar = new SqlCeCommand("SELECT COUNT(*) FROM " + table.Rows[i][0], main.connection))
                        {
                            int count = 0;
                            int.TryParse(cmdScalar.ExecuteScalar().ToString(), out count);
                            Log.n("Database: Tabell navn: " + table.Rows[i][0] + " Antall rader: " + count.ToString("#,##0"), Color.Blue, true);
                        }
                    }
                }
                Log.n("Database: Status ferdig. Se logg.");
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.e("Unntak ved PrintStatus(): " + ex.Message);
            }
        }

        private void CheckOldDatabaseElements()
        {
            try
            {
                if (main.connection.TableExists("tblObsolete") && !main.connection.TableExists("tblUkurans"))
                {
                    MessageBox.Show("Under validering av databasen har vi oppdaget gamle tabeller "
                        + "som ikke kan konverteres til denne versjon av programmet.\nLager beholdning "
                        + "må importeres på nytt.", "KGSA - Oppgradering av database", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Log.n("Database: Sletter gammel lager tabell (tblObsolete)..");
                    var cmd = new SqlCeCommand("DROP TABLE [tblObsolete]", main.connection);
                    cmd.ExecuteNonQuery();
                    Log.n("Database: tblObsolete fjernet.");
                }
                else if (main.connection.TableExists("tblObsolete") && main.connection.TableExists("tblUkurans"))
                {
                    var cmd = new SqlCeCommand("DROP TABLE [tblObsolete]", main.connection);
                    cmd.ExecuteNonQuery();
                    Log.n("Database: tblObsolete fjernet. (Fjernet automatisk, ny tabell finnes)");
                }

                if (main.connection.TableExists("tblVaregruppe"))
                {
                    var cmd = new SqlCeCommand("DROP TABLE [tblVaregruppe]", main.connection);
                    cmd.ExecuteNonQuery();
                    Log.n("Database: tblVaregruppe fjernet.");
                }

                if (!main.connection.FieldExists("tblSelgerkoder", "Navn")) // tblSelgerkoder mangler kolonnen Navn..
                {
                    if (MessageBox.Show("Under validering av databasen har vi oppdaget manglende kolonne(r) i selgerkode tabellen.\nEn oppgradering av databasen er nødvendig\nSkal vi starte oppgraderingen nå og forsøke å flytte over alle gamle selgerkoder?", "KGSA - Oppgradering av database", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes)
                    {
                        Log.n("Database: Henter ut gamle selgerkoder..");
                        DataTable dt = main.database.GetSqlDataTable("SELECT * FROM tblSelgerkoder");
                        Log.n("Database: Fant " + dt.Rows.Count + " selgerkoder.");

                        Log.n("Database: Oppretter ny tabell for tblSelgerkoder..");

                        tableSelgerkoder.Reset();

                        try
                        {
                            Log.n("Database: Flytter eksisterende selgerkoder tilbake..");
                            if (dt != null)
                            {
                                main.salesCodes = new SalesCodes(main);
                                int teller = 0;
                                for (int i = 0; i < dt.Rows.Count; i++)
                                {
                                    try
                                    {
                                        string sk = dt.Rows[i]["Selgerkode"].ToString();
                                        string kat = dt.Rows[i]["Kategori"].ToString();
                                        string prov = dt.Rows[i]["Provisjon"].ToString();
                                        object g;
                                        int finans = 0, mod = 0, strom = 0, rtgsa = 0;

                                        g = dt.Rows[i]["FinansKrav"];
                                        if (!DBNull.Value.Equals(g))
                                            finans = Convert.ToInt32(g);
                                        g = dt.Rows[i]["ModKrav"];
                                        if (!DBNull.Value.Equals(g))
                                            mod = Convert.ToInt32(g);
                                        g = dt.Rows[i]["StromKrav"];
                                        if (!DBNull.Value.Equals(g))
                                            strom = Convert.ToInt32(g);
                                        g = dt.Rows[i]["RtgsaKrav"];
                                        if (!DBNull.Value.Equals(g))
                                            rtgsa = Convert.ToInt32(g);

                                        if (main.salesCodes.AddAll(sk, kat, prov, finans, mod, strom, rtgsa))
                                            teller++;
                                        else
                                            Log.n("Klarte ikke konvertere " + dt.Rows[i]["Selgerkode"].ToString() + "..");
                                    }
                                    catch
                                    {
                                        Log.n("Klarte ikke konvertere " + dt.Rows[i]["Selgerkode"].ToString() + "..");
                                    }
                                }
                                Log.n("Database: Fullført flytting av " + teller + " selgerkoder.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Unhandled(ex);
                            Log.n("En feil oppstod under flytting av gamle selgerkoder til ny tabell. Se logg for detaljer.", Color.Red);
                            MessageBox.Show("En feil oppstod under konvertering av gamle selgerkoder til ny tabell. Noen selgerkoder kan være tapt.\n\nSorry!", "KGSA - Feil", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }

                        Log.n("Database: Tabell oppgradert.", Color.Green);
                    }
                    else
                        MessageBox.Show("Merk at programmet vil ikke fungere som det skal til databasen er oppgradert.\nDu vil bli påminnet om oppgraderingen neste gang programmet starter på nytt", "KGSA - Informasjon", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                if (File.Exists(FormMain.settingsPath + @"\emailRecipients.txt"))
                {
                    if (MessageBox.Show("Gammel database oppdaget.\nManglende tabeller er opprettet.\n\nØnsker du å importere gamle e-post adresser\ntil den nye adresseboken?", "KGSA - Informasjon", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes)
                    {
                        KgsaEmail email = new KgsaEmail(main);
                        email.ImportOldEmailList(FormMain.settingsPath + @"\emailRecipients.txt");

                        FormEmailAddressbook form = new FormEmailAddressbook();
                        form.ShowDialog();
                    }
                    try
                    {
                        File.Delete(FormMain.settingsPath + @"\emailRecipients.txt");
                    }
                    catch
                    {
                        Log.Alert("Fikke ikke slettet gammel epostliste, filen var låst.");
                    }
                }
            }
            catch (SqlCeException ex)
            {
                FormError errorMsg = new FormError("SQLCE uhåndtert unntak oppstod ved CheckOldDatabaseElements()", ex);
                errorMsg.ShowDialog();
                return;
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Generelt uhåndtert unntak oppstod ved CheckOldDatabaseElements()", ex);
                errorMsg.ShowDialog();
                return;
            }
        }

        public int GetCount(string sql)
        {
            try
            {
                if (main.connection.State != ConnectionState.Open)
                {
                    Log.e("Database: Forbindelsen til databasen er lukket. Vent til forbindelsen er gjenopprettet eller start om programmet.");
                    return 0;
                }

                TimeWatch tw = new TimeWatch();
                tw.Start();
                Log.DebugSql("SQL GetCount: " + sql);

                int count = 0;
                using (SqlCeCommand command = new SqlCeCommand(sql, main.connection))
                {
                    count = (Int32)command.ExecuteScalar();
                }

                Log.DebugSql("SQL GetCount: Retrieved row count (" + count + ") after " + tw.Stop() + " seconds.");
                return count;
            }
            catch (SqlCeException ex)
            {
                FormError errorMsg = new FormError("SQLCE uhåndtert unntak oppstod ved main.database.GetCount(string sql)", ex);
                errorMsg.ShowDialog();
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Generelt uhåndtert unntak oppstod ved main.database.GetCount(string sql)", ex);
                errorMsg.ShowDialog();
            }
            return 0;
        }

        public DataTable GetSqlDataTable(string sql)
        {
            try
            {
                if (main.connection.State != ConnectionState.Open)
                {
                    Log.n("Database: Forbindelsen til databasen er lukket. Vent til forbindelsen er gjenopprettet eller start om programmet.", Color.Red);
                    return null;
                }

                string hash = sql.GetHashCode().ToString();
                TimeWatch tw = new TimeWatch();
                tw.Start();
                Log.DebugSql("SQL Reader (" + hash + "): " + sql);

                if (main.appConfig.useSqlCache)
                {
                    if (hash.Length > 1 && sql.Length > 5)
                    {
                        for (int i = 0; i < sqlCache.Tables.Count; i++)
                        {
                            if (sqlCache.Tables[i].TableName.Equals(hash))
                            {
                                Log.DebugSql("SQL Reader (" + hash + "): Hentet MELLOMLAGRET tabell med " + sqlCache.Tables[i].Rows.Count + " linjer etter " + tw.Stop() + " sekunder.");
                                return sqlCache.Tables[i];
                            }
                        }
                    }
                }

                SqlCeCommand command = new SqlCeCommand(sql, main.connection);
                SqlCeDataReader reader = command.ExecuteReader();
                DataTable table = new DataTable();
                table.Load(reader);
                reader.Dispose();
                command.Dispose();
                if (main.appConfig.useSqlCache)
                {
                    bool found = false;
                    if (hash.Length > 1 && sql.Length > 5 && table.Rows.Count > 0)
                    {
                        for (int i = 0; i < sqlCache.Tables.Count; i++)
                        {
                            if (sqlCache.Tables[i].TableName.Equals(hash))
                                found = true;
                        }
                        if (!found)
                        {
                            table.TableName = hash;
                            sqlCache.Tables.Add(table.Copy());
                        }
                    }
                }

                Log.DebugSql("SQL Reader (" + hash + "): Hentet tabell med " + table.Rows.Count + " linjer etter " + tw.Stop() + " sekunder.");
                return table;
            }
            catch (SqlCeException ex)
            {
                FormError errorMsg = new FormError("SQLCE uhåndtert unntak oppstod ved main.database.GetSqlDataTable()", ex);
                errorMsg.ShowDialog();
                return null;
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Generelt uhåndtert unntak oppstod ved main.database.GetSqlDataTable()", ex);
                errorMsg.ShowDialog();
                return null;
            }
        }

        //public DateTime GetStartOfLastWholeWeek(DateTime time)
        //{
        //    DateTime startOfLastWeek = time;
        //    if (main.appConfig.ignoreSunday)
        //    {
        //        while (startOfLastWeek.DayOfWeek != DayOfWeek.Monday || (time - startOfLastWeek).TotalDays < 7)
        //            startOfLastWeek = startOfLastWeek.AddDays(-1);
        //    }
        //    else
        //    {
        //        while (startOfLastWeek.DayOfWeek != DayOfWeek.Sunday || (time - startOfLastWeek).TotalDays < 7)
        //            startOfLastWeek = startOfLastWeek.AddDays(-1);
        //    }

        //    return startOfLastWeek;
        //}

        // This presumes that weeks start with Monday.
        // Week 1 is the 1st week of the year with a Thursday in it.
        public int GetIso8601WeekOfYear(DateTime time)
        {
            // Seriously cheat.  If its Monday, Tuesday or Wednesday, then it'll 
            // be the same week# as whatever Thursday, Friday or Saturday are,
            // and we always get those right
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                time = time.AddDays(3);
            }

            // Return the week of our adjusted day
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

        public void ClearCacheTables()
        {
            this.sqlceDatabase = new DataSet();
            this.sqlCache = new DataSet();
            Log.d("Database: Cache cleared!");
        }

        public DataTable CallMonthTable(DateTime date, string avdeling)
        {
            return CallMonthTable(date, Convert.ToInt32(avdeling));
        }

        public DataTable CallMonthTable(DateTime date, int avdeling)
        {
            if (sqlceDatabase == null)
                sqlceDatabase = new DataSet();

            try
            {
                for (int i = 0; i < sqlceDatabase.Tables.Count; i++)
                {
                    if (sqlceDatabase.Tables[i].Rows.Count == 0)
                        sqlceDatabase.Tables.RemoveAt(i);
                    else
                    {
                        if (sqlceDatabase.Tables[i].Rows[0]["Avdeling"].ToString() == avdeling.ToString())
                        {
                            DateTime oldDate = Convert.ToDateTime(sqlceDatabase.Tables[i].Rows[0]["Dato"]);
                            if (oldDate.Month == date.Month && oldDate.Year == date.Year)
                            {
                                Log.d("Table cache hit (" + avdeling + " - " + date.ToString("MMMM yyyy") + ")");
                                return sqlceDatabase.Tables[i];
                            }
                        }
                    }
                }

                return CreateNewMonthTable(date, avdeling);;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.n("Feil i SQL mellomlager, forsøker ny spørring..", null, true);
            }

            DateTime first = FormMain.GetFirstDayOfMonth(date);
            DateTime last = FormMain.GetLastDayOfMonth(date);
            return GetSqlDataTable("SELECT * FROM tblSalg WHERE Avdeling = " + avdeling + " AND (Dato >= '" + first.ToString("yyy-MM-dd") + "') AND (Dato <= '" + last.ToString("yyy-MM-dd") + "')");
        }

        public DataTable CreateNewMonthTable(DateTime date, int avdeling)
        {
            try
            {
                Log.d("Table cache NO-hit (" + avdeling + " - " + date.ToString("MMMM yyyy") + ")");
                DateTime first = FormMain.GetFirstDayOfMonth(date);
                DateTime last = FormMain.GetLastDayOfMonth(date);
                DataTable month = GetSqlDataTable("SELECT * FROM tblSalg WHERE Avdeling = " + avdeling + " AND (Dato >= '" + first.ToString("yyy-MM-dd") + "') AND (Dato <= '" + last.ToString("yyy-MM-dd") + "')");
                if (month != null)
                {
                    month.Columns.Add("SalgsprisExMva", typeof(Double), "Salgspris / Mva");
                    if (month.Rows.Count > 0)
                    {
                        if (sqlceDatabase != null)
                        {
                            month.TableName = avdeling.ToString() + ":" + date.ToShortDateString();
                            sqlceDatabase.Tables.Add(month.Copy());
                        }
                        return month;  // Ny er laget, returner!
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
            return CreateEmptyMonthTable();
        }

        public static DataTable CreateEmptyMonthTable()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("SalgID", typeof(int));
            dataTable.Columns.Add("Selgerkode", typeof(string));
            dataTable.Columns.Add("Varegruppe", typeof(int));
            dataTable.Columns.Add("Varekode", typeof(string));
            dataTable.Columns.Add("Dato", typeof(DateTime));
            dataTable.Columns.Add("Antall", typeof(int));
            dataTable.Columns.Add("Btokr", typeof(decimal));
            dataTable.Columns.Add("Avdeling", typeof(int));
            dataTable.Columns.Add("Salgspris", typeof(decimal));
            dataTable.Columns.Add("Bilagsnr", typeof(int));
            dataTable.Columns.Add("Mva", typeof(decimal));
            dataTable.Columns.Add("SalgsprisExMva", typeof(Double), "Salgspris / Mva");

            return dataTable;
        }
    }

    public static class DateTimeExtensions
    {
        public static DateTime EndOfLastWorkWeek(this DateTime date, bool ignoreSunday)
        {
            if (date.DayOfWeek == DayOfWeek.Saturday)
                return date;
            else if (date.DayOfWeek == DayOfWeek.Sunday && ignoreSunday)
                return date.AddDays(-1);
            else if (date.DayOfWeek == DayOfWeek.Monday && !ignoreSunday)
                return date.AddDays(-1);
            else if (ignoreSunday)
                return date.AddDays(-7).EndOfWeek().AddDays(-1);
            else
                return date.AddDays(-7).EndOfWeek();
        }
        
        public static DateTime StartOfWeek(this DateTime dt)
        {
            int diff = 0;

            if (dt.DayOfWeek == DayOfWeek.Monday)
                diff = 0;
            else if (dt.DayOfWeek == DayOfWeek.Tuesday)
                diff = -1;
            else if (dt.DayOfWeek == DayOfWeek.Wednesday)
                diff = -2;
            else if (dt.DayOfWeek == DayOfWeek.Thursday)
                diff = -3;
            else if (dt.DayOfWeek == DayOfWeek.Friday)
                diff = -4;
            else if (dt.DayOfWeek == DayOfWeek.Saturday)
                diff = -5;
            else if (dt.DayOfWeek == DayOfWeek.Sunday)
                diff = -6;

            return dt.AddDays(diff).ChangeTime(0, 0, 0, 0);
        }

        public static DateTime EndOfWeek(this DateTime dt)
        {
            int diff = 0;
            
            if (dt.DayOfWeek == DayOfWeek.Monday)
                diff = 6;
            else if (dt.DayOfWeek == DayOfWeek.Tuesday)
                diff = 5;
            else if (dt.DayOfWeek == DayOfWeek.Wednesday)
                diff = 4;
            else if (dt.DayOfWeek == DayOfWeek.Thursday)
                diff = 3;
            else if (dt.DayOfWeek == DayOfWeek.Friday)
                diff = 2;
            else if (dt.DayOfWeek == DayOfWeek.Saturday)
                diff = 1;
            else if (dt.DayOfWeek == DayOfWeek.Sunday)
                diff = 0;

            return dt.AddDays(diff).ChangeTime(23, 59, 59, 999);
        }

        public static DateTime ChangeTime(this DateTime dt, int hours, int minutes, int seconds, int milliseconds)
        {
            return new DateTime(
                dt.Year,
                dt.Month,
                dt.Day,
                hours,
                minutes,
                seconds,
                milliseconds,
                dt.Kind);
        }
    }

    //public enum KgsaDocument
    //{
    //    // Ranking
    //    Toppselgere,
    //    Oversikt,
    //    Butikk,
    //    Knowhow,
    //    Data,
    //    Lydogbilde,
    //    Tele,
    //    Lister,
    //    Vinnprodukter,
    //    // Avdelinger
    //    Tjenester,
    //    Snittpris,
    //    // Budsjett
    //    Budsjett_mda,
    //    Budsjett_lydogbilde,
    //    Budsjett_sda,
    //    Budsjett_tele,
    //    Budsjett_data,
    //    Budsjett_cross,
    //    Budsjett_kasse,
    //    Budsjett_aftersales,
    //    Budsjett_mdasda,
    //    Budsjett_butikk,
    //    // Lager
    //    Lagerstatus,
    //    Ukuransliste,
    //    Marginer,
    //    Lagerimport,
    //    // Service
    //    Service_oversikt,
    //    Service_aktive,
    //    // Annet
    //    Ranking_periode,
    //    Ranking_rapport,
    //    Ukjent

    //}

    //public class KgsaDocument
    //{
    //    Dictionary<KgsaDocument, string> kgsaDocumentNames = new Dictionary<KgsaDocument, string>()
    //    {
    //         // Ranking
    //        { KgsaDocument.Toppselgere, "Toppselgere" },
    //        { KgsaDocument.Oversikt, "Oversikt" },
    //        { KgsaDocument.Butikk, "Butikk" },
    //        { KgsaDocument.Knowhow, "KnowHow" },
    //        { KgsaDocument.Data, "Data" },
    //        { KgsaDocument.Lydogbilde, "Lyd og Bilde" },
    //        { KgsaDocument.Tele, "Tele" },
    //        { KgsaDocument.Lister, "Lister" },
    //        { KgsaDocument.Vinnprodukter, "Vinnprodukter" },
    //        // Avdelinger
    //        { KgsaDocument.Tjenester, "Tjenester" },
    //        { KgsaDocument.Snittpris, "Snittpriser" },
    //        // Budsjett
    //        { KgsaDocument.Budsjett_mda, "Budsjett Hvitervarer" },
    //        { KgsaDocument.Budsjett_lydogbilde, "Budsjett Lyd og Bilde" },
    //        { KgsaDocument.Budsjett_sda, "Budsjett Småvarer" },
    //        { KgsaDocument.Budsjett_tele, "Budsjett Tele" },
    //        { KgsaDocument.Budsjett_data, "Budsjett Data" },
    //        { KgsaDocument.Budsjett_cross, "Budsjett Crossere" },
    //        { KgsaDocument.Budsjett_kasse, "Budsjett Kass" },
    //        { KgsaDocument.Budsjett_aftersales, "Budsjett Aftersales" },
    //        { KgsaDocument.Budsjett_mdasda, "Budsjett MDA og SDA" },
    //        { KgsaDocument.Budsjett_butikk, "Budsjett Butikken" },
    //        // Lager
    //        { KgsaDocument.Lagerstatus, "Lagerstatus" },
    //        { KgsaDocument.Ukuransliste, "Ukuransliste" },
    //        { KgsaDocument.Marginer, "Lagermarginer" },
    //        { KgsaDocument.Lagerimport, "Lagerimport" },
    //        // Service
    //        { KgsaDocument.Service_oversikt, "Serviceoveriskt" },
    //        { KgsaDocument.Service_aktive, "Aktive servicer" },
    //        // Annet
    //        { KgsaDocument.Ranking_periode, "Ranking periode" },
    //        { KgsaDocument.Ranking_rapport, "Ranking rapport" },
    //        { KgsaDocument.Ukjent, "Ukjent" }

    //    };

    //    public string GetName(KgsaDocument type)
    //    {
    //        foreach (KeyValuePair<KgsaDocument, string> entry in kgsaDocumentNames)
    //        {
    //            if (entry.Key == type)
    //                return entry.Value;
    //        }
    //        return "Ukjent";
    //    }

    //    public KgsaDocument GetType(string name)
    //    {
    //        foreach (KeyValuePair<KgsaDocument, string> entry in kgsaDocumentNames)
    //        {
    //            if (entry.Value.Equals(name))
    //                return entry.Key;
    //        }
    //        return KgsaDocument.Ukjent;
    //    }

    //}
}
