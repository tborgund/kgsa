using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;
using System.Threading;

namespace KGSA
{
    public class Selgerkoder : IDisposable
    {
        FormMain main;
        private string navn;
        private string kategori;
        private string provisjon;
        private bool disposed;
        private DataTable dtCache;
        private string[] arrayAll = new string[] { "[tom]" };
        private string[] arrayVaregruppe = new string[] { "[tom]" };

        public Selgerkoder()
        { }

        public Selgerkoder(FormMain form, bool preload = false)
        {
            this.main = form;
            if (preload)
            {
                BackgroundWorker bwLoad = new BackgroundWorker();
                bwLoad.DoWork += new DoWorkEventHandler(bwLoad_DoWork);
                bwLoad.WorkerSupportsCancellation = true;
                bwLoad.RunWorkerAsync();
            }
        }

        private int ObjectToNumber(object obj)
        {
            if (DBNull.Value == obj)
                return 0;
            else
                return Convert.ToInt32(obj);
        }

        public int GetKrav(string skArg, string skKat, string katArg, bool antall)
        {
            try
            {
                if (dtCache == null)
                    dtCache = main.database.GetSqlDataTable("SELECT * FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'");

                int krav = 0;

                for (int i = 0; i < dtCache.Rows.Count; i++)
                {
                    if (dtCache.Rows[i]["Selgerkode"].ToString() == skArg)
                    {
                        if (katArg == "Finans")
                            krav = ObjectToNumber(dtCache.Rows[i]["FinansKrav"]);
                        else if (katArg == "Mod")
                            krav = ObjectToNumber(dtCache.Rows[i]["ModKrav"]);
                        else if (katArg == "Strom")
                            krav = ObjectToNumber(dtCache.Rows[i]["StromKrav"]);
                        else if (katArg == "Rtgsa")
                            krav = ObjectToNumber(dtCache.Rows[i]["RtgsaKrav"]);
                        break;
                    }
                }

                return krav;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return 0;
        }

        public string GetKategori(string skArg)
        {
            try
            {
                if (dtCache == null)
                    dtCache = main.database.GetSqlDataTable("SELECT * FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'");

                for (int i = 0; i < dtCache.Rows.Count; i++)
                {
                    if (dtCache.Rows[i]["Selgerkode"].ToString() == skArg)
                        return dtCache.Rows[i]["Kategori"].ToString();
                }
                return "";
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return "";
        }

        public string GetNavn(string skArg)
        {
            try
            {
                if (skArg == "TOTALT" || skArg == "Andre")
                    return skArg;
                if (skArg == "INT")
                    return "Internett";

                if (dtCache == null)
                    dtCache = main.database.GetSqlDataTable("SELECT * FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'");

                skArg = skArg.Trim();
                if (main.connection.FieldExists("tblSelgerkoder", "Navn"))
                {
                    for (int i = 0; i < dtCache.Rows.Count; i++)
                    {
                        if (!DBNull.Value.Equals(dtCache.Rows[i]["Navn"]))
                            if (dtCache.Rows[i]["Selgerkode"].ToString() == skArg)
                                return dtCache.Rows[i]["Navn"].ToString();
                    }
                }
                return skArg;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return "";
        }

        private void updateData()
        {
            dtCache = main.database.GetSqlDataTable("SELECT * FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'");
            arrayAll = listSelgere();
        }

        private void bwLoad_DoWork(object sender, DoWorkEventArgs e)
        {
            System.Threading.Thread.Sleep(3000);
            updateData();
        }

        private string[] listSelgere()
        {
            try
            {
                string SQL = "SELECT Selgerkode, SUM(Antall) AS Antall FROM tblSalg WHERE (Avdeling = '"
                    + main.appConfig.Avdeling + "') AND (Dato >= '" + main.appConfig.dbTo.AddMonths(-2).ToString("yyy-MM-dd")
                    + "') AND (Dato <= '" + main.appConfig.dbTo.ToString("yyy-MM-dd") + "') GROUP BY Selgerkode ORDER BY Antall DESC";
                DataTable dtSel = main.database.GetSqlDataTable(SQL);
                if (dtSel != null)
                {
                    int size = dtSel.Rows.Count;
                    string[] array = new string[size];
                    int i;
                    for (i = 0; i < size; i++)
                    {
                        array[i] = dtSel.Rows[i][0].ToString();
                    }
                    return array;
                }
                string[] tom = new string[0];
                tom[0] = "{tom}";
                return tom;
            }
            catch (SqlCeException)
            {
                return new string[1] { "[sql error]" };
            }
            catch (Exception)
            {
                return new string[1] { "[error]" };
            }
        }

        private string[] listSelgereAndre(int avdArg)
        {
            try
            {
                string SQL = "SELECT Selgerkode, SUM(Antall) AS Antall FROM tblSalg WHERE (Avdeling = '"
                    + avdArg + "') AND (Dato >= '" + main.appConfig.dbTo.AddMonths(-2).ToString("yyy-MM-dd")
                    + "') AND (Dato <= '" + main.appConfig.dbTo.ToString("yyy-MM-dd") + "') GROUP BY Selgerkode ORDER BY Antall DESC";
                DataTable dtSel = main.database.GetSqlDataTable(SQL);
                if (dtSel != null)
                {
                    int size = dtSel.Rows.Count;
                    string[] array = new string[size];
                    int i;
                    for (i = 0; i < size; i++)
                    {
                        array[i] = dtSel.Rows[i][0].ToString();
                    }
                    return array;
                }
                string[] tom = new string[0];
                tom[0] = "{tom}";
                return tom;
            }
            catch (SqlCeException)
            {
                return new string[1] { "[sql error]" };
            }
            catch (Exception)
            {
                return new string[1] { "[error]" };
            }
        }

        /// <summary>
        /// Legg til selgerkode
        /// </summary>
        /// <param name="name">Selgerkoden</param>
        /// <param name="kat">Katergori</param>
        /// <param name="provisjon">Provisjon</param>
        /// <returns></returns>
        public bool Add(string name, string kat, string provisjon)
        {
            navn = name.Trim();
            kategori = kat;

            if (checkDuplicate())
                return false; // Denne selgerkoden finnes allerede

            try
            {
                using (SqlCeCommand cmd = new SqlCeCommand("insert into tblSelgerkoder(Selgerkode, Avdeling, Kategori, Provisjon) values (@Val1, @val2, @val3, @val4)", main.connection))
                {
                    cmd.Parameters.AddWithValue("@Val1", navn);
                    cmd.Parameters.AddWithValue("@Val2", main.appConfig.Avdeling);
                    cmd.Parameters.AddWithValue("@Val3", kategori);
                    cmd.Parameters.AddWithValue("@Val4", provisjon);
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.ExecuteNonQuery();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool AddAll(string name, string kat, string provisjon, int finansKrav, int modKrav, int stromKrav, int rtgsaKrav)
        {
            navn = name.Trim();
            kategori = kat;

            if (checkDuplicate())
                return false; // Denne selgerkoden finnes allerede

            try
            {
                using (SqlCeCommand cmd = new SqlCeCommand("insert into tblSelgerkoder(Selgerkode, Avdeling, Kategori, Provisjon, FinansKrav, ModKrav, StromKrav, RtgsaKrav) values (@Val1, @val2, @val3, @val4, @val5, @val6, @val7, @val8)", main.connection))
                {
                    cmd.Parameters.AddWithValue("@Val1", navn);
                    cmd.Parameters.AddWithValue("@Val2", main.appConfig.Avdeling);
                    cmd.Parameters.AddWithValue("@Val3", kategori);
                    cmd.Parameters.AddWithValue("@Val4", provisjon);
                    cmd.Parameters.AddWithValue("@Val5", finansKrav);
                    cmd.Parameters.AddWithValue("@Val6", modKrav);
                    cmd.Parameters.AddWithValue("@Val7", stromKrav);
                    cmd.Parameters.AddWithValue("@Val8", rtgsaKrav);
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.ExecuteNonQuery();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void Update()
        {
            dtCache = main.database.GetSqlDataTable("SELECT * FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'");
        }

        /// <summary>
        /// Metode for å sjekke antall selgerkoder.
        /// </summary>
        /// <returns>Antall selgerkoder i databasen</returns>
        public int Count()
        {
            if (dtCache == null)
                dtCache = main.database.GetSqlDataTable("SELECT * FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'");

            var rows = dtCache.Rows.Count;
            return rows;
        }

        /// <summary>
        /// Henter ut selgerkoder i string format.
        /// </summary>
        /// <param name="katArg">Kategori</param>
        /// <param name="tekniker">Med eller uten teknikere</param>
        public string[] GetSelgerkoder(string katArg = "", bool tekniker = false)
        {
            if (dtCache == null)
                dtCache = main.database.GetSqlDataTable("SELECT * FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'");
            // Selgerkode, Kategori, Provisjon
            List<string> selgere = new List<string> ();

            if (!String.IsNullOrEmpty(katArg))
            {
                if (katArg == "Computing")
                    katArg = "Data";
                if (katArg == "Telecom")
                    katArg = "Tele";

                if (dtCache.Rows.Count > 0 && tekniker)
                    for (int i = 0; i < dtCache.Rows.Count; i++)
                        if (dtCache.Rows[i]["Kategori"].ToString() == katArg || (dtCache.Rows[i]["Kategori"].ToString() == "Teknikere" && tekniker))
                            selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
                if (dtCache.Rows.Count > 0 && !tekniker)
                    for (int i = 0; i < dtCache.Rows.Count; i++)
                        if (dtCache.Rows[i]["Kategori"].ToString() == katArg || (katArg == "Aftersales" && dtCache.Rows[i]["Kategori"].ToString() == "Teknikere"))
                            selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
                        //if (dtCache.Rows[i]["Kategori"].ToString() == katArg || (dtCache.Rows[i]["Kategori"].ToString() == "Cross" && (katArg == "Data" || katArg == "Tele" || katArg == "AudioVideo")) || (katArg == "Aftersales" && dtCache.Rows[i]["Kategori"].ToString() == "Teknikere"))
                        //    selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
                if (dtCache.Rows.Count > 0 && !tekniker)
                    for (int i = 0; i < dtCache.Rows.Count; i++)
                        if (dtCache.Rows[i]["Kategori"].ToString() == "Cross" && (katArg == "Data" || katArg == "Tele" || katArg == "AudioVideo"))
                            selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
                if (dtCache.Rows.Count > 0 && tekniker)
                    for (int i = 0; i < dtCache.Rows.Count; i++)
                        if (dtCache.Rows[i]["Kategori"].ToString() == "Cross" && (katArg == "Data" || katArg == "Tele" || katArg == "AudioVideo"))
                            selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
            }
            else
            {
                for (int i = 0; i < dtCache.Rows.Count; i++)
                    selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
            }

            string[] sk = selgere.ToArray();
            
            return sk;
        }

        public string[] GetBudgetSelgerkoder(BudgetCategory cat)
        {
            if (dtCache == null)
                dtCache = main.database.GetSqlDataTable("SELECT * FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'");

            List<string> selgere = new List<string>();

            string name = BudgetCategoryClass.TypeToName(cat);

            if (cat == BudgetCategory.Butikk)
            {
                for (int i = 0; i < dtCache.Rows.Count; i++)
                    if (dtCache.Rows[i]["Kategori"].ToString() != "Teknikere" || dtCache.Rows[i]["Kategori"].ToString() != "Aftersales")
                        selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
            }
            else
            {
                for (int i = 0; i < dtCache.Rows.Count; i++)
                    if (dtCache.Rows[i]["Kategori"].ToString() == name)
                        selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
            }

            if (cat == BudgetCategory.Aftersales)
            {
                for (int i = 0; i < dtCache.Rows.Count; i++)
                    if (dtCache.Rows[i]["Kategori"].ToString() == "Teknikere")
                        selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
            }

            if (cat == BudgetCategory.MDASDA)
            {
                for (int i = 0; i < dtCache.Rows.Count; i++)
                    if (dtCache.Rows[i]["Kategori"].ToString() == "MDA" || dtCache.Rows[i]["Kategori"].ToString() == "SDA")
                        selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
            }

            if (main.appConfig.budgetInclAllSalesRepUnderCross && cat == BudgetCategory.Cross)
            {
                for (int i = 0; i < dtCache.Rows.Count; i++)
                    if (dtCache.Rows[i]["Kategori"].ToString() == "AudioVideo" || dtCache.Rows[i]["Kategori"].ToString() == "Tele" || dtCache.Rows[i]["Kategori"].ToString() == "Data")
                        selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
            }

            if (!main.appConfig.budgetIsolateCrossSalesRep && (cat == BudgetCategory.AudioVideo || cat == BudgetCategory.Tele || cat == BudgetCategory.Data))
            {
                for (int i = 0; i < dtCache.Rows.Count; i++)
                    if (dtCache.Rows[i]["Kategori"].ToString() == "Cross")
                        selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
            }

            string[] sk = selgere.ToArray();

            return sk;
        }

        public string[] GetVinnSelgerkoder(BudgetCategory cat)
        {
            if (dtCache == null)
                dtCache = main.database.GetSqlDataTable("SELECT * FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'");

            List<string> selgere = new List<string>();
            if (cat == BudgetCategory.MDASDA)
            {
                for (int i = 0; i < dtCache.Rows.Count; i++)
                    if (dtCache.Rows[i]["Kategori"].ToString() == "MDA" || dtCache.Rows[i]["Kategori"].ToString() == "SDA")
                        selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
            }
            else if (cat == BudgetCategory.Cross)
            {
                for (int i = 0; i < dtCache.Rows.Count; i++)
                    if (dtCache.Rows[i]["Kategori"].ToString() == "AudioVideo" || dtCache.Rows[i]["Kategori"].ToString() == "Tele" || dtCache.Rows[i]["Kategori"].ToString() == "Data" || dtCache.Rows[i]["Kategori"].ToString() == "Cross")
                        selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
            }
            else if (cat == BudgetCategory.Kasse)
            {
                for (int i = 0; i < dtCache.Rows.Count; i++)
                    if (dtCache.Rows[i]["Kategori"].ToString() == "Kasse")
                        selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
            }
            else if (cat == BudgetCategory.Aftersales)
            {
                for (int i = 0; i < dtCache.Rows.Count; i++)
                    if (dtCache.Rows[i]["Kategori"].ToString() == "Aftersales" || dtCache.Rows[i]["Kategori"].ToString() == "Teknikere")
                        selgere.Add(dtCache.Rows[i]["Selgerkode"].ToString());
            }
            string[] sk = selgere.ToArray();

            return sk;
        }

        /// <summary>
        /// Henter en liste på selgere fra avdelingen de siste to månedene.
        /// </summary>
        /// <returns>Array med selgere</returns>
        public string[] GetAlleSelgerkoder(int avdArg = 0)
        {
            if (avdArg > 1000)
                return listSelgereAndre(avdArg);
            return arrayAll;
        }

        public string[] GetVarekoder()
        {
            return arrayVaregruppe;
        }

        /// <summary>
        /// Hent provisjon type
        /// </summary>
        /// <param name="skArg">Selgerkoden</param>
        /// <returns></returns>
        public string GetProvisjon(string skArg)
        {
            if (dtCache == null)
                dtCache = main.database.GetSqlDataTable("SELECT * FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'");

            string prov = "";
            if (dtCache.Rows.Count > 0)
            {
                DataRow[] row = dtCache.Select("[Selgerkode] = '" + skArg + "'");
                if (!DBNull.Value.Equals(row))
                    foreach (var item in row)
                        prov = item["Provisjon"].ToString();
                else
                    prov = "";
            }
            return prov;
        }

        /// <summary>
        /// Hent selgerkoden med TeknikerAlle provisjon hvis avdelingen har
        /// </summary>
        /// <returns>Selgerkoden eller ingen hvis butikken ikke har TeknikerAlle</returns>
        public string GetTeknikerAlle()
        {
            try
            {
                string sql = "SELECT TOP 1 Selgerkode FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "' AND Provisjon = 'TeknikerAlle'";
                var cmd = new SqlCeCommand(sql, main.connection);
                string result = (string)cmd.ExecuteScalar();
                cmd.Dispose();

                if (!String.IsNullOrEmpty(result))
                    return result;

                return "";
            }
            catch
            {
            }
            return "";
        }

        /// <summary>
        /// Sletter selgerkode for denne butikken
        /// </summary>
        public void Clear()
        {
            SqlCeCommand cmd = main.connection.CreateCommand();
            cmd.CommandText = "DELETE FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'";
            cmd.ExecuteScalar();
            if (dtCache != null)
                dtCache.Rows.Clear();
        }

        private bool checkDuplicate()
        {
            SqlCeCommand cmd = main.connection.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM tblSelgerkoder WHERE Selgerkode = '" + navn + "' AND Avdeling = '" + main.appConfig.Avdeling + "'";
            int result = ((int)cmd.ExecuteScalar());
            if (result == 0)
                return false;
            return true;
        }

        /// <summary>
        /// Slett selgerkoder som ligger dobbelt
        /// </summary>
        /// <returns>Rapport i streng format</returns>
        public string DeleteDuplicates()
        {
            var command = new SqlCeCommand("SELECT * FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'", main.connection);
            var dt = main.database.GetSqlDataTable("SELECT * FROM tblSelgerkoder WHERE Avdeling = '" + main.appConfig.Avdeling + "'");

            int result = 0; bool TeknikerAlle = false; string returnMsg = "";
            object r = dt.Compute("Count(Selgerkode)", "[Provisjon] = 'TeknikerAlle'");
                if (!DBNull.Value.Equals(r))
            result = Convert.ToInt32(r);
                else
            result = 0;
            if (result > 0)
                TeknikerAlle = true;

            var da = new SqlCeDataAdapter(command);
            var ds = new DataSet();
            da.Fill(ds, "tblSelgerkoder");
            var ca = new SqlCeCommandBuilder(da);

            bool slettTeknikerAlle = false;
            for (int i = 0; i < ds.Tables["tblSelgerkoder"].Rows.Count; i++)
            {
                try
                {
                    string selger = ds.Tables["tblSelgerkoder"].Rows[i]["Provisjon"].ToString().Trim();
                    if (selger.Length > 0)
                        ds.Tables["tblSelgerkoder"].Rows[i]["Provisjon"] = selger;
                }
                catch { }

                provisjon = ds.Tables["tblSelgerkoder"].Rows[i]["Provisjon"].ToString();
                navn = ds.Tables["tblSelgerkoder"].Rows[i]["Selgerkode"].ToString();
                kategori = ds.Tables["tblSelgerkoder"].Rows[i]["Kategori"].ToString();

                if (TeknikerAlle && provisjon == "TeknikerAlle")
                {
                    slettTeknikerAlle = true;
                    continue;
                }

                result = 0;
                r = dt.Compute("Count(Selgerkode)", "[Selgerkode] = '" + navn + "'");
                if (!DBNull.Value.Equals(r))
                    result = Convert.ToInt32(r);
                else
                    result = 0;

                if (result > 1)
                {
                    returnMsg = "Selgerkoder: Fant duplikat. Slettet selgerkode " + navn + " med provisjon " + provisjon;
                    ds.Tables["tblSelgerkoder"].Rows[i].Delete();
                    break;
                }

                if (TeknikerAlle && provisjon == "Tekniker")
                {
                    returnMsg = "Selgerkoder: Med TeknikerAlle valgt på en selgerkode, kan ikke andre motta tekniker-provisjon. Slettet selgerkode " + navn + " med provisjon " + provisjon;
                    ds.Tables["tblSelgerkoder"].Rows[i].Delete();
                    break;
                }
                if (TeknikerAlle && provisjon == "TeknikerAlle" && slettTeknikerAlle)
                {
                    returnMsg = "Selgerkoder: Fant duplikat av TeknikerAlle provisjon. Bare en selgerkode kan motta tekniker-provisjon på alle salg. Slettet selgerkode " + navn + " med provisjon " + provisjon;
                    ds.Tables["tblSelgerkoder"].Rows[i].Delete();
                    break;
                }
            }

            da.Update(ds, "tblSelgerkoder");

            return returnMsg;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    //dispose managed ressources
                }
            }
            //dispose unmanaged ressources
            dtCache = null;
            disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
