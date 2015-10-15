using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace KGSA
{
    public class Vinnprodukt
    {
        FormMain main;
        DataTable dtCache;
        public bool aborted = false;

        public Vinnprodukt(FormMain form)
        {
            this.main = form;
        }

        public void Update()
        {
            if (dtCache == null)
                dtCache = main.database.GetSqlDataTable("SELECT * FROM tblVinnprodukt WHERE Avdeling = " + main.appConfig.Avdeling);
        }

        public List<VinnproduktItem> GetList()
        {
            try
            {
                if (dtCache == null)
                    dtCache = main.database.GetSqlDataTable("SELECT * FROM tblVinnprodukt WHERE Avdeling = " + main.appConfig.Avdeling);

                List<VinnproduktItem> list = new List<VinnproduktItem>() { };
                for (int i = 0; i < dtCache.Rows.Count; i++)
                {
                    var prod = new VinnproduktItem();
                    prod.varekode = dtCache.Rows[i]["Varekode"].ToString();
                    prod.poeng = Convert.ToDecimal(dtCache.Rows[i]["Poeng"]);
                    prod.kategori = dtCache.Rows[i]["Kategori"].ToString();
                    prod.id = Convert.ToInt32(dtCache.Rows[i]["Id"]);
                    prod.expire = Convert.ToDateTime(dtCache.Rows[i]["DatoExpire"]);
                    prod.dato = Convert.ToDateTime(dtCache.Rows[i]["DatoOpprettet"]);
                    prod.start = Convert.ToDateTime(dtCache.Rows[i]["DatoStart"]);
                    list.Add(prod);
                }

                return list;
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<VinnproduktItem>() { };
            }
        }

        public bool Add(string varekode, string kategori, decimal poeng, DateTime slutt, DateTime start)
        {
            if (CheckDuplicate(varekode, poeng, slutt, start))
                return false; // Denne varekoden finnes allerede

            try
            {
                var con = new SqlCeConnection(FormMain.SqlConStr);
                con.Open();
                using (SqlCeCommand cmd = new SqlCeCommand("insert into tblVinnprodukt(Avdeling, Varekode, Kategori, Poeng, DatoOpprettet, DatoExpire, DatoStart) values (@Val1, @val2, @val3, @val4, @val5, @val6, @val7)", con))
                {
                    cmd.Parameters.AddWithValue("@Val1", main.appConfig.Avdeling);
                    cmd.Parameters.AddWithValue("@Val2", varekode);
                    cmd.Parameters.AddWithValue("@Val3", kategori);
                    cmd.Parameters.AddWithValue("@Val4", poeng);
                    cmd.Parameters.AddWithValue("@Val5", DateTime.Now);
                    cmd.Parameters.AddWithValue("@Val6", slutt);
                    cmd.Parameters.AddWithValue("@Val7", start);
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.ExecuteNonQuery();
                }
                con.Close();
                con.Dispose();

                return true;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return false;
        }

        private bool CheckDuplicate(string varekode, decimal poeng, DateTime slutt, DateTime start)
        {
            DataTable dt = main.database.GetSqlDataTable("SELECT * FROM tblVinnprodukt WHERE Avdeling = '" + main.appConfig.Avdeling + "' AND Varekode = '" + varekode + "'");

            if (dt == null)
                return false;

            if (dt.Rows.Count > 0)
            {
                var response = Logg.Alert("Varekoden \"" + varekode + "\" finnes allerede!\n\nBytte ut eksisterende med ny?\n\nNye verdier:\nPoeng: " + poeng + "\nStart: " + start.ToShortDateString() + "\nSlutt: " + slutt.ToShortDateString() +
                    "\n\nEksisterende:\nPoeng: " + dt.Rows[0]["Poeng"].ToString() + "\nStart: " + Convert.ToDateTime(dt.Rows[0]["DatoStart"]).ToShortDateString() + "\nSlutt: " + Convert.ToDateTime(dt.Rows[0]["DatoExpire"]).ToShortDateString(), "Overskriv", System.Windows.Forms.MessageBoxButtons.YesNoCancel, System.Windows.Forms.MessageBoxIcon.Question);
                if (response == System.Windows.Forms.DialogResult.Yes)
                {
                    Logg.Log("Bytter ut varekoden '" + varekode + "' med nye verdier.");
                    Delete(varekode);
                    return false;
                }
                else if (response == System.Windows.Forms.DialogResult.Cancel)
                    aborted = true;
                else if (response == System.Windows.Forms.DialogResult.No)
                    return true;
            }

            return false;
        }

        public bool Delete(string varekode)
        {
            try
            {
                int result = 0;
                var con = new SqlCeConnection(FormMain.SqlConStr);
                con.Open();
                using (SqlCeCommand com = new SqlCeCommand("DELETE FROM tblVinnprodukt WHERE Avdeling = " + main.appConfig.Avdeling + " AND Varekode = '" + varekode + "'", con))
                {
                    result = com.ExecuteNonQuery();
                }
                con.Close();
                con.Dispose();

                if (result > 0)
                {
                    dtCache = main.database.GetSqlDataTable("SELECT * FROM tblVinnprodukt WHERE Avdeling = " + main.appConfig.Avdeling);
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return false;
            }
        }

        public bool Remove(string id)
        {
            try
            {
                int result = 0;
                var con = new SqlCeConnection(FormMain.SqlConStr);
                con.Open();
                using (SqlCeCommand com = new SqlCeCommand("DELETE FROM tblVinnprodukt WHERE Id = " + id, con))
                {
                    result = com.ExecuteNonQuery();
                }
                con.Close();
                con.Dispose();

                if (result > 0)
                {
                    dtCache = main.database.GetSqlDataTable("SELECT * FROM tblVinnprodukt WHERE Avdeling = " + main.appConfig.Avdeling);
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return false;
            }
        }

        public void ClearAll()
        {
            try
            {
                var con = new SqlCeConnection(FormMain.SqlConStr);
                con.Open();
                SqlCeCommand command = new SqlCeCommand("DELETE FROM tblVinnprodukt WHERE Avdeling = " + main.appConfig.Avdeling, con);
                command.ExecuteNonQuery();
                Logg.Log("Vinnprodukt listen er tømt.");
                con.Close();
                con.Dispose();
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public bool ExportToCsv(string file)
        {
            try
            {
                retry:

                if (main.IsBusy(true))
                {
                    Logg.Alert("Programmet er opptatt med andre oppgaver, vent litt og prøv igjen.");
                    return false;
                }

                if (dtCache == null)
                    dtCache = main.database.GetSqlDataTable("SELECT * FROM tblVinnprodukt WHERE Avdeling = " + main.appConfig.Avdeling);

                StringBuilder sb = new StringBuilder();

                string[] columnNames = dtCache.Columns.Cast<DataColumn>().
                                                  Select(column => column.ColumnName).
                                                  ToArray();

                sb.AppendLine("Varekode;Poeng;Kategori;Fra;Til");

                //sb.AppendLine(string.Join(";", columnNames));

                foreach (DataRow row in dtCache.Rows)
                {
                    string[] fields = row.ItemArray.Select(field => field.ToString()).
                                                    ToArray();

                    sb.AppendLine(fields[3] + ";" + fields[4] + ";" + fields[2] + ";" + Convert.ToDateTime(fields[7]).ToShortDateString() + ";" + Convert.ToDateTime(fields[6]).ToShortDateString());

                    //sb.AppendLine(string.Join(";", fields));
                }

                try
                {
                    File.WriteAllText(file, sb.ToString());
                }
                catch(IOException ex)
                {
                    Logg.Unhandled(ex);
                    if (Logg.Alert("Ingen tilgang til " + file + ".\n\nForsøk igjen?", "Ingen tilgang", System.Windows.Forms.MessageBoxButtons.YesNoCancel, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes)
                        goto retry;
                    else
                        return false;
                }

                Logg.Log("Fullført lagring av CSV. file://" + file.Replace(' ', (char)160), Color.Green);
                return true;
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex, true);
            }
            return false;
        }

        public bool ImportFromCsv(string file)
        {
            try
            {
                if (main.IsBusy(true))
                {
                    Logg.Alert("Programmet er opptatt med andre oppgaver, vent litt og prøv igjen.");
                    return false;
                }

                if (!File.Exists(file))
                {
                    Logg.Log("Fant ikke CSV.", Color.Red);
                    return false;
                }

                string[] Lines;
                try
                {
                    Lines = File.ReadAllLines(file);
                }
                catch(IOException)
                {
                    Logg.Alert("Filen er i bruk av et annet program eller er utilgjengelig", "Ingen tilgang", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return false;
                }
                string[] Fields;
                Fields = Lines[0].Split(new char[] { ';' });
                int Cols = Fields.GetLength(0);
                DataTable dt = new DataTable();
                //1st row must be column names; force lower case to ensure matching later on.
                for (int i = 0; i < Cols; i++)
                    dt.Columns.Add(Fields[i].ToLower(), typeof(string));
                DataRow Row;
                for (int i = 1; i < Lines.GetLength(0); i++)
                {
                    Fields = Lines[i].Split(new char[] { ';' });
                    Row = dt.NewRow();
                    for (int f = 0; f < Cols; f++)
                        Row[f] = Fields[f];
                    dt.Rows.Add(Row);
                }

                if (dt.Rows.Count > 0)
                {
                    Logg.Status("Legger til Vinnprodukter. Vent..");
                    int teller = 0;
                    int dupli = 0;
                    string dupliVarekoder = "";
                    DateTime datoSlutt = DateTime.Now.AddYears(1);
                    DateTime datoStart = DateTime.Now.AddYears(-1);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string varekode = "";
                        if (dt.Columns.Contains("varekode"))
                            varekode = dt.Rows[i]["varekode"].ToString();

                        int poeng = 0;
                        if (dt.Columns.Contains("poeng"))
                            int.TryParse(dt.Rows[i]["poeng"].ToString(), out poeng);

                        string kategori = "Alle";
                        if (dt.Columns.Contains("kategori"))
                        {
                            string kat = dt.Rows[i]["kategori"].ToString();
                            if (kat == "Alle" || kat == "MDA" || kat == "AudioVideo" || kat == "SDA" || kat == "Tele" || kat == "Data")
                                kategori = kat;
                        }

                        if (dt.Columns.Contains("til"))
                        {
                            string slutt = dt.Rows[i]["til"].ToString();
                            DateTime.TryParse(slutt, out datoSlutt);
                        }

                        if (dt.Columns.Contains("fra"))
                        {
                            string start = dt.Rows[i]["fra"].ToString();
                            DateTime.TryParse(start, out datoStart);
                        }

                        if (varekode.Length > 3)
                            if (Add(varekode, kategori, poeng, datoSlutt, datoStart))
                                teller++;
                            else
                            {
                                dupliVarekoder += varekode + ", ";
                                dupli++;
                            }

                        if (main.vinnprodukt.aborted)
                            break;
                    }

                    if (main.vinnprodukt.aborted)
                    {
                        Logg.Log("Import avbrutt.", Color.Red);
                        main.vinnprodukt.aborted = false;
                    }
                    else
                    {
                        if (dupli > 0)
                        {
                            Logg.Log("Varekoder som fantes allerede eller inneholdte feil: " + dupliVarekoder, null, true);
                            Logg.Alert("Lagt til " + teller + " vinnprodukter.\n" + dupli + " fantes allerede eller inneholdte feil.\nSe logg for detaljer.", "Fullført import", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                        }
                        else
                            Logg.Alert("Lagt til " + teller + " vinnprodukter", "Fullført import", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }

                    // Oppdater tabel etter endring
                    Update();

                    return true;
                }
                else
                {
                    Logg.Alert("CSV var av feil format eller inneholdte ingen data.", "Import feil", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                Logg.Alert("Kunne ikke tolke CSV.\nFeil format!", "Import feil", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                Logg.Unhandled(ex);
            }
            return false;
        }
    }

    public class VinnproduktItem
    {
        public int id { get; set; }
        public string varekode { get; set; }
        public decimal poeng { get; set; }
        public string kategori {get; set;}
        public DateTime dato { get; set; }
        public DateTime expire { get; set; }
        public DateTime start { get; set; }
        public VinnproduktItem()
        {
        }
        public VinnproduktItem(string varekode, decimal poeng, string kategori)
        {
            this.varekode = varekode;
            this.poeng = poeng;
            this.kategori = kategori;
        }
    }
}
