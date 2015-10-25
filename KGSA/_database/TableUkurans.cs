using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
   public class TableUkurans
    {
        FormMain main;
        public static string TABLE_NAME = "tblUkurans";

        public static string KEY_ID = "Id";
        public static string KEY_AVDELING = "Avdeling";
        public static string KEY_VAREKODE = "Varekode";
        public static string KEY_ANTALL = "Antall";
        public static string KEY_KOST = "Kost";
        public static string KEY_DATO = "Dato";
        public static string KEY_UKURANS = "UkuransVerdi";
        public static string KEY_UKURANSPROSENT = "UkuransProsent";

        public static int INDEX_AVDELING = 0;
        public static int INDEX_VAREKODE = 1;
        public static int INDEX_ANTALL = 2;
        public static int INDEX_KOST = 3;
        public static int INDEX_DATO = 4;
        public static int INDEX_UKURANS = 5;
        public static int INDEX_UKURANSPROSENT = 6;

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_AVDELING + "] int NOT NULL "
            + ", [" + KEY_VAREKODE + "] nvarchar(25) NOT NULL "
            + ", [" + KEY_ANTALL + "] int NOT NULL "
            + ", [" + KEY_KOST + "] money NULL "
            + ", [" + KEY_DATO + "] datetime NULL "
            + ", [" + KEY_UKURANS + "] money NULL "
            + ", [" + KEY_UKURANSPROSENT + "] money NULL "
            + ");";
        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableUkurans(FormMain form)
        {
            this.main = form;
        }

        public void Create()
        {
            if (!main.connection.TableExists(TABLE_NAME))
            {
                var cmdCreate = new SqlCeCommand(sqlCreateTable, main.connection);
                cmdCreate.ExecuteNonQuery();

                var cmdAlter = new SqlCeCommand(sqlAlter, main.connection);
                cmdAlter.ExecuteNonQuery();
            }
            Log.d("Table " + TABLE_NAME + " ready!");
        }

        public void Reset()
        {
            if (main.connection.TableExists(TABLE_NAME))
            {
                var cmdDrop = new SqlCeCommand(sqlDrop, main.connection);
                cmdDrop.ExecuteNonQuery();
            }
            Create();
            main.appConfig.dbObsoleteUpdated = FormMain.rangeMin;
            Log.d("Table " + TABLE_NAME + " cleared and ready!");
        }

        public DataTable GetDataTable()
        {
            var table = new DataTable();
            table.Columns.Add(KEY_AVDELING, typeof(int));
            table.Columns.Add(KEY_VAREKODE, typeof(string));
            table.Columns.Add(KEY_ANTALL, typeof(int));
            table.Columns.Add(KEY_KOST, typeof(decimal));
            table.Columns.Add(KEY_DATO, typeof(DateTime));
            table.Columns.Add(KEY_UKURANS, typeof(decimal));
            table.Columns.Add(KEY_UKURANSPROSENT, typeof(decimal));
            return table;
        }

        public DataTable GetProductCodesInStock(int avdeling)
        {
            string sql = "SELECT " + KEY_VAREKODE + ", SUM(" + KEY_ANTALL + ") AS Antall FROM " + TABLE_NAME
                + " WHERE " + KEY_AVDELING + " = " + avdeling + " GROUP BY " + KEY_VAREKODE;
            return main.database.GetSqlDataTable(sql);
        }

        public DataTable GetInventory(int avdeling, BackgroundWorker bw = null)
        {
            try
            {
                string sql = "SELECT " + KEY_AVDELING + ", " + KEY_VAREKODE + ", " + KEY_ANTALL + ", "
                    + KEY_KOST + ", " + KEY_DATO + ", " + KEY_UKURANS + ", " + KEY_UKURANSPROSENT + " FROM " + TABLE_NAME
                    + " WHERE " + KEY_AVDELING + " = " + avdeling + " OR " + KEY_AVDELING + " = " + (avdeling + 1000);


                DataTable table = main.database.GetSqlDataTable(sql);
                if (table == null)
                    throw new NullReferenceException("Datatable tableUkurans returned NULL");

                table.Columns.Add(TableVareinfo.KEY_TEKST, typeof(string));
                table.Columns.Add(TableVareinfo.KEY_KAT, typeof(int));
                table.Columns.Add(TableVareinfo.KEY_KATNAVN, typeof(string));
                table.Columns.Add(TableVareinfo.KEY_GRUPPE, typeof(int));
                table.Columns.Add(TableVareinfo.KEY_GRUPPENAVN, typeof(string));
                table.Columns.Add(TableVareinfo.KEY_MODGRUPPE, typeof(int));
                table.Columns.Add(TableVareinfo.KEY_MODGRUPPENAVN, typeof(string));
                table.Columns.Add(TableVareinfo.KEY_MERKE, typeof(int));
                table.Columns.Add(TableVareinfo.KEY_MERKENAVN, typeof(string));
                table.Columns.Add(TableVareinfo.KEY_SALGSPRIS, typeof(decimal));

                DataTable tableInfo = main.database.tableVareinfo.GetAllProducts();
                if (tableInfo == null)
                    throw new NullReferenceException("Datatable tableVareinfo returned NULL");

                int count = table.Rows.Count;
                for (int i = 0; i < count; i++)
                {
                    if (bw.CancellationPending)
                        return null;

                    if (i % 88 == 0)
                    {
                        bw.ReportProgress(i, new StatusProgress(count, null, 0, 100));
                        main.processing.SetText = "Søker i produkt-databasen..: "
                            + string.Format("{0:n0}", i) + " / " + string.Format("{0:n0}", count);
                    }

                    DataRow[] result = tableInfo.Select(TableVareinfo.KEY_VAREKODE + " = '" + table.Rows[i][TableVareinfo.KEY_VAREKODE] + "'");
                    if (result.Count() > 0)
                    {
                        table.Rows[i][TableVareinfo.KEY_TEKST] = result[0][TableVareinfo.KEY_TEKST];
                        table.Rows[i][TableVareinfo.KEY_KAT] = result[0][TableVareinfo.KEY_KAT];
                        table.Rows[i][TableVareinfo.KEY_KATNAVN] = result[0][TableVareinfo.KEY_KATNAVN];
                        table.Rows[i][TableVareinfo.KEY_GRUPPE] = result[0][TableVareinfo.KEY_GRUPPE];
                        table.Rows[i][TableVareinfo.KEY_GRUPPENAVN] = result[0][TableVareinfo.KEY_GRUPPENAVN];
                        table.Rows[i][TableVareinfo.KEY_MODGRUPPE] = result[0][TableVareinfo.KEY_MODGRUPPE];
                        table.Rows[i][TableVareinfo.KEY_MODGRUPPENAVN] = result[0][TableVareinfo.KEY_MODGRUPPENAVN];
                        table.Rows[i][TableVareinfo.KEY_MERKE] = result[0][TableVareinfo.KEY_MERKE];
                        table.Rows[i][TableVareinfo.KEY_MERKENAVN] = result[0][TableVareinfo.KEY_MERKENAVN];
                        table.Rows[i][TableVareinfo.KEY_SALGSPRIS] = result[0][TableVareinfo.KEY_SALGSPRIS];
                    }
                }

                return table;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.e("Kritisk feil i tableUkurans.GetInventory. Se logg for detaljer.");
            }
            return null;
        }
    }
}
