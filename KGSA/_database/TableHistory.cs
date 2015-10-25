using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TableHistory
    {
        FormMain main;
        static string TABLE_NAME = "tblHistory";

        static string KEY_ID = "Id";
        static string KEY_AVDELING = "Avdeling";
        static string KEY_DATO = "Dato";
        static string KEY_KATEGORI = "Kategori";
        static string KEY_LAGERVERDI = "Lagerverdi";
        static string KEY_LAGERANTALL = "Lagerantall";
        static string KEY_UKURANS_ANTALL = "Ukuransantall";
        static string KEY_UKURANS_VERDI = "Ukuransverdi";
        static string KEY_UKURANS_PROSENT = "Ukuransprosent";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_AVDELING + "] smallint NOT NULL "
            + ", [" + KEY_DATO + "] datetime NOT NULL "
            + ", [" + KEY_KATEGORI + "] nvarchar(15) NOT NULL "
            + ", [" + KEY_LAGERVERDI + "] money NOT NULL "
            + ", [" + KEY_LAGERANTALL + "] int NOT NULL "
            + ", [" + KEY_UKURANS_ANTALL + "] int NOT NULL "
            + ", [" + KEY_UKURANS_VERDI + "] money NOT NULL "
            + ", [" + KEY_UKURANS_PROSENT + "] money NOT NULL "
            + ");";
        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableHistory(FormMain form)
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
            main.appConfig.dbStoreFrom = FormMain.rangeMin;
            main.appConfig.dbStoreTo = FormMain.rangeMin;
            Log.d("Table " + TABLE_NAME + " cleared and ready!");
        }

        public DataTable GetDataTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add(KEY_AVDELING, typeof(int));
            table.Columns.Add(KEY_DATO, typeof(DateTime));
            table.Columns.Add(KEY_KATEGORI, typeof(string));
            table.Columns.Add(KEY_LAGERVERDI, typeof(decimal));
            table.Columns.Add(KEY_LAGERANTALL, typeof(int));
            table.Columns.Add(KEY_UKURANS_ANTALL, typeof(int));
            table.Columns.Add(KEY_UKURANS_VERDI, typeof(decimal));
            table.Columns.Add(KEY_UKURANS_PROSENT, typeof(decimal));
            return table;
        }
    }
}
