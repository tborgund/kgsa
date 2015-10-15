using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
   public class TableUkurans
    {
        FormMain main;
        static string TABLE_NAME = "tblUkurans";

        static string KEY_ID = "Id";
        static string KEY_AVDELING = "Avdeling";
        static string KEY_VAREKODE = "Varekode";
        static string KEY_ANTALL = "Antall";
        static string KEY_KOST = "Kost";
        static string KEY_DATO = "Dato";
        static string KEY_UKURANS = "UkuransVerdi";
        static string KEY_UKURANSPROSENT = "UkuransProsent";

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
            Logg.Debug("Table " + TABLE_NAME + " ready!");
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
            Logg.Debug("Table " + TABLE_NAME + " cleared and ready!");
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
    }
}
