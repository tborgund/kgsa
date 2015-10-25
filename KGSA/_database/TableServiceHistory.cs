using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TableServiceHistory
    {
        FormMain main;
        static string TABLE_NAME = "tblServiceHistory";

        static string KEY_ID = "Id";
        static string KEY_AVDELING = "Avdeling";
        static string KEY_DATO = "Dato";
        static string KEY_TOTALT = "Totalt";
        static string KEY_AKTIVE = "Aktive";
        static string KEY_FERDIG = "Ferdig";
        static string KEY_TAT = "Tat";
        static string KEY_OVER_14 = "Over14";
        static string KEY_OVER_14_PROSENT = "Over14prosent";
        static string KEY_OVER_21 = "Over21";
        static string KEY_OVER_21_PROSENT = "Over21prosent";
        static string KEY_TILARBEID = "Tilarbeid";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_AVDELING + "] int NOT NULL "
            + ", [" + KEY_DATO + "] datetime NOT NULL "
            + ", [" + KEY_TOTALT + "] int NOT NULL "
            + ", [" + KEY_AKTIVE + "] int NOT NULL "
            + ", [" + KEY_FERDIG + "] int NOT NULL "
            + ", [" + KEY_TAT + "] numeric(9,5) NOT NULL "
            + ", [" + KEY_OVER_14 + "] int NOT NULL "
            + ", [" + KEY_OVER_14_PROSENT + "] float NOT NULL "
            + ", [" + KEY_OVER_21 + "] int NOT NULL "
            + ", [" + KEY_OVER_21_PROSENT + "] float NOT NULL "
            + ", [" + KEY_TILARBEID + "] float NOT NULL "
            + ");";
        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableServiceHistory(FormMain form)
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
            Log.d("Table " + TABLE_NAME + " cleared and ready!");
        }

        public DataTable GetDataTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add(KEY_AVDELING, typeof(int));
            table.Columns.Add(KEY_DATO, typeof(DateTime));
            table.Columns.Add(KEY_TOTALT, typeof(int));
            table.Columns.Add(KEY_AKTIVE, typeof(int));
            table.Columns.Add(KEY_FERDIG, typeof(int));
            table.Columns.Add(KEY_TAT, typeof(decimal));
            table.Columns.Add(KEY_OVER_14, typeof(int));
            table.Columns.Add(KEY_OVER_14_PROSENT, typeof(decimal));
            table.Columns.Add(KEY_OVER_21, typeof(int));
            table.Columns.Add(KEY_OVER_21_PROSENT, typeof(decimal));
            table.Columns.Add(KEY_TILARBEID, typeof(bool));
            return table;
        }
    }
}
