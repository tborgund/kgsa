using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TableServiceLogg
    {
        FormMain main;
        static string TABLE_NAME = "tblServiceLogg";

        static string KEY_ID = "Id";
        static string KEY_SERVICE_ID = "ServiceID";
        static string KEY_DATO_TID = "DatoTid";
        static string KEY_KODE = "Kode";
        static string KEY_TEKST = "Tekst";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_SERVICE_ID + "] int NOT NULL "
            + ", [" + KEY_DATO_TID + "] datetime NOT NULL "
            + ", [" + KEY_KODE + "] nvarchar(20) NOT NULL "
            + ", [" + KEY_TEKST + "] nvarchar(100) NOT NULL "
            + ");";
        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableServiceLogg(FormMain form)
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
            table.Columns.Add(KEY_SERVICE_ID, typeof(int));
            table.Columns.Add(KEY_DATO_TID, typeof(DateTime));
            table.Columns.Add(KEY_KODE, typeof(string));
            table.Columns.Add(KEY_TEKST, typeof(string));
            return table;
        }
    }
}
