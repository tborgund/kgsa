using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TableEmail
    {
        FormMain main;
        static string TABLE_NAME = "tblEmail";

        static string KEY_ID = "Id";
        static string KEY_NAME = "Name";
        static string KEY_ADDRESS = "Address";
        static string KEY_TYPE = "Type";
        static string KEY_QUICK = "Quick";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_NAME + "] nvarchar(100) NOT NULL "
            + ", [" + KEY_ADDRESS + "] nvarchar(100) NOT NULL "
            + ", [" + KEY_TYPE + "] nvarchar(100) NOT NULL "
            + ", [" + KEY_QUICK + "] bit NOT NULL "
            + ");";

        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableEmail(FormMain form)
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
            table.Columns.Add(KEY_NAME, typeof(string));
            table.Columns.Add(KEY_ADDRESS, typeof(string));
            table.Columns.Add(KEY_TYPE, typeof(string));
            table.Columns.Add(KEY_QUICK, typeof(bool));
            return table;
        }
    }
}
