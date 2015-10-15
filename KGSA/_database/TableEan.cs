using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TableEan
    {
        FormMain main;
        static string TABLE_NAME = "tblEan";

        static string KEY_ID = "Id";
        static string KEY_BARCODE = "Barcode";
        static string KEY_VAREKODE = "Varekode";
        static string KEY_VARETEKST = "Varetekst";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_BARCODE + "] nvarchar(13) NOT NULL "
            + ", [" + KEY_VAREKODE + "] nvarchar(25) NULL "
            + ", [" + KEY_VARETEKST + "] nvarchar(50) NULL "
            + ");";
        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableEan(FormMain form)
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
            Logg.Debug("Table " + TABLE_NAME + " cleared and ready!");
        }

        public DataTable GetDataTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add(KEY_BARCODE, typeof(string));
            table.Columns.Add(KEY_VAREKODE, typeof(string));
            table.Columns.Add(KEY_VARETEKST, typeof(string));
            return table;
        }
    }
}
