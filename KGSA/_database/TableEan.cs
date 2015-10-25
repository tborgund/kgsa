using System;
using System.Data;
using System.Data.SqlServerCe;

namespace KGSA
{
    public class TableEan
    {
        FormMain main;
        public static string TABLE_NAME = "tblEan";
        public static string KEY_ID = "Id";
        public static string KEY_BARCODE = "Barcode";
        public static string KEY_PRODUCT_CODE = "ProductCode";
        public static string KEY_PRODUCT_TEXT = "ProductText";

        public static int INDEX_BARCODE = 0;
        public static int INDEX_PRODUCT_CODE = 1;
        public static int INDEX_PRODUCT_TEXT = 2;

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_BARCODE + "] nvarchar(13) NOT NULL "
            + ", [" + KEY_PRODUCT_CODE + "] nvarchar(25) NULL "
            + ", [" + KEY_PRODUCT_TEXT + "] nvarchar(50) NULL "
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
            table.Columns.Add(KEY_BARCODE, typeof(string));
            table.Columns.Add(KEY_PRODUCT_CODE, typeof(string));
            table.Columns.Add(KEY_PRODUCT_TEXT, typeof(string));
            return table;
        }

        public DataTable GetAllRows()
        {
            string sql = "SELECT " + KEY_BARCODE + ", " + KEY_PRODUCT_CODE + ", " + KEY_PRODUCT_TEXT + " FROM " + TABLE_NAME;
            return main.database.GetSqlDataTable(sql);
        }
    }
}
