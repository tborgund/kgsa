using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Globalization;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TablePrisguide
    {
        FormMain main;
        public static string TABLE_NAME = "tblPrisguide";

        public static string KEY_ID = "Id";
        public static string KEY_PRISGUIDE_ID = "PrisguideId";
        public static string KEY_PRISGUIDE_STATUS = "PrisguideStatus";
        public static string KEY_DATE = "Date";
        public static string KEY_AVDELING = "Avdeling";
        public static string KEY_POSITION = "Position";
        public static string KEY_PRODUCT_CODE = "ProductCode";
        public static string KEY_PRODUCT_STOCK = "ProductStock";
        public static string KEY_PRODUCT_PRIZE_INTERNET = "ProductPrizeInternet";
        public static string KEY_PRODUCT_STOCK_INTERNET = "ProductStockInternet";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( " +
            "  [" + KEY_ID + "] int IDENTITY (1,1) NOT NULL " +
            ", [" + KEY_PRISGUIDE_ID + "] int NOT NULL " +
            ", [" + KEY_PRISGUIDE_STATUS + "] int NOT NULL " +
            ", [" + KEY_DATE + "] datetime NOT NULL " +
            ", [" + KEY_AVDELING + "] int NOT NULL " +
            ", [" + KEY_POSITION + "] int NOT NULL " +
            ", [" + KEY_PRODUCT_CODE + "] nvarchar(25) NOT NULL " +
            ", [" + KEY_PRODUCT_STOCK + "] int NOT NULL " +
            ", [" + KEY_PRODUCT_PRIZE_INTERNET + "] money NOT NULL " +
            ", [" + KEY_PRODUCT_STOCK_INTERNET + "] int NOT NULL " +
            ");";
        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";
        static string sqlIndexDate = "CREATE INDEX [" + KEY_DATE + "] ON [" + TABLE_NAME + "] ([" + KEY_DATE + "] ASC);";

        public TablePrisguide(FormMain form)
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

                var cmdIndexDate = new SqlCeCommand(sqlIndexDate, main.connection);
                cmdIndexDate.ExecuteNonQuery();
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
            table.Columns.Add(KEY_PRISGUIDE_ID, typeof(int));
            table.Columns.Add(KEY_PRISGUIDE_STATUS, typeof(int));
            table.Columns.Add(KEY_DATE, typeof(DateTime));
            table.Columns.Add(KEY_AVDELING, typeof(int));
            table.Columns.Add(KEY_POSITION, typeof(int));
            table.Columns.Add(KEY_PRODUCT_CODE, typeof(string));
            table.Columns.Add(KEY_PRODUCT_STOCK, typeof(int));
            table.Columns.Add(KEY_PRODUCT_PRIZE_INTERNET, typeof(decimal));
            table.Columns.Add(KEY_PRODUCT_STOCK_INTERNET, typeof(int));
            return table;
        }

        public void RemoveDate(DateTime date)
        {
            string sql = "DELETE FROM " + TABLE_NAME + " WHERE " + KEY_AVDELING + " = " + main.appConfig.Avdeling
                + " AND CONVERT(NVARCHAR(10),Date,121) >= CONVERT(NVARCHAR(10),'" + date.ToString("yyyy-MM-dd")
                + "',121) AND CONVERT(NVARCHAR(10),Date,121) <= CONVERT(NVARCHAR(10),'" + date.ToString("yyyy-MM-dd") + "',121)";
            Log.d("SQL: " + sql);
            SqlCeCommand command = new SqlCeCommand(sql, main.connection);
            var result = command.ExecuteNonQuery();
            Log.d(TABLE_NAME + ": Slettet " + result + " oppføringer for dato: " + date.ToShortDateString());
        }

        public DataTable GetPrisguideTable(int avdeling, DateTime date)
        {
            string sql = "SELECT * FROM " + TABLE_NAME + " WHERE " + KEY_AVDELING + " = " + avdeling
                + " AND CONVERT(NVARCHAR(10),Date,121) >= CONVERT(NVARCHAR(10),'" + date.ToString("yyyy-MM-dd") 
                + "',121) AND CONVERT(NVARCHAR(10),Date,121) <= CONVERT(NVARCHAR(10),'" + date.ToString("yyyy-MM-dd") + "',121)";
            return main.database.GetSqlDataTable(sql);
        }

        public DataTable GetPrisguideList(int avdeling)
        {
            string sql = "SELECT TOP(10) " + TABLE_NAME + "." + KEY_DATE + ", COUNT(" + TABLE_NAME + "." + KEY_DATE + ") AS NumberOfProducts, " +
                " COUNT(CASE " + KEY_POSITION + " WHEN 0 THEN 0 ELSE 1 END) AS Total, " +
                " SUM(CASE " + KEY_PRODUCT_STOCK + " WHEN 0 THEN 1 ELSE 0 END) AS NoStock, " +
                " SUM(CASE " + KEY_PRODUCT_STOCK_INTERNET + " WHEN 0 THEN 1 ELSE 0 END) AS NoInetStock " +
                " FROM " + TABLE_NAME + " WHERE " + KEY_AVDELING + " = " + avdeling + " GROUP BY " + KEY_DATE;
            return main.database.GetSqlDataTable(sql);
        }

        public DateTime GetLatestDate(int avdeling)
        {
            string sql = "SELECT TOP(1) " + KEY_DATE + " FROM " + TABLE_NAME + " WHERE " + KEY_AVDELING + " = " + avdeling + " ORDER BY " + KEY_DATE + " DESC";
            try
            {
                using (SqlCeCommand command = new SqlCeCommand(sql, main.connection))
                {
                    DateTime date = FormMain.rangeMin;
                    var result = command.ExecuteScalar();
                    if (result != null)
                        date = Convert.ToDateTime(result);

                    return date;
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
            return FormMain.rangeMin;
        }
    }
}
