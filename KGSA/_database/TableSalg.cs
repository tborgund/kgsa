using System;
using System.Data;
using System.Data.SqlServerCe;

namespace KGSA
{
   public class TableSalg
    {
        FormMain main;
        static string TABLE_NAME = "tblSalg";

        public static string KEY_ID = "SalgID";
        public static string KEY_SELGERKODE = "Selgerkode";
        public static string KEY_VAREGRUPPE = "Varegruppe";
        public static string KEY_VAREKODE = "Varekode";
        public static string KEY_DATO = "Dato";
        public static string KEY_ANTALL = "Antall";
        public static string KEY_BTOKR = "Btokr";
        public static string KEY_AVDELING = "Avdeling";
        public static string KEY_SALGSPRIS = "Salgspris";
        public static string KEY_BILAGSNR = "Bilagsnr";
        public static string KEY_MVA = "Mva";
        public static string KEY_SALGSPRIS_EX_MVA = "SalgsprisExMva";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_SELGERKODE + "] nchar(10) NOT NULL "
            + ", [" + KEY_VAREGRUPPE + "] smallint NOT NULL "
            + ", [" + KEY_VAREKODE + "] nchar(20) NOT NULL "
            + ", [" + KEY_DATO + "] datetime NOT NULL "
            + ", [" + KEY_ANTALL + "] int NULL "
            + ", [" + KEY_BTOKR + "] money NULL "
            + ", [" + KEY_AVDELING + "] int NOT NULL "
            + ", [" + KEY_SALGSPRIS + "] money NULL "
            + ", [" + KEY_BILAGSNR + "] int NULL "
            + ", [" + KEY_MVA + "] money NULL "
            + ");";
        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlIndexDato = "CREATE INDEX [Dato] ON [tblSalg] ([Dato] ASC);";
        static string sqlIndexAvdeling = "CREATE INDEX [Avdeling] ON [tblSalg] ([Avdeling] ASC);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableSalg(FormMain form)
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

                var cmdIndexDato = new SqlCeCommand(sqlIndexDato, main.connection);
                cmdIndexDato.ExecuteNonQuery();

                var cmdIndexAvdeling = new SqlCeCommand(sqlIndexAvdeling, main.connection);
                cmdIndexAvdeling.ExecuteNonQuery();
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
            var table = new DataTable();
            table.Columns.Add(KEY_SELGERKODE, typeof(string));
            table.Columns.Add(KEY_VAREGRUPPE, typeof(int));
            table.Columns.Add(KEY_VAREKODE, typeof(string));
            table.Columns.Add(KEY_DATO, typeof(DateTime));
            table.Columns.Add(KEY_ANTALL, typeof(int));
            table.Columns.Add(KEY_BTOKR, typeof(decimal));
            table.Columns.Add(KEY_AVDELING, typeof(int));
            table.Columns.Add(KEY_SALGSPRIS, typeof(decimal));
            table.Columns.Add(KEY_BILAGSNR, typeof(int));
            table.Columns.Add(KEY_MVA, typeof(decimal));
            return table;
        }

        public DataTable GetWeeklySales(int avdeling, DateTime date)
        {
            string sql = "SELECT * FROM " + TABLE_NAME + " WHERE " + KEY_AVDELING + " = " + avdeling
                + " AND CONVERT(NVARCHAR(10)," + KEY_DATO + ",121) >= CONVERT(NVARCHAR(10),'" + date.StartOfWeek().ToString("yyyy-MM-dd")
                + "',121) AND CONVERT(NVARCHAR(10)," + KEY_DATO + ",121) <= CONVERT(NVARCHAR(10),'" + date.EndOfWeek().ToString("yyyy-MM-dd") + "',121)";
            DataTable table = main.database.GetSqlDataTable(sql);
            if (table != null)
                table.Columns.Add("SalgsprisExMva", typeof(double), "Salgspris / Mva");
            return table;
        }

        public DataTable GetMonthlySales(int avdeling, DateTime date)
        {
            var firstDay = new DateTime(date.Year, date.Month, 1);
            var lastDay = firstDay.AddMonths(1).AddDays(-1);

            string sql = "SELECT * FROM " + TABLE_NAME + " WHERE " + KEY_AVDELING + " = " + avdeling
                + " AND CONVERT(NVARCHAR(10)," + KEY_DATO + ",121) >= CONVERT(NVARCHAR(10),'" + firstDay.ToString("yyyy-MM-dd")
                + "',121) AND CONVERT(NVARCHAR(10)," + KEY_DATO + ",121) <= CONVERT(NVARCHAR(10),'" + lastDay.ToString("yyyy-MM-dd") + "',121)";
            DataTable table = main.database.GetSqlDataTable(sql);
            if (table != null)
                table.Columns.Add("SalgsprisExMva", typeof(double), "Salgspris / Mva");
            return table;
        }
    }
}
