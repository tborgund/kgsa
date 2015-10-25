using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Globalization;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TableDailyBudget
    {
        FormMain main;
        public static string TABLE_NAME = "tblDailyBudget";

        public static string KEY_ID = "Id";
        public static string KEY_DATE = "Date";
        public static string KEY_AVDELING = "Avdeling";
        public static string KEY_BUDGET_ID = "BudgetId";
        public static string KEY_BUDGET_TYPE = "BudgetType";
        public static string KEY_BUDGET_SALES = "BudgetSales";
        public static string KEY_BUDGET_GM = "BudgetGm";
        public static string KEY_BUDGET_GM_PERCENT = "BudgetGmPercent";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( " +
            "  [" + KEY_ID + "] int IDENTITY (1,1) NOT NULL " +
            ", [" + KEY_DATE + "] datetime NOT NULL " +
            ", [" + KEY_AVDELING + "] int NOT NULL " +
            ", [" + KEY_BUDGET_ID + "] int NOT NULL " +
            ", [" + KEY_BUDGET_TYPE + "] nvarchar(15) NOT NULL " +
            ", [" + KEY_BUDGET_SALES + "] money NOT NULL " +
            ", [" + KEY_BUDGET_GM + "] money NOT NULL " +
            ", [" + KEY_BUDGET_GM_PERCENT + "] money NOT NULL " +
            ");";
        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";
        static string sqlIndexDate = "CREATE INDEX [" + KEY_DATE + "] ON [" + TABLE_NAME + "] ([" + KEY_DATE + "] ASC);";

        public TableDailyBudget(FormMain form)
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
            table.Columns.Add(KEY_DATE, typeof(DateTime));
            table.Columns.Add(KEY_AVDELING, typeof(int));
            table.Columns.Add(KEY_BUDGET_ID, typeof(int));
            table.Columns.Add(KEY_BUDGET_TYPE, typeof(string));
            table.Columns.Add(KEY_BUDGET_SALES, typeof(decimal));
            table.Columns.Add(KEY_BUDGET_GM, typeof(decimal));
            table.Columns.Add(KEY_BUDGET_GM_PERCENT, typeof(decimal));
            return table;
        }

        public void RemoveDate(int avdeling, DateTime date)
        {
            string sql = "DELETE FROM " + TABLE_NAME + " WHERE " + KEY_AVDELING + " = " + avdeling
                + " AND CONVERT(NVARCHAR(10),Date,121) >= CONVERT(NVARCHAR(10),'" + date.ToString("yyyy-MM-dd")
                + "',121) AND CONVERT(NVARCHAR(10),Date,121) <= CONVERT(NVARCHAR(10),'" + date.ToString("yyyy-MM-dd") + "',121)";
            Log.d("SQL: " + sql);
            SqlCeCommand command = new SqlCeCommand(sql, main.connection);
            var result = command.ExecuteNonQuery();
            Log.d(TABLE_NAME + ": Slettet " + result + " oppføringer for dato: " + date.ToShortDateString());
        }

        public void RemoveBudgetId(int avdeling, int budgetId)
        {
            string sql = "DELETE FROM " + TABLE_NAME + " WHERE " + KEY_AVDELING + " = " + avdeling
                + " AND " + KEY_BUDGET_ID + " = " + budgetId;
            Log.d("SQL: " + sql);
            SqlCeCommand command = new SqlCeCommand(sql, main.connection);
            var result = command.ExecuteNonQuery();
            Log.d(TABLE_NAME + ": Slettet " + result + " oppføringer for budsjett id: " + budgetId);
        }

        public DataTable GetBudgetFromDate(int avdeling, DateTime date)
        {
            string sql = "SELECT * FROM " + TABLE_NAME + " WHERE " + KEY_AVDELING + " = " + avdeling
                + " AND CONVERT(NVARCHAR(10),Date,121) >= CONVERT(NVARCHAR(10),'" + date.ToString("yyyy-MM-dd")
                + "',121) AND CONVERT(NVARCHAR(10),Date,121) <= CONVERT(NVARCHAR(10),'" + date.ToString("yyyy-MM-dd") + "',121)";
            return main.database.GetSqlDataTable(sql);
        }

        public DataTable GetBudgetFromBudgetId(int avdeling, int budgetId)
        {
            string sql = "SELECT * FROM " + TABLE_NAME + " WHERE " + KEY_AVDELING + " = " + avdeling
                + " AND " + KEY_BUDGET_ID + " = " + budgetId;
            return main.database.GetSqlDataTable(sql);
        }
    }
}
