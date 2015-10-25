using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TableBudgetSelger
    {
        FormMain main;
        static string TABLE_NAME = "tblBudgetSelger";

        static string KEY_ID = "Id";
        static string KEY_BUDGET_ID = "BudgetId";
        static string KEY_SELGERKODE = "Selgerkode";
        static string KEY_TIMER = "Timer";
        static string KEY_DAGER = "Dager";
        static string KEY_MULTIPLIKATOR = "Multiplikator";
        static string KEY_COMMENT = "Comment";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_BUDGET_ID + "] int NOT NULL "
            + ", [" + KEY_SELGERKODE + "] nvarchar(25) NOT NULL "
            + ", [" + KEY_TIMER + "] int NOT NULL "
            + ", [" + KEY_DAGER + "] int NOT NULL "
            + ", [" + KEY_MULTIPLIKATOR + "] money NOT NULL "
            + ", [" + KEY_COMMENT + "] nvarchar(100) NOT NULL "
            + ");";
        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableBudgetSelger(FormMain form)
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
            table.Columns.Add(KEY_BUDGET_ID, typeof(int));
            table.Columns.Add(KEY_SELGERKODE, typeof(string));
            table.Columns.Add(KEY_TIMER, typeof(int));
            table.Columns.Add(KEY_DAGER, typeof(int));
            table.Columns.Add(KEY_MULTIPLIKATOR, typeof(decimal));
            table.Columns.Add(KEY_COMMENT, typeof(string));
            return table;
        }
    }
}
