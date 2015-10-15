using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TableBudgetTimer
    {
        FormMain main;
        static string TABLE_NAME = "tblBudgetTimer";

        static string KEY_ID = "Id";
        static string KEY_BUDGET_ID = "BudgetId";
        static string KEY_SELGERKODE = "Selgerkode";
        static string KEY_1 = "1";
        static string KEY_2 = "2";
        static string KEY_3 = "3";
        static string KEY_4 = "4";
        static string KEY_5 = "5";
        static string KEY_6 = "6";
        static string KEY_7 = "7";
        static string KEY_8 = "8";
        static string KEY_9 = "9";
        static string KEY_10 = "10";
        static string KEY_11 = "11";
        static string KEY_12 = "12";
        static string KEY_13 = "13";
        static string KEY_14 = "14";
        static string KEY_15 = "15";
        static string KEY_16 = "16";
        static string KEY_17 = "17";
        static string KEY_18 = "18";
        static string KEY_19 = "19";
        static string KEY_20 = "20";
        static string KEY_21 = "21";
        static string KEY_22 = "22";
        static string KEY_23 = "23";
        static string KEY_24 = "24";
        static string KEY_25 = "25";
        static string KEY_26 = "26";
        static string KEY_27 = "27";
        static string KEY_28 = "28";
        static string KEY_29 = "29";
        static string KEY_30 = "30";
        static string KEY_31 = "31";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_BUDGET_ID + "] int NOT NULL "
            + ", [" + KEY_SELGERKODE + "] nvarchar(15) NOT NULL "
            + ", [" + KEY_1 + "] money NULL "
            + ", [" + KEY_2 + "] money NULL "
            + ", [" + KEY_3 + "] money NULL "
            + ", [" + KEY_4 + "] money NULL "
            + ", [" + KEY_5 + "] money NULL "
            + ", [" + KEY_6 + "] money NULL "
            + ", [" + KEY_7 + "] money NULL "
            + ", [" + KEY_8 + "] money NULL "
            + ", [" + KEY_9 + "] money NULL "
            + ", [" + KEY_10 + "] money NULL "
            + ", [" + KEY_11 + "] money NULL "
            + ", [" + KEY_12 + "] money NULL "
            + ", [" + KEY_13 + "] money NULL "
            + ", [" + KEY_14 + "] money NULL "
            + ", [" + KEY_15 + "] money NULL "
            + ", [" + KEY_16 + "] money NULL "
            + ", [" + KEY_17 + "] money NULL "
            + ", [" + KEY_18 + "] money NULL "
            + ", [" + KEY_19 + "] money NULL "
            + ", [" + KEY_20 + "] money NULL "
            + ", [" + KEY_21 + "] money NULL "
            + ", [" + KEY_22 + "] money NULL "
            + ", [" + KEY_23 + "] money NULL "
            + ", [" + KEY_24 + "] money NULL "
            + ", [" + KEY_25 + "] money NULL "
            + ", [" + KEY_26 + "] money NULL "
            + ", [" + KEY_27 + "] money NULL "
            + ", [" + KEY_28 + "] money NULL "
            + ", [" + KEY_29 + "] money NULL "
            + ", [" + KEY_30 + "] money NULL "
            + ", [" + KEY_31 + "] money NULL "
            + ");";
        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableBudgetTimer(FormMain form)
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
            table.Columns.Add(KEY_BUDGET_ID, typeof(int));
            table.Columns.Add(KEY_SELGERKODE, typeof(string));
            table.Columns.Add(KEY_1, typeof(decimal));
            table.Columns.Add(KEY_2, typeof(decimal));
            table.Columns.Add(KEY_3, typeof(decimal));
            table.Columns.Add(KEY_4, typeof(decimal));
            table.Columns.Add(KEY_5, typeof(decimal));
            table.Columns.Add(KEY_6, typeof(decimal));
            table.Columns.Add(KEY_7, typeof(decimal));
            table.Columns.Add(KEY_8, typeof(decimal));
            table.Columns.Add(KEY_9, typeof(decimal));
            table.Columns.Add(KEY_10, typeof(decimal));
            table.Columns.Add(KEY_11, typeof(decimal));
            table.Columns.Add(KEY_12, typeof(decimal));
            table.Columns.Add(KEY_13, typeof(decimal));
            table.Columns.Add(KEY_14, typeof(decimal));
            table.Columns.Add(KEY_15, typeof(decimal));
            table.Columns.Add(KEY_16, typeof(decimal));
            table.Columns.Add(KEY_17, typeof(decimal));
            table.Columns.Add(KEY_18, typeof(decimal));
            table.Columns.Add(KEY_19, typeof(decimal));
            table.Columns.Add(KEY_20, typeof(decimal));
            table.Columns.Add(KEY_21, typeof(decimal));
            table.Columns.Add(KEY_22, typeof(decimal));
            table.Columns.Add(KEY_23, typeof(decimal));
            table.Columns.Add(KEY_24, typeof(decimal));
            table.Columns.Add(KEY_25, typeof(decimal));
            table.Columns.Add(KEY_26, typeof(decimal));
            table.Columns.Add(KEY_27, typeof(decimal));
            table.Columns.Add(KEY_28, typeof(decimal));
            table.Columns.Add(KEY_29, typeof(decimal));
            table.Columns.Add(KEY_30, typeof(decimal));
            table.Columns.Add(KEY_31, typeof(decimal));
            return table;
        }
    }
}
