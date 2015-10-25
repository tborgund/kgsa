using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TableBudget
    {
        FormMain main;
        static string TABLE_NAME = "tblBudget";

        static string KEY_ID = "Id";
        static string KEY_AVDELING = "Avdeling";
        static string KEY_KATEGORI = "Kategori";
        static string KEY_DATE = "Date";
        static string KEY_DAGER = "Dager";
        static string KEY_OMSETNING = "Omsetning";
        static string KEY_INNTJENING = "Inntjening";
        static string KEY_MARGIN = "Margin";
        static string KEY_TA = "TA";
        static string KEY_TA_TYPE = "TA_Type";
        static string KEY_STROM = "Strom";
        static string KEY_STROM_TYPE = "Strom_Type";
        static string KEY_FINANS = "Finans";
        static string KEY_FINANS_TYPE = "Finans_Type";
        static string KEY_RTGSA = "Rtgsa";
        static string KEY_RTGSA_TYPE = "Rtgsa_Type";
        static string KEY_ACC = "Acc";
        static string KEY_ACC_TYPE = "Acc_Type";
        static string KEY_VINN = "Vinn";
        static string KEY_VINN_TYPE = "Vinn_Type";
        static string KEY_UPDATED = "Updated";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_AVDELING + "] smallint NOT NULL "
            + ", [" + KEY_KATEGORI + "] nvarchar(15) NOT NULL "
            + ", [" + KEY_DATE + "] datetime NOT NULL "
            + ", [" + KEY_DAGER + "] int NOT NULL "
            + ", [" + KEY_OMSETNING + "] money NOT NULL "
            + ", [" + KEY_INNTJENING + "] money NOT NULL "
            + ", [" + KEY_MARGIN + "] money NOT NULL "
            + ", [" + KEY_TA + "] money NOT NULL "
            + ", [" + KEY_TA_TYPE + "] nvarchar(15) NOT NULL "
            + ", [" + KEY_STROM + "] money NOT NULL "
            + ", [" + KEY_STROM_TYPE + "] nvarchar(15) NOT NULL "
            + ", [" + KEY_FINANS + "] money NOT NULL "
            + ", [" + KEY_FINANS_TYPE + "] nvarchar(15) NOT NULL "
            + ", [" + KEY_RTGSA + "] money NOT NULL "
            + ", [" + KEY_RTGSA_TYPE + "] nvarchar(15) NOT NULL "
            + ", [" + KEY_ACC + "] money NOT NULL "
            + ", [" + KEY_ACC_TYPE + "] nvarchar(15) NOT NULL "
            + ", [" + KEY_VINN + "] money NOT NULL "
            + ", [" + KEY_VINN_TYPE + "] nvarchar(15) NOT NULL "
            + ", [" + KEY_UPDATED + "] datetime NOT NULL "
            + ");";
        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableBudget(FormMain form)
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
            table.Columns.Add(KEY_KATEGORI, typeof(string));
            table.Columns.Add(KEY_DATE, typeof(DateTime));
            table.Columns.Add(KEY_DAGER, typeof(int));
            table.Columns.Add(KEY_OMSETNING, typeof(decimal));
            table.Columns.Add(KEY_INNTJENING, typeof(decimal));
            table.Columns.Add(KEY_MARGIN, typeof(decimal));
            table.Columns.Add(KEY_TA, typeof(decimal));
            table.Columns.Add(KEY_TA_TYPE, typeof(string));
            table.Columns.Add(KEY_STROM, typeof(decimal));
            table.Columns.Add(KEY_STROM_TYPE, typeof(string));
            table.Columns.Add(KEY_FINANS, typeof(decimal));
            table.Columns.Add(KEY_FINANS_TYPE, typeof(string));
            table.Columns.Add(KEY_RTGSA, typeof(decimal));
            table.Columns.Add(KEY_RTGSA_TYPE, typeof(string));
            table.Columns.Add(KEY_ACC, typeof(decimal));
            table.Columns.Add(KEY_ACC_TYPE, typeof(string));
            table.Columns.Add(KEY_VINN, typeof(decimal));
            table.Columns.Add(KEY_VINN_TYPE, typeof(string));
            table.Columns.Add(KEY_UPDATED, typeof(DateTime));
            return table;
        }
    }
}
