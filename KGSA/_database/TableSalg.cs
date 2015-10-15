using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
   public class TableSalg
    {
        FormMain main;
        static string TABLE_NAME = "tblSalg";

        static string KEY_ID = "SalgID";
        static string KEY_SELGERKODE = "Selgerkode";
        static string KEY_VAREGRUPPE = "Varegruppe";
        static string KEY_VAREKODE = "Varekode";
        static string KEY_DATO = "Dato";
        static string KEY_ANTALL = "Antall";
        static string KEY_BTOKR = "Btokr";
        static string KEY_AVDELING = "Avdeling";
        static string KEY_SALGSPRIS = "Salgspris";
        static string KEY_BILAGSNR = "Bilagsnr";
        static string KEY_MVA = "Mva";

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
    }
}
