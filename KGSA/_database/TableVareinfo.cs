using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TableVareinfo
    {
        FormMain main;
        static string TABLE_NAME = "tblVareinfo";

        static string KEY_ID = "Id";
        static string KEY_VAREKODE = "Varekode";
        static string KEY_TEKST = "Varetekst";
        static string KEY_KAT = "Kategori";
        static string KEY_KATNAVN = "KategoriNavn";
        static string KEY_GRUPPE = "Varegruppe";
        static string KEY_GRUPPENAVN = "VaregruppeNavn";
        static string KEY_MODGRUPPE = "Modgruppe";
        static string KEY_MODGRUPPENAVN = "ModgruppeNavn";
        static string KEY_MERKE = "Merke";
        static string KEY_MERKENAVN = "MerkeNavn";
        static string KEY_DATO = "Dato";
        static string KEY_SALGSPRIS = "Salgspris";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_VAREKODE + "] nvarchar(25) NOT NULL "
            + ", [" + KEY_TEKST + "] nvarchar(100) NOT NULL "
            + ", [" + KEY_KAT + "] int NULL "
            + ", [" + KEY_KATNAVN + "] nvarchar(30) NOT NULL "
            + ", [" + KEY_GRUPPE + "] int NOT NULL "
            + ", [" + KEY_GRUPPENAVN + "] nvarchar(30) NOT NULL "
            + ", [" + KEY_MODGRUPPE + "] int NOT NULL "
            + ", [" + KEY_MODGRUPPENAVN + "] nvarchar(30) NOT NULL "
            + ", [" + KEY_MERKE + "] int NOT NULL "
            + ", [" + KEY_MERKENAVN + "] nvarchar(30) NOT NULL "
            + ", [" + KEY_DATO + "] datetime NOT NULL "
            + ", [" + KEY_SALGSPRIS + "] money NULL "
            + ");";
        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlIndex = "CREATE INDEX [" + KEY_VAREKODE + "] ON [" + TABLE_NAME + "] ([" + KEY_VAREKODE + "] ASC);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableVareinfo(FormMain form)
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

                var cmdIndex = new SqlCeCommand(sqlIndex, main.connection);
                cmdIndex.ExecuteNonQuery();
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
            table.Columns.Add(KEY_VAREKODE, typeof(string));
            table.Columns.Add(KEY_TEKST, typeof(string));
            table.Columns.Add(KEY_KAT, typeof(int));
            table.Columns.Add(KEY_KATNAVN, typeof(string));
            table.Columns.Add(KEY_GRUPPE, typeof(int));
            table.Columns.Add(KEY_GRUPPENAVN, typeof(string));
            table.Columns.Add(KEY_MODGRUPPE, typeof(int));
            table.Columns.Add(KEY_MODGRUPPENAVN, typeof(string));
            table.Columns.Add(KEY_MERKE, typeof(int));
            table.Columns.Add(KEY_MERKENAVN, typeof(string));
            table.Columns.Add(KEY_DATO, typeof(DateTime));
            table.Columns.Add(KEY_SALGSPRIS, typeof(decimal));
            return table;
        }
    }
}
