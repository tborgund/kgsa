using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
   public class TableVinnprodukt
    {
        FormMain main;
        static string TABLE_NAME = "tblVinnprodukt";

        static string KEY_ID = "Id";
        static string KEY_AVDELING = "Avdeling";
        static string KEY_KATEGORI = "Kategori";
        static string KEY_VAREKODE = "Varekode";
        static string KEY_POENG = "Poeng";
        static string KEY_DATO_OPPRETTET = "DatoOpprettet";
        static string KEY_DATO_EXPIRE = "DatoExpire";
        static string KEY_DATO_START = "DatoStart";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_AVDELING + "] smallint NOT NULL "
            + ", [" + KEY_KATEGORI + "] nvarchar(15) NOT NULL "
            + ", [" + KEY_VAREKODE + "] nvarchar(25) NOT NULL "
            + ", [" + KEY_POENG + "] int NOT NULL "
            + ", [" + KEY_DATO_OPPRETTET + "] datetime NOT NULL "
            + ", [" + KEY_DATO_EXPIRE + "] datetime NOT NULL "
            + ", [" + KEY_DATO_START + "] datetime NOT NULL "
            + ");";

        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableVinnprodukt(FormMain form)
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
            var table = new DataTable();
            table.Columns.Add(KEY_AVDELING, typeof(int));
            table.Columns.Add(KEY_KATEGORI, typeof(string));
            table.Columns.Add(KEY_VAREKODE, typeof(string));
            table.Columns.Add(KEY_POENG, typeof(int));
            table.Columns.Add(KEY_DATO_OPPRETTET, typeof(DateTime));
            table.Columns.Add(KEY_DATO_EXPIRE, typeof(DateTime));
            table.Columns.Add(KEY_DATO_START, typeof(DateTime));
            return table;
        }
    }
}
