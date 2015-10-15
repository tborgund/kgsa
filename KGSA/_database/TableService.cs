using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TableService
    {
        FormMain main;
        static string TABLE_NAME = "tblService";

        static string KEY_ID = "ServiceID";
        static string KEY_AVDELING = "Avdeling";
        static string KEY_ORDRENR = "Ordrenr";
        static string KEY_NAVN = "Navn";
        static string KEY_DATO_MOTTAT = "DatoMottat";
        static string KEY_DATO_IARBEID = "DatoIarbeid";
        static string KEY_DATO_FERDIG = "DatoFerdig";
        static string KEY_DATO_UTLEVERT = "DatoUtlevert";
        static string KEY_STATUS = "Status";
        static string KEY_SELGERKODE = "Selgerkode";
        static string KEY_VERKSTED = "Verksted";
        static string KEY_FERDIGBEHANDLET = "FerdigBehandlet";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
            + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
            + ", [" + KEY_AVDELING + "] int NOT NULL "
            + ", [" + KEY_ORDRENR + "] int NOT NULL "
            + ", [" + KEY_NAVN + "] nvarchar(40) NOT NULL "
            + ", [" + KEY_DATO_MOTTAT + "] datetime NOT NULL "
            + ", [" + KEY_DATO_IARBEID + "] datetime NULL "
            + ", [" + KEY_DATO_FERDIG + "] datetime NULL "
            + ", [" + KEY_DATO_UTLEVERT + "] datetime NULL "
            + ", [" + KEY_STATUS + "] nvarchar(100) NULL "
            + ", [" + KEY_SELGERKODE + "] nvarchar(20) NOT NULL "
            + ", [" + KEY_VERKSTED + "] nvarchar(30) NULL "
            + ", [" + KEY_FERDIGBEHANDLET + "] bit NULL "
            + ");";
        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableService(FormMain form)
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
            table.Columns.Add(KEY_AVDELING, typeof(int));
            table.Columns.Add(KEY_ORDRENR, typeof(int));
            table.Columns.Add(KEY_NAVN, typeof(string));
            table.Columns.Add(KEY_DATO_MOTTAT, typeof(DateTime));
            table.Columns.Add(KEY_DATO_IARBEID, typeof(DateTime));
            table.Columns.Add(KEY_DATO_FERDIG, typeof(DateTime));
            table.Columns.Add(KEY_DATO_UTLEVERT, typeof(DateTime));
            table.Columns.Add(KEY_STATUS, typeof(string));
            table.Columns.Add(KEY_VERKSTED, typeof(string));
            table.Columns.Add(KEY_FERDIGBEHANDLET, typeof(bool));
            return table;
        }
    }
}
