using System.Data;
using System.Data.SqlServerCe;

namespace KGSA
{
   public class TableSelgerkoder
    {
        FormMain main;
        static string TABLE_NAME = "tblSelgerkoder";

        public static string KEY_ID = "Id";
        public static string KEY_SELGERKODE = "Selgerkode";
        public static string KEY_AVDELING = "Avdeling";
        public static string KEY_KATEGORI = "Kategori";
        public static string KEY_PROVISJON = "Provisjon";
        public static string KEY_FINANSKRAV = "FinansKrav";
        public static string KEY_MODKRAV = "ModKrav";
        public static string KEY_STROMKRAV = "StromKrav";
        public static string KEY_RTGSAKRAV = "RtgsaKrav";
        public static string KEY_NAVN = "Navn";

        static string sqlCreateTable = "CREATE TABLE [" + TABLE_NAME + "] ( "
             + "[" + KEY_ID + "] int IDENTITY (1,1) NOT NULL "
             + ", [" + KEY_SELGERKODE + "] nvarchar(15) NOT NULL "
             + ", [" + KEY_AVDELING + "] int NOT NULL "
             + ", [" + KEY_KATEGORI + "] nvarchar(15) NOT NULL "
             + ", [" + KEY_PROVISJON + "] nvarchar(15) NULL "
             + ", [" + KEY_FINANSKRAV + "] int NULL "
             + ", [" + KEY_MODKRAV + "] int NULL "
             + ", [" + KEY_STROMKRAV + "] int NULL "
             + ", [" + KEY_RTGSAKRAV + "] int NULL "
             + ", [" + KEY_NAVN + "] nvarchar(15) NULL "
             + ");";

        static string sqlAlter = "ALTER TABLE [" + TABLE_NAME + "] ADD CONSTRAINT [PK_" + TABLE_NAME + "] PRIMARY KEY ([" + KEY_ID + "]);";
        static string sqlDrop = "DROP TABLE [" + TABLE_NAME + "];";

        public TableSelgerkoder(FormMain form)
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
            var table = new DataTable();
            table.Columns.Add(KEY_SELGERKODE, typeof(string));
            table.Columns.Add(KEY_AVDELING, typeof(int));
            table.Columns.Add(KEY_KATEGORI, typeof(string));
            table.Columns.Add(KEY_PROVISJON, typeof(string));
            table.Columns.Add(KEY_FINANSKRAV, typeof(int));
            table.Columns.Add(KEY_MODKRAV, typeof(int));
            table.Columns.Add(KEY_STROMKRAV, typeof(int));
            table.Columns.Add(KEY_RTGSAKRAV, typeof(int));
            table.Columns.Add(KEY_NAVN, typeof(string));
            return table;
        }

        public DataTable GetSalesCodesTable(int avdeling)
        {
            string sql = "SELECT " + KEY_SELGERKODE + ", " + KEY_KATEGORI + ", " + KEY_NAVN + " FROM " + TABLE_NAME +
                " WHERE " + KEY_AVDELING + " = " + avdeling + " ORDER BY " + KEY_KATEGORI;
            return main.database.GetSqlDataTable(sql);
        }
    }
}
