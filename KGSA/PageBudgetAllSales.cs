using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;

namespace KGSA
{
    public class PageBudgetAllSales : PageGenerator
    {
        public PageBudgetAllSales(FormMain form, bool runInBackground, BackgroundWorker bw, System.Windows.Forms.WebBrowser webBrowser)
        {
            this.main = form;
            this.runningInBackground = runInBackground;
            this.worker = bw;
            this.browser = webBrowser;
        }

        public bool BuildPage(BudgetCategory cat, string strHash, string htmlPage, DateTime date)
        {
            pickedDate = date;
            string katArg = BudgetCategoryClass.TypeToName(cat);
            bool abort = main.HarSisteVersjonBudget(cat, strHash);
            try
            {
                if (!runningInBackground && !abort) main.timewatch.Start();
                if (!runningInBackground) main.savedBudgetPage = cat;
                if (!abort)
                {
                    Logg.Log("Oppdaterer [" + BudgetCategoryClass.TypeToName(cat) + "]..");
                    OpenPage_Loading();

                    doc = new List<string>();

                    main.openXml.DeleteDocument(BudgetCategoryClass.TypeToName(cat), pickedDate);

                    AddPage_Start(true, "Budsjett (" + main.avdeling.Get(main.appConfig.Avdeling) + ")");
                    AddPage_Title("Budsjett (" + main.avdeling.Get(main.appConfig.Avdeling) + ")");

                    ShowProgress();

                    MakeLastWeekSalesList();

                    AddPage_End();

                    if (FormMain.stopRanking)
                    {
                        main.ClearHash(katArg);
                        Logg.Log("Lasting avbrutt", Color.Red);
                        OpenPage_Stopped();
                        FormMain.stopRanking = false;
                    }
                    else
                    {
                        File.WriteAllLines(htmlPage, doc.ToArray(), Encoding.Unicode);
                        OpenPage(htmlPage);
                        if (!runningInBackground)
                            Logg.Log("Side [" + katArg + "] tok " + main.timewatch.Stop() + " sekunder.", Color.Black, true);
                    }
                }
                else
                    OpenPage(htmlPage);
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                if (!runningInBackground)
                {
                    OpenPage_Error();
                    FormError errorMsg = new FormError("Feil ved generering av side for [" + katArg + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
            return false;
        }

        private void MakeLastWeekSalesList()
        {
            try
            {
                DateTime dateStartLastWeek = main.database.GetStartOfLastWholeWeek(pickedDate);

                Logg.Debug("Henter selger tabell for sist uke..");
                DataTable table = MakeTableForWeek(main.appConfig.Avdeling, dateStartLastWeek);
                if (table == null || table.Rows.Count == 0)
                {
                    AddWarning("Mangler selgerkoder og/eller transaksjoner for uke " + main.database.GetIso8601WeekOfYear(dateStartLastWeek));
                    return;
                }

                main.openXml.SaveDocument(table, BudgetCategoryClass.TypeToName(BudgetCategory.AlleSelgere),
                    "Salgskvalitet - Uke " + main.database.GetIso8601WeekOfYear(dateStartLastWeek), dateStartLastWeek, "Budsjett  (" + main.avdeling.Get(main.appConfig.Avdeling) + ")");

                AddTable_Start("Salgskvalitet - Uke " + main.database.GetIso8601WeekOfYear(dateStartLastWeek) + " ("
                    + dateStartLastWeek.StartOfWeek().ToString("dddd d. MMMM", FormMain.norway) + " - "
                    + dateStartLastWeek.EndOfWeek().ToString("dddd d. MMMM", FormMain.norway));

                AddTable_Header_Start();
                AddTable_Header_Name("Navn", 100, "", Sorter_Type_Text);
                AddTable_Header_Name("Avd", 60, "", Sorter_Type_Text);

                AddTable_Header_Name("Oms.eks.mva", 80, "", Sorter_Type_Digit);
                AddTable_Header_Name("Bto %", 80, "", Sorter_Type_Procent);
                AddTable_Header_Name("Inntjent", 80, "", Sorter_Type_Digit);
                AddTable_Header_Name("Antall timer", 100, "border - left:2px solid #000;", Sorter_Type_Digit);
                AddTable_Header_Name("Inntjent pr.time", 100, "", Sorter_Type_Digit);
                AddTable_Header_Name("Oms pr.time", 100, "", Sorter_Type_Digit);
                AddTable_Header_End();

                AddTable_Body_Start();

                foreach (DataRow row in table.Rows)
                {
                    AddTable_Row_Start();
                    AddTable_Row_Cell(row[INDEX_NAVN].ToString(), "", Class_Style_Numbers_Text_Cat);
                    AddTable_Row_Cell(row[INDEX_AVD].ToString(), "", Class_Style_Numbers_Text_Cat);

                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(row[INDEX_OMSET_EKS_MVA], 0, "", true), "", Class_Style_Numbers_Gen);
                    AddTable_Row_Cell(main.tools.NumberStyle_Percent(row[INDEX_BTO_PROSENT], 1, true, false, 0), "text-align: right;", Class_Style_Numbers_Percent);
                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(row[INDEX_INNTJENT], 0, "", true), "", Class_Style_Numbers_Gen);
                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(0), "border-left:2px solid #000;", Class_Style_Numbers_Gen);
                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(0), "", Class_Style_Numbers_Gen);
                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(0), "", Class_Style_Numbers_Gen);
                    AddTable_Row_End();
                }

                AddTable_End();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                Logg.Log("Feil ved generering av side: " + ex.Message, Color.Red);
            }
            return;
        }

        private DataTable MakeTableForWeek(int avdeling, DateTime date)
        {
            try
            {
                DataTable tableSales = main.database.tableSalg.GetWeeklySales(main.appConfig.Avdeling, date);
                if (tableSales == null || tableSales.Rows.Count == 0)
                {
                    Logg.Log("Ingen salg funnet for uken " + main.database.GetIso8601WeekOfYear(date), Color.Red);
                    return null;
                }

                DataTable tableSalesReps = main.database.tableSelgerkoder.GetSalesCodesTable(avdeling);
                if (tableSalesReps == null || tableSalesReps.Rows.Count == 0)
                {
                    Logg.Log("Ingen selgere funnet i databasen", Color.Red);
                    return null;
                }

                DataTable table = new DataTable();
                table.Columns.Add(new DataColumn(KEY_NAVN, typeof(string)));
                table.Columns.Add(new DataColumn(KEY_AVD, typeof(string)));
                table.Columns.Add(new DataColumn(KEY_OMSET_EKS_MVA, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_BTO_PROSENT, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_INNTJENT, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_ANTALL_TIMER, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_INNTJENT_PR_TIME, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_OMSET_PR_TIME, typeof(decimal)));

                foreach (DataRow salesRepRow in tableSalesReps.Rows)
                {
                    string sk = salesRepRow[TableSelgerkoder.KEY_SELGERKODE].ToString();
                    string navn = salesRepRow[TableSelgerkoder.KEY_NAVN].ToString();
                    string kat = salesRepRow[TableSelgerkoder.KEY_KATEGORI].ToString();
                    
                    decimal omset_eks_mva = 0, omset = 0, inntjen = 0;

                    omset = Compute(tableSales, "SUM(" + TableSalg.KEY_SALGSPRIS + ")", TableSalg.KEY_SELGERKODE + " like '" + sk + "'");
                    if (omset <= 0)
                        continue;

                    omset_eks_mva = Compute(tableSales, "SUM(" + TableSalg.KEY_SALGSPRIS_EX_MVA + ")", TableSalg.KEY_SELGERKODE + " like '" + sk + "'");
                    inntjen = Compute(tableSales, "SUM(" + TableSalg.KEY_BTOKR + ")", TableSalg.KEY_SELGERKODE + " like '" + sk + "'");

                    DataRow row = table.NewRow();
                    if (String.IsNullOrEmpty(navn))
                        row[INDEX_NAVN] = sk;
                    else
                        row[INDEX_NAVN] = sk + " " + navn;
                    row[INDEX_AVD] = kat;
                    row[INDEX_OMSET_EKS_MVA] = omset_eks_mva;
                    if (omset_eks_mva != 0)
                        row[INDEX_BTO_PROSENT] = Math.Round(inntjen / omset_eks_mva, 8);
                    else
                        row[INDEX_BTO_PROSENT] = 0;
                    row[INDEX_INNTJENT] = inntjen;
                    table.Rows.Add(row);
                }

                return table;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return null;
        }

        private string KEY_NAVN = "Navn:";
        private int INDEX_NAVN = 0;
        private string KEY_AVD = "Avd:";
        private int INDEX_AVD = 1;
        private string KEY_OMSET_EKS_MVA = "Omset.eks.mva:";
        private int INDEX_OMSET_EKS_MVA = 2;
        private string KEY_BTO_PROSENT = "Bto %:";
        private int INDEX_BTO_PROSENT = 3;
        private string KEY_INNTJENT = "Inntjent:";
        private int INDEX_INNTJENT = 4;
        private string KEY_ANTALL_TIMER = "Antall timer:";
        //private int INDEX_ANTALL_TIMER = 5;
        private string KEY_INNTJENT_PR_TIME = "Inntjent pr.time:";
        //private int INDEX_INNTJENT_PR_TIME = 6;
        private string KEY_OMSET_PR_TIME = "Omset.pr.time:";
        //private int INDEX_OMSET_PR_TIME = 7;
    }
}
