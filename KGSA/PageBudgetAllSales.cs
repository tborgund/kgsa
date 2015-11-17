using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class PageBudgetAllSales : PageGenerator
    {
        public PageBudgetAllSales(FormMain form, bool runInBackground, BackgroundWorker bw, System.Windows.Forms.WebBrowser webBrowser)
        {
            main = form;
            runningInBackground = runInBackground;
            worker = bw;
            browser = webBrowser;
        }

        public DataTable MakeTableForMonth(int avdeling, DateTime date, bool includeAll = false)
        {
            try
            {
                DataTable tableSales = main.database.tableSalg.GetMonthlySales(main.appConfig.Avdeling, date);
                if (tableSales == null || tableSales.Rows.Count == 0)
                {
                    Log.e("Ingen salg funnet for måned " + date.ToString("MMMM yyyy"));
                    return new DataTable();
                }

                DataTable tableSalesReps = main.database.tableSelgerkoder.GetSalesCodesTable(avdeling);
                if (tableSalesReps == null || tableSalesReps.Rows.Count == 0)
                {
                    Log.n("Ingen selgere funnet i databasen", Color.Red);
                    return new DataTable();
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

                table.Columns.Add(new DataColumn(KEY_ANTALL_BILAG, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_SNITT_OMSET_BILAG, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_SNITT_INNTJEN_BILAG, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_FINANS_ANTALL, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_FINANS_INNTJEN, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_FINANS_SOM, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_STROM_ANTALL, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_STROM_INNTJEN, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_STROM_SOM, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_TA_ANTALL, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_TA_INNTJEN, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_TA_OMSET, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_TA_SOB, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_KH_ANTALL, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_KH_INNTJEN, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_KH_OMSET, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_KH_SOM, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_ACC_ANTALL, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_ACC_INNTJEN, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_ACC_OMSET, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_ACC_SOM, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_ACC_SOB, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_HOVED_PROD_ANTALL, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_HOVED_PROD_SNITT_OMSET, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_HOVED_PROD_SNITT_INNTJEN, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_KUPPVARER_ANTALL, typeof(decimal)));

                int[] accessoriesGrpList = main.appConfig.GetAccessorieGroups(0);
                int[] mainproductsGrpList = main.appConfig.GetMainproductGroups(0);
                List<VarekodeList> varekoderAlle = main.appConfig.varekoder.ToList();

                foreach (DataRow salesRepRow in tableSalesReps.Rows)
                {
                    string sk = salesRepRow[TableSelgerkoder.KEY_SELGERKODE].ToString();
                    string navn = salesRepRow[TableSelgerkoder.KEY_NAVN].ToString();
                    string kat = salesRepRow[TableSelgerkoder.KEY_KATEGORI].ToString();

                    decimal omset_eks_mva = 0, omset = 0, inntjen = 0;

                    var rows = tableSales.Select(TableSalg.KEY_SELGERKODE + " = '" + sk + "'");
                    DataTable tableSel = rows.Any() ? rows.CopyToDataTable() : tableSales.Clone();

                    omset = Compute(tableSales, "SUM(" + TableSalg.KEY_SALGSPRIS + ")", TableSalg.KEY_SELGERKODE + " like '" + sk + "'");
                    if (omset <= 0)
                        continue;

                    omset_eks_mva = Compute(tableSales, "SUM(" + TableSalg.KEY_SALGSPRIS_EX_MVA + ")", TableSalg.KEY_SELGERKODE + " like '" + sk + "'");
                    inntjen = Compute(tableSales, "SUM(" + TableSalg.KEY_BTOKR + ")", TableSalg.KEY_SELGERKODE + " like '" + sk + "'");


                    decimal sFinansAntallSel = 0, sFinansInntjenSel = 0, sAntallSalgSel = 0, sAntallBilagSel = 0;
                    decimal sStromInntjenSel = 0, sStromAntallSel = 0, sModInntjenSel = 0, sModOmsetSel = 0, sModAntallSel = 0, sKuppvarerSel = 0;
                    decimal sAccessoriesAntallSel = 0, sAccessoriesInntjenSel = 0, sAccessoriesOmsetSel = 0, sSnittAntallSel = 0, sSnittInntjenSel = 0;
                    decimal sSnittOmsetSel = 0, sTjenOmsetSel = 0, sTjenInntjenSel = 0, sAntallTjenSel = 0;

                    sAntallSalgSel = Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", null);
                    sAntallBilagSel = tableSel.AsEnumerable().Select(g => g.Field<int>(TableSalg.KEY_BILAGSNR)).Distinct().Count();
                    sFinansAntallSel = Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = 961");
                    sFinansInntjenSel = Compute(tableSel, "Sum(" + TableSalg.KEY_BTOKR + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = 961");
                    sStromInntjenSel = Compute(tableSel, "Sum(" + TableSalg.KEY_BTOKR + ")", "([" + TableSalg.KEY_VAREKODE + "] LIKE 'ELSTROM*' OR [" + TableSalg.KEY_VAREKODE + "] LIKE 'ELRABATT*')");
                    sStromAntallSel = Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREKODE + "] LIKE 'ELSTROM*'");
                    sModInntjenSel = Compute(tableSel, "Sum(" + TableSalg.KEY_BTOKR + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] % 100 = 83 AND [" + TableSalg.KEY_VAREKODE + "] LIKE 'MOD*'");
                    sModOmsetSel = Compute(tableSel, "Sum(" + TableSalg.KEY_SALGSPRIS + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] % 100 = 83 AND [" + TableSalg.KEY_VAREKODE + "] LIKE 'MOD*'");
                    sModAntallSel = Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] % 100 = 83 AND [" + TableSalg.KEY_VAREKODE + "] LIKE 'MOD*'");
                    sKuppvarerSel = Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREKODE + "] = 'KV'");

                    foreach (int ac in accessoriesGrpList)
                    {
                        sAccessoriesAntallSel += Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = " + ac);
                        sAccessoriesInntjenSel += Compute(tableSel, "Sum(" + TableSalg.KEY_BTOKR + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = " + ac);
                        sAccessoriesOmsetSel += Compute(tableSel, "Sum(" + TableSalg.KEY_SALGSPRIS + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = " + ac);
                    }

                    foreach (int grp in mainproductsGrpList)
                    {
                        sSnittAntallSel += Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = " + grp);
                        sSnittInntjenSel += Compute(tableSel, "Sum(" + TableSalg.KEY_BTOKR + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = " + grp);
                        sSnittOmsetSel += Compute(tableSel, "Sum(" + TableSalg.KEY_SALGSPRIS + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = " + grp);
                    }

                    foreach (var varekode in varekoderAlle)
                    {
                        sTjenOmsetSel += Compute(tableSel, "Sum(" + TableSalg.KEY_SALGSPRIS + ")", "[" + TableSalg.KEY_VAREKODE + "]='" + varekode.kode + "'");
                        sTjenInntjenSel += Compute(tableSel, "Sum(" + TableSalg.KEY_BTOKR + ")", "[" + TableSalg.KEY_VAREKODE + "]='" + varekode.kode + "'");
                        if (varekode.synlig)
                            sAntallTjenSel += Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREKODE + "]='" + varekode.kode + "'");
                    }

                    DataRow row = table.NewRow();
                    if (String.IsNullOrEmpty(navn))
                        row[KEY_NAVN] = sk;
                    else
                        row[KEY_NAVN] = sk + " " + navn;
                    row[KEY_AVD] = kat;
                    row[KEY_OMSET_EKS_MVA] = omset_eks_mva;
                    if (omset_eks_mva != 0)
                        row[KEY_BTO_PROSENT] = Math.Round(inntjen / omset_eks_mva, 8);
                    else
                        row[KEY_BTO_PROSENT] = 0;
                    row[KEY_INNTJENT] = inntjen;

                    if (sAntallSalgSel != 0)
                        row[KEY_SNITT_OMSET_BILAG] = omset / sAntallSalgSel;
                    else
                        row[KEY_SNITT_OMSET_BILAG] = 0;

                    row[KEY_ANTALL_BILAG] = sAntallBilagSel;
                    if (sAntallBilagSel != 0)
                        row[KEY_SNITT_OMSET_BILAG] = Math.Round(omset / sAntallBilagSel, 2);
                    else
                        row[KEY_SNITT_OMSET_BILAG] = 0;
                    if (sAntallBilagSel != 0)
                        row[KEY_SNITT_INNTJEN_BILAG] = Math.Round(inntjen / sAntallBilagSel, 2);
                    else
                        row[KEY_SNITT_INNTJEN_BILAG] = 0;
                    row[KEY_KH_ANTALL] = sAntallTjenSel;
                    row[KEY_KH_OMSET] = sTjenOmsetSel;
                    row[KEY_KH_INNTJEN] = sTjenInntjenSel;
                    if (inntjen != 0)
                        row[KEY_KH_SOM] = Math.Round(sTjenInntjenSel / inntjen, 2);
                    else
                        row[KEY_KH_SOM] = 0;
                    row[KEY_STROM_INNTJEN] = sStromInntjenSel;
                    row[KEY_STROM_ANTALL] = sStromAntallSel;
                    if (inntjen != 0)
                        row[KEY_STROM_SOM] = Math.Round(sStromInntjenSel / inntjen, 2);
                    else
                        row[KEY_STROM_SOM] = 0;
                    row[KEY_TA_INNTJEN] = sModInntjenSel;
                    row[KEY_TA_OMSET] = sModOmsetSel;
                    row[KEY_TA_ANTALL] = sModAntallSel;
                    if (omset_eks_mva != 0)
                        row[KEY_TA_SOB] = Math.Round(sModOmsetSel / omset_eks_mva, 2);
                    else
                        row[KEY_TA_SOB] = 0;
                    row[KEY_FINANS_INNTJEN] = sFinansInntjenSel;
                    row[KEY_FINANS_ANTALL] = sFinansAntallSel;
                    if (inntjen != 0)
                        row[KEY_FINANS_SOM] = Math.Round(sFinansInntjenSel / inntjen, 2);
                    else
                        row[KEY_FINANS_SOM] = 0;
                    row[KEY_KUPPVARER_ANTALL] = sKuppvarerSel;

                    row[KEY_ACC_ANTALL] = sAccessoriesAntallSel;
                    row[KEY_ACC_INNTJEN] = sAccessoriesInntjenSel;
                    row[KEY_ACC_OMSET] = sAccessoriesOmsetSel;
                    if (inntjen != 0)
                        row[KEY_ACC_SOM] = Math.Round(sAccessoriesInntjenSel / inntjen, 2);
                    else
                        row[KEY_ACC_SOM] = 0;
                    if (sAccessoriesOmsetSel != 0)
                        row[KEY_ACC_SOB] = Math.Round(sAccessoriesOmsetSel / omset_eks_mva, 2);
                    else
                        row[KEY_ACC_SOB] = 0;

                    row[KEY_HOVED_PROD_ANTALL] = sSnittAntallSel;
                    if (sSnittAntallSel != 0)
                    {
                        row[KEY_HOVED_PROD_SNITT_INNTJEN] = sSnittInntjenSel / sSnittAntallSel;
                        row[KEY_HOVED_PROD_SNITT_OMSET] = sSnittOmsetSel / sSnittAntallSel;
                    }
                    else
                    {
                        row[KEY_HOVED_PROD_SNITT_INNTJEN] = 0;
                        row[KEY_HOVED_PROD_SNITT_OMSET] = 0;
                    }

                    table.Rows.Add(row);
                }

                return table;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.e("Kritisk feil i MakeTableForWeek(): " + ex.Message);
            }
            return null;
        }

        public DataTable MakeTableForWeek(int avdeling, DateTime date)
        {
            try
            {
                DataTable tableSales = main.database.tableSalg.GetWeeklySales(main.appConfig.Avdeling, date);
                if (tableSales == null || tableSales.Rows.Count == 0)
                {
                    Log.n("Ingen salg funnet for uken " + main.database.GetIso8601WeekOfYear(date), Color.Red);
                    return new DataTable();
                }

                DataTable tableSalesReps = main.database.tableSelgerkoder.GetSalesCodesTable(avdeling);
                if (tableSalesReps == null || tableSalesReps.Rows.Count == 0)
                {
                    Log.n("Ingen selgere funnet i databasen", Color.Red);
                    return new DataTable();
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

                table.Columns.Add(new DataColumn(KEY_ANTALL_BILAG, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_SNITT_OMSET_BILAG, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_SNITT_INNTJEN_BILAG, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_FINANS_ANTALL, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_FINANS_INNTJEN, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_FINANS_SOM, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_STROM_ANTALL, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_STROM_INNTJEN, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_STROM_SOM, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_TA_ANTALL, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_TA_INNTJEN, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_TA_OMSET, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_TA_SOB, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_KH_ANTALL, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_KH_INNTJEN, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_KH_OMSET, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_KH_SOM, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_ACC_ANTALL, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_ACC_INNTJEN, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_ACC_OMSET, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_ACC_SOM, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_ACC_SOB, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_HOVED_PROD_ANTALL, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_HOVED_PROD_SNITT_OMSET, typeof(decimal)));
                table.Columns.Add(new DataColumn(KEY_HOVED_PROD_SNITT_INNTJEN, typeof(decimal)));

                table.Columns.Add(new DataColumn(KEY_KUPPVARER_ANTALL, typeof(decimal)));

                int[] accessoriesGrpList = main.appConfig.GetAccessorieGroups(0);
                int[] mainproductsGrpList = main.appConfig.GetMainproductGroups(0);
                List<VarekodeList> varekoderAlle = main.appConfig.varekoder.ToList();

                foreach (DataRow salesRepRow in tableSalesReps.Rows)
                {
                    string sk = salesRepRow[TableSelgerkoder.KEY_SELGERKODE].ToString();
                    string navn = salesRepRow[TableSelgerkoder.KEY_NAVN].ToString();
                    string kat = salesRepRow[TableSelgerkoder.KEY_KATEGORI].ToString();
                    
                    decimal omset_eks_mva = 0, omset = 0, inntjen = 0;

                    var rows = tableSales.Select(TableSalg.KEY_SELGERKODE + " = '" + sk + "'");
                    DataTable tableSel = rows.Any() ? rows.CopyToDataTable() : tableSales.Clone();

                    omset = Compute(tableSel, "SUM(" + TableSalg.KEY_SALGSPRIS + ")", TableSalg.KEY_SELGERKODE + " like '" + sk + "'");

                    omset_eks_mva = Compute(tableSel, "SUM(" + TableSalg.KEY_SALGSPRIS_EX_MVA + ")", TableSalg.KEY_SELGERKODE + " like '" + sk + "'");
                    inntjen = Compute(tableSel, "SUM(" + TableSalg.KEY_BTOKR + ")", TableSalg.KEY_SELGERKODE + " like '" + sk + "'");

                    decimal sFinansAntallSel = 0, sFinansInntjenSel = 0, sAntallSalgSel = 0, sAntallBilagSel = 0;
                    decimal sStromInntjenSel = 0, sStromAntallSel = 0, sModInntjenSel = 0, sModOmsetSel = 0, sModAntallSel = 0, sKuppvarerSel = 0;
                    decimal sAccessoriesAntallSel = 0, sAccessoriesInntjenSel = 0, sAccessoriesOmsetSel = 0, sSnittAntallSel = 0, sSnittInntjenSel = 0;
                    decimal sSnittOmsetSel = 0, sTjenOmsetSel = 0, sTjenInntjenSel = 0, sAntallTjenSel = 0;

                    sAntallSalgSel = Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", null);
                    if (sAntallSalgSel <= 0)
                        continue;

                    sAntallBilagSel = tableSel.AsEnumerable().Select(g => g.Field<int>(TableSalg.KEY_BILAGSNR)).Distinct().Count();
                    sFinansAntallSel = Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = 961");
                    sFinansInntjenSel = Compute(tableSel, "Sum(" + TableSalg.KEY_BTOKR + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = 961");
                    sStromInntjenSel = Compute(tableSel, "Sum(" + TableSalg.KEY_BTOKR + ")", "([" + TableSalg.KEY_VAREKODE + "] LIKE 'ELSTROM*' OR [" + TableSalg.KEY_VAREKODE + "] LIKE 'ELRABATT*')");
                    sStromAntallSel = Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREKODE + "] LIKE 'ELSTROM*'");
                    sModInntjenSel = Compute(tableSel, "Sum(" + TableSalg.KEY_BTOKR + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] % 100 = 83 AND [" + TableSalg.KEY_VAREKODE + "] LIKE 'MOD*'");
                    sModOmsetSel = Compute(tableSel, "Sum(" + TableSalg.KEY_SALGSPRIS + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] % 100 = 83 AND [" + TableSalg.KEY_VAREKODE + "] LIKE 'MOD*'");
                    sModAntallSel = Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] % 100 = 83 AND [" + TableSalg.KEY_VAREKODE + "] LIKE 'MOD*'");
                    sKuppvarerSel = Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREKODE + "] = 'KV'");

                    foreach (int ac in accessoriesGrpList)
                    {
                        sAccessoriesAntallSel += Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = " + ac);
                        sAccessoriesInntjenSel += Compute(tableSel, "Sum(" + TableSalg.KEY_BTOKR + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = " + ac);
                        sAccessoriesOmsetSel += Compute(tableSel, "Sum(" + TableSalg.KEY_SALGSPRIS + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = " + ac);
                    }

                    foreach (int grp in mainproductsGrpList)
                    {
                        sSnittAntallSel += Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = " + grp);
                        sSnittInntjenSel += Compute(tableSel, "Sum(" + TableSalg.KEY_BTOKR + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = " + grp);
                        sSnittOmsetSel += Compute(tableSel, "Sum(" + TableSalg.KEY_SALGSPRIS + ")", "[" + TableSalg.KEY_VAREGRUPPE + "] = " + grp);
                    }

                    foreach (var varekode in varekoderAlle)
                    {
                        sTjenOmsetSel += Compute(tableSel, "Sum(" + TableSalg.KEY_SALGSPRIS + ")", "[" + TableSalg.KEY_VAREKODE + "]='" + varekode.kode + "'");
                        sTjenInntjenSel += Compute(tableSel, "Sum(" + TableSalg.KEY_BTOKR + ")", "[" + TableSalg.KEY_VAREKODE + "]='" + varekode.kode + "'");
                        if (varekode.synlig)
                            sAntallTjenSel += Compute(tableSel, "Sum(" + TableSalg.KEY_ANTALL + ")", "[" + TableSalg.KEY_VAREKODE + "]='" + varekode.kode + "'");
                    }


                    DataRow row = table.NewRow();
                    if (String.IsNullOrEmpty(navn))
                        row[KEY_NAVN] = sk;
                    else
                        row[KEY_NAVN] = sk + " " + navn;
                    row[KEY_AVD] = kat;
                    row[KEY_OMSET_EKS_MVA] = omset_eks_mva;
                    if (omset_eks_mva != 0)
                        row[KEY_BTO_PROSENT] = Math.Round(inntjen / omset_eks_mva, 8);
                    else
                        row[KEY_BTO_PROSENT] = 0;
                    row[KEY_INNTJENT] = inntjen;

                    if (sAntallSalgSel != 0)
                        row[KEY_SNITT_OMSET_BILAG] = omset / sAntallSalgSel;
                    else
                        row[KEY_SNITT_OMSET_BILAG] = 0;

                    row[KEY_ANTALL_BILAG] = sAntallBilagSel;
                    if (sAntallBilagSel != 0)
                        row[KEY_SNITT_OMSET_BILAG] = Math.Round(omset / sAntallBilagSel, 2);
                    else
                        row[KEY_SNITT_OMSET_BILAG] = 0;
                    if (sAntallBilagSel != 0)
                        row[KEY_SNITT_INNTJEN_BILAG] = Math.Round(inntjen / sAntallBilagSel, 2);
                    else
                        row[KEY_SNITT_INNTJEN_BILAG] = 0;
                    row[KEY_KH_ANTALL] = sAntallTjenSel;
                    row[KEY_KH_OMSET] = sTjenOmsetSel;
                    row[KEY_KH_INNTJEN] = sTjenInntjenSel;
                    if (inntjen != 0)
                        row[KEY_KH_SOM] = Math.Round(sTjenInntjenSel / inntjen, 2);
                    else
                        row[KEY_KH_SOM] = 0;
                    row[KEY_STROM_INNTJEN] = sStromInntjenSel;
                    row[KEY_STROM_ANTALL] = sStromAntallSel;
                    if (inntjen != 0)
                        row[KEY_STROM_SOM] = Math.Round(sStromInntjenSel / inntjen, 2);
                    else
                        row[KEY_STROM_SOM] = 0;
                    row[KEY_TA_INNTJEN] = sModInntjenSel;
                    row[KEY_TA_OMSET] = sModOmsetSel;
                    row[KEY_TA_ANTALL] = sModAntallSel;
                    if (omset_eks_mva != 0)
                        row[KEY_TA_SOB] = Math.Round(sModOmsetSel / omset_eks_mva, 2);
                    else
                        row[KEY_TA_SOB] = 0;
                    row[KEY_FINANS_INNTJEN] = sFinansInntjenSel;
                    row[KEY_FINANS_ANTALL] = sFinansAntallSel;
                    if (inntjen != 0)
                        row[KEY_FINANS_SOM] = Math.Round(sFinansInntjenSel / inntjen, 2);
                    else
                        row[KEY_FINANS_SOM] = 0;
                    row[KEY_KUPPVARER_ANTALL] = sKuppvarerSel;

                    row[KEY_ACC_ANTALL] = sAccessoriesAntallSel;
                    row[KEY_ACC_INNTJEN] = sAccessoriesInntjenSel;
                    row[KEY_ACC_OMSET] = sAccessoriesOmsetSel;
                    if (inntjen != 0)
                        row[KEY_ACC_SOM] = Math.Round(sAccessoriesInntjenSel / inntjen, 2);
                    else
                        row[KEY_ACC_SOM] = 0;
                    if (sAccessoriesOmsetSel != 0)
                        row[KEY_ACC_SOB] = Math.Round(sAccessoriesOmsetSel / omset_eks_mva, 2);
                    else
                        row[KEY_ACC_SOB] = 0;

                    row[KEY_HOVED_PROD_ANTALL] = sSnittAntallSel;
                    if (sSnittAntallSel != 0)
                    {
                        row[KEY_HOVED_PROD_SNITT_INNTJEN] = sSnittInntjenSel / sSnittAntallSel;
                        row[KEY_HOVED_PROD_SNITT_OMSET] = sSnittOmsetSel / sSnittAntallSel;
                    }
                    else
                    {
                        row[KEY_HOVED_PROD_SNITT_INNTJEN] = 0;
                        row[KEY_HOVED_PROD_SNITT_OMSET] = 0;
                    }

                    table.Rows.Add(row);
                }

                return table;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.e("Kritisk feil i MakeTableForWeek(): " + ex.Message);
            }
            return null;
        }

        private string KEY_NAVN = "Navn:";
        private string KEY_AVD = "Avd:";
        private string KEY_OMSET_EKS_MVA = "Omset.eks.mva:";
        private string KEY_BTO_PROSENT = "Bto %:";
        private string KEY_INNTJENT = "Inntjent:";

        private string KEY_ANTALL_TIMER = "Antall timer:";

        private string KEY_INNTJENT_PR_TIME = "Inntjent pr.time:";
        private string KEY_OMSET_PR_TIME = "Omset.pr.time:";

        private string KEY_ANTALL_BILAG = "Antall bilag";
        private string KEY_SNITT_OMSET_BILAG = "Snitt omset.";
        private string KEY_SNITT_INNTJEN_BILAG = "Snitt inntjen.";

        private string KEY_ACC_ANTALL = "Tilbehør.antall";
        private string KEY_ACC_INNTJEN = "Tilbehør.Inntjen.";
        private string KEY_ACC_OMSET = "Tilbehør.Omsetn.";
        private string KEY_ACC_SOM = "Tilbehør.SoM.";
        private string KEY_ACC_SOB = "Tilbehør.SoB.";

        private string KEY_HOVED_PROD_ANTALL = "Hovedprodukt.Antall";
        private string KEY_HOVED_PROD_SNITT_OMSET = "Hovedprodukt.Snitt.Omset.";
        private string KEY_HOVED_PROD_SNITT_INNTJEN = "Hovedprodukt.Snitt.Inntjen.";

        private string KEY_KUPPVARER_ANTALL = "Kuppvarer.Antall";

        private string KEY_FINANS_ANTALL = "Finans.Antall";
        private string KEY_FINANS_INNTJEN = "Finans.Inntjen.";
        private string KEY_FINANS_SOM = "Finans.SoM.";

        private string KEY_TA_ANTALL = "Forsikring.Antall";
        private string KEY_TA_INNTJEN = "Forsikring.Inntjen.";
        private string KEY_TA_OMSET = "Forsikring.Omset.";
        private string KEY_TA_SOB = "Forsikring.SoB.";

        private string KEY_STROM_ANTALL = "Strøm.Antall";
        private string KEY_STROM_INNTJEN = "Strøm.Inntjen.";
        private string KEY_STROM_SOM = "Strøm.SoM.";

        private string KEY_KH_ANTALL = "KnowHow.Antall";
        private string KEY_KH_INNTJEN = "KnowHow.Inntjen.";
        private string KEY_KH_OMSET = "KnowHow.Omset.";
        private string KEY_KH_SOM = "KnowHow.SoM.";
    }
}
