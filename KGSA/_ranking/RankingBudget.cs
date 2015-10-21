using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using KGSA.Properties;
using System.Threading;
using System.ComponentModel;
using System.Text;
using System.Diagnostics;

namespace KGSA
{
    public class RankingBudget : Ranking
    {
        private DataTable dtBudgetOverall;
        public List<VarekodeList> varekoderAlle;
        public IEnumerable<string> varekoderAlleAlias;
        public BudgetInfo budgetInfo;
        private BudgetObj budget;

        public RankingBudget() { }

        public RankingBudget(FormMain form, DateTime dtFraArg, DateTime dtTilArg, BudgetCategory cat)
        {
            try
            {
                this.main = form;
                dtFra = dtFraArg;
                dtTil = dtTilArg;
                velgerPeriode = FormMain.datoPeriodeVelger;

                this.varekoderAlle = main.appConfig.varekoder.ToList();
                this.varekoderAlleAlias = varekoderAlle.Where(item => item.synlig == true).Select(x => x.alias).Distinct();

                budget = new BudgetObj(main);

                budgetInfo = budget.GetBudgetInfo(dtTil, cat);
                if (budgetInfo != null)
                    if (budgetInfo.selgere != null)
                        if (budgetInfo.selgere.Count == 0)
                            budgetInfo = null;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                budgetInfo = null;
            }
        }

        private DataTable MakeTableBudget(BudgetCategory cat)
        {
            try
            {
                DataTable dtWork = ReadyTableOverall();

                var rowsGet = main.database.CallMonthTable(dtTil, main.appConfig.Avdeling).Select(BudgetCategoryClass.GetSqlCategoryString(cat) + " AND (Dato >= '" + dtFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtTil.ToString("yyy-MM-dd") + "')");
                sqlce = rowsGet.Any() ? rowsGet.CopyToDataTable() : sqlce.Clone();

                DateTime dtMainFra = dtFra;
                DateTime dtMainTil = dtTil;

                object g;
                decimal Omset_totaltSelgere = 0, OmsetExMva_totaltSelgere = 0, Inntjen_totaltSelgere = 0;

                budgetInfo.chartdata_inntjen = new List<BudgetChartData>() { };
                budgetInfo.chartdata_omset = new List<BudgetChartData>() { };
                budgetInfo.chartdata_kvalitet = new List<BudgetChartData>() { };
                budgetInfo.chartdata_effektivitet = new List<BudgetChartData>() { };
                budgetInfo.comparelist = new List<BudgetCompareData>() { };


                if (cat == BudgetCategory.MDA || cat == BudgetCategory.AudioVideo || cat == BudgetCategory.SDA || cat == BudgetCategory.Tele ||
                    cat == BudgetCategory.Data || cat == BudgetCategory.Cross || cat == BudgetCategory.MDASDA)
                {
                    // Egen avdeling
                    if (sqlce.Rows.Count > 0)
                    {
                        decimal omset_totaltAvdeling = 0, omsetExMva_totaltAvdeling = 0, inntjen_totaltAvdeling = 0, margin_totaltAvdeling = 0;

                        g = sqlce.Compute("Sum(Salgspris)", "");
                        if (!DBNull.Value.Equals(g))
                            omset_totaltAvdeling = Convert.ToDecimal(g);

                        g = sqlce.Compute("Sum(SalgsprisExMva)", "");
                        if (!DBNull.Value.Equals(g))
                            omsetExMva_totaltAvdeling = Convert.ToDecimal(g);

                        g = sqlce.Compute("Sum(Btokr)", "");
                        if (!DBNull.Value.Equals(g))
                            inntjen_totaltAvdeling = Convert.ToDecimal(g);

                        if (omset_totaltAvdeling != 0)
                            margin_totaltAvdeling = inntjen_totaltAvdeling / omsetExMva_totaltAvdeling;
                        else
                            margin_totaltAvdeling = 0;

                        budgetInfo.comparelist.Add(new BudgetCompareData("own", "Din avdeling", omset_totaltAvdeling, inntjen_totaltAvdeling, margin_totaltAvdeling));
                    }

                    // Egen avdeling, sist år
                    var rowLastYear = main.database.CallMonthTable(dtTil.AddYears(-1), main.appConfig.Avdeling).Select(BudgetCategoryClass.GetSqlCategoryString(cat) + " AND (Dato >= '" + dtFra.AddYears(-1).ToString("yyy-MM-dd") + "' AND Dato <= '" + dtTil.AddYears(-1).ToString("yyy-MM-dd") + "')");
                    DataTable dtLastYear = rowLastYear.Any() ? rowLastYear.CopyToDataTable() : sqlce.Clone();
                    if (dtLastYear.Rows.Count > 0)
                    {
                        decimal compare_omset = 0, compare_inntjen = 0, compare_margin = 0, compare_omsetexmva = 0;

                        g = dtLastYear.Compute("Sum(Salgspris)", "");
                        if (!DBNull.Value.Equals(g))
                            compare_omset = Convert.ToDecimal(g);

                        g = dtLastYear.Compute("Sum(SalgsprisExMva)", "");
                        if (!DBNull.Value.Equals(g))
                            compare_omsetexmva = Convert.ToDecimal(g);

                        g = dtLastYear.Compute("Sum(Btokr)", "");
                        if (!DBNull.Value.Equals(g))
                            compare_inntjen = Convert.ToDecimal(g);

                        if (compare_omsetexmva != 0)
                            compare_margin = compare_inntjen / compare_omsetexmva;
                        else
                            compare_margin = 0;

                        budgetInfo.comparelist.Add(new BudgetCompareData("compareYear", "MTD i fjor", compare_omset, compare_inntjen, compare_margin));
                    }

                    // Egen avdeling, siste måned
                    var rowLastMonth = main.database.CallMonthTable(dtTil.AddMonths(-1), main.appConfig.Avdeling).Select(BudgetCategoryClass.GetSqlCategoryString(cat) + " AND (Dato >= '" + dtFra.AddMonths(-1).ToString("yyy-MM-dd") + "' AND Dato <= '" + dtTil.AddMonths(-1).ToString("yyy-MM-dd") + "')");
                    DataTable dtLastMonth = rowLastMonth.Any() ? rowLastMonth.CopyToDataTable() : sqlce.Clone();
                    if (dtLastMonth.Rows.Count > 0)
                    {
                        decimal compare_omset = 0, compare_inntjen = 0, compare_margin = 0, compare_omsetexmva = 0;

                        g = dtLastMonth.Compute("Sum(Salgspris)", "");
                        if (!DBNull.Value.Equals(g))
                            compare_omset = Convert.ToDecimal(g);

                        g = dtLastMonth.Compute("Sum(SalgsprisExMva)", "");
                        if (!DBNull.Value.Equals(g))
                            compare_omsetexmva = Convert.ToDecimal(g);

                        g = dtLastMonth.Compute("Sum(Btokr)", "");
                        if (!DBNull.Value.Equals(g))
                            compare_inntjen = Convert.ToDecimal(g);

                        if (compare_omsetexmva != 0)
                            compare_margin = compare_inntjen / compare_omsetexmva;
                        else
                            compare_margin = 0;

                        budgetInfo.comparelist.Add(new BudgetCompareData("compareMonth", "MTD " + dtTil.AddMonths(-1).ToString("MMMM"), compare_omset, compare_inntjen, compare_margin));
                    }

                    // Andre avdelinger, opp til 5 favoritter
                    favoritter = FormMain.Favoritter.ToArray();
                    for (int d = 1; d < favoritter.Length; d++)
                    {
                        if (d > 4) // Max 5 favoritter
                            break;
                        decimal fav_omset = 0, fav_inntjen = 0, fav_margin = 0, fav_omsetexmva = 0;
                        var rowFav = main.database.CallMonthTable(dtTil, favoritter[d]).Select(BudgetCategoryClass.GetSqlCategoryString(cat) + " AND (Dato >= '" + dtFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtTil.ToString("yyy-MM-dd") + "')");
                        DataTable dtFav = rowFav.Any() ? rowFav.CopyToDataTable() : sqlce.Clone();

                        g = dtFav.Compute("Sum(Salgspris)", "");
                        if (!DBNull.Value.Equals(g))
                            fav_omset = Convert.ToDecimal(g);

                        g = dtFav.Compute("Sum(SalgsprisExMva)", "");
                        if (!DBNull.Value.Equals(g))
                            fav_omsetexmva = Convert.ToDecimal(g);

                        g = dtFav.Compute("Sum(Btokr)", "");
                        if (!DBNull.Value.Equals(g))
                            fav_inntjen = Convert.ToDecimal(g);

                        if (fav_omsetexmva != 0)
                            fav_margin = fav_inntjen / fav_omsetexmva;
                        else
                            fav_margin = 0;

                        budgetInfo.comparelist.Add(new BudgetCompareData("favorite", favoritter[d], fav_omset, fav_inntjen, fav_margin));
                    }
                }

                // Totalen for alle selgerene
                string strSel = ""; DataRow[] rows;
                for (int i = 0; i < budgetInfo.selgere.Count; i++)
                    strSel += " OR Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'";
                strSel = strSel.Substring(4, strSel.Length - 4);
                rows = sqlce.Select(strSel + " AND (Dato >= '" + dtFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtTil.ToString("yyy-MM-dd") + "')");
                DataTable dtSel = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();

                g = dtSel.Compute("Sum(Salgspris)", "");
                if (!DBNull.Value.Equals(g))
                    Omset_totaltSelgere = Convert.ToDecimal(g);

                g = dtSel.Compute("Sum(SalgsprisExMva)", "");
                if (!DBNull.Value.Equals(g))
                    OmsetExMva_totaltSelgere = Convert.ToDecimal(g);

                g = dtSel.Compute("Sum(Btokr)", "");
                if (!DBNull.Value.Equals(g))
                    Inntjen_totaltSelgere = Convert.ToDecimal(g);

                DataRow dtRow = dtWork.NewRow();
                dtRow["Kategori"] = "Overall";
                dtRow["Selgerkode"] = "Totalt";
                dtRow["Actual_omset"] = Omset_totaltSelgere;
                dtRow["Actual_inntjen"] = Inntjen_totaltSelgere;
                dtRow["Actual_omsetExMva"] = Math.Round(OmsetExMva_totaltSelgere, 2);
                if (OmsetExMva_totaltSelgere != 0)
                    dtRow["Actual_margin"] = Inntjen_totaltSelgere / OmsetExMva_totaltSelgere;
                else
                    dtRow["Actual_margin"] = 0;

                dtRow["Target_omset"] = budgetInfo.omsetning * budgetInfo.timeElapsedCoefficient;
                dtRow["Target_inntjen"] = budgetInfo.inntjening * budgetInfo.timeElapsedCoefficient;
                dtRow["Target_omsetExMva"] = 0;
                dtRow["Target_margin"] = budgetInfo.margin;

                dtRow["Diff_omset"] = Omset_totaltSelgere - (decimal)dtRow["Target_omset"];
                dtRow["Diff_inntjen"] = Inntjen_totaltSelgere - (decimal)dtRow["Target_inntjen"];
                dtRow["Diff_omsetExMva"] = 0;
                dtRow["Diff_margin"] = (decimal)dtRow["Actual_margin"] - budgetInfo.margin;
                dtRow["Sort_value"] = -9999999;


                for (int i = 0; i < budgetInfo.selgere.Count; i++)
                {
                    var rowsSel = sqlce.Select("Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "' AND (Dato >= '" + dtFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtTil.ToString("yyy-MM-dd") + "')");
                    dt = rowsSel.Any() ? rowsSel.CopyToDataTable() : sqlce.Clone();

                    decimal Omset_selger = 0, OmsetExMva_selger = 0, Inntjen_selger = 0;

                    g = dt.Compute("Sum(Salgspris)", "");
                    if (!DBNull.Value.Equals(g))
                        Omset_selger = Convert.ToDecimal(g);

                    g = dt.Compute("Sum(SalgsprisExMva)", "");
                    if (!DBNull.Value.Equals(g))
                        OmsetExMva_selger = Convert.ToDecimal(g);

                    g = dt.Compute("Sum(Btokr)", "");
                    if (!DBNull.Value.Equals(g))
                        Inntjen_selger = Convert.ToDecimal(g);

                    DataRow dtRowSel = dtWork.NewRow();
                    dtRowSel["Kategori"] = "Overall";
                    dtRowSel["Selgerkode"] = budgetInfo.selgere[i].selgerkode;
                    dtRowSel["Actual_omset"] = Omset_selger;
                    dtRowSel["Actual_inntjen"] = Inntjen_selger;
                    dtRowSel["Actual_omsetExMva"] = OmsetExMva_selger;
                    if (OmsetExMva_selger != 0)
                        dtRowSel["Actual_margin"] = Inntjen_selger / OmsetExMva_selger;
                    else
                        dtRowSel["Actual_margin"] = 0;

                    dtRowSel["Target_omset"] = Math.Round(budgetInfo.omsetning * budgetInfo.timeElapsedCoefficient * budgetInfo.selgere[i].weight, 2);
                    dtRowSel["Target_inntjen"] = Math.Round(budgetInfo.inntjening * budgetInfo.timeElapsedCoefficient * budgetInfo.selgere[i].weight, 2);
                    dtRowSel["Target_omsetExMva"] = 0;
                    dtRowSel["Target_margin"] = budgetInfo.margin;

                    dtRowSel["Diff_omset"] = Omset_selger - (decimal)dtRowSel["Target_omset"];
                    dtRowSel["Diff_inntjen"] = Inntjen_selger - (decimal)dtRowSel["Target_inntjen"];
                    dtRowSel["Diff_omsetExMva"] = Math.Round(OmsetExMva_selger, 2) - (decimal)dtRowSel["Target_omsetExMva"];
                    dtRowSel["Diff_margin"] = (decimal)dtRowSel["Actual_margin"] - budgetInfo.margin;
                    dtRowSel["Sort_value"] = Inntjen_selger;
                    dtWork.Rows.Add(dtRowSel);

                    budgetInfo.chartdata_inntjen.Add(new BudgetChartData(budgetInfo.selgere[i].selgerkode, (decimal)dtRowSel["Target_inntjen"], Inntjen_selger));
                    budgetInfo.chartdata_omset.Add(new BudgetChartData(budgetInfo.selgere[i].selgerkode, (decimal)dtRowSel["Target_omset"], Omset_selger));

                    int timer = budgetInfo.selgere[i].timer;
                    if (timer == 0)
                        timer = 1;
                    budgetInfo.chartdata_kvalitet.Add(new BudgetChartData(budgetInfo.selgere[i].selgerkode, 0, Inntjen_selger / timer));
                    budgetInfo.chartdata_effektivitet.Add(new BudgetChartData(budgetInfo.selgere[i].selgerkode, 0, Omset_selger / timer));
                }

                budgetInfo.chartdata_inntjen = budgetInfo.chartdata_inntjen.OrderByDescending(x => x.actual).ToList();
                budgetInfo.chartdata_omset = budgetInfo.chartdata_omset.OrderByDescending(x => x.actual).ToList();

                budgetInfo.chartdata_kvalitet = budgetInfo.chartdata_kvalitet.OrderByDescending(x => x.actual).ToList();
                budgetInfo.chartdata_effektivitet = budgetInfo.chartdata_effektivitet.OrderByDescending(x => x.actual).ToList();

                dtWork.Rows.Add(dtRow);

                DataView dv = dtWork.DefaultView;
                dv.Sort = "Sort_value desc";


                var barchart = new BudgetBarChartData();
                barchart.actual_inntjen = Inntjen_totaltSelgere;
                barchart.actual_omset = Omset_totaltSelgere;
                if (OmsetExMva_totaltSelgere != 0)
                    barchart.actual_margin = OmsetExMva_totaltSelgere / Inntjen_totaltSelgere;
                else
                    barchart.actual_margin = 0;
                barchart.target_inntjen = budgetInfo.inntjening * budgetInfo.timeElapsedCoefficient;
                barchart.target_omset = budgetInfo.omsetning * budgetInfo.timeElapsedCoefficient;
                barchart.target_margin = budgetInfo.margin;
                if (budgetInfo.inntjening != 0)
                    barchart.result_inntjen = Math.Round(Inntjen_totaltSelgere / (budgetInfo.inntjening * budgetInfo.timeElapsedCoefficient) * 100, 2);
                else
                    barchart.result_inntjen = 0;
                if (budgetInfo.omsetning != 0)
                    barchart.result_omset = Math.Round(Omset_totaltSelgere / (budgetInfo.omsetning * budgetInfo.timeElapsedCoefficient) * 100, 2);
                else
                    barchart.result_omset = 0;
                budgetInfo.barchart = barchart;

                return dv.ToTable();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        private DataTable MakeTableBudgetProduct(BudgetCategory cat, BudgetType product, BudgetValueType type)
        {
            try
            {
                DataTable dtWork = ReadyTableProductNew();

                sqlce = main.database.CallMonthTable(dtTil, main.appConfig.Avdeling);
                if (sqlce.Rows.Count == 0)
                    return dtWork;

                string strSel = ""; DataRow[] rowsSel;
                for (int i = 0; i < budgetInfo.selgere.Count; i++)
                    strSel += " OR Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'";
                strSel = strSel.Substring(4, strSel.Length - 4);
                rowsSel = sqlce.Select(strSel + " AND (Dato >= '" + dtFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtTil.ToString("yyy-MM-dd") + "')");
                DataTable dtSel = rowsSel.Any() ? rowsSel.CopyToDataTable() : sqlce.Clone();

                object r;
                decimal omset_total = 0, inntjen_total = 0, omsetExMva_total = 0;
                int d = 0;
                if (cat == BudgetCategory.MDA)
                    d = 1;
                else if (cat == BudgetCategory.AudioVideo)
                    d = 2;
                else if (cat == BudgetCategory.SDA)
                    d = 3;
                else if (cat == BudgetCategory.Tele)
                    d = 4;
                else if (cat == BudgetCategory.Data)
                    d = 5;
                decimal target_value = 0;
                if (product == BudgetType.Acc)
                    target_value = budgetInfo.acc;
                else if (product == BudgetType.Finans)
                    target_value = budgetInfo.finans;
                else if (product == BudgetType.Rtgsa)
                    target_value = budgetInfo.rtgsa;
                else if (product == BudgetType.Strom)
                    target_value = budgetInfo.strom;
                else if (product == BudgetType.TA)
                    target_value = budgetInfo.ta;
                else if (product == BudgetType.Vinnprodukt)
                    target_value = budgetInfo.vinn;

                inntjen_total = Compute(dtSel, "Sum(Btokr)");
                omset_total = Compute(dtSel, "Sum(Salgspris)");
                omsetExMva_total = Compute(dtSel, "Sum(SalgsprisExMva)");

                DataRow dtRowAvd = dtWork.NewRow();
                dtRowAvd["Kategori"] = product;
                dtRowAvd["Selgerkode"] = "Totalt";

                dtRowAvd["Actual_inntjen"] = inntjen_total;
                dtRowAvd["Actual_omset"] = omset_total;
                dtRowAvd["Actual_omsetExMva"] = omsetExMva_total;
                if (omsetExMva_total != 0)
                    dtRowAvd["Actual_margin"] = Math.Round(inntjen_total / omsetExMva_total, 2);
                else
                    dtRowAvd["Actual_margin"] = 0;

                if (sqlce.Rows.Count > 0)
                {
                    decimal inntjen_prod = 0, antall_prod = 0, omset_prod = 0, hovedprod_antall = 0, vinnprodukt_poeng = 0;
                    KgsaBudgetProduct Actual = new KgsaBudgetProduct(product, type);
                    KgsaBudgetField actual_1 = new KgsaBudgetField();
                    KgsaBudgetField actual_2 = new KgsaBudgetField();
                    KgsaBudgetField actual_3 = new KgsaBudgetField();

                    KgsaBudgetProduct Target = new KgsaBudgetProduct(product, type);
                    KgsaBudgetField target_1 = new KgsaBudgetField();
                    KgsaBudgetField target_2 = new KgsaBudgetField();

                    KgsaBudgetProduct Difference = new KgsaBudgetProduct(product, type);
                    KgsaBudgetField difference_1 = new KgsaBudgetField();
                    KgsaBudgetField difference_2 = new KgsaBudgetField();

                    if (product == BudgetType.Vinnprodukt)
                    {
                        decimal antall = 0;
                        List<VinnproduktItem> items = main.vinnprodukt.GetList();
                        foreach (VinnproduktItem item in items)
                        {
                            inntjen_prod += Compute(dtSel, "Sum(Btokr)", "[Varekode] = '" + item.varekode + "'");
                            omset_prod += Compute(dtSel, "Sum(Salgspris)", "[Varekode] = '" + item.varekode + "'");
                            antall = Compute(dtSel, "Sum(Antall)", "[Varekode] = '" + item.varekode + "'");
                            antall_prod += antall;
                            vinnprodukt_poeng += item.poeng * antall;
                        }
                    }
                    else if (product == BudgetType.Acc)
                    {
                        int[] accessoriesGrpList = main.appConfig.GetAccessorieGroups(0);
                        foreach (int ac in accessoriesGrpList)
                        {
                            inntjen_prod += Compute(dtSel, "Sum(Btokr)", "[Varegruppe] = " + ac);
                            omset_prod += Compute(dtSel, "Sum(Salgspris)", "[Varegruppe] = " + ac);
                            antall_prod += Compute(dtSel, "Sum(Antall)", "[Varegruppe] = " + ac);
                        }
                    }
                    else if (product == BudgetType.Finans)
                    {
                        var rowf = dtSel.Select("[Varegruppe] = 961");
                        for (int f = 0; f < rowf.Length; f++)
                        {
                            var rows2 = sqlce.Select("[Bilagsnr] = " + rowf[f]["Bilagsnr"]);
                            DataTable dtFinans = rows2.Any() ? rows2.CopyToDataTable() : sqlce.Clone();
                            dtFinans.DefaultView.Sort = "Salgspris DESC";
                            int gruppe = Convert.ToInt32(dtFinans.Rows[0]["Varegruppe"].ToString().Substring(0, 1));
                            if (gruppe == d || d == 0)
                            {
                                r = sqlce.Compute("Sum(Btokr)", "[Varegruppe] = 961 AND [Bilagsnr] = " + dtFinans.Rows[0]["Bilagsnr"].ToString());
                                if (!DBNull.Value.Equals(r))
                                    inntjen_prod += Convert.ToDecimal(r);

                                r = sqlce.Compute("Sum(Antall)", "[Varegruppe] = 961 AND [Bilagsnr] = " + dtFinans.Rows[0]["Bilagsnr"].ToString());
                                if (!DBNull.Value.Equals(r))
                                    antall_prod += Convert.ToInt32(r);
                            }
                        }
                    }
                    else if (product == BudgetType.Rtgsa)
                    {
                        r = dtSel.Compute("Sum(Antall)", "([Varegruppe]=531 OR [Varegruppe]=533 OR [Varegruppe]=534 OR [Varegruppe]=224 OR [Varegruppe]=431)");
                        if (!DBNull.Value.Equals(r))
                            hovedprod_antall = Convert.ToInt32(r);

                        foreach (var varekode in varekoderAlle)
                        {
                            r = dtSel.Compute("Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                            if (!DBNull.Value.Equals(r))
                                inntjen_prod += Convert.ToDecimal(r);

                            r = dtSel.Compute("Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");
                            if (!DBNull.Value.Equals(r))
                                omset_prod += Convert.ToDecimal(r);

                            if (varekode.synlig)
                            {
                                r = dtSel.Compute("Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                                if (!DBNull.Value.Equals(r))
                                    antall_prod += Convert.ToInt32(r);
                            }
                        }
                    }
                    else if (product == BudgetType.Strom)
                    {
                        r = dtSel.Compute("Sum(Btokr)", "[Varekode] LIKE 'ELSTROM*' OR [Varekode] LIKE 'ELRABATT*'");
                        if (!DBNull.Value.Equals(r))
                            inntjen_prod = Convert.ToDecimal(r);

                        r = dtSel.Compute("Sum(Salgspris)", "[Varekode] LIKE 'ELSTROM*' OR [Varekode] LIKE 'ELRABATT*'");
                        if (!DBNull.Value.Equals(r))
                            omset_prod = Convert.ToDecimal(r);

                        r = dtSel.Compute("Sum(Antall)", "[Varekode] LIKE 'ELSTROM*'");
                        if (!DBNull.Value.Equals(r))
                            antall_prod = Convert.ToInt32(r);
                    }
                    else if (product == BudgetType.TA)
                    {
                        r = dtSel.Compute("Sum(Btokr)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                        if (!DBNull.Value.Equals(r))
                            inntjen_prod = Convert.ToDecimal(r);

                        r = dtSel.Compute("Sum(Salgspris)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                        if (!DBNull.Value.Equals(r))
                            omset_prod = Convert.ToDecimal(r);

                        r = dtSel.Compute("Sum(Antall)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                        if (!DBNull.Value.Equals(r))
                            antall_prod = Convert.ToInt32(r);
                    }

                    if (type == BudgetValueType.Poeng)
                    {
                        actual_1.value = vinnprodukt_poeng;
                        actual_1.type = BudgetValueType.Poeng;
                        actual_2.value = antall_prod;
                        actual_2.type = BudgetValueType.Antall;
                        actual_3.value = inntjen_prod;
                        actual_3.type = BudgetValueType.Inntjening;

                        target_1.value = target_value * budgetInfo.timeElapsedCoefficient;
                        target_1.type = BudgetValueType.Poeng;
                        target_2.value = (target_value * budgetInfo.timeElapsedCoefficient) / budgetInfo.daysElapsed;
                        target_2.type = BudgetValueType.AntallPerDag;

                        difference_1.value = actual_1.value - target_1.value;
                        difference_1.type = BudgetValueType.Poeng;
                        difference_2.value = ((antall_prod * budgetInfo.timeElapsedCoefficient) / budgetInfo.daysElapsed) - target_2.value;
                        difference_2.type = BudgetValueType.AntallPerDag;
                    }
                    else if (type == BudgetValueType.Hitrate)
                    {
                        if (hovedprod_antall != 0)
                            actual_1.value = antall_prod / hovedprod_antall;
                        else
                            actual_1.value = 0;
                        actual_1.type = BudgetValueType.Hitrate;
                        actual_2.value = antall_prod;
                        actual_2.type = BudgetValueType.Antall;
                        actual_3.value = inntjen_prod;
                        actual_3.type = BudgetValueType.Inntjening;

                        target_1.value = target_value;
                        target_1.type = BudgetValueType.Hitrate;
                        target_2.value = (hovedprod_antall * target_value) / budgetInfo.daysElapsed;
                        target_2.type = BudgetValueType.AntallPerDag;

                        difference_1.value = actual_1.value - target_1.value;
                        difference_1.type = BudgetValueType.Hitrate;
                        difference_2.value = ((hovedprod_antall * actual_1.value) / budgetInfo.daysElapsed) - target_2.value;
                        difference_2.type = BudgetValueType.AntallPerDag;
                    }
                    else if (type == BudgetValueType.Antall)
                    {
                        actual_1.value = antall_prod;
                        actual_1.type = BudgetValueType.Antall;
                        actual_2.value = inntjen_prod;
                        actual_2.type = BudgetValueType.Inntjening;
                        if (inntjen_total != 0)
                            actual_3.value = inntjen_prod / inntjen_total;
                        else
                            actual_3.value = 0;
                        actual_3.type = BudgetValueType.SoM;

                        target_1.value = target_value * budgetInfo.timeElapsedCoefficient;
                        target_1.type = BudgetValueType.Antall;
                        target_2.value = (target_value * budgetInfo.timeElapsedCoefficient) / budgetInfo.daysElapsed;
                        target_2.type = BudgetValueType.AntallPerDag;

                        difference_1.value = actual_1.value - target_1.value;
                        difference_1.type = BudgetValueType.Antall;
                        difference_2.value = ((antall_prod * budgetInfo.timeElapsedCoefficient) / budgetInfo.daysElapsed) - target_2.value;
                        difference_2.type = BudgetValueType.AntallPerDag;
                    }
                    if (type == BudgetValueType.Inntjening)
                    {
                        actual_1.value = antall_prod;
                        actual_1.type = BudgetValueType.Antall;
                        actual_2.value = inntjen_prod;
                        actual_2.type = BudgetValueType.Inntjening;
                        if (inntjen_total != 0)
                            actual_3.value = inntjen_prod / inntjen_total;
                        else
                            actual_3.value = 0;
                        actual_3.type = BudgetValueType.SoM;

                        target_1.value = target_value * budgetInfo.timeElapsedCoefficient;
                        target_1.type = BudgetValueType.Inntjening;
                        target_2.value = (target_value * budgetInfo.timeElapsedCoefficient) / budgetInfo.daysElapsed;
                        target_2.type = BudgetValueType.InntjeningPerDag;

                        difference_1.value = actual_2.value - target_1.value;
                        difference_1.type = BudgetValueType.Inntjening;
                        difference_2.value = (inntjen_prod / budgetInfo.daysElapsed) - target_2.value;
                        difference_2.type = BudgetValueType.InntjeningPerDag;
                    }
                    if (type == BudgetValueType.Omsetning)
                    {
                        actual_1.value = antall_prod;
                        actual_1.type = BudgetValueType.Antall;
                        actual_2.value = omset_prod;
                        actual_2.type = BudgetValueType.Omsetning;
                        if (omset_prod != 0)
                            actual_3.value = omset_total / omset_prod;
                        else
                            actual_3.value = 0;
                        actual_3.type = BudgetValueType.SoB;

                        target_1.value = target_value * budgetInfo.timeElapsedCoefficient;
                        target_1.type = BudgetValueType.Omsetning;
                        target_2.value = (target_value * budgetInfo.timeElapsedCoefficient) / budgetInfo.daysElapsed;
                        target_2.type = BudgetValueType.OmsetningPerDag;

                        difference_1.value = actual_2.value - target_1.value;
                        difference_1.type = BudgetValueType.Omsetning;
                        difference_2.value = (omset_prod / budgetInfo.daysElapsed) - target_2.value;
                        difference_2.type = BudgetValueType.OmsetningPerDag;
                    }
                    if (type == BudgetValueType.SoB)
                    {
                        actual_1.value = antall_prod;
                        actual_1.type = BudgetValueType.Antall;
                        actual_2.value = omset_prod;
                        actual_2.type = BudgetValueType.Omsetning;
                        if (omset_total != 0)
                            actual_3.value = omset_prod / omset_total;
                        else
                            actual_3.value = 0;
                        actual_3.type = BudgetValueType.SoB;

                        target_1.value = target_value;
                        target_1.type = BudgetValueType.SoB;
                        target_2.value = omset_total * target_value;
                        target_2.type = BudgetValueType.Omsetning;

                        difference_1.value = actual_3.value - target_1.value;
                        difference_1.type = BudgetValueType.SoB;
                        difference_2.value = actual_2.value - target_2.value;
                        difference_2.type = BudgetValueType.Omsetning;
                    }
                    if (type == BudgetValueType.SoM)
                    {
                        actual_1.value = antall_prod;
                        actual_1.type = BudgetValueType.Antall;
                        actual_2.value = inntjen_prod;
                        actual_2.type = BudgetValueType.Inntjening;
                        if (inntjen_total != 0)
                            actual_3.value = inntjen_prod / inntjen_total;
                        else
                            actual_3.value = 0;
                        actual_3.type = BudgetValueType.SoM;

                        target_1.value = target_value;
                        target_1.type = BudgetValueType.SoM;
                        target_2.value = target_1.value * inntjen_total;
                        target_2.type = BudgetValueType.Inntjening;

                        difference_1.value = actual_3.value - target_1.value;
                        difference_1.type = BudgetValueType.SoM;
                        difference_2.value = actual_2.value - target_2.value;
                        difference_2.type = BudgetValueType.Inntjening;
                    }

                    Actual.fields.Add(actual_1); // kolonne 1 av 3
                    Actual.fields.Add(actual_2); // kolonne 2 av 3
                    Actual.fields.Add(actual_3); // kolonne 3 av 3

                    Target.fields.Add(target_1); // kolonne 1 av 3
                    Target.fields.Add(target_2); // kolonne 2 av 3

                    Difference.fields.Add(difference_1); // kolonne 1 av 3
                    Difference.fields.Add(difference_2); // kolonne 2 av 3

                    dtRowAvd["Field1"] = PrepFields(Actual, false, true);
                    dtRowAvd["Field2"] = PrepFields(Target);
                    dtRowAvd["Field3"] = PrepFields(Difference, true);

                    dtRowAvd["Sort_value"] = -9999999;
                }

                budgetInfo.chartdata = new List<BudgetChartData>() { };

                for (int i = 0; i < budgetInfo.selgere.Count; i++)
                {
                    decimal omset_sel = 0, inntjen_sel = 0, omsetExMva_sel = 0;

                    inntjen_sel = Compute(sqlce, "Sum(Btokr)", "Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                    omset_sel = Compute(sqlce, "Sum(Salgspris)", "Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                    omsetExMva_sel = Compute(sqlce, "Sum(SalgsprisExMva)", "Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");

                    DataRow dtRowSel = dtWork.NewRow();
                    dtRowSel["Kategori"] = product;
                    dtRowSel["Selgerkode"] = budgetInfo.selgere[i].selgerkode;

                    dtRowSel["Actual_inntjen"] = inntjen_sel;
                    dtRowSel["Actual_omset"] = omset_sel;
                    dtRowSel["Actual_omsetExMva"] = omsetExMva_sel;
                    if (omsetExMva_sel != 0)
                        dtRowSel["Actual_margin"] = Math.Round(inntjen_sel / omsetExMva_sel, 2);
                    else
                        dtRowSel["Actual_margin"] = 0;

                    if (sqlce.Rows.Count > 0)
                    {
                        decimal inntjen_prod = 0, antall_prod = 0, omset_prod = 0, hovedprod_antall = 0, vinnprodukt_poeng = 0;
                        KgsaBudgetProduct Actual = new KgsaBudgetProduct(product, type);
                        KgsaBudgetField actual_1 = new KgsaBudgetField();
                        KgsaBudgetField actual_2 = new KgsaBudgetField();
                        KgsaBudgetField actual_3 = new KgsaBudgetField();
                        KgsaBudgetProduct Target = new KgsaBudgetProduct(product, type);
                        KgsaBudgetField target_1 = new KgsaBudgetField();
                        KgsaBudgetField target_2 = new KgsaBudgetField();
                        KgsaBudgetProduct Difference = new KgsaBudgetProduct(product, type);
                        KgsaBudgetField diff_1 = new KgsaBudgetField();
                        KgsaBudgetField diff_2 = new KgsaBudgetField();

                        if (product == BudgetType.Vinnprodukt)
                        {
                            decimal antall = 0;
                            List<VinnproduktItem> items = main.vinnprodukt.GetList();
                            foreach (VinnproduktItem item in items)
                            {
                                inntjen_prod += Compute(sqlce, "Sum(Btokr)", "[Varekode] = '" + item.varekode + "' AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                                omset_prod += Compute(sqlce, "Sum(Salgspris)", "[Varekode] = '" + item.varekode + "' AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                                antall = Compute(sqlce, "Sum(Antall)", "[Varekode] = '" + item.varekode + "' AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                                antall_prod += antall;
                                vinnprodukt_poeng += item.poeng * antall;
                            }
                        }
                        else if (product == BudgetType.Acc)
                        {
                            int[] accessoriesGrpList = main.appConfig.GetAccessorieGroups(0);
                            foreach (int ac in accessoriesGrpList)
                            {
                                r = sqlce.Compute("Sum(Antall)", "[Varegruppe] = " + ac + " AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                                if (!DBNull.Value.Equals(r))
                                    antall_prod += Convert.ToInt32(r);

                                r = sqlce.Compute("Sum(Btokr)", "[Varegruppe] = " + ac + " AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                                if (!DBNull.Value.Equals(r))
                                    inntjen_prod += Convert.ToDecimal(r);

                                r = sqlce.Compute("Sum(Salgspris)", "[Varegruppe] = " + ac + " AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                                if (!DBNull.Value.Equals(r))
                                    omset_prod += Convert.ToDecimal(r);
                            }
                        }
                        else if (product == BudgetType.Finans)
                        {
                            var rowf = sqlce.Select("[Varegruppe] = 961 AND [Selgerkode] = '" + budgetInfo.selgere[i].selgerkode + "'");
                            for (int f = 0; f < rowf.Length; f++)
                            {
                                var rows2 = sqlce.Select("[Bilagsnr] = " + rowf[f]["Bilagsnr"]);
                                DataTable dtFinans = rows2.Any() ? rows2.CopyToDataTable() : sqlce.Clone();
                                dtFinans.DefaultView.Sort = "Salgspris DESC";
                                int gruppe = Convert.ToInt32(dtFinans.Rows[0]["Varegruppe"].ToString().Substring(0, 1));
                                if (gruppe == d || d == 0)
                                {
                                    r = sqlce.Compute("Sum(Btokr)", "[Varegruppe] = 961 AND [Selgerkode] = '" + budgetInfo.selgere[i].selgerkode + "' AND [Bilagsnr] = " + dtFinans.Rows[0]["Bilagsnr"].ToString());
                                    if (!DBNull.Value.Equals(r))
                                        inntjen_prod += Convert.ToDecimal(r);

                                    r = sqlce.Compute("Sum(Antall)", "[Varegruppe] = 961 AND [Selgerkode] = '" + budgetInfo.selgere[i].selgerkode + "' AND [Bilagsnr] = " + dtFinans.Rows[0]["Bilagsnr"].ToString());
                                    if (!DBNull.Value.Equals(r))
                                        antall_prod += Convert.ToInt32(r);
                                }
                            }
                        }
                        else if (product == BudgetType.Rtgsa)
                        {
                            r = sqlce.Compute("Sum(Antall)", "([Varegruppe]=531 OR [Varegruppe]=533 OR [Varegruppe]=534 OR [Varegruppe]=224 OR [Varegruppe]=431) AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                            if (!DBNull.Value.Equals(r))
                                hovedprod_antall = Convert.ToInt32(r);

                            foreach (var varekode in varekoderAlle)
                            {
                                r = sqlce.Compute("Sum(Salgspris)", "[Varekode]='" + varekode.kode + "' AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                                if (!DBNull.Value.Equals(r))
                                    omset_prod += Convert.ToDecimal(r);

                                r = sqlce.Compute("Sum(Btokr)", "[Varekode]='" + varekode.kode + "' AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                                if (!DBNull.Value.Equals(r))
                                    inntjen_prod += Convert.ToDecimal(r);

                                if (varekode.synlig)
                                {
                                    r = sqlce.Compute("Sum(Antall)", "[Varekode]='" + varekode.kode + "' AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                                    if (!DBNull.Value.Equals(r))
                                        antall_prod += Convert.ToInt32(r);
                                }
                            }
                        }
                        else if (product == BudgetType.Strom)
                        {
                            r = sqlce.Compute("Sum(Btokr)", "([Varekode] LIKE 'ELSTROM*' OR [Varekode] LIKE 'ELRABATT*') AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                            if (!DBNull.Value.Equals(r))
                                inntjen_prod = Convert.ToDecimal(r);

                            r = sqlce.Compute("Sum(Salgspris)", "([Varekode] LIKE 'ELSTROM*' OR [Varekode] LIKE 'ELRABATT*') AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                            if (!DBNull.Value.Equals(r))
                                omset_prod = Convert.ToDecimal(r);

                            r = sqlce.Compute("Sum(Antall)", "[Varekode] LIKE 'ELSTROM*' AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                            if (!DBNull.Value.Equals(r))
                                antall_prod = Convert.ToInt32(r);
                        }
                        else if (product == BudgetType.TA)
                        {
                            r = sqlce.Compute("Sum(Btokr)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*' AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                            if (!DBNull.Value.Equals(r))
                                inntjen_prod = Convert.ToDecimal(r);

                            r = sqlce.Compute("Sum(Salgspris)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*' AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                            if (!DBNull.Value.Equals(r))
                                omset_prod = Convert.ToDecimal(r);

                            r = sqlce.Compute("Sum(Antall)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*' AND Selgerkode = '" + budgetInfo.selgere[i].selgerkode + "'");
                            if (!DBNull.Value.Equals(r))
                                antall_prod = Convert.ToInt32(r);
                        }

                        if (type == BudgetValueType.Poeng)
                        {
                            actual_1.value = vinnprodukt_poeng;
                            actual_1.type = BudgetValueType.Poeng;
                            actual_2.value = antall_prod;
                            actual_2.type = BudgetValueType.Antall;
                            actual_3.value = inntjen_prod;
                            actual_3.type = BudgetValueType.Inntjening;

                            target_1.value = target_value * budgetInfo.timeElapsedCoefficient * budgetInfo.selgere[i].weight;
                            target_1.type = BudgetValueType.Poeng;
                            target_2.value = ((target_value * budgetInfo.timeElapsedCoefficient) / budgetInfo.daysElapsed) * budgetInfo.selgere[i].weight;
                            target_2.type = BudgetValueType.AntallPerDag;

                            diff_1.value = actual_1.value - target_1.value;
                            diff_1.type = BudgetValueType.Poeng;
                            diff_2.value = ((antall_prod * budgetInfo.timeElapsedCoefficient) / budgetInfo.daysElapsed) - target_2.value;
                            diff_2.type = BudgetValueType.AntallPerDag;
                        }
                        else  if (type == BudgetValueType.Hitrate)
                        {
                            if (hovedprod_antall != 0)
                                actual_1.value = antall_prod / hovedprod_antall;
                            else
                                actual_1.value = 0;
                            actual_1.type = BudgetValueType.Hitrate;
                            actual_2.value = antall_prod;
                            actual_2.type = BudgetValueType.Antall;
                            actual_3.value = inntjen_prod;
                            actual_3.type = BudgetValueType.Inntjening;

                            target_1.value = target_value;
                            target_1.type = BudgetValueType.Hitrate;
                            target_2.value = (hovedprod_antall * target_value) / budgetInfo.daysElapsed;
                            target_2.type = BudgetValueType.AntallPerDag;

                            diff_1.value = actual_1.value - target_1.value;
                            diff_1.type = BudgetValueType.Hitrate;
                            diff_2.value = ((hovedprod_antall * actual_1.value) / budgetInfo.daysElapsed) - target_2.value;
                            diff_2.type = BudgetValueType.AntallPerDag;
                        }
                        else if (type == BudgetValueType.Antall)
                        {
                            actual_1.value = antall_prod;
                            actual_1.type = BudgetValueType.Antall;
                            actual_2.value = inntjen_prod;
                            actual_2.type = BudgetValueType.Inntjening;
                            if (inntjen_sel != 0)
                                actual_3.value = inntjen_prod / inntjen_sel;
                            else
                                actual_3.value = 0;
                            actual_3.type = BudgetValueType.SoM;

                            target_1.value = target_value * budgetInfo.timeElapsedCoefficient * budgetInfo.selgere[i].weight;
                            target_1.type = BudgetValueType.Antall;
                            target_2.value = ((target_value * budgetInfo.timeElapsedCoefficient) / budgetInfo.daysElapsed) * budgetInfo.selgere[i].weight;
                            target_2.type = BudgetValueType.AntallPerDag;

                            diff_1.value = actual_1.value - target_1.value;
                            diff_1.type = BudgetValueType.Antall;
                            diff_2.value = antall_prod / budgetInfo.daysElapsed - target_2.value;
                            diff_2.type = BudgetValueType.AntallPerDag;
                        }
                        else if (type == BudgetValueType.Inntjening)
                        {
                            actual_1.value = antall_prod;
                            actual_1.type = BudgetValueType.Antall;
                            actual_2.value = inntjen_prod;
                            actual_2.type = BudgetValueType.Inntjening;
                            if (inntjen_sel != 0)
                                actual_3.value = inntjen_prod / inntjen_sel;
                            else
                                actual_3.value = 0;
                            actual_3.type = BudgetValueType.SoM;

                            target_1.value = target_value * budgetInfo.timeElapsedCoefficient * budgetInfo.selgere[i].weight;
                            target_1.type = BudgetValueType.Inntjening;
                            target_2.value = ((target_value * budgetInfo.timeElapsedCoefficient) / budgetInfo.daysElapsed) * budgetInfo.selgere[i].weight;
                            target_2.type = BudgetValueType.InntjeningPerDag;

                            diff_1.value = actual_2.value - target_1.value;
                            diff_1.type = BudgetValueType.Inntjening;
                            diff_2.value = inntjen_prod / budgetInfo.daysElapsed - target_2.value;
                            diff_2.type = BudgetValueType.InntjeningPerDag;
                        }
                        else if (type == BudgetValueType.Omsetning)
                        {
                            actual_1.value = antall_prod;
                            actual_1.type = BudgetValueType.Antall;
                            actual_2.value = omset_prod;
                            actual_2.type = BudgetValueType.Omsetning;
                            if (omset_sel != 0)
                                actual_3.value = omset_prod / omset_sel;
                            else
                                actual_3.value = 0;
                            actual_3.type = BudgetValueType.SoB;

                            target_1.value = target_value * budgetInfo.timeElapsedCoefficient * budgetInfo.selgere[i].weight;
                            target_1.type = BudgetValueType.Omsetning;
                            target_2.value = ((target_value * budgetInfo.timeElapsedCoefficient) / budgetInfo.daysElapsed) * budgetInfo.selgere[i].weight;
                            target_2.type = BudgetValueType.OmsetningPerDag;

                            diff_1.value = actual_2.value - target_1.value;
                            diff_1.type = BudgetValueType.Omsetning;
                            diff_2.value = omset_prod / budgetInfo.daysElapsed - target_2.value;
                            diff_2.type = BudgetValueType.OmsetningPerDag;
                        }
                        else if (type == BudgetValueType.SoB)
                        {
                            actual_1.value = antall_prod;
                            actual_1.type = BudgetValueType.Antall;
                            actual_2.value = omset_prod;
                            actual_2.type = BudgetValueType.Omsetning;
                            if (omset_sel != 0)
                                actual_3.value = omset_prod / omset_sel;
                            else
                                actual_3.value = 0;
                            actual_3.type = BudgetValueType.SoB;

                            target_1.value = target_value;
                            target_1.type = BudgetValueType.SoB;
                            target_2.value = omset_sel * target_value;
                            target_2.type = BudgetValueType.Omsetning;

                            diff_1.value = actual_3.value - target_1.value;
                            diff_1.type = BudgetValueType.SoB;
                            diff_2.value = actual_2.value - target_2.value;
                            diff_2.type = BudgetValueType.Omsetning;
                        }
                        else if (type == BudgetValueType.SoM)
                        {
                            actual_1.value = antall_prod;
                            actual_1.type = BudgetValueType.Antall;
                            actual_2.value = inntjen_prod;
                            actual_2.type = BudgetValueType.Inntjening;
                            if (inntjen_sel != 0)
                                actual_3.value = inntjen_prod / inntjen_sel;
                            else
                                actual_3.value = 0;
                            actual_3.type = BudgetValueType.SoM;

                            target_1.value = target_value;
                            target_1.type = BudgetValueType.SoM;
                            target_2.value = inntjen_sel * target_value;
                            target_2.type = BudgetValueType.Inntjening;

                            diff_1.value = actual_3.value - target_1.value;
                            diff_1.type = BudgetValueType.SoM;
                            diff_2.value = actual_2.value - target_2.value;
                            diff_2.type = BudgetValueType.Inntjening;
                        }
                        Actual.fields.Add(actual_1); // kolonne 1 av 3
                        Actual.fields.Add(actual_2); // kolonne 2 av 3
                        Actual.fields.Add(actual_3); // kolonne 3 av 3

                        Target.fields.Add(target_1); // kolonne 1 av 2
                        Target.fields.Add(target_2); // kolonne 2 av 2

                        Difference.fields.Add(diff_1); // kolonne 1 av 2
                        Difference.fields.Add(diff_2); // kolonne 2 av 2

                        dtRowSel["Field1"] = PrepFields(Actual, false, true);
                        dtRowSel["Field2"] = PrepFields(Target);
                        dtRowSel["Field3"] = PrepFields(Difference, true);

                        if (type == BudgetValueType.Antall || type == BudgetValueType.Poeng)
                        {
                            dtRowSel["Sort_value"] = actual_1.value;
                            budgetInfo.chartdata.Add(new BudgetChartData(budgetInfo.selgere[i].selgerkode, target_1.value, actual_1.value));
                        }
                        else if (type == BudgetValueType.Inntjening || type == BudgetValueType.Omsetning)
                        {
                            dtRowSel["Sort_value"] = actual_2.value;
                            budgetInfo.chartdata.Add(new BudgetChartData(budgetInfo.selgere[i].selgerkode, target_1.value, actual_2.value));
                        }
                        else if (type == BudgetValueType.SoM || type == BudgetValueType.SoB)
                        {
                            dtRowSel["Sort_value"] = actual_3.value;
                            budgetInfo.chartdata.Add(new BudgetChartData(budgetInfo.selgere[i].selgerkode, target_1.value, actual_3.value));
                        }
                        else if (type == BudgetValueType.Hitrate)
                        {
                            dtRowSel["Sort_value"] = actual_1.value;
                            budgetInfo.chartdata.Add(new BudgetChartData(budgetInfo.selgere[i].selgerkode, target_1.value, actual_1.value));
                        }
                    }
                    dtWork.Rows.Add(dtRowSel);
                }

                dtWork.Rows.Add(dtRowAvd);
                DataView dv = dtWork.DefaultView;
                dv.Sort = "Sort_value desc";

                budgetInfo.chartdata = budgetInfo.chartdata.OrderByDescending(x => x.actual).ToList();

                return dv.ToTable();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        private KgsaBudgetProduct PrepFields(KgsaBudgetProduct obj, bool positiveIsGreen = false, bool markMainFiel = false)
        {
            var mainType = obj.mainType;
            foreach(KgsaBudgetField field in obj.fields)
            {
                if (field.type == BudgetValueType.Poeng)
                {
                    field.value = Math.Round(field.value, 0);
                    field.header = "Poeng";
                    field.suffix = "";
                    field.style = "numbers-small";
                    field.sorter = "digit";
                    field.filter = "#,##0";
                }
                else if (field.type == BudgetValueType.Antall)
                {
                    field.value = Math.Round(field.value, 0);
                    field.header = "Antall";
                    field.suffix = "";
                    field.style = "numbers-small";
                    field.sorter = "digit";
                    field.filter = "#,##0";
                }
                else if (field.type == BudgetValueType.AntallPerDag)
                {
                    field.value = Math.Round(field.value, 1);
                    field.header = "Antall pr.dag";
                    field.suffix = " stk /dag";
                    field.style = "numbers-small";
                    field.sorter = "digit";
                    field.filter = "";
                }
                else if (field.type == BudgetValueType.Empty)
                {
                    field.header = "";
                    field.suffix = "";
                    field.style = "numbers-gen";
                    field.sorter = "digit";
                    field.filter = "";
                }
                else if (field.type == BudgetValueType.Hitrate)
                {
                    field.value = Math.Round(field.value * 100, 2);
                    field.header = "Hitrate";
                    field.suffix = " %";
                    field.style = "numbers-percent";
                    field.sorter = "procent";
                    field.filter = "";
                }
                else if (field.type == BudgetValueType.Inntjening)
                {
                    field.value = Math.Round(field.value, 0);
                    field.header = "Inntjen.";
                    field.suffix = "";
                    field.style = "numbers-gen";
                    field.sorter = "digit";
                    field.filter = "#,##0";
                }
                else if (field.type == BudgetValueType.InntjeningPerDag)
                {
                    field.value = Math.Round(field.value, 2);
                    field.header = "Inntjen. pr.dag";
                    field.suffix = " kr /dag";
                    field.style = "numbers-gen";
                    field.sorter = "digit";
                    field.filter = "#,##0.00";
                }
                else if (field.type == BudgetValueType.Omsetning)
                {
                    field.value = Math.Round(field.value, 0);
                    field.header = "Omset.";
                    field.suffix = "";
                    field.style = "numbers-gen";
                    field.sorter = "digit";
                    field.filter = "#,##0.00";
                }
                else if (field.type == BudgetValueType.OmsetningPerDag)
                {
                    field.value = Math.Round(field.value, 2);
                    field.header = "Omset. pr.dag";
                    field.suffix = " kr /dag";
                    field.style = "numbers-gen";
                    field.sorter = "digit";
                    field.filter = "";
                }
                else if (field.type == BudgetValueType.SoB)
                {
                    field.value = Math.Round(field.value * 100, 2);
                    field.header = "SoB";
                    field.suffix = " %";
                    field.style = "numbers-percent";
                    field.sorter = "procent";
                    field.filter = "";
                }
                else if (field.type == BudgetValueType.SoM)
                {
                    field.value = Math.Round(field.value * 100, 2);
                    field.header = "SoM";
                    field.suffix = " %";
                    field.style = "numbers-percent";
                    field.sorter = "procent";
                    field.filter = "";
                }

                if (mainType == field.type && markMainFiel)
                {
                    if (obj.product == BudgetType.Vinnprodukt)
                        field.style = "numbers-gen";
                    else if (obj.product == BudgetType.Acc)
                        field.style = "numbers-finans";
                    else if (obj.product == BudgetType.Finans)
                        field.style = "numbers-finans";
                    else if (obj.product == BudgetType.Rtgsa)
                        field.style = "numbers-service";
                    else if (obj.product == BudgetType.Strom)
                        field.style = "numbers-strom";
                    else if (obj.product == BudgetType.TA)
                        field.style = "numbers-moderna";
                    else if (obj.product == BudgetType.Inntjening)
                        field.style = "numbers-gen";
                    else if (obj.product == BudgetType.Omsetning)
                        field.style = "numbers-gen";
                }

                if (field.value < 0)
                    field.text = "<span style='color:red'>" + field.value.ToString(field.filter) + field.suffix + "</span>";
                else if (field.value > 0)
                {
                    if (positiveIsGreen)
                        field.text = "<span style='color:green'>+" + field.value.ToString(field.filter) + field.suffix + "</span>";
                    else
                        field.text = field.value.ToString(field.filter) + field.suffix;
                }
                else
                    field.text = main.appConfig.visningNull;
            }

            return obj;
        }

        public List<string> GetTableHtml(BudgetCategory cat, BackgroundWorker bw = null)
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                var hashId = random.Next(999, 99999);

                string urlID = "link";
                doc.Add("<br>");

                dtBudgetOverall = MakeTableBudget(cat);
                if (dtBudgetOverall.Rows.Count == 0)
                {
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                    return doc;
                }

                DateTime date = dtTil;
                if (date > main.appConfig.dbTo)
                    date = main.appConfig.dbTo;

                doc.Add("<h3>Resultat for " + BudgetCategoryClass.TypeToName(cat) + " budsjett oppdatert "
                    + date.ToString("dddd d. MMMM yyyy", FormMain.norway) + " og timeantall oppdatert "
                    + budgetInfo.updated.ToString("dddd d. MMMM yyyy", FormMain.norway) + "</h2>");

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table style='font: bold 12px/18px Arial, Sans-serif;' >");
                doc.Add("<tr><td style='background-color:#D9D9D9;text-align:center;padding: 0px 5px 0px 5px;' width=124 >" + BudgetCategoryClass.TypeToName(cat).ToUpper() + "</td>");
                doc.Add("<td style='background-color:#ff5b2e;text-align:center;padding: 0px 5px 0px 5px;border-left:1px solid black;' width=282 >Resultat</td>");
                doc.Add("<td style='background-color:#1163b0;text-align:center;padding: 0px 5px 0px 5px;border-left:1px solid black;' width=237 >Mål</td>");
                doc.Add("<td style='background-color:orange;text-align:center;padding: 0px 5px 0px 5px;border-left:1px solid black;' width=242 >Differanse</td>");
                doc.Add("</tr>");
                doc.Add("</table>");

                doc.Add("<table class='tablesorter'>");
                doc.Add("<thead><tr>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=130 >Selgerkode</td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" style='border-left:1px solid black;background-color:#ff5b2e;' width=95 >Omset.</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" style='background-color:#ff5b2e;' width=95 >Inntjen.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" style='background-color:#ff5b2e;' width=90 >Margin</td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" style='border-left:1px solid black;background-color:#1163b0;' width=80 >Omset.</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" style='background-color:#1163b0;' width=80 >Inntjen.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" style='background-color:#1163b0;' width=75 >Margin</td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" style='border-left:1px solid black;background-color:orange;' width=80 >Omset.</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" style='background-color:orange;' width=80 >Inntjen.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" style='background-color:orange;' width=80 >Margin</td>");
                doc.Add("</tr></thead>");
                doc.Add("<tbody>");

                for (int i = 0; i < dtBudgetOverall.Rows.Count; i++)
                {
                    string strLine = dtBudgetOverall.Rows[i]["Selgerkode"].ToString() == "Totalt" ? " style='border-top:2px solid black;' " : "";

                    if (dtBudgetOverall.Rows[i]["Selgerkode"].ToString() == "Totalt")
                        doc.Add("</tbody><tfoot><tr><td class='text-cat' " + strLine + ">" + main.salesCodes.GetNavn(dtBudgetOverall.Rows[i]["Selgerkode"].ToString()) + "</td>");
                    else
                        doc.Add("<tr><td class='text-cat' " + strLine + "><a href='#" + urlID + "b" + dtBudgetOverall.Rows[i]["Selgerkode"].ToString() + "'>" + main.salesCodes.GetNavn(dtBudgetOverall.Rows[i]["Selgerkode"].ToString()) + "</a></td>");

                    doc.Add("<td class='numbers-gen'" + strLine + ">" + budget.BudgetPlusMinus((decimal)dtBudgetOverall.Rows[i]["Actual_Omset"]) + "</td>");
                    doc.Add("<td class='numbers-gen'" + strLine + ">" + budget.BudgetPlusMinus((decimal)dtBudgetOverall.Rows[i]["Actual_Inntjen"]) + "</td>");
                    doc.Add("<td class='numbers-percent'" + strLine + ">" + budget.BudgetPlusMinus((decimal)dtBudgetOverall.Rows[i]["Actual_Margin"] * 100, 2, false, " %") + "</td>");

                    doc.Add("<td class='numbers-gen'" + strLine + ">" + budget.BudgetPlusMinus((decimal)dtBudgetOverall.Rows[i]["Target_Omset"]) + "</td>");
                    doc.Add("<td class='numbers-gen'" + strLine + ">" + budget.BudgetPlusMinus((decimal)dtBudgetOverall.Rows[i]["Target_Inntjen"]) + "</td>");
                    doc.Add("<td class='numbers-percent'" + strLine + ">" + budget.BudgetPlusMinus((decimal)dtBudgetOverall.Rows[i]["Target_Margin"] * 100, 2, false, " %") + "</td>");

                    doc.Add("<td class='numbers-gen'" + strLine + ">" + budget.BudgetPlusMinus((decimal)dtBudgetOverall.Rows[i]["Diff_Omset"], 0, true) + "</td>");
                    doc.Add("<td class='numbers-gen'" + strLine + ">" + budget.BudgetPlusMinus((decimal)dtBudgetOverall.Rows[i]["Diff_Inntjen"], 0, true) + "</td>");
                    doc.Add("<td class='numbers-percent'" + strLine + ">" + budget.BudgetPlusMinus((decimal)dtBudgetOverall.Rows[i]["Diff_Margin"] * 100, 2, true, " %") + "</td>");

                    doc.Add("</tr>");
                }
                doc.Add("</tfoot>");

                if (budgetInfo.comparelist != null)
                {
                    // Detailed look START
                    doc.Add("<tfoot><tr><td colspan='5' style='padding: 0px 0px 0px 0px;border:none;'>");
                    doc.Add("<table class='tablesorter'><thead><tr>");
                    doc.Add("<th class=\"{sorter: 'text'}\" style='border-top:#000 1px solid;border-left:#000 1px solid;' width=160 >Avdeling (" + BudgetCategoryClass.TypeToName(cat) + ")</td>");
                    doc.Add("<th class=\"{sorter: 'digit'}\" style='border-top:#000 1px solid;' width=103 >Omset.</td>");
                    doc.Add("<th class=\"{sorter: 'digit'}\" style='border-top:#000 1px solid;' width=103 >Inntjen.</td>");
                    doc.Add("<th class=\"{sorter: 'procent'}\" style='border-top:#000 1px solid;border-right:#000 1px solid;' width=94 >Margin</td>");
                    doc.Add("</tr></thead><tbody>");
                    
                    for (int i = 0; i < budgetInfo.comparelist.Count; i++)
                    {
                        if (budgetInfo.comparelist[i].type == "favorite")
                            doc.Add("<tr><td class='text-cat' style='border-bottom: #ccc 1px solid;border-left: #ccc 1px solid;'>" + avdeling.Get(Convert.ToInt32(budgetInfo.comparelist[i].navn)).Replace(" ", "&nbsp;") + "</td>");
                        else
                            doc.Add("<tr><td class='text-cat' style='border-bottom: #ccc 1px solid;border-left: #ccc 1px solid;'>" + budgetInfo.comparelist[i].navn + "</td>");

                        doc.Add("<td class='numbers-gen' style='border-bottom: #ccc 1px solid;'>" + budget.BudgetPlusMinus((decimal)budgetInfo.comparelist[i].omset) + "</td>");
                        doc.Add("<td class='numbers-gen' style='border-bottom: #ccc 1px solid;'>" + budget.BudgetPlusMinus((decimal)budgetInfo.comparelist[i].inntjen) + "</td>");
                        doc.Add("<td class='numbers-percent' style='border-bottom: #ccc 1px solid;border-right: #ccc 1px solid;'>" + budget.BudgetPlusMinus((decimal)budgetInfo.comparelist[i].margin * 100, 2, false, " %") + "</td></tr>");
                    }

                    doc.Add("</tr></body></table>");
                    budget.SaveBarChartImage(FormMain.settingsPath + @"\graphBudgetOverall" + BudgetCategoryClass.TypeToName(cat) + ".png", budgetInfo);
                    doc.Add("</td><td colspan='5' style='padding: 0px 0px 0px 0px;border:none;'><img src='graphBudgetOverall" + BudgetCategoryClass.TypeToName(cat) + ".png' style='width:410px;height:auto;'></td></tr>");
                    doc.Add("</tfoot>");
                    // Detailed look END
                }

                doc.Add("</table>");
                doc.Add("</td></tr></table>");

                decimal count = budgetInfo.selgere.Count;
                if (count < main.appConfig.budgetChartMinPosts)
                    count = main.appConfig.budgetChartMinPosts;
                decimal chartHeight = 100 + main.appConfig.budgetChartPostWidth * count;

                hashId = random.Next(999, 99999);
                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<div class='no-break'>");
                doc.Add("<h3>Inntjening</h3>");
                doc.Add("<table class='toggleAll' style='background-color:white;border:1px solid black;' id='" + hashId + "'><tr><td>");
                budget.SaveChartImage(main.appConfig.graphResX, (int)chartHeight, FormMain.settingsPath + @"\graphBudgetInntjen" + BudgetCategoryClass.TypeToName(cat) + ".png", BudgetValueType.Inntjening, BudgetType.Inntjening, budgetInfo.chartdata_inntjen);
                doc.Add("<img src='graphBudgetInntjen" + BudgetCategoryClass.TypeToName(cat) + ".png' style='width:900px;height:auto;'><br>");
                doc.Add("</td></tr></table>");
                doc.Add("</div>");

                doc.Add("<div class='no-break'>");
                doc.Add("<h3>Omsetning</h3>");
                doc.Add("<table class='toggleAll' style='background-color:white;border:1px solid black;' id='" + hashId + "'><tr><td>");
                budget.SaveChartImage(main.appConfig.graphResX, (int)chartHeight, FormMain.settingsPath + @"\graphBudgetOmset" + BudgetCategoryClass.TypeToName(cat) + ".png", BudgetValueType.Omsetning, BudgetType.Omsetning, budgetInfo.chartdata_omset);
                doc.Add("<img src='graphBudgetOmset" + BudgetCategoryClass.TypeToName(cat) + ".png' style='width:900px;height:auto;'><br>");
                doc.Add("</td></tr></table>");
                doc.Add("</div>");

                if (main.appConfig.budgetChartShowEfficiency || main.appConfig.budgetChartShowQuality)
                {
                    hashId = random.Next(999, 99999);
                    doc.Add("<div class='toolbox hidePdf'>");
                    doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                    doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                    doc.Add("</div>");

                    if (main.appConfig.budgetChartShowQuality)
                    {
                        doc.Add("<div class='no-break'>");
                        doc.Add("<h3>Kvalitet: Inntjening pr. time</h3>");
                        doc.Add("<table class='toggleAll' style='background-color:white;border:1px solid black;' id='" + hashId + "'><tr><td>");
                        budget.SaveChartImage(main.appConfig.graphResX, (int)chartHeight, FormMain.settingsPath + @"\graphBudgetKvalitet" + BudgetCategoryClass.TypeToName(cat) + ".png", BudgetValueType.Inntjening, BudgetType.Kvalitet, budgetInfo.chartdata_kvalitet);
                        doc.Add("<img src='graphBudgetKvalitet" + BudgetCategoryClass.TypeToName(cat) + ".png' style='width:900px;height:auto;'><br>");
                        doc.Add("</td></tr></table>");
                        doc.Add("</div>");
                    }
                    if (main.appConfig.budgetChartShowEfficiency)
                    {
                        doc.Add("<div class='no-break'>");
                        doc.Add("<h3>Effektivitet: Omsetning pr. time</h3>");
                        doc.Add("<table class='toggleAll' style='background-color:white;border:1px solid black;' id='" + hashId + "'><tr><td>");
                        budget.SaveChartImage(main.appConfig.graphResX, (int)chartHeight, FormMain.settingsPath + @"\graphBudgetEffektivitet" + BudgetCategoryClass.TypeToName(cat) + ".png", BudgetValueType.Omsetning, BudgetType.Effektivitet, budgetInfo.chartdata_effektivitet);
                        doc.Add("<img src='graphBudgetEffektivitet" + BudgetCategoryClass.TypeToName(cat) + ".png' style='width:900px;height:auto;'><br>");
                        doc.Add("</td></tr></table>");
                        doc.Add("</div>");
                    }
                }

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableHtmlProduct(BudgetCategory cat, BudgetType product)
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                var hashId = random.Next(999, 99999);

                string urlID = "link";
                doc.Add("<br>");
                doc.Add("<a name='" + budget.ProductToString(product) + "'></a>"); 

                budgetInfo = budget.GetBudgetInfo(dtTil, cat);
                if (budgetInfo == null)
                {
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Budsjett for " + BudgetCategoryClass.TypeToName(cat) + " mangler for valgt måned.</span><br>");
                    return doc;
                }
                if (budgetInfo.selgere.Count == 0)
                {
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Ingen selgerkoder er satt opp for " + BudgetCategoryClass.TypeToName(cat) + " budsjettet.</span><br>");
                    return doc;
                }
                BudgetValueType type = budgetInfo.ProductToType(product);
                DataTable dtBudgetProduct = MakeTableBudgetProduct(cat, product, type);
                string productColor = budget.GetProductColor(product);

                DateTime date = dtTil;
                if (date > main.appConfig.dbTo)
                    date = main.appConfig.dbTo;

                doc.Add("<h2>Resultat for " + budget.ProductToString(product) + " - " + BudgetCategoryClass.TypeToName(cat) + " budsjett oppdatert " + date.ToString("dddd d. MMMM yyyy", FormMain.norway) + " og timeantall oppdatert " + budgetInfo.updated.ToString("dddd d. MMMM yyyy", FormMain.norway) + "</h2>");

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");

                doc.Add("<table style='font: bold 12px/18px Arial, Sans-serif;' >");
                doc.Add("<tr><td style='background-color:#D9D9D9;text-align:center;padding: 0px 5px 0px 5px;' width=144 >" + BudgetCategoryClass.TypeToName(cat).ToUpper() + "</td>");
                doc.Add("<td style='background-color:" + productColor + ";text-align:center;padding: 0px 5px 0px 5px;border-left:1px solid black;' width=362 >Resultat</td>");
                doc.Add("<td style='background-color:#1163b0;text-align:center;padding: 0px 5px 0px 5px;border-left:1px solid black;' width=208 >Mål</td>");
                doc.Add("<td style='background-color:orange;text-align:center;padding: 0px 5px 0px 5px;border-left:1px solid black;' width=188 >Differanse</td>");
                doc.Add("</tr>");
                doc.Add("</table>");

                doc.Add("<table class='tablesorter'>");
                doc.Add("<thead><tr>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=150 >" + budget.ProductToString(product) + "</td>");
                KgsaBudgetProduct proActual = (KgsaBudgetProduct)dtBudgetProduct.Rows[0]["Field1"];
                doc.Add("<th class=\"{sorter: '" + proActual.fields[0].sorter + "'}\" style='border-left:1px solid black;background-color:" + productColor + ";' width=130 >" + proActual.fields[0].header + "</td>");
                doc.Add("<th class=\"{sorter: '" + proActual.fields[1].sorter + "'}\" style='background-color:" + productColor + ";' width=130 >" + proActual.fields[1].header + "</td>");
                doc.Add("<th class=\"{sorter: '" + proActual.fields[2].sorter + "'}\" style='background-color:" + productColor + ";' width=100 >" + proActual.fields[2].header + "</td>");

                KgsaBudgetProduct proTarget = (KgsaBudgetProduct)dtBudgetProduct.Rows[0]["Field2"];
                doc.Add("<th class=\"{sorter: '" + proTarget.fields[0].sorter + "'}\" style='border-left:1px solid black;background-color:#1163b0;' width=120 >" + proTarget.fields[0].header + "</td>");
                doc.Add("<th class=\"{sorter: '" + proTarget.fields[1].sorter + "'}\" style='background-color:#1163b0;' width=90 >" + proTarget.fields[1].header + "</td>");

                KgsaBudgetProduct proDiff = (KgsaBudgetProduct)dtBudgetProduct.Rows[0]["Field3"];
                doc.Add("<th class=\"{sorter: '" + proDiff.fields[0].sorter + "'}\" style='border-left:1px solid black;background-color:orange;' width=100 >" + proDiff.fields[0].header + "</td>");
                doc.Add("<th class=\"{sorter: '" + proDiff.fields[1].sorter + "'}\" style='background-color:orange;' width=90 >" + proDiff.fields[1].header + "</td>");
                doc.Add("</tr></thead>");
                doc.Add("<tbody>");

                string sorterStyle = type == BudgetValueType.Antall || type == BudgetValueType.Inntjening || type == BudgetValueType.Omsetning ? "numbers-gen" : "percent";
                for (int i = 0; i < dtBudgetProduct.Rows.Count; i++)
                {
                    string strLine = dtBudgetOverall.Rows[i]["Selgerkode"].ToString() == "Totalt" ? " style='border-top:2px solid black;border-bottom:none;' " : "";

                    if (dtBudgetOverall.Rows[i]["Selgerkode"].ToString() == "Totalt")
                        doc.Add("</tbody><tfoot><tr><td class='text-cat' " + strLine + ">" + main.salesCodes.GetNavn(dtBudgetProduct.Rows[i]["Selgerkode"].ToString()) + "</td>");
                    else
                        doc.Add("<tr><td class='text-cat' " + strLine + "><a href='#" + urlID + "b" + dtBudgetProduct.Rows[i]["Selgerkode"].ToString() + "'>" + main.salesCodes.GetNavn(dtBudgetProduct.Rows[i]["Selgerkode"].ToString()) + "</a></td>");

                    proActual = (KgsaBudgetProduct)dtBudgetProduct.Rows[i]["Field1"];
                    doc.Add("<td class='" + proActual.fields[0].style + "'" + strLine + ">" + proActual.fields[0].text + "</td>");
                    doc.Add("<td class='" + proActual.fields[1].style + "'" + strLine + ">" + proActual.fields[1].text + "</td>");
                    doc.Add("<td class='" + proActual.fields[2].style + "'" + strLine + ">" + proActual.fields[2].text + "</td>");

                    proTarget = (KgsaBudgetProduct)dtBudgetProduct.Rows[i]["Field2"];
                    doc.Add("<td class='" + proTarget.fields[0].style + "'" + strLine + ">" + proTarget.fields[0].text + "</td>");
                    doc.Add("<td class='" + proTarget.fields[1].style + "'" + strLine + ">" + proTarget.fields[1].text + "</td>");

                    proDiff = (KgsaBudgetProduct)dtBudgetProduct.Rows[i]["Field3"];
                    doc.Add("<td class='" + proDiff.fields[0].style + "'" + strLine + ">" + proDiff.fields[0].text + "</td>");
                    doc.Add("<td class='" + proDiff.fields[1].style + "'" + strLine + ">" + proDiff.fields[1].text + "</td>");

                    doc.Add("</tr>");
                }
                doc.Add("</tfoot></table>");
                doc.Add("</td></tr></table>");

                decimal count = budgetInfo.selgere.Count;
                if (count < main.appConfig.budgetChartMinPosts)
                    count = main.appConfig.budgetChartMinPosts;
                decimal chartHeight = 100 + main.appConfig.budgetChartPostWidth * count;

                hashId = random.Next(999, 99999);
                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<div class='no-break'>");
                doc.Add("<h3>" + budget.ProductToString(product) + "</h3>");
                doc.Add("<table class='toggleAll' style='background-color:white;border:1px solid black;' id='" + hashId + "'><tr><td>");
                budget.SaveChartImage(main.appConfig.graphResX, (int)chartHeight, FormMain.settingsPath + @"\graphBudget" + product + BudgetCategoryClass.TypeToName(cat) + ".png", type, product, budgetInfo.chartdata);
                doc.Add("<img src='graphBudget" + product + BudgetCategoryClass.TypeToName(cat) + ".png' style='width:900px;height:auto;'><br>");
                doc.Add("</td></tr></table>");
                doc.Add("</div>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        public DataTable ReadyTableOverall()
        {
            var dt = new DataTable();
            dt.Columns.Add("Kategori", typeof(string));

            dt.Columns.Add("Selgerkode", typeof(string));

            dt.Columns.Add("Actual_omset", typeof(decimal));
            dt.Columns.Add("Actual_omsetExMva", typeof(decimal));
            dt.Columns.Add("Actual_inntjen", typeof(decimal));
            dt.Columns.Add("Actual_margin", typeof(decimal));

            dt.Columns.Add("Target_omset", typeof(decimal));
            dt.Columns.Add("Target_omsetExMva", typeof(decimal));
            dt.Columns.Add("Target_inntjen", typeof(decimal));
            dt.Columns.Add("Target_margin", typeof(decimal));

            dt.Columns.Add("Diff_omset", typeof(decimal));
            dt.Columns.Add("Diff_omsetExMva", typeof(decimal));
            dt.Columns.Add("Diff_inntjen", typeof(decimal));
            dt.Columns.Add("Diff_margin", typeof(decimal));

            dt.Columns.Add("Sort_value", typeof(decimal));

            return dt;
        }

        public DataTable ReadyTableProductNew()
        {
            var dt = new DataTable();
            dt.Columns.Add("Kategori", typeof(string));

            dt.Columns.Add("Selgerkode", typeof(string));

            dt.Columns.Add("Actual_inntjen", typeof(decimal));
            dt.Columns.Add("Actual_omset", typeof(decimal));
            dt.Columns.Add("Actual_omsetExMva", typeof(decimal));
            dt.Columns.Add("Actual_margin", typeof(decimal));

            dt.Columns.Add("Field1", typeof(KgsaBudgetProduct));
            dt.Columns.Add("Field2", typeof(KgsaBudgetProduct));
            dt.Columns.Add("Field3", typeof(KgsaBudgetProduct));

            dt.Columns.Add("Sort_value", typeof(decimal));

            return dt;
        }

        public string ValueToDisplay(string strArg, BudgetValueType type, bool green = false)
        {
            try
            {
                decimal value = 0;
                if (!decimal.TryParse(strArg, out value))
                    return main.appConfig.visningNull;

                string filter = "";
                string suffix = "";
                if (type == BudgetValueType.Hitrate || type == BudgetValueType.SoB || type == BudgetValueType.SoM)
                {
                    value = Math.Round(value, 2);
                    suffix = " %";
                }
                else
                {
                    value = Math.Round(value, 0);
                    filter = "#,##0";
                }
                string valueStr = main.appConfig.visningNull;

                if (value < 0)
                    valueStr = "<span style='color:red'>" + value.ToString(filter) + suffix + "</span>";
                if (value > 0)
                {
                    if (green)
                        valueStr = "<span style='color:green'>" + value.ToString(filter) + suffix + "</span>";
                    else
                        valueStr = value.ToString(filter) + suffix;
                }

                return valueStr;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return main.appConfig.visningNull;
        }
    }


    public class KgsaBudgetProduct
    {
        public BudgetType product { get; set; }
        public BudgetValueType mainType { get; set; }
        public List<KgsaBudgetField> fields { get; set; }
        public KgsaBudgetProduct(BudgetType product, BudgetValueType type)
        {
            this.product = product;
            this.mainType = type;
            fields = new List<KgsaBudgetField>() { };
        }
    }

    public class KgsaBudgetField
    {
        public decimal value { get; set; }
        public BudgetValueType type { get; set; }
        public string suffix { get; set; }
        public string header { get; set; }
        public string style { get; set; }
        public string sorter { get; set; }
        public string filter { get; set; }
        public string text { get; set; }
    }

}
