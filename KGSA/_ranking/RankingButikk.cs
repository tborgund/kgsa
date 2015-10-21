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
using System.ComponentModel;
using System.Threading;

namespace KGSA
{
    public class RankingButikk : Ranking
    {
        public List<VarekodeList> varekoderAlle;
        public IEnumerable<string> varekoderAlleAlias;
        public RankingButikk() { }

        public RankingButikk(FormMain form, DateTime dtFraArg, DateTime dtTilArg, DateTime dtPickArg)
        {
            this.main = form;
            dtFra = dtFraArg;
            dtTil = dtTilArg;
            dtPick = dtPickArg;
            velgerPeriode = FormMain.datoPeriodeVelger;

            this.varekoderAlle = main.appConfig.varekoder.ToList();
            this.varekoderAlleAlias = varekoderAlle.Where(item => item.synlig == true).Select(x => x.alias).Distinct();
        }

        private DataTable MakeTableButikk(string strArg)
        {
            try
            {
                DateTime dtMainFra;
                DateTime dtMainTil;

                if (strArg == "dag")
                {
                    dtMainFra = dtPick;
                    dtMainTil = dtPick;

                    var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select("Dato = '" + dtMainTil.ToString("yyy-MM-dd") + "'");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                }
                else
                {
                    dtMainFra = dtFra;
                    dtMainTil = dtTil;

                    var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select("(Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                }
                if (strArg == "compare")
                {
                    dtMainFra = dtFra.AddYears(-1);
                    dtMainTil = dtTil.AddYears(-1);

                    var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select("(Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                }
                else if (strArg == "lastmonth")
                {
                    dtMainFra = dtFra.AddMonths(-1);
                    dtMainTil = dtTil.AddMonths(-1);

                    var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select("(Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                }
                else if (strArg == "lastweek")
                {
                    DateTime startOfLastWeek = main.database.GetStartOfLastWholeWeek(dtTil); ;
                    dtMainFra = startOfLastWeek;
                    dtMainTil = startOfLastWeek.AddDays(6);

                    var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select("(Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                }
                if (velgerPeriode)
                {
                    dtMainFra = dtFra;
                    dtMainTil = dtTil;

                    sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE Avdeling = " + main.appConfig.Avdeling + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                }

                DataTable dtWork = ReadyTableButikk();

                if (sqlce.Rows.Count == 0)
                    return dtWork;

                for (int d = 1; d <= 9; d++)
                {
                    if ((main.appConfig.rankingCompareLastyear == 1 && strArg == "compare") || (main.appConfig.rankingCompareLastmonth == 1 && strArg == "lastmonth") || strArg == "lastweek")
                        break;
                    if (StopRankingPending())
                        return dtWork;
                    if (d == 7)
                        d = 9;
                    // A V D E L I N G

                    string sKat = "";
                    if (d == 1)
                        sKat = "MDA";
                    if (d == 2)
                        sKat = "AudioVideo";
                    if (d == 3)
                        sKat = "SDA";
                    if (d == 4)
                        sKat = "Telecom";
                    if (d == 5)
                        sKat = "Computing";
                    if (d == 6)
                        sKat = "Kitchen";
                    if (d == 9)
                        sKat = "Other";

                    DataRow dtRow = dtWork.NewRow();

                    decimal sInntjen = 0, sOmset = 0, sOmsetExMva = 0, sTjenInntjen = 0, sTjenOmset = 0, sAntallTjen = 0, sAntallSalg = 0;
                    decimal sStromAntall = 0, sStromInntjen = 0, sModAntall = 0, sModInntjen = 0, sFinansAntall = 0, sFinansInntjen = 0, sModOmset = 0;
                    object r;
                    string[] selgere = main.salesCodes.GetSalesCodes(sKat);

                    if (main.appConfig.importSetting.StartsWith("Full"))
                    {
                        r = sqlce.Compute("Sum(Antall)", "[Varegruppe] >= '" + (d * 100) + "' AND [Varegruppe] < '" + ((d + 1) * 100) + "'");
                        if (!DBNull.Value.Equals(r))
                            sAntallSalg = Convert.ToInt32(r);

                        r = sqlce.Compute("Sum(Salgspris)", "[Varegruppe] >= '" + (d * 100) + "' AND [Varegruppe] < '" + ((d + 1) * 100) + "'");
                        if (!DBNull.Value.Equals(r))
                            sOmset = Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(SalgsprisExMva)", "[Varegruppe] >= '" + (d * 100) + "' AND [Varegruppe] < '" + ((d + 1) * 100) + "'");
                        if (!DBNull.Value.Equals(r))
                            sOmsetExMva = Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Btokr)", "[Varegruppe] >= '" + (d * 100) + "' AND [Varegruppe] < '" + ((d + 1) * 100) + "'");
                        if (!DBNull.Value.Equals(r))
                            sInntjen = Convert.ToDecimal(r);

                        var rows = sqlce.Select("[Varegruppe] = 961");
                        for (int f = 0; f < rows.Length; f++)
                        {
                            var rows2 = sqlce.Select("[Bilagsnr] = " + rows[f]["Bilagsnr"]);
                            DataTable dtFinans = rows2.Any() ? rows2.CopyToDataTable() : sqlce.Clone();
                            dtFinans.DefaultView.Sort = "Salgspris DESC";
                            int gruppe = Convert.ToInt32(dtFinans.Rows[0]["Varegruppe"].ToString().Substring(0, 1));
                            if (gruppe == d)
                            {
                                r = sqlce.Compute("Sum(Antall)", "[Varegruppe] = 961 AND [Bilagsnr] = " + dtFinans.Rows[0]["Bilagsnr"].ToString());
                                if (!DBNull.Value.Equals(r))
                                    sFinansAntall += Convert.ToInt32(r);

                                r = sqlce.Compute("Sum(Btokr)", "[Varegruppe] = 961 AND [Bilagsnr] = " + dtFinans.Rows[0]["Bilagsnr"].ToString());
                                if (!DBNull.Value.Equals(r))
                                    sFinansInntjen += Convert.ToDecimal(r);
                            }
                        }

                        r = sqlce.Compute("Sum(Btokr)", "[Varegruppe] >= '" + (d * 100) + "' AND [Varegruppe] < '" + ((d + 1) * 100) + "' AND ([Varekode] LIKE 'ELSTROM*' OR [Varekode] LIKE 'ELRABATT*')");
                        if (!DBNull.Value.Equals(r))
                            sStromInntjen = Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Antall)", "[Varegruppe] >= '" + (d * 100) + "' AND [Varegruppe] < '" + ((d + 1) * 100) + "' AND [Varekode] LIKE 'ELSTROM*'");
                        if (!DBNull.Value.Equals(r))
                            sStromAntall = Convert.ToInt32(r);

                        r = sqlce.Compute("Sum(Btokr)", "[Varegruppe] = " + d + "83 AND [Varekode] LIKE 'MOD*'");
                        if (!DBNull.Value.Equals(r))
                            sModInntjen = Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Salgspris)", "[Varegruppe] = " + d + "83 AND [Varekode] LIKE 'MOD*'");
                        if (!DBNull.Value.Equals(r))
                            sModOmset = Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Antall)", "[Varegruppe] = " + d + "83 AND [Varekode] LIKE 'MOD*'");
                        if (!DBNull.Value.Equals(r))
                            sModAntall = Convert.ToInt32(r);
                    }

                    foreach (var varekode in varekoderAlle)
                    {
                        r = sqlce.Compute("Sum(Salgspris)", "[Varegruppe] = " + d + "80 AND [Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sTjenOmset += Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Btokr)", "[Varegruppe] = " + d + "80 AND [Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sTjenInntjen += Convert.ToDecimal(r);
                    }

                    foreach (var varekode in varekoderAlle)
                    {
                        if (!varekode.synlig)
                            continue;

                        r = sqlce.Compute("Sum(Antall)", "[Varegruppe] = " + d + "80 AND [Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sAntallTjen += Convert.ToInt32(r);
                    }

                    if (d == 6 && sOmset == 0) // hopp over kjøkken hvis ikke butikken har det
                        continue;

                    dtRow["Kat"] = sKat;
                    dtRow["Salg"] = sAntallSalg;
                    dtRow["Omset"] = sOmset;
                    dtRow["Inntjen"] = sInntjen;
                    dtRow["OmsetExMva"] = sOmsetExMva;
                    if (sOmsetExMva != 0)
                        dtRow["Prosent"] = Math.Round(sInntjen / sOmsetExMva * 100, 2);
                    else
                        dtRow["Prosent"] = 0;
                    dtRow["AntallTjen"] = sAntallTjen;
                    dtRow["TjenOmset"] = sTjenOmset;
                    dtRow["TjenInntjen"] = sTjenInntjen;
                    if (sInntjen != 0)
                        dtRow["TjenMargin"] = Math.Round(sTjenInntjen / sInntjen * 100, 2);
                    else
                        dtRow["TjenMargin"] = 0;
                    dtRow["StromInntjen"] = sStromInntjen;
                    dtRow["StromAntall"] = sStromAntall;
                    if (sInntjen != 0)
                        dtRow["StromMargin"] = Math.Round(sStromInntjen / sInntjen * 100, 2);
                    else
                        dtRow["StromMargin"] = 0;
                    dtRow["ModInntjen"] = sModInntjen;
                    dtRow["ModOmset"] = sModOmset;
                    dtRow["ModAntall"] = sModAntall;
                    if (sOmsetExMva != 0)
                        dtRow["ModMargin"] = Math.Round(sModOmset / sOmsetExMva * 100, 2); // Math.Round(sModInntjen / sOmset * 100, 2);
                    else
                        dtRow["ModMargin"] = 0;
                    dtRow["FinansInntjen"] = sFinansInntjen;
                    dtRow["FinansAntall"] = sFinansAntall;
                    if (sInntjen != 0)
                        dtRow["FinansMargin"] = Math.Round(sFinansInntjen / sInntjen * 100, 2);
                    else
                        dtRow["FinansMargin"] = 0;
                    dtWork.Rows.Add(dtRow);
                }

                // ------------- T O T A L T ---------------
                DataRow dtTotalt = dtWork.NewRow();
                decimal tInntjen = 0, tOmset = 0, tOmsetExMva = 0, tTjenInntjen = 0, tTjenOmset = 0, tAntallTjen = 0, tAntallSalg = 0;
                decimal tStromAntall = 0, tStromInntjen = 0, tModAntall = 0, tModInntjen = 0, tFinansAntall = 0, tFinansInntjen = 0, tModOmset = 0;
                object g;

                if (main.appConfig.importSetting.StartsWith("Full"))
                {
                    g = sqlce.Compute("Sum(Antall)", "");
                    if (!DBNull.Value.Equals(g))
                        tAntallSalg = Convert.ToInt32(g);

                    g = sqlce.Compute("Sum(Salgspris)", "");
                    if (!DBNull.Value.Equals(g))
                        tOmset = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(SalgsprisExMva)", "");
                    if (!DBNull.Value.Equals(g))
                        tOmsetExMva = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Btokr)", "");
                    if (!DBNull.Value.Equals(g))
                        tInntjen = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Btokr)", "[Varekode] LIKE 'ELSTROM*' OR [Varekode] LIKE 'ELRABATT*'");
                    if (!DBNull.Value.Equals(g))
                        tStromInntjen = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Antall)", "[Varekode] LIKE 'ELSTROM*'");
                    if (!DBNull.Value.Equals(g))
                        tStromAntall = Convert.ToInt32(g);

                    g = sqlce.Compute("Sum(Btokr)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                    if (!DBNull.Value.Equals(g))
                        tModInntjen = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Salgspris)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                    if (!DBNull.Value.Equals(g))
                        tModOmset = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Antall)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                    if (!DBNull.Value.Equals(g))
                        tModAntall = Convert.ToInt32(g);

                    g = sqlce.Compute("Sum(Btokr)", "[Varegruppe] = 961");
                    if (!DBNull.Value.Equals(g))
                        tFinansInntjen = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Antall)", "[Varegruppe] = 961");
                    if (!DBNull.Value.Equals(g))
                        tFinansAntall = Convert.ToInt32(g);
                }

                foreach (var varekode in varekoderAlle)
                {
                    g = sqlce.Compute("Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");
                    if (!DBNull.Value.Equals(g))
                        tTjenOmset += Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                    if (!DBNull.Value.Equals(g))
                        tTjenInntjen += Convert.ToDecimal(g);
                }

                foreach (var varekode in varekoderAlle)
                {
                    if (!varekode.synlig)
                        continue;

                    g = sqlce.Compute("Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                    if (!DBNull.Value.Equals(g))
                        tAntallTjen += Convert.ToInt32(g);
                }

                dtTotalt["Kat"] = "TOTALT";
                dtTotalt["Salg"] = tAntallSalg;
                dtTotalt["Omset"] = tOmset;
                dtTotalt["Inntjen"] = tInntjen;
                dtTotalt["OmsetExMva"] = tOmsetExMva;
                if (tOmsetExMva != 0)
                    dtTotalt["Prosent"] = Math.Round(tInntjen / tOmsetExMva * 100, 2);
                else
                    dtTotalt["Prosent"] = 0;
                dtTotalt["AntallTjen"] = tAntallTjen;
                dtTotalt["TjenOmset"] = tTjenOmset;
                dtTotalt["TjenInntjen"] = tTjenInntjen;
                if (tInntjen != 0)
                    dtTotalt["TjenMargin"] = Math.Round(tTjenInntjen / tInntjen * 100, 2);
                else
                    dtTotalt["TjenMargin"] = 0;
                dtTotalt["StromInntjen"] = tStromInntjen;
                dtTotalt["StromAntall"] = tStromAntall;
                if (tInntjen != 0)
                    dtTotalt["StromMargin"] = Math.Round(tStromInntjen / tInntjen * 100, 2);
                else
                    dtTotalt["StromMargin"] = 0;
                dtTotalt["ModInntjen"] = tModInntjen;
                dtTotalt["ModAntall"] = tModAntall;
                dtTotalt["ModOmset"] = tModOmset;
                if (tOmsetExMva != 0)
                    dtTotalt["ModMargin"] = Math.Round(tModOmset / tOmsetExMva * 100, 2); // Math.Round(tModInntjen / tOmset * 100, 2);
                else
                    dtTotalt["ModMargin"] = 0;
                dtTotalt["FinansInntjen"] = tFinansInntjen;
                dtTotalt["FinansAntall"] = tFinansAntall;
                if (tInntjen != 0)
                    dtTotalt["FinansMargin"] = Math.Round(tFinansInntjen / tInntjen * 100, 2);
                else
                    dtTotalt["FinansMargin"] = 0;
                dtWork.Rows.Add(dtTotalt);
                sqlce.Dispose();

                return dtWork;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        private DataTable MakeTableButikkLastWeek()
        {
            try
            {
                DateTime dtMainFra;
                DateTime dtMainTil;

                DateTime startOfLastWeek = main.database.GetStartOfLastWholeWeek(dtTil); ;
                dtMainFra = startOfLastWeek;
                dtMainTil = startOfLastWeek.AddDays(6);

                sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE Avdeling = " + main.appConfig.Avdeling + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                if (sqlce != null)
                    sqlce.Columns.Add("SalgsprisExMva", typeof(Double), "Salgspris / Mva");

                DataTable dtWork = ReadyTableButikk();
                if (sqlce.Rows.Count == 0)
                    return dtWork;

                // ------------- T O T A L T ---------------
                DataRow dtTotalt = dtWork.NewRow();
                decimal tInntjen = 0, tOmset = 0, tOmsetExMva = 0, tTjenInntjen = 0, tTjenOmset = 0, tAntallTjen = 0, tAntallSalg = 0;
                decimal tStromAntall = 0, tStromInntjen = 0, tModAntall = 0, tModInntjen = 0, tFinansAntall = 0, tFinansInntjen = 0, tModOmset = 0;
                decimal tAccessoriesAntall = 0, tAccessoriesInntjen = 0, tAccessoriesOmset = 0, tKuppvarer = 0, tSnittAntall = 0, tSnittInntjen = 0, tSnittOmset = 0, tProdukter = 0;
                object g;

                if (main.appConfig.importSetting.StartsWith("Full"))
                {
                    g = sqlce.Compute("Sum(Antall)", "");
                    if (!DBNull.Value.Equals(g))
                        tAntallSalg = Convert.ToInt32(g);

                    g = sqlce.Compute("Sum(Salgspris)", "");
                    if (!DBNull.Value.Equals(g))
                        tOmset = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(SalgsprisExMva)", "");
                    if (!DBNull.Value.Equals(g))
                        tOmsetExMva = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Btokr)", "");
                    if (!DBNull.Value.Equals(g))
                        tInntjen = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Btokr)", "[Varekode] LIKE 'ELSTROM*' OR [Varekode] LIKE 'ELRABATT*'");
                    if (!DBNull.Value.Equals(g))
                        tStromInntjen = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Antall)", "[Varekode] LIKE 'ELSTROM*'");
                    if (!DBNull.Value.Equals(g))
                        tStromAntall = Convert.ToInt32(g);

                    g = sqlce.Compute("Sum(Btokr)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                    if (!DBNull.Value.Equals(g))
                        tModInntjen = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Salgspris)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                    if (!DBNull.Value.Equals(g))
                        tModOmset = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Antall)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                    if (!DBNull.Value.Equals(g))
                        tModAntall = Convert.ToInt32(g);

                    g = sqlce.Compute("Sum(Btokr)", "[Varegruppe] = 961");
                    if (!DBNull.Value.Equals(g))
                        tFinansInntjen = Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Antall)", "[Varegruppe] = 961");
                    if (!DBNull.Value.Equals(g))
                        tFinansAntall = Convert.ToInt32(g);
                }

                tProdukter = Compute(sqlce, "Sum(Antall)", "[Varegruppe]=531 OR [Varegruppe]=533 OR [Varegruppe]=534 OR [Varegruppe]=224 OR [Varegruppe]=431");
                foreach (var varekode in varekoderAlle)
                {
                    if (!varekode.synlig)
                        continue;

                    tAntallTjen += Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                }

                foreach (var varekode in varekoderAlle)
                {
                    g = sqlce.Compute("Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");
                    if (!DBNull.Value.Equals(g))
                        tTjenOmset += Convert.ToDecimal(g);

                    g = sqlce.Compute("Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                    if (!DBNull.Value.Equals(g))
                        tTjenInntjen += Convert.ToDecimal(g);
                }

                tKuppvarer = Compute(sqlce, "Count(Varekode)", "[Varekode] = 'KV'");

                int[] accessoriesGrpList = main.appConfig.GetAccessorieGroups(0);
                foreach (int ac in accessoriesGrpList)
                {
                    tAccessoriesAntall += Compute(sqlce, "Sum(Antall)", "[Varegruppe] = " + ac);
                    tAccessoriesInntjen += Compute(sqlce, "Sum(Btokr)", "[Varegruppe] = " + ac);
                    tAccessoriesOmset += Compute(sqlce, "Sum(Salgspris)", "[Varegruppe] = " + ac);
                }

                int[] mainproductsGrpList = main.appConfig.GetMainproductGroups(0);
                foreach (int grp in mainproductsGrpList)
                {
                    tSnittAntall += Compute(sqlce, "Sum(Antall)", "[Varegruppe] = " + grp);
                    tSnittInntjen += Compute(sqlce, "Sum(Btokr)", "[Varegruppe] = " + grp);
                    tSnittOmset += Compute(sqlce, "Sum(Salgspris)", "[Varegruppe] = " + grp);
                }

                dtTotalt["Kat"] = "TOTALT";
                dtTotalt["Salg"] = tAntallSalg;
                dtTotalt["Omset"] = tOmset;
                dtTotalt["Inntjen"] = tInntjen;
                dtTotalt["OmsetExMva"] = tOmsetExMva;
                if (tOmsetExMva != 0)
                    dtTotalt["Prosent"] = Math.Round(tInntjen / tOmsetExMva * 100, 2);
                else
                    dtTotalt["Prosent"] = 0;
                if (tAntallSalg != 0)
                    dtTotalt["SnittOmsetAlle"] = tOmset / tAntallSalg;
                else
                    dtTotalt["SnittOmsetAlle"] = 0;
                dtTotalt["AntallTjen"] = tAntallTjen;
                dtTotalt["TjenOmset"] = tTjenOmset;
                dtTotalt["TjenInntjen"] = tTjenInntjen;
                if (tInntjen != 0)
                    dtTotalt["TjenMargin"] = Math.Round(tTjenInntjen / tInntjen * 100, 2);
                else
                    dtTotalt["TjenMargin"] = 0;
                dtTotalt["TjenHitrate"] = CalcHitrate(Convert.ToDecimal(tAntallTjen), tProdukter);
                dtTotalt["StromInntjen"] = tStromInntjen;
                dtTotalt["StromAntall"] = tStromAntall;
                if (tInntjen != 0)
                    dtTotalt["StromMargin"] = Math.Round(tStromInntjen / tInntjen * 100, 2);
                else
                    dtTotalt["StromMargin"] = 0;
                dtTotalt["ModInntjen"] = tModInntjen;
                dtTotalt["ModAntall"] = tModAntall;
                dtTotalt["ModOmset"] = tModOmset;
                if (tOmsetExMva != 0)
                    dtTotalt["ModMargin"] = Math.Round(tModOmset / tOmsetExMva * 100, 2); // Math.Round(tModInntjen / tOmset * 100, 2);
                else
                    dtTotalt["ModMargin"] = 0;
                dtTotalt["FinansInntjen"] = tFinansInntjen;
                dtTotalt["FinansAntall"] = tFinansAntall;
                if (tInntjen != 0)
                    dtTotalt["FinansMargin"] = Math.Round(tFinansInntjen / tInntjen * 100, 2);
                else
                    dtTotalt["FinansMargin"] = 0;

                dtTotalt["AccessoriesAntall"] = tAccessoriesAntall;
                dtTotalt["AccessoriesInntjen"] = tAccessoriesInntjen;
                dtTotalt["AccessoriesOmset"] = tAccessoriesOmset;
                if (tOmsetExMva != 0)
                    dtTotalt["AccessoriesSoB"] = Math.Round(tAccessoriesOmset / tOmsetExMva * 100, 2);
                else
                    dtTotalt["AccessoriesSoB"] = 0;

                dtTotalt["Kuppvarer"] = tKuppvarer;

                dtTotalt["SnittAntall"] = tSnittAntall;
                if (tSnittAntall != 0)
                {
                    dtTotalt["SnittInntjen"] = tSnittInntjen / tSnittAntall;
                    dtTotalt["SnittOmset"] = tSnittOmset / tSnittAntall;
                }
                else
                {
                    dtTotalt["SnittInntjen"] = 0;
                    dtTotalt["SnittOmset"] = 0;
                }

                dtWork.Rows.Add(dtTotalt);
                sqlce.Dispose();

                return dtWork;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableHtml(string strArg)
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                var hashId = random.Next(999, 99999);

                string urlID = "link";
                if (strArg == "måned")
                {
                    urlID += "m";
                    dtMonth = MakeTableButikk("måned");
                    if (dtMonth.Rows.Count > 0)
                        dt = dtMonth;
                    else
                    {
                        doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                        return doc;
                    }
                }
                if (strArg == "dag")
                {
                    urlID += "d";
                    dtDay = MakeTableButikk("dag");
                    if (dtDay.Rows.Count > 0)
                        dt = dtDay;
                    else
                    {
                        doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                        return doc;
                    }
                }
                if (strArg == "compare")
                {
                    urlID = "linkx";
                    outerclass = "OutertableCompare";
                    if (dtCompare.Rows.Count > 0)
                        dt = dtCompare;
                    else
                    {
                        doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                        return doc;
                    }
                }
                else if (strArg == "lastmonth")
                {
                    urlID = "linkx";
                    outerclass = "OutertableCompareLastMonth";
                    if (dtCompareLastMonth.Rows.Count > 0)
                        dt = dtCompareLastMonth;
                    else
                    {
                        doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                        return doc;
                    }
                }
                else if (strArg == "lastweek")
                {
                    urlID = "linkx";
                    outerclass = "OutertableLastWeek";
                    if (dtLastWeek.Rows.Count > 0)
                        dt = dtLastWeek;
                    else
                    {
                        doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                        return doc;
                    }
                }

                main.openXml.SaveDocument(dt, "Butikk", strArg, dtPick, strArg.ToUpper() + " - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.AddRange(MakeTableHeaderButikk());
                doc.Add("<tbody>");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows.Count == i + 1) // Vi er på siste row
                        doc.Add("</tbody><tfoot><tr><td class='text-cat'><a href='#" + urlID + "t" + "'>" + dt.Rows[i]["Kat"] + "</a></td>");
                    else
                        doc.Add("<tr><td class='text-cat'><a href='#" + urlID + "b" + dt.Rows[i]["Kat"].ToString() + "'>" + dt.Rows[i]["Kat"] + "</a></td>");

                    if (main.appConfig.importSetting.StartsWith("Full"))
                    {
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Salg"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Omset"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Inntjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["Prosent"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-finans'>" + PlusMinus(dt.Rows[i]["FinansAntall"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["FinansInntjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["FinansMargin"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-moderna'>" + PlusMinus(dt.Rows[i]["ModAntall"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["ModOmset"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["ModMargin"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-strom'>" + PlusMinus(dt.Rows[i]["StromAntall"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["StromInntjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["StromMargin"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-service'>" + PlusMinus(dt.Rows[i]["AntallTjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["TjenInntjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["TjenMargin"].ToString()) + "</td>");
                    }
                    if (!main.appConfig.importSetting.StartsWith("Full"))
                    {
                        doc.Add("<td class='numbers-service'>" + PlusMinus(dt.Rows[i]["AntallTjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["TjenOmset"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["TjenInntjen"].ToString()) + "</td>");
                    }
                    doc.Add("</tr>");
                }
                doc.Add("</tfoot></table></td></tr></table>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { "Error" };
            }
        }

        public List<string> GetTableHtmlLastWeek()
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                var hashId = random.Next(999, 99999);

                string urlID = "linkx";
                outerclass = "OutertableLastWeek";
                if (dtLastWeek.Rows.Count > 0)
                    dt = dtLastWeek;
                else
                {
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                    return doc;
                }

                main.openXml.SaveDocument(dt, "Butikk", "Sist uke", dtPick, "SIST UKE - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.AddRange(MakeTableHeaderButikkLastWeek());
                doc.Add("<tbody>");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows.Count == i + 1) // Vi er på siste row
                        doc.Add("</tbody><tfoot><tr><td class='text-cat'><a href='#" + urlID + "t" + "'>" + dt.Rows[i]["Kat"] + "</a></td>");
                    else
                        doc.Add("<tr><td class='text-cat'><a href='#" + urlID + "b" + dt.Rows[i]["Kat"].ToString() + "'>" + dt.Rows[i]["Kat"] + "</a></td>");

                    if (main.appConfig.importSetting.StartsWith("Full"))
                    {
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Salg"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Omset"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Inntjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["Prosent"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["SnittOmsetAlle"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-finans'>" + PlusMinus(dt.Rows[i]["FinansAntall"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["FinansMargin"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-moderna'>" + PlusMinus(dt.Rows[i]["ModAntall"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["ModMargin"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-strom'>" + PlusMinus(dt.Rows[i]["StromAntall"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["StromMargin"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-service'>" + PlusMinus(dt.Rows[i]["AntallTjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["TjenMargin"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["TjenHitrate"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-finans'>" + PlusMinus(dt.Rows[i]["AccessoriesAntall"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["AccessoriesSoB"].ToString()) + "</td>");
                    }
                    if (!main.appConfig.importSetting.StartsWith("Full"))
                    {
                        doc.Add("<td class='numbers-service'>" + PlusMinus(dt.Rows[i]["AntallTjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["TjenOmset"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["TjenInntjen"].ToString()) + "</td>");
                    }
                    doc.Add("</tr>");
                }
                doc.Add("</tfoot></table></td></tr></table>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { "Error" };
            }
        }

        public List<string> GetTableCompareHtml()
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                dtCompare = MakeTableButikk("compare");
                if (dtCompare.Rows.Count > 0)
                    doc.AddRange(GetTableHtml("compare"));
                else
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { "Error" };
            }
        }

        public List<string> GetTableCompareLastMonthHtml()
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                dtCompareLastMonth = MakeTableButikk("lastmonth");
                if (dtCompareLastMonth.Rows.Count > 0)
                    doc.AddRange(GetTableHtml("lastmonth"));
                else
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { "Error" };
            }
        }

        public List<string> GetTableLastWholeWeek()
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                dtLastWeek = MakeTableButikkLastWeek();
                if (dtLastWeek.Rows.Count > 0)
                    doc.AddRange(GetTableHtmlLastWeek());
                else
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { "Error" };
            }
        }

        private DataTable MakeTableButikkAvd()
        {
            try
            {
                favoritter = FormMain.Favoritter.ToArray();
                DataTable dtWork = ReadyTableButikk();

                int noResults = 0;
                for (int d = 0; d < favoritter.Length; d++)
                {
                    if (StopRankingPending())
                        return dtWork;

                    var rows = main.database.CallMonthTable(dtTil, favoritter[d]).Select("(Dato >= '" + dtFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();

                    if (sqlce.Rows.Count == 0)
                        noResults++;

                    decimal sInntjen = 0, sOmset = 0, sOmsetExMva = 0, sTjenInntjen = 0, sTjenOmset = 0, sAntallTjen = 0, sAntallSalg = 0;
                    decimal sStromAntall = 0, sStromInntjen = 0, sModAntall = 0, sModInntjen = 0, sFinansAntall = 0, sFinansInntjen = 0, sModOmset = 0;
                    object r;

                    if (main.appConfig.importSetting.StartsWith("Full"))
                    {
                        r = sqlce.Compute("Sum(Antall)", "");
                        if (!DBNull.Value.Equals(r))
                            sAntallSalg = Convert.ToInt32(r);

                        r = sqlce.Compute("Sum(Salgspris)", "");
                        if (!DBNull.Value.Equals(r))
                            sOmset = Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(SalgsprisExMva)", "");
                        if (!DBNull.Value.Equals(r))
                            sOmsetExMva = Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Btokr)", "");
                        if (!DBNull.Value.Equals(r))
                            sInntjen = Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Btokr)", "[Varekode] LIKE 'ELSTROM*' OR [Varekode] LIKE 'ELRABATT*'");
                        if (!DBNull.Value.Equals(r))
                            sStromInntjen = Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Antall)", "[Varekode] LIKE 'ELSTROM*'");
                        if (!DBNull.Value.Equals(r))
                            sStromAntall = Convert.ToInt32(r);

                        r = sqlce.Compute("Sum(Btokr)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                        if (!DBNull.Value.Equals(r))
                            sModInntjen = Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Salgspris)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                        if (!DBNull.Value.Equals(r))
                            sModOmset = Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Antall)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                        if (!DBNull.Value.Equals(r))
                            sModAntall = Convert.ToInt32(r);

                        r = sqlce.Compute("Sum(Btokr)", "[Varegruppe] = 961");
                        if (!DBNull.Value.Equals(r))
                            sFinansInntjen = Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Antall)", "[Varegruppe] = 961");
                        if (!DBNull.Value.Equals(r))
                            sFinansAntall = Convert.ToInt32(r);
                    }

                    foreach (var varekode in varekoderAlle)
                    {
                        r = sqlce.Compute("Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sTjenOmset += Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sTjenInntjen += Convert.ToDecimal(r);
                    }

                    foreach (var varekode in varekoderAlle)
                    {
                        if (!varekode.synlig)
                            continue;

                        r = sqlce.Compute("Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sAntallTjen += Convert.ToInt32(r);
                    }

                    DataRow dtRow = dtWork.NewRow();
                    dtRow["Kat"] = favoritter[d];
                    dtRow["Salg"] = sAntallSalg;
                    dtRow["Omset"] = sOmset;
                    dtRow["Inntjen"] = sInntjen;
                    dtRow["OmsetExMva"] = sOmsetExMva;
                    if (sOmset != 0)
                        dtRow["Prosent"] = Math.Round(sInntjen / sOmsetExMva * 100, 2);
                    else
                        dtRow["Prosent"] = 0;
                    dtRow["AntallTjen"] = sAntallTjen;
                    dtRow["TjenOmset"] = sTjenOmset;
                    dtRow["TjenInntjen"] = sTjenInntjen;
                    if (sInntjen != 0)
                        dtRow["TjenMargin"] = Math.Round(sTjenInntjen / sInntjen * 100, 2);
                    else
                        dtRow["TjenMargin"] = 0;
                    dtRow["StromInntjen"] = sStromInntjen;
                    dtRow["StromAntall"] = sStromAntall;
                    if (sInntjen != 0)
                        dtRow["StromMargin"] = Math.Round(sStromInntjen / sInntjen * 100, 2);
                    else
                        dtRow["StromMargin"] = 0;
                    dtRow["ModInntjen"] = sModInntjen;
                    dtRow["ModOmset"] = sModOmset;
                    dtRow["ModAntall"] = sModAntall;
                    if (sOmsetExMva != 0)
                        dtRow["ModMargin"] = Math.Round(sModOmset / sOmsetExMva * 100, 2); // Math.Round(sModInntjen / sOmset * 100, 2);
                    else
                        dtRow["ModMargin"] = 0;
                    dtRow["FinansInntjen"] = sFinansInntjen;
                    dtRow["FinansAntall"] = sFinansAntall;
                    if (sInntjen != 0)
                        dtRow["FinansMargin"] = Math.Round(sFinansInntjen / sInntjen * 100, 2);
                    else
                        dtRow["FinansMargin"] = 0;
                    dtWork.Rows.Add(dtRow);
                    sqlce.Dispose();
                }
                if (noResults == favoritter.Length)
                    dtWork.Clear();

                return dtWork;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetAvdHtml()
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                dtAvd = MakeTableButikkAvd();
                if (dtAvd.Rows.Count > 0)
                    dt = dtAvd;
                else
                {
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                    return doc;
                }
                var hashId = random.Next(999, 99999);

                main.openXml.SaveDocument(dt, "Butikk", "Favoritter", dtPick, "FAVORITTER - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='OutertableNormal toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.AddRange(MakeTableHeaderFav());

                for (int i = 0; i < dtAvd.Rows.Count; i++)
                {
                    doc.Add("<tr><td class='text-cat'>" + avdeling.Get(Convert.ToInt32(dtAvd.Rows[i]["Kat"])).Replace(" ", "&nbsp;") + "</td>");
                    if (main.appConfig.importSetting.StartsWith("Full"))
                    {
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dtAvd.Rows[i]["Salg"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dtAvd.Rows[i]["Omset"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dtAvd.Rows[i]["Inntjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dtAvd.Rows[i]["Prosent"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-finans'>" + PlusMinus(dtAvd.Rows[i]["FinansAntall"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dtAvd.Rows[i]["FinansInntjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dtAvd.Rows[i]["FinansMargin"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-moderna'>" + PlusMinus(dtAvd.Rows[i]["ModAntall"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dtAvd.Rows[i]["ModOmset"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dtAvd.Rows[i]["ModMargin"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-strom'>" + PlusMinus(dtAvd.Rows[i]["StromAntall"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dtAvd.Rows[i]["StromInntjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dtAvd.Rows[i]["StromMargin"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-service'>" + PlusMinus(dtAvd.Rows[i]["AntallTjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dtAvd.Rows[i]["TjenInntjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + PercentShare(dtAvd.Rows[i]["TjenMargin"].ToString()) + "</td>");
                    }
                    if (!main.appConfig.importSetting.StartsWith("Full"))
                    {
                        doc.Add("<td class='numbers-service'>" + PlusMinus(dtAvd.Rows[i]["AntallTjen"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dtAvd.Rows[i]["TjenOmset"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dtAvd.Rows[i]["TjenInntjen"].ToString()) + "</td>");
                    }
                    doc.Add("</tr>");
                }
                doc.Add("</tbody></table></td></tr></table>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { "Error" };
            }
        }

        private List<string> MakeTableHeaderFav()
        {
            List<string> doc = new List<string> { };

            doc.Add("<thead><tr>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Avdeling</td>");
            if (main.appConfig.importSetting.StartsWith("Full"))
            {
                doc.Add("<th class=\"{sorter: 'digit'}\" width=50 >Salg</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Omsetn.</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Inntjen.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=60 ><abbr title='Btokr. inntjen / Omset. ex. mva. alle varer'>Margin</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#f5954e;'>Finans</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#f5954e;'>Inntjen.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#f5954e;'><abbr title='Btokr. inntjen. Finans / Btokr inntjen. alle varer'>SoM</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#6699ff;'>TA</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#6699ff;'>Omset.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#6699ff;'><abbr title='Omset. TA / Omset. ex. mva. alle varer'>SoB</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#FAF39E;'>Strøm</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#FAF39E;'>Inntjen.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#FAF39E;'><abbr title='Btokr. inntjen. Strøm / Btokr. inntjen. alle varer'>SoM</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>Tjen.</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 style='background:#80c34a;'>Inntjen.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#80c34a;'><abbr title='Btokr. inntjen. Tjenester / Btokr. inntjen. alle varer'>SoM</abbr></td>");
            }
            if (!main.appConfig.importSetting.StartsWith("Full"))
            {
                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>Tjen.</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 style='background:#80c34a;'>Tjen.Omsetn.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=80 style='background:#80c34a;'>Tjen.Inntjen.</td>");
            }
            doc.Add("</tr></thead><tbody>");

            return doc;
        }

        public List<string> MakeTableHeaderButikk()
        {
            List<string> doc = new List<string> { };

            doc.Add("<thead><tr>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Kategori</td>");
            if (main.appConfig.importSetting.StartsWith("Full"))
            {
                doc.Add("<th class=\"{sorter: 'digit'}\" width=50 >Salg</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Omsetn.</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Inntjen.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=60 ><abbr title='Btokr. inntjen / Omset. ex. mva. alle varer'>Margin</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#f5954e;'>Finans</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#f5954e;'>Inntjen.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#f5954e;'><abbr title='Btokr. inntjen. Finans / Btokr inntjen. alle varer'>SoM</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#6699ff;'>TA</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#6699ff;'>Omset.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#6699ff;'><abbr title='Omset. TA / Omset. ex. mva. alle varer'>SoB</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#FAF39E;'>Strøm</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#FAF39E;'>Inntjen.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#FAF39E;'><abbr title='Btokr. inntjen. Strøm / Btokr. inntjen. alle varer'>SoM</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>RTG/SA</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 style='background:#80c34a;'>Inntjen.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#80c34a;'><abbr title='Btokr. inntjen. Tjenester / Btokr. inntjen. alle varer'>SoM</abbr></td>");
            }
            if (!main.appConfig.importSetting.StartsWith("Full"))
            {
                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>RTG/SA</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 style='background:#80c34a;'>Omsetn.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=80 style='background:#80c34a;'>Inntjen.</td>");
            }
            doc.Add("</tr></thead>");
            return doc;
        }

        public List<string> MakeTableHeaderButikkLastWeek()
        {
            List<string> doc = new List<string> { };

            doc.Add("<thead><tr>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Kategori</td>");
            if (main.appConfig.importSetting.StartsWith("Full"))
            {
                doc.Add("<th class=\"{sorter: 'digit'}\" width=50 >Salg</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Omsetn.</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Inntjen.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=60 ><abbr title='Btokr. inntjen / Omset. ex. mva. alle varer'>Margin</abbr></td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Snittsalg.</td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#f5954e;'>Finans</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#f5954e;'><abbr title='Btokr. inntjen. Finans / Btokr inntjen. alle varer'>SoM</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#6699ff;'>TA</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#6699ff;'><abbr title='Omset. TA / Omset. ex. mva. alle varer'>SoB</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#FAF39E;'>Strøm</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#FAF39E;'><abbr title='Btokr. inntjen. Strøm / Btokr. inntjen. alle varer'>SoM</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>RTG/SA</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#80c34a;'><abbr title='Btokr. inntjen. Tjenester / Btokr. inntjen. alle varer'>SoM</abbr></td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#80c34a;'><abbr title='Antall produkter med salgbare tjenester solgt / Totalt antall solgte tjenester'>Hitrate</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#f36565;'>Tilbehør</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#f36565;'><abbr title='Omset. Tilbehør / Omset. ex. mva. alle varer'>SoB</abbr></td>");
            }
            if (!main.appConfig.importSetting.StartsWith("Full"))
            {
                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>RTG/SA</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 style='background:#80c34a;'>Omsetn.</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=80 style='background:#80c34a;'>Inntjen.</td>");
            }
            doc.Add("</tr></thead>");
            return doc;
        }

    }
}