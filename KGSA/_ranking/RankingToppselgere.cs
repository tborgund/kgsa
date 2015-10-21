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
using System.Diagnostics;

namespace KGSA
{
    public class RankingToppselgere : Ranking
    {
        public List<VarekodeList> varekoderAlle;
        public IEnumerable<string> varekoderAlleAlias;
        private DataTable tableListMtd;
        private DataTable tableListMtdAvd;
        private DataTable tableListLastYearMtd;
        private DataTable tableListLastYear;
        private DataTable tableListPreviousOpenDay;
        private DataTable tableListPreviousOpenDayFav;
        public RankingToppselgere() { }

        public RankingToppselgere(FormMain form, DateTime dtFraArg, DateTime dtTilArg)
        {
            this.main = form;
            dtFra = dtFraArg;
            dtTil = dtTilArg;
            velgerPeriode = FormMain.datoPeriodeVelger;

            if (FormMain.selgerkodeList != null)
                if (FormMain.selgerkodeList.Count == 0)
                    FormMain.selgerkodeList.AddRange(main.salesCodes.GetAlleSelgerkoder(main.appConfig.Avdeling));

            this.varekoderAlle = main.appConfig.varekoder.ToList();
            this.varekoderAlleAlias = varekoderAlle.Where(item => item.synlig == true).Select(x => x.alias).Distinct();
        }

        private DataTable MakeTableList(DateTime from, DateTime to, bool total = false)
        {
            try
            {
                DataTable dtWork = ReadyTableList();

                DataTable sqlce = main.database.GetSqlDataTable("SELECT Selgerkode, Varekode, Varegruppe, Dato, Antall, Btokr, Salgspris, Mva FROM tblSalg WHERE Avdeling = " + main.appConfig.Avdeling + " AND (Dato >= '" + from.ToString("yyy-MM-dd") + "' AND Dato <= '" + to.ToString("yyy-MM-dd") + "')");
                sqlce.Columns.Add("SalgsprisExMva", typeof(Double), "Salgspris / Mva");

                if (sqlce.Rows.Count == 0)
                    return dtWork;

                var selgere = FormMain.selgerkodeList.Distinct().ToList();

                for (int i = 0; i < selgere.Count; i++)
                {
                    // S E L G E R E
                    if (StopRankingPending())
                        return dtWork;

                    var rows = sqlce.Select("Selgerkode = '" + selgere[i] + "'");
                    DataTable dt = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();

                    object r;
                    DataRow dtRow = dtWork.NewRow();
                    decimal sInntjenSel = 0, sOmsetSel = 0, sTjenInntjenSel = 0, sTjenOmsetSel = 0, sAntallTjenSel = 0, sAntallSalgSel = 0;
                    decimal sStromAntallSel = 0, sStromInntjenSel = 0, sModAntallSel = 0, sModInntjenSel = 0, sFinansAntallSel = 0, sFinansInntjenSel = 0, sModOmsetSel = 0;
                    decimal sAccessoriesAntallSel = 0, sAccessoriesOmsetSel = 0, sAccessoriesInntjenSel = 0;
                    decimal sOmsetExMvaSel = 0;

                    r = dt.Compute("Sum(Antall)", null);
                    if (!DBNull.Value.Equals(r))
                        sAntallSalgSel += Convert.ToInt32(r);

                    r = dt.Compute("Sum(Salgspris)", null);
                    if (!DBNull.Value.Equals(r))
                        sOmsetSel = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(Btokr)", null);
                    if (!DBNull.Value.Equals(r))
                        sInntjenSel = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(SalgsprisExMva)", null);
                    if (!DBNull.Value.Equals(r))
                        sOmsetExMvaSel = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(Antall)", "[Varegruppe] = 961");
                    if (!DBNull.Value.Equals(r))
                        sFinansAntallSel = Convert.ToInt32(r);

                    r = dt.Compute("Sum(Btokr)", "[Varegruppe] = 961");
                    if (!DBNull.Value.Equals(r))
                        sFinansInntjenSel = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(Btokr)", "([Varekode] LIKE 'ELSTROM*' OR [Varekode] LIKE 'ELRABATT*')");
                    if (!DBNull.Value.Equals(r))
                        sStromInntjenSel = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(Antall)", "[Varekode] LIKE 'ELSTROM*'");
                    if (!DBNull.Value.Equals(r))
                        sStromAntallSel = Convert.ToInt32(r);

                    r = dt.Compute("Sum(Btokr)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                    if (!DBNull.Value.Equals(r))
                        sModInntjenSel = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(Salgspris)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                    if (!DBNull.Value.Equals(r))
                        sModOmsetSel = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(Antall)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                    if (!DBNull.Value.Equals(r))
                        sModAntallSel = Convert.ToInt32(r);

                    foreach (var varekode in varekoderAlle)
                    {
                        r = dt.Compute("Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sTjenOmsetSel += Convert.ToDecimal(r);

                        r = dt.Compute("Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sTjenInntjenSel += Convert.ToDecimal(r);

                        if (varekode.synlig)
                        {
                            r = dt.Compute("Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                            if (!DBNull.Value.Equals(r))
                                sAntallTjenSel += Convert.ToInt32(r);
                        }
                    }

                    if (main.appConfig.listerVisAccessories)
                    {
                        int[] accessoriesGrpList = main.appConfig.GetAccessorieGroups(0);
                        foreach (int ac in accessoriesGrpList)
                        {
                            r = dt.Compute("Sum(Antall)", "[Varegruppe] = " + ac);
                            if (!DBNull.Value.Equals(r))
                                sAccessoriesAntallSel += Convert.ToInt32(r);

                            r = dt.Compute("Sum(Btokr)", "[Varegruppe] = " + ac);
                            if (!DBNull.Value.Equals(r))
                                sAccessoriesInntjenSel += Convert.ToDecimal(r);

                            r = dt.Compute("Sum(Salgspris)", "[Varegruppe] = " + ac);
                            if (!DBNull.Value.Equals(r))
                                sAccessoriesOmsetSel += Convert.ToDecimal(r);
                        }
                    }

                    if (sInntjenSel != 0)
                    {
                        dtRow["Selgerkode"] = selgere[i].Trim();
                        dtRow["AntallSalg"] = sAntallSalgSel;
                        dtRow["Omset"] = sOmsetSel;
                        dtRow["Bto_Margin"] = Math.Round(sOmsetExMvaSel, 2);
                        dtRow["Inntjen"] = sInntjenSel;
                        dtRow["Rtgsa_Antall"] = sAntallTjenSel;
                        dtRow["Rtgsa_Omset"] = sTjenOmsetSel;
                        dtRow["Rtgsa_Inntjen"] = sTjenInntjenSel;
                        dtRow["Strom_Antall"] = sStromAntallSel;
                        dtRow["Strom_Inntjen"] = sStromInntjenSel;
                        dtRow["TA_Antall"] = sModAntallSel;
                        dtRow["TA_Inntjen"] = sModInntjenSel;
                        dtRow["TA_Omset"] = sModOmsetSel;
                        dtRow["Finans_Antall"] = sFinansAntallSel;
                        dtRow["Finans_Inntjen"] = sFinansInntjenSel;
                        dtRow["Accessories_Antall"] = sAccessoriesAntallSel;
                        dtRow["Accessories_Inntjen"] = sAccessoriesInntjenSel;
                        dtRow["Accessories_Omset"] = sAccessoriesOmsetSel;


                        if (sOmsetExMvaSel != 0)
                            dtRow["Bto_Margin"] = Math.Round(sInntjenSel / sOmsetExMvaSel * 100, 2);
                        else
                            dtRow["Bto_Margin"] = 0;

                        if (sInntjenSel != 0)
                            dtRow["Rtgsa_Margin"] = Math.Round(sTjenInntjenSel / sInntjenSel * 100, 2);
                        else
                            dtRow["Rtgsa_Margin"] = 0;

                        if (sInntjenSel != 0)
                            dtRow["Strom_Margin"] = Math.Round(sStromInntjenSel / sInntjenSel * 100, 2);
                        else
                            dtRow["Strom_Margin"] = 0;

                        if (sOmsetExMvaSel != 0)
                            dtRow["TA_Margin"] = Math.Round(sModOmsetSel / sOmsetExMvaSel * 100, 2);
                        else
                            dtRow["TA_Margin"] = 0;

                        if (sInntjenSel != 0)
                            dtRow["Finans_Margin"] = Math.Round(sFinansInntjenSel / sInntjenSel * 100, 2);
                        else
                            dtRow["Finans_Margin"] = 0;

                        if (sInntjenSel != 0)
                            dtRow["Accessories_Margin"] = Math.Round(sAccessoriesInntjenSel / sInntjenSel * 100, 2);
                        else
                            dtRow["Accessories_Margin"] = 0;

                        dtWork.Rows.Add(dtRow);
                    }
                }

                if (total)
                {
                    // ------------- T O T A L T ---------------
                    DataRow dtTotalt = dtWork.NewRow();
                    decimal tInntjen = 0, tOmset = 0, tTjenInntjen = 0, tTjenOmset = 0, tAntallTjen = 0, tAntallSalg = 0;
                    decimal tStromAntall = 0, tStromInntjen = 0, tModAntall = 0, tModInntjen = 0, tFinansAntall = 0, tFinansInntjen = 0, tModOmset = 0;
                    decimal tAccessoriesAntall = 0, tAccessoriesInntjen = 0, tAccessoriesOmset = 0;
                    decimal tOmsetExMva = 0;
                    object g;

                    if (main.appConfig.importSetting.StartsWith("Full"))
                    {
                        g = sqlce.Compute("Sum(Antall)", "");
                        if (!DBNull.Value.Equals(g))
                            tAntallSalg += Convert.ToInt32(g);

                        g = sqlce.Compute("Sum(Salgspris)", "");
                        if (!DBNull.Value.Equals(g))
                            tOmset = Convert.ToDecimal(g);

                        g = sqlce.Compute("Sum(Btokr)", "");
                        if (!DBNull.Value.Equals(g))
                            tInntjen = Convert.ToDecimal(g);

                        g = sqlce.Compute("Sum(SalgsprisExMva)", "");
                        if (!DBNull.Value.Equals(g))
                            tOmsetExMva = Convert.ToDecimal(g);

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

                        int[] accessoriesGrpList = main.appConfig.GetAccessorieGroups(0);
                        foreach (int ac in accessoriesGrpList)
                        {
                            g = sqlce.Compute("Sum(Antall)", "[Varegruppe] = " + ac);
                            if (!DBNull.Value.Equals(g))
                                tAccessoriesAntall += Convert.ToInt32(g);

                            g = sqlce.Compute("Sum(Btokr)", "[Varegruppe] = " + ac);
                            if (!DBNull.Value.Equals(g))
                                tAccessoriesInntjen += Convert.ToDecimal(g);

                            g = sqlce.Compute("Sum(Salgspris)", "[Varegruppe] = " + ac);
                            if (!DBNull.Value.Equals(g))
                                tAccessoriesOmset += Convert.ToDecimal(g);
                        }
                    }

                    foreach (var varekode in varekoderAlle)
                    {
                        g = sqlce.Compute("Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(g))
                            tTjenOmset += Convert.ToDecimal(g);

                        g = sqlce.Compute("Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(g))
                            tTjenInntjen += Convert.ToDecimal(g);

                        if (varekode.synlig)
                        {
                            g = sqlce.Compute("Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                            if (!DBNull.Value.Equals(g))
                                tAntallTjen += Convert.ToInt32(g);
                        }
                    }

                    dtTotalt["Selgerkode"] = "TOTALT";
                    dtTotalt["AntallSalg"] = tAntallSalg;
                    dtTotalt["Omset"] = tOmset;
                    dtTotalt["Inntjen"] = tInntjen;
                    dtTotalt["Rtgsa_Antall"] = tAntallTjen;
                    dtTotalt["Rtgsa_Inntjen"] = tTjenInntjen;
                    dtTotalt["Rtgsa_Omset"] = tTjenOmset;
                    dtTotalt["Strom_Antall"] = tStromAntall;
                    dtTotalt["Strom_Inntjen"] = tStromInntjen;
                    dtTotalt["TA_Antall"] = tModAntall;
                    dtTotalt["TA_Inntjen"] = tModInntjen;
                    dtTotalt["TA_Omset"] = tModOmset;
                    dtTotalt["Finans_Antall"] = tFinansAntall;
                    dtTotalt["Finans_Inntjen"] = tFinansInntjen;
                    dtTotalt["Accessories_Antall"] = tAccessoriesAntall;
                    dtTotalt["Accessories_Inntjen"] = tAccessoriesInntjen;
                    dtTotalt["Accessories_Omset"] = tAccessoriesOmset;

                    if (tOmsetExMva != 0)
                        dtTotalt["Bto_Margin"] = Math.Round(tInntjen / tOmsetExMva * 100, 2);
                    else
                        dtTotalt["Bto_Margin"] = 0;

                    if (tInntjen != 0)
                        dtTotalt["Rtgsa_Margin"] = Math.Round(tTjenInntjen / tInntjen * 100, 2);
                    else
                        dtTotalt["Rtgsa_Margin"] = 0;

                    if (tInntjen != 0)
                        dtTotalt["Strom_Margin"] = Math.Round(tStromInntjen / tInntjen * 100, 2);
                    else
                        dtTotalt["Strom_Margin"] = 0;

                    if (tOmsetExMva != 0)
                        dtTotalt["TA_Margin"] = Math.Round(tModOmset / tOmsetExMva * 100, 2);
                    else
                        dtTotalt["TA_Margin"] = 0;

                    if (tInntjen != 0)
                        dtTotalt["Finans_Margin"] = Math.Round(tFinansInntjen / tInntjen * 100, 2);
                    else
                        dtTotalt["Finans_Margin"] = 0;

                    if (tInntjen != 0)
                        dtTotalt["Accessories_Margin"] = Math.Round(tAccessoriesInntjen / tInntjen * 100, 2);
                    else
                        dtTotalt["Accessories_Margin"] = 0;

                    dtWork.Rows.Add(dtTotalt);
                }

                sqlce.Dispose();
                return dtWork;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        private DataTable MakeTableFav(DateTime from, DateTime to, bool total = false)
        {
            try
            {
                DataTable dtWork = ReadyTableList();

                string sql = "SELECT Avdeling, Varekode, Varegruppe, Antall, Btokr, Salgspris, Mva FROM tblSalg WHERE (Avdeling = 0 ";
                foreach (string a in FormMain.Favoritter)
                    sql += " OR Avdeling = " + a + " ";
                sql += ") AND (Dato >= '" + from.ToString("yyy-MM-dd") + "' AND Dato <= '" + to.ToString("yyy-MM-dd") + "')";

                DataTable sqlce = main.database.GetSqlDataTable(sql);
                sqlce.Columns.Add("SalgsprisExMva", typeof(Double), "Salgspris / Mva");

                if (sqlce.Rows.Count == 0)
                    return dtWork;

                for (int i = 0; i < FormMain.Favoritter.Count; i++)
                {
                    // AVDELINGER
                    if (StopRankingPending())
                        return dtWork;

                    var rows = sqlce.Select("Avdeling = " + FormMain.Favoritter[i]);
                    DataTable dt = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();

                    object r;
                    DataRow dtRow = dtWork.NewRow();
                    decimal sInntjen = 0, sOmset = 0, sTjenInntjen = 0, sTjenOmset = 0, sAntallTjen = 0, sAntallSalg = 0;
                    decimal sStromAntall = 0, sStromInntjen = 0, sModAntall = 0, sModInntjen = 0, sFinansAntall = 0, sFinansInntjen = 0, sModOmset = 0;
                    decimal sAccessoriesAntall = 0, sAccessoriesOmset = 0, sAccessoriesInntjen = 0;
                    decimal sOmsetExMva = 0;

                    r = dt.Compute("Sum(Antall)", null);
                    if (!DBNull.Value.Equals(r))
                        sAntallSalg += Convert.ToInt32(r);

                    r = dt.Compute("Sum(Salgspris)", null);
                    if (!DBNull.Value.Equals(r))
                        sOmset = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(Btokr)", null);
                    if (!DBNull.Value.Equals(r))
                        sInntjen = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(SalgsprisExMva)", null);
                    if (!DBNull.Value.Equals(r))
                        sOmsetExMva = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(Antall)", "[Varegruppe] = 961");
                    if (!DBNull.Value.Equals(r))
                        sFinansAntall = Convert.ToInt32(r);

                    r = dt.Compute("Sum(Btokr)", "[Varegruppe] = 961");
                    if (!DBNull.Value.Equals(r))
                        sFinansInntjen = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(Btokr)", "([Varekode] LIKE 'ELSTROM*' OR [Varekode] LIKE 'ELRABATT*')");
                    if (!DBNull.Value.Equals(r))
                        sStromInntjen = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(Antall)", "[Varekode] LIKE 'ELSTROM*'");
                    if (!DBNull.Value.Equals(r))
                        sStromAntall = Convert.ToInt32(r);

                    r = dt.Compute("Sum(Btokr)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                    if (!DBNull.Value.Equals(r))
                        sModInntjen = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(Salgspris)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                    if (!DBNull.Value.Equals(r))
                        sModOmset = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(Antall)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                    if (!DBNull.Value.Equals(r))
                        sModAntall = Convert.ToInt32(r);

                    foreach (var varekode in varekoderAlle)
                    {
                        r = dt.Compute("Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sTjenOmset += Convert.ToDecimal(r);

                        r = dt.Compute("Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sTjenInntjen += Convert.ToDecimal(r);

                        if (varekode.synlig)
                        {
                            r = dt.Compute("Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                            if (!DBNull.Value.Equals(r))
                                sAntallTjen += Convert.ToInt32(r);
                        }
                    }

                    if (main.appConfig.listerVisAccessories)
                    {
                        int[] accessoriesGrpList = main.appConfig.GetAccessorieGroups(0);
                        foreach (int ac in accessoriesGrpList)
                        {
                            r = dt.Compute("Sum(Antall)", "[Varegruppe] = " + ac);
                            if (!DBNull.Value.Equals(r))
                                sAccessoriesAntall += Convert.ToInt32(r);

                            r = dt.Compute("Sum(Btokr)", "[Varegruppe] = " + ac);
                            if (!DBNull.Value.Equals(r))
                                sAccessoriesInntjen += Convert.ToDecimal(r);

                            r = dt.Compute("Sum(Salgspris)", "[Varegruppe] = " + ac);
                            if (!DBNull.Value.Equals(r))
                                sAccessoriesOmset += Convert.ToDecimal(r);
                        }
                    }

                    if (sInntjen != 0)
                    {
                        dtRow["Selgerkode"] = FormMain.Favoritter[i];
                        dtRow["AntallSalg"] = sAntallSalg;
                        dtRow["Omset"] = sOmset;
                        dtRow["Inntjen"] = sInntjen;
                        dtRow["Rtgsa_Antall"] = sAntallTjen;
                        dtRow["Rtgsa_Omset"] = sTjenOmset;
                        dtRow["Rtgsa_Inntjen"] = sTjenInntjen;
                        dtRow["Strom_Antall"] = sStromAntall;
                        dtRow["Strom_Inntjen"] = sStromInntjen;
                        dtRow["TA_Antall"] = sModAntall;
                        dtRow["TA_Inntjen"] = sModInntjen;
                        dtRow["TA_Omset"] = sModOmset;
                        dtRow["Finans_Antall"] = sFinansAntall;
                        dtRow["Finans_Inntjen"] = sFinansInntjen;
                        dtRow["Accessories_Antall"] = sAccessoriesAntall;
                        dtRow["Accessories_Inntjen"] = sAccessoriesInntjen;
                        dtRow["Accessories_Omset"] = sAccessoriesOmset;

                        if (sOmsetExMva != 0)
                            dtRow["Bto_Margin"] = Math.Round(sInntjen / sOmsetExMva * 100, 2);
                        else
                            dtRow["Bto_Margin"] = 0;

                        if (sInntjen != 0)
                            dtRow["Rtgsa_Margin"] = Math.Round(sTjenInntjen / sInntjen * 100, 2);
                        else
                            dtRow["Rtgsa_Margin"] = 0;

                        if (sInntjen != 0)
                            dtRow["Strom_Margin"] = Math.Round(sStromInntjen / sInntjen * 100, 2);
                        else
                            dtRow["Strom_Margin"] = 0;

                        if (sOmsetExMva != 0)
                            dtRow["TA_Margin"] = Math.Round(sModOmset / sOmsetExMva * 100, 2);
                        else
                            dtRow["TA_Margin"] = 0;

                        if (sInntjen != 0)
                            dtRow["Finans_Margin"] = Math.Round(sFinansInntjen / sInntjen * 100, 2);
                        else
                            dtRow["Finans_Margin"] = 0;

                        if (sInntjen != 0)
                            dtRow["Accessories_Margin"] = Math.Round(sAccessoriesInntjen / sInntjen * 100, 2);
                        else
                            dtRow["Accessories_Margin"] = 0;

                        dtWork.Rows.Add(dtRow);
                    }
                }


                sqlce.Dispose();
                return dtWork;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetToppListAll(DateTime date)
        {
            try
            {
                var doc = new List<string> { };
                if (StopRankingPending()) { return doc; }

                if (main.appConfig.bestofCompareChange && !velgerPeriode)
                {
                    DateTime previousOpenDay = FindLastTransactionday(date.AddDays(-3), date);
                    if (main.appConfig.ignoreSunday && previousOpenDay.DayOfWeek == DayOfWeek.Sunday)
                        previousOpenDay = DateTime.Now;

                    if (previousOpenDay.Month != date.Month)
                        previousOpenDay = DateTime.Now;

                    if (previousOpenDay.Date != DateTime.Now.Date)
                        tableListPreviousOpenDay = MakeTableList(previousOpenDay, previousOpenDay, true);
                }

                if (StopRankingPending()) { return doc; }

                if (!velgerPeriode)
                    tableListMtd = MakeTableList(FormMain.GetFirstDayOfMonth(date), date, true);
                else
                    tableListMtd = MakeTableList(dtFra, date, true);

                if (StopRankingPending()) { return doc; }

                if (main.appConfig.rankingCompareLastyear > 0)
                    tableListLastYearMtd = MakeTableList(FormMain.GetFirstDayOfMonth(date.AddYears(-1)), date.AddYears(-1), true);

                if (StopRankingPending()) { return doc; }

                if (main.appConfig.rankingCompareLastyear > 0)
                    tableListLastYear = MakeTableList(FormMain.GetFirstDayOfMonth(date.AddYears(-1)), FormMain.GetLastDayOfMonth(date.AddYears(-1)), true);

                if (tableListMtd.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner for denne måneden!</span><br>");
                    return doc;
                }

                var hashId = random.Next(999, 99999);
                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<div class='toggleAll' id='" + hashId + "'>");
                doc.Add("<div class='bestofhold'>");

                doc.AddRange(GetToppList("Inntjening", "Inntjen", "Omset", !main.appConfig.bestofSortInntjenSecondary, main.appConfig.bestofTallInntjen, "#D9D9D9", "Inntjen.", "Omset."));

                doc.AddRange(GetToppList("Finans", "Finans_Antall", "Finans_Inntjen", !main.appConfig.bestofSortTjenesterSecondary, main.appConfig.bestofTallFinans, "#f5954e", "Antall", "Inntjen."));

                doc.Add("</div>");
                doc.Add("<div class='bestofhold'>");

                doc.AddRange(GetToppList("TA", "TA_Antall", "TA_Omset", !main.appConfig.bestofSortTjenesterSecondary, main.appConfig.bestofTallTA, "#6699ff", "Antall", "Omset."));

                doc.AddRange(GetToppList("Strøm", "Strom_Antall", "Strom_Inntjen", !main.appConfig.bestofSortTjenesterSecondary, main.appConfig.bestofTallStrom, "#FAF39E", "Antall", "Inntjen."));

                doc.Add("</div>");
                doc.Add("<div class='bestofhold'>");

                doc.AddRange(GetToppList("RTG/SA", "Rtgsa_Antall", "Rtgsa_Inntjen", !main.appConfig.bestofSortTjenesterSecondary, main.appConfig.bestofTallTjen, "#80c34a", "Antall", "Inntjen."));

                doc.Add("</div>");
                doc.Add("</div>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { };
            }
        }

        public List<string> GetToppListAllLastOpenDay(DateTime date)
        {
            try
            {
                var doc = new List<string> { };
                if (StopRankingPending()) { return doc; }


                DateTime previousOpenDay = FindLastTransactionday(date.AddDays(-3), date);
                if (main.appConfig.ignoreSunday && previousOpenDay.DayOfWeek == DayOfWeek.Sunday)
                    previousOpenDay = DateTime.Now;

                if (previousOpenDay.Date != DateTime.Now.Date)
                    tableListPreviousOpenDay = MakeTableList(previousOpenDay, previousOpenDay, true);

                if (StopRankingPending()) { return doc; }

                if (tableListPreviousOpenDay != null && tableListPreviousOpenDay.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner for denne dagen!</span><br>");
                    return doc;
                }

                var hashId = random.Next(999, 99999);
                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<div class='toggleAll' id='" + hashId + "'>");
                doc.Add("<div class='bestofhold'>");

                doc.AddRange(GetTopp("Inntjening", "Inntjen", "Omset", !main.appConfig.bestofSortInntjenSecondary, "#D9D9D9", "tjente", " kr"));

                doc.AddRange(GetTopp("Finans", "Finans_Antall", "Finans_Inntjen", !main.appConfig.bestofSortTjenesterSecondary, "#f5954e", "solgte"));

                doc.Add("</div>");
                doc.Add("<div class='bestofhold'>");

                doc.AddRange(GetTopp("TA", "TA_Antall", "TA_Omset", !main.appConfig.bestofSortTjenesterSecondary, "#6699ff", "solgte"));

                doc.AddRange(GetTopp("Strøm", "Strom_Antall", "Strom_Inntjen", !main.appConfig.bestofSortTjenesterSecondary, "#FAF39E", "solgte"));

                doc.Add("</div>");
                doc.Add("<div class='bestofhold'>");

                doc.AddRange(GetTopp("RTG/SA", "Rtgsa_Antall", "Rtgsa_Inntjen", !main.appConfig.bestofSortTjenesterSecondary, "#80c34a", "solgte"));

                doc.Add("</div>");
                doc.Add("</div>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { };
            }
        }

        public List<string> GetToppList(string argCaption, string argTableColumnPrimary, string argTableColumnSecondary, bool argOrderPrimary, int argLength, string argColor, string argTitlePrimaryColumn, string argTitleSecondaryColumn)
        {
            try
            {
                var doc = new List<string> { };
                string urlID = "linkm"; // Marker linker for måned-periode

                DataView dv = new DataView(tableListMtd);
                if (argOrderPrimary)
                    dv.Sort = argTableColumnPrimary + " DESC, " + argTableColumnSecondary + " DESC";
                else
                    dv.Sort = argTableColumnSecondary + " DESC, " + argTableColumnPrimary + " DESC";
                DataTable sortedTable = dv.ToTable();

                doc.Add("<div class='bestof'>Topp " + argLength + " selgere " + argCaption);
                if (main.appConfig.bestofHoppoverKasse && argCaption == "Inntjening")
                    doc.Add("*<br>");
                else
                    doc.Add("<br>");
                doc.Add("<table class='tblBestOf'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.Add("<thead><tr>");

                doc.Add("<th class=\"{sorter: 'text'}\" width=95 style='background:" + argColor + ";'>Selger</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=70 style='background:" + argColor + ";'>" + argTitlePrimaryColumn + "</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 style='background:" + argColor + ";'>" + argTitleSecondaryColumn + "</td>");

                doc.Add("</tr></thead>");
                doc.Add("<tbody>");

                string end = ""; decimal expTot = 0; string expStrTot = "";
                for (int i = 0; i < argLength; i++)
                {
                    if (argLength == i + 1)
                        end = " style='border-bottom: #000 1px solid;' ";

                    if (sortedTable.Rows[i]["Selgerkode"].ToString() == "TOTALT" || (main.appConfig.bestofHoppoverKasse && main.salesCodes.GetKategori(sortedTable.Rows[i]["Selgerkode"].ToString()) == "Kasse" && argCaption == "Inntjening"))
                    {
                        argLength++;
                        continue;
                    }

                    if (sortedTable.Rows.Count >= i && (Convert.ToInt32(sortedTable.Rows[i][argTableColumnSecondary]) != 0 && Convert.ToInt32(sortedTable.Rows[i][argTableColumnPrimary]) != 0))
                    {
                        decimal expDec = 0; string expStr = "";
                        if (!velgerPeriode && main.appConfig.bestofCompareChange && tableListPreviousOpenDay != null && (argCaption == "TA" || argCaption == "RTG/SA" || argCaption == "Finans" || argCaption == "Strøm"))
                        {
                            for (int b = 0; b < tableListPreviousOpenDay.Rows.Count; b++)
                            {
                                if (sortedTable.Rows[i][0].ToString() == tableListPreviousOpenDay.Rows[b][0].ToString())
                                {
                                    expDec = Convert.ToDecimal(tableListPreviousOpenDay.Rows[b][argTableColumnPrimary]);
                                    if (expDec > 0)
                                        expStr = "<span style='color:green;font-size:xx-small;'>(+" + ForkortTall(expDec) + ")</span> ";
                                    else if (expDec < 0)
                                        expStr = "<span style='color:red;font-size:xx-small;'>(-" + ForkortTall(expDec) + ")</span> ";
                                    break;
                                }
                            }
                            expTot += expDec;
                        }

                        doc.Add("<tr><td class='text-cat'" + end + "><a href='#" + urlID + "s" + sortedTable.Rows[i][0].ToString() + "'>" + main.salesCodes.GetNavn(sortedTable.Rows[i][0].ToString()) + "</a></td>");
                        doc.Add("<td class='numbers-gen'" + end + ">" + expStr + ForkortTall(Convert.ToDecimal(sortedTable.Rows[i][argTableColumnPrimary])) + "</td>");
                        doc.Add("<td class='numbers-gen'" + end + ">" + ForkortTall(Convert.ToDecimal(sortedTable.Rows[i][argTableColumnSecondary])) + "</td>");
                        doc.Add("</tr>");
                    }
                    else
                    {
                        doc.Add("<tr><td class='text-cat'" + end + ">&nbsp;</td>");
                        doc.Add("<td class='numbers-gen'" + end + "></td>");
                        doc.Add("<td class='numbers-gen'" + end + "></td>");
                        doc.Add("</tr>");
                    }
                }

                if (!velgerPeriode && main.appConfig.bestofCompareChange && tableListPreviousOpenDay != null && (argCaption == "TA" || argCaption == "RTG/SA" || argCaption == "Finans" || argCaption == "Strøm"))
                {

                    expTot = Convert.ToDecimal(tableListPreviousOpenDay.Rows[tableListPreviousOpenDay.Rows.Count - 1][argTableColumnPrimary]);
                    if (expTot > 0)
                        expStrTot = "<span style='color:green;font-size:xx-small;'>(+" + ForkortTall(expTot) + ")</span> ";
                    else if (expTot < 0)
                        expStrTot = "<span style='color:red;font-size:xx-small;'>(-" + ForkortTall(expTot) + ")</span> ";
                }

                doc.Add("<tr><td class='text-cat linebreak'><b>TOTALT</b></td></td>");
                doc.Add("<td class='numbers-gen linebreak'>" + expStrTot + ForkortTall(Convert.ToDecimal(tableListMtd.Rows[tableListMtd.Rows.Count - 1][argTableColumnPrimary])) + "</td>");
                doc.Add("<td class='numbers-gen linebreak'>" + ForkortTall(Convert.ToDecimal(tableListMtd.Rows[tableListMtd.Rows.Count - 1][argTableColumnSecondary])) + "</td>");
                doc.Add("</tr>");
                if (!velgerPeriode && main.appConfig.rankingCompareLastyear > 0 && tableListLastYearMtd.Rows.Count > 0 && tableListLastYear.Rows.Count > 0)
                {
                    doc.Add("<tr><td class='text-cat'><b>" + dtTil.AddYears(-1).Year + " MTD</b></td></td>");
                    doc.Add("<td class='numbers-gen'>" + ForkortTall(Convert.ToDecimal(tableListLastYearMtd.Rows[tableListLastYearMtd.Rows.Count - 1][argTableColumnPrimary])) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + ForkortTall(Convert.ToDecimal(tableListLastYearMtd.Rows[tableListLastYearMtd.Rows.Count - 1][argTableColumnSecondary])) + "</td>");
                    doc.Add("</tr>");

                    doc.Add("<tr><td class='text-cat'><b>" + dtTil.AddYears(-1).Year + " TOT</b></td></td>");
                    doc.Add("<td class='numbers-gen'>" + ForkortTall(Convert.ToDecimal(tableListLastYear.Rows[tableListLastYear.Rows.Count - 1][argTableColumnPrimary])) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + ForkortTall(Convert.ToDecimal(tableListLastYear.Rows[tableListLastYear.Rows.Count - 1][argTableColumnSecondary])) + "</td>");
                    doc.Add("</tr>");
                }

                if (argCaption == "Inntjening")
                    argCaption = "Inn.";
                if (!velgerPeriode && main.appConfig.rankingCompareLastyear > 0 && tableListLastYearMtd.Rows.Count > 0 && tableListLastYear.Rows.Count > 0 && main.appConfig.bestofVisBesteLastYear)
                {
                    DataView dvCompare;
                    if (main.appConfig.bestofVisBesteLastYearTotal)
                        dvCompare = new DataView(tableListLastYear);
                    else
                        dvCompare = new DataView(tableListLastYearMtd);

                    if (argOrderPrimary)
                        dvCompare.Sort = argTableColumnPrimary + " DESC, " + argTableColumnSecondary + " DESC";
                    else
                        dvCompare.Sort = argTableColumnSecondary + " DESC, " + argTableColumnPrimary + " DESC";
                    DataTable sortedTableCompare = dvCompare.ToTable();

                    var rows = sortedTableCompare.Select("Selgerkode LIKE 'TOTALT'");
                    foreach (var row in rows)
                        row.Delete();

                    doc.Add("<tfoot><tr><td colspan=3>");
                    if (main.appConfig.bestofVisBesteLastYearTotal)
                        doc.Add("Best i fjor: " + main.salesCodes.GetNavn(sortedTableCompare.Rows[0][0].ToString()) + " med " + ForkortTall(Convert.ToDecimal(sortedTableCompare.Rows[0][argTableColumnPrimary])) + " " + argCaption);
                    else
                        doc.Add("Best i fjor MTD: " + main.salesCodes.GetNavn(sortedTableCompare.Rows[0][0].ToString()) + " med " + ForkortTall(Convert.ToDecimal(sortedTableCompare.Rows[0][argTableColumnPrimary])) + " " + argCaption);
                    doc.Add("</td></tr></tfoot>");
                }

                doc.Add("</table></td></tr></table></div>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { };
            }
        }

        public List<string> GetToppListLarge(DateTime date, string kat, bool save = false)
        {
            try
            {
                var doc = new List<string> { };
                if (StopRankingPending()) { return doc; }

                if (save)
                {
                    if (main.appConfig.bestofCompareChange && !velgerPeriode)
                    {
                        DateTime previousOpenDay = FindLastTransactionday(date.AddDays(-3), date);
                        if (main.appConfig.ignoreSunday && previousOpenDay.DayOfWeek == DayOfWeek.Sunday)
                            previousOpenDay = DateTime.Now;

                        if (previousOpenDay.Month != date.Month)
                            previousOpenDay = DateTime.Now;

                        if (previousOpenDay.Date != DateTime.Now.Date)
                            tableListPreviousOpenDay = MakeTableList(previousOpenDay, previousOpenDay, true);
                    }

                    if (main.appConfig.bestofCompareChange && !velgerPeriode && main.appConfig.favVis && FormMain.Favoritter.Count > 1)
                    {
                        DateTime previousOpenDay = FindLastTransactionday(date.AddDays(-3), date);
                        if (main.appConfig.ignoreSunday && previousOpenDay.DayOfWeek == DayOfWeek.Sunday)
                            previousOpenDay = DateTime.Now;

                        if (previousOpenDay.Month != date.Month)
                            previousOpenDay = DateTime.Now;

                        if (previousOpenDay.Date != DateTime.Now.Date)
                            tableListPreviousOpenDayFav = MakeTableFav(previousOpenDay, previousOpenDay, true);
                    }


                    if (StopRankingPending()) { return doc; }

                    if (!velgerPeriode)
                        tableListMtd = MakeTableList(FormMain.GetFirstDayOfMonth(date), date, true);
                    else
                        tableListMtd = MakeTableList(dtFra, date, true);


                    if (StopRankingPending()) { return doc; }

                    if (!velgerPeriode && main.appConfig.favVis)
                        tableListMtdAvd = MakeTableFav(FormMain.GetFirstDayOfMonth(date), date, true);
                    else
                        tableListMtdAvd = MakeTableFav(dtFra, date, true);

                    if (StopRankingPending()) { return doc; }

                    if (main.appConfig.rankingCompareLastyear > 0)
                        tableListLastYearMtd = MakeTableList(FormMain.GetFirstDayOfMonth(date.AddYears(-1)), date.AddYears(-1), true);

                    if (StopRankingPending()) { return doc; }

                    if (main.appConfig.rankingCompareLastyear > 0)
                        tableListLastYear = MakeTableList(FormMain.GetFirstDayOfMonth(date.AddYears(-1)), FormMain.GetLastDayOfMonth(date.AddYears(-1)), true);
                }
                else
                {

                    if (tableListMtd.Rows.Count == 0)
                    {
                        doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner for denne måneden!</span><br>");
                        return doc;
                    }

                    if (kat == "Inntjen")
                        doc.AddRange(GetToppListLarge("Inntjening", "Inntjen", "Omset", "Bto_Margin", !main.appConfig.bestofSortInntjenSecondary, main.appConfig.listerMaxLinjer, "#D9D9D9", "Inntjen.", "Omset.", "Margin"));
                    else if (kat == "Finans")
                        doc.AddRange(GetToppListLarge("Finans", "Finans_Antall", "Finans_Inntjen", "Finans_Margin", !main.appConfig.bestofSortTjenesterSecondary, main.appConfig.listerMaxLinjer, "#f5954e", "Antall", "Inntjen.", "SoM"));
                    else if (kat == "TA")
                        doc.AddRange(GetToppListLarge("TA", "TA_Antall", "TA_Omset", "TA_Margin", !main.appConfig.bestofSortTjenesterSecondary, main.appConfig.listerMaxLinjer, "#6699ff", "Antall", "Omset.", "SoB"));
                    else if (kat == "Strom")
                        doc.AddRange(GetToppListLarge("Strøm", "Strom_Antall", "Strom_Inntjen", "Strom_Margin", !main.appConfig.bestofSortTjenesterSecondary, main.appConfig.listerMaxLinjer, "#FAF39E", "Antall", "Inntjen.", "SoM"));
                    else if (kat == "RTGSA")
                        doc.AddRange(GetToppListLarge("RTGSA", "Rtgsa_Antall", "Rtgsa_Inntjen", "Rtgsa_Margin", !main.appConfig.bestofSortTjenesterSecondary, main.appConfig.listerMaxLinjer, "#80c34a", "Antall", "Inntjen.", "SoM"));
                    else if (kat == "Tilbehør")
                        doc.AddRange(GetToppListLarge("Tilbehør", "Accessories_Antall", "Accessories_Inntjen", "Accessories_Margin", main.appConfig.bestofSortTjenesterSecondary, main.appConfig.listerMaxLinjer, "#f36565", "Antall", "Inntjen.", "SoM"));
                    else
                        doc.Add(kat + " eksisterer ikke!");
                }

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { };
            }
        }

        public List<string> GetToppListLarge(string argCaption, string argTableColumnPrimary, string argTableColumnSecondary, string argTableColumnTertiary, bool argOrderPrimary, int argLength, string argColor, string argTitlePrimaryColumn, string argTitleSecondaryColumn, string argTitleTertiaryColumn)
        {
            try
            {
                var doc = new List<string> { };
                string urlID = "linkm"; // Marker linker for måned-periode
                var hashId = random.Next(999, 99999);

                DataView dv = new DataView(tableListMtd);
                if (argOrderPrimary)
                    dv.Sort = argTableColumnPrimary + " DESC, " + argTableColumnSecondary + " DESC";
                else
                    dv.Sort = argTableColumnSecondary + " DESC, " + argTableColumnPrimary + " DESC";
                DataTable sortedTable = dv.ToTable();

                main.openXml.SaveDocument(dt, "Lister", argCaption, dtPick, argCaption.ToUpper() + " - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='tblBestOf toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.Add("<thead><tr>");

                doc.Add("<th class=\"{sorter: 'text'}\" width='230' style='background:" + argColor + ";'>Selger</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width='227' style='background:" + argColor + ";'>" + argTitlePrimaryColumn + "</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width='227' style='background:" + argColor + ";'>" + argTitleSecondaryColumn + "</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width='226' style='background:" + argColor + ";'>" + argTitleTertiaryColumn + "</td>");

                doc.Add("</tr></thead>");
                doc.Add("<tbody>");

                decimal expTot = 0; string expStrTot = "";
                decimal c1 = 0, c2 = 0, c3 = 0;
                for (int i = 0; i < sortedTable.Rows.Count; i++)
                {
                    if (Convert.ToDecimal(sortedTable.Rows[i][argTableColumnPrimary]) == 0)
                        break;

                    if (sortedTable.Rows[i]["Selgerkode"].ToString() == "TOTALT")
                    {
                        argLength++;
                        continue;
                    }

                    decimal expDec = 0; string expStr = "";
                    if (!velgerPeriode && main.appConfig.bestofCompareChange && tableListPreviousOpenDay != null && (argCaption == "TA" || argCaption == "RTGSA" || argCaption == "Finans" || argCaption == "Strøm" || argCaption == "Tilbehør"))
                    {
                        for (int b = 0; b < tableListPreviousOpenDay.Rows.Count; b++)
                        {
                            if (sortedTable.Rows[i][0].ToString() == tableListPreviousOpenDay.Rows[b][0].ToString())
                            {
                                expDec = Convert.ToDecimal(tableListPreviousOpenDay.Rows[b][argTableColumnPrimary]);
                                if (expDec > 0)
                                    expStr = "<span style='color:green;font-size:12pt;'>(+" + ForkortTall(expDec) + ")</span> ";
                                else if (expDec < 0)
                                    expStr = "<span style='color:red;font-size:12pt;'>(-" + ForkortTall(expDec) + ")</span> ";
                                break;
                            }
                        }
                        expTot += expDec;
                    }

                    if (argLength < i)
                    {
                        c1 += Convert.ToDecimal(sortedTable.Rows[i][argTableColumnPrimary]);
                        c2 += Convert.ToDecimal(sortedTable.Rows[i][argTableColumnSecondary]);
                        c3 += Convert.ToDecimal(sortedTable.Rows[i][argTableColumnTertiary]);

                        if (i == sortedTable.Rows.Count - 1)
                        {
                            c3 = c3 / (i - argLength);
                            doc.Add("<tr><td class='text-cat' style=\"font-size:14pt;\">Mer..</td>");
                            doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\">" + expStr + ForkortTall(c1) + "</td>");
                            doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\">" + ForkortTall(c2) + "</td>");
                            doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\">" + PercentShare(c3.ToString()) + "</td>");
                            doc.Add("</tr>");
                        }
                        continue;
                    }
                    doc.Add("<tr><td class='text-cat' style=\"font-size:14pt;\"><a href='#" + urlID + "s" + sortedTable.Rows[i][0].ToString() + "'>" + main.salesCodes.GetNavn(sortedTable.Rows[i][0].ToString()) + "</a></td>");
                    doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\">" + expStr + ForkortTall(Convert.ToDecimal(sortedTable.Rows[i][argTableColumnPrimary])) + "</td>");
                    doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\">" + ForkortTall(Convert.ToDecimal(sortedTable.Rows[i][argTableColumnSecondary])) + "</td>");
                    doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\">" + PercentShare(Convert.ToDecimal(sortedTable.Rows[i][argTableColumnTertiary]).ToString()) + "</td>");
                    doc.Add("</tr>");
                }

                if (!velgerPeriode && main.appConfig.bestofCompareChange && tableListPreviousOpenDay != null && (argCaption == "TA" || argCaption == "RTGSA" || argCaption == "Finans" || argCaption == "Strøm" || argCaption == "Tilbehør"))
                {
                    expTot = Convert.ToDecimal(tableListPreviousOpenDay.Rows[tableListPreviousOpenDay.Rows.Count - 1][argTableColumnPrimary]);
                    if (expTot > 0)
                        expStrTot = "<span style='color:green;font-size:12pt;'>(+" + ForkortTall(expTot) + ")</span> ";
                    else if (expTot < 0)
                        expStrTot = "<span style='color:red;font-size:12pt;'>(-" + ForkortTall(expTot) + ")</span> ";
                }

                doc.Add("<tfoot><tr><td class='text-cat linebreak' style=\"font-size:14pt;\"><b>TOTALT</b></td></td>");
                doc.Add("<td class='numbers-gen linebreak' style=\"font-size:14pt;\">" + expStrTot + ForkortTall(Convert.ToDecimal(tableListMtd.Rows[tableListMtd.Rows.Count - 1][argTableColumnPrimary])) + "</td>");
                doc.Add("<td class='numbers-gen linebreak' style=\"font-size:14pt;\">" + ForkortTall(Convert.ToDecimal(tableListMtd.Rows[tableListMtd.Rows.Count - 1][argTableColumnSecondary])) + "</td>");
                doc.Add("<td class='numbers-gen linebreak' style=\"font-size:14pt;\">" + PercentShare(Convert.ToDecimal(tableListMtd.Rows[tableListMtd.Rows.Count - 1][argTableColumnTertiary]).ToString()) + "</td>");
                doc.Add("</tr></tfoot>");
                doc.Add("</table></td></tr></table>");

                if (!velgerPeriode && main.appConfig.rankingCompareLastyear > 0 && tableListLastYearMtd.Rows.Count > 0 && tableListLastYear.Rows.Count > 0)
                {
                    doc.Add("<table class='tblBestOf'><tr><td>");
                    doc.Add("<table class='tablesorter'>");
                    doc.Add("<tr><td class='text-cat' style=\"font-size:14pt;\" width=224 ><b>" + dtTil.AddYears(-1).Year + " MTD</b></td></td>");
                    doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\" width=220 >" + ForkortTall(Convert.ToDecimal(tableListLastYearMtd.Rows[tableListLastYearMtd.Rows.Count - 1][argTableColumnPrimary])) + "</td>");
                    doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\" width=220 >" + ForkortTall(Convert.ToDecimal(tableListLastYearMtd.Rows[tableListLastYearMtd.Rows.Count - 1][argTableColumnSecondary])) + "</td>");
                    doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\" width=219 >" + PercentShare(Convert.ToDecimal(tableListLastYearMtd.Rows[tableListLastYearMtd.Rows.Count - 1][argTableColumnTertiary]).ToString()) + "</td>");
                    doc.Add("</tr>");

                    doc.Add("<tr><td class='text-cat' style=\"font-size:14pt;\" width=224 ><b>" + dtTil.AddYears(-1).Year + " TOTALT</b></td></td>");
                    doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\" width=220 >" + ForkortTall(Convert.ToDecimal(tableListLastYear.Rows[tableListLastYear.Rows.Count - 1][argTableColumnPrimary])) + "</td>");
                    doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\" width=220 >" + ForkortTall(Convert.ToDecimal(tableListLastYear.Rows[tableListLastYear.Rows.Count - 1][argTableColumnSecondary])) + "</td>");
                    doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\" width=219 >" + PercentShare(Convert.ToDecimal(tableListLastYear.Rows[tableListLastYear.Rows.Count - 1][argTableColumnTertiary]).ToString()) + "</td>");
                    doc.Add("</tr>");
                    doc.Add("</tbody>");
                    doc.Add("</table></td></tr></table>");
                }


                if (main.appConfig.favVis && FormMain.Favoritter.Count > 1)
                {
                    doc.Add("<table class='tblBestOf'><tr><td>");
                    doc.Add("<table class='tablesorter'>");
                    doc.Add("<tbody>");
                    decimal expTotFav = 0;
                    for (int i = 1; i < tableListMtdAvd.Rows.Count; i++)
                    {
                        if (tableListMtdAvd.Rows[i]["Selgerkode"].ToString() == "TOTALT")
                        {
                            argLength++;
                            continue;
                        }

                        decimal expDecFav = 0; string expStrFav = "";
                        if (!velgerPeriode && main.appConfig.bestofCompareChange && tableListPreviousOpenDayFav != null && (argCaption == "TA" || argCaption == "RTGSA" || argCaption == "Finans" || argCaption == "Strøm" || argCaption == "Tilbehør"))
                        {
                            for (int b = 0; b < tableListPreviousOpenDayFav.Rows.Count; b++)
                            {
                                if (tableListMtdAvd.Rows[i][0].ToString() == tableListPreviousOpenDayFav.Rows[b][0].ToString())
                                {
                                    expDecFav = Convert.ToDecimal(tableListPreviousOpenDayFav.Rows[b][argTableColumnPrimary]);
                                    if (expDecFav > 0)
                                        expStrFav = "<span style='color:green;font-size:12pt;'>(+" + ForkortTall(expDecFav) + ")</span> ";
                                    else if (expDecFav < 0)
                                        expStrFav = "<span style='color:red;font-size:12pt;'>(-" + ForkortTall(expDecFav) + ")</span> ";
                                    break;
                                }
                            }
                            expTotFav += expDecFav;
                        }

                        doc.Add("<tr><td class='text-cat' style=\"font-size:14pt;\" width=224 >" + avdeling.Get(Convert.ToInt32(tableListMtdAvd.Rows[i][0])).Replace(" ", "&nbsp;") + "</a></td>");
                        doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\" width=220 >" + expStrFav + ForkortTall(Convert.ToDecimal(tableListMtdAvd.Rows[i][argTableColumnPrimary])) + "</td>");
                        doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\" width=220 >" + ForkortTall(Convert.ToDecimal(tableListMtdAvd.Rows[i][argTableColumnSecondary])) + "</td>");
                        doc.Add("<td class='numbers-gen' style=\"font-size:14pt;\" width=219 >" + PercentShare(Convert.ToDecimal(tableListMtdAvd.Rows[i][argTableColumnTertiary]).ToString()) + "</td>");
                        doc.Add("</tr>");
                    }
                    doc.Add("</tbody>");
                    doc.Add("</table></td></tr></table>");
                }

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { };
            }
        }

        public List<string> GetTopp(string argCaption, string argTableColumnPrimary, string argTableColumnSecondary, bool argOrderPrimary, string argColor, string argCustom, string argCustom2 = "")
        {
            try
            {
                var doc = new List<string> { };

                if (tableListPreviousOpenDay == null || tableListPreviousOpenDay.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner</span><br>");
                    return doc;
                }

                DataView dv = new DataView(tableListPreviousOpenDay);
                if (argOrderPrimary)
                    dv.Sort = argTableColumnPrimary + " DESC, " + argTableColumnSecondary + " DESC";
                else
                    dv.Sort = argTableColumnSecondary + " DESC, " + argTableColumnPrimary + " DESC";
                DataTable sortedTable = dv.ToTable();

                doc.Add("<div class='bestofStar'>Beste " + argCaption + " selger<br>");
                doc.Add("<div class='tblBestOfBehind'><br><br><br>" + argCaption + "</div><div class='tblBestOfStar'><div class='tblBestOfText'><br><br><br><br><br><br>");

                bool match = false;
                for (int i = 0; i < sortedTable.Rows.Count; i++)
                {
                    if (sortedTable.Rows[i]["Selgerkode"].ToString() != "TOTALT" && Convert.ToDecimal(sortedTable.Rows[i][argTableColumnPrimary]) > 0
                        && !(main.appConfig.bestofHoppoverKasse && main.salesCodes.GetKategori(sortedTable.Rows[i]["Selgerkode"].ToString()) == "Kasse" && argCaption == "Inntjening"))
                    {
                        doc.Add("<span style='font-size:x-large;'>" + main.salesCodes.GetNavn(sortedTable.Rows[i]["Selgerkode"].ToString())
                            + "</span><br>" + argCustom + " " + ForkortTall(Convert.ToDecimal(sortedTable.Rows[i][argTableColumnPrimary])) + argCustom2
                            + "<br><span style='font-size:small'>" + ForkortTall(Convert.ToDecimal(sortedTable.Rows[i][argTableColumnSecondary])) + " kr</span>");
                        if (main.appConfig.bestofHoppoverKasse && argCaption == "Inntjening")
                            doc.Add("*");
                        match = true;
                        break;
                    }
                }
                if (!match)
                    doc.Add("<br>Ingen solgte " + argCaption + "!");

                doc.Add("</div></div></div>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { };
            }
        }

        public DataTable ReadyTableList()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Selgerkode", typeof(string));
            dataTable.Columns.Add("Antallsalg", typeof(int));
            dataTable.Columns.Add("Omset", typeof(decimal));
            dataTable.Columns.Add("Inntjen", typeof(decimal));
            dataTable.Columns.Add("Bto_Margin", typeof(decimal));
            dataTable.Columns.Add("Rtgsa_Antall", typeof(int));
            dataTable.Columns.Add("Rtgsa_Inntjen", typeof(decimal));
            dataTable.Columns.Add("Rtgsa_Omset", typeof(decimal));
            dataTable.Columns.Add("Rtgsa_Margin", typeof(decimal));
            dataTable.Columns.Add("Strom_Antall", typeof(int));
            dataTable.Columns.Add("Strom_Inntjen", typeof(decimal));
            dataTable.Columns.Add("Strom_Margin", typeof(decimal));
            dataTable.Columns.Add("TA_Antall", typeof(int));
            dataTable.Columns.Add("TA_Inntjen", typeof(decimal));
            dataTable.Columns.Add("TA_Omset", typeof(decimal));
            dataTable.Columns.Add("TA_Margin", typeof(decimal));
            dataTable.Columns.Add("Finans_Antall", typeof(int));
            dataTable.Columns.Add("Finans_Inntjen", typeof(decimal));
            dataTable.Columns.Add("Finans_Margin", typeof(decimal));
            dataTable.Columns.Add("Accessories_Antall", typeof(int));
            dataTable.Columns.Add("Accessories_Inntjen", typeof(decimal));
            dataTable.Columns.Add("Accessories_Omset", typeof(decimal));
            dataTable.Columns.Add("Accessories_Margin", typeof(decimal));
            return dataTable;
        }


    }
}