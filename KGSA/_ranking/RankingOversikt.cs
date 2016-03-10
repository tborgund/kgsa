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
    public class RankingOversikt : Ranking
    {
        private DataTable dtDetaljer;
        public List<VarekodeList> varekoderAlle;
        public IEnumerable<string> varekoderAlleAlias;
        public RankingOversikt() { }

        public RankingOversikt(FormMain form, DateTime dtFraArg, DateTime dtTilArg)
        {
            this.main = form;
            dtFra = dtFraArg;
            dtTil = dtTilArg;
            velgerPeriode = FormMain.datoPeriodeVelger;

            this.varekoderAlle = main.appConfig.varekoder.ToList();
            this.varekoderAlleAlias = varekoderAlle.Where(item => item.synlig == true).Select(x => x.alias).Distinct();

            dtDetaljer = MakeTableOversikt(false, false, false, false);
        }

        private DataTable MakeTableOversikt(bool skipSelgere, bool skipKategori, bool skipTotalt, bool compare, bool compareMtd = false)
        {
            try
            {
                DataTable dtWork = ReadyTableOversikt();

                DateTime dtMainFra = dtFra;
                DateTime dtMainTil = dtTil;

                if (compare)
                {
                    dtMainFra = dtFra.AddYears(-1);
                    if (compareMtd)
                        dtMainTil = dtTil.AddYears(-1);
                    else
                        dtMainTil = FormMain.GetLastDayOfMonth(dtTil.AddYears(-1));
                }

                var rowsGet = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select("(Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                sqlce = rowsGet.Any() ? rowsGet.CopyToDataTable() : sqlce.Clone();
                if (sqlce.Rows.Count == 0)
                    return dtWork;

                decimal dager = dtMainTil.Day;
                if (dtMainFra.Month != dtMainTil.Month && dtMainFra.Year != dtMainTil.Year)
                    dager = (dtMainTil - dtMainFra).Days;
                decimal dagerImnd = FormMain.GetLastDayOfMonth(dtMainTil).Day;
                int finansKravAccTot = 0, modKravAccTot = 0, stromKravAccTot = 0, rtgsaKravAccTot = 0;

                decimal multi = 0;
                if (dager != 0 && main.appConfig.oversiktKravMtd)
                    multi = Math.Round(dager / dagerImnd, 6);
                if (!main.appConfig.oversiktKravMtd)
                    multi = 1;

                for (int d = 1; d <= 9; d++)
                {
                    // A V D E L I N G
                    if (StopRankingPending())
                        return dtWork;

                    string sKat = "";
                    if (d == 1)
                        sKat = "MDA";
                    if (d == 2)
                        sKat = "AudioVideo";
                    if (d == 3)
                        sKat = "SDA";
                    if (d == 4)
                        sKat = "Tele";
                    if (d == 5)
                        sKat = "Data";
                    if (d == 6)
                        sKat = "Kjøkken";
                    if (d == 7)
                        sKat = "Kasse";
                    if (d == 8)
                        sKat = "Aftersales";
                    if (d == 9)
                        sKat = "Cross";

                    string[] selgere = main.salesCodes.GetSalesCodes(sKat, false);
                    decimal sInntjen = 0, sOmset = 0, sOmsetExMva = 0, sTjenInntjen = 0, sTjenOmset = 0, sAntallTjen = 0, sAntallSalg = 0, sAccessoriesAntall = 0, sAccessoriesInntjen = 0, sAccessoriesOmset = 0;
                    decimal sStromAntall = 0, sStromInntjen = 0, sModAntall = 0, sModInntjen = 0, sFinansAntall = 0, sFinansInntjen = 0, sModOmset = 0, sKupppvarer = 0;
                    decimal sSnittAntall = 0, sSnittOmset = 0, sSnittInntjen = 0, sAntallBilag = 0;

                    int finansKravAcc = 0, modKravAcc = 0, stromKravAcc = 0, rtgsaKravAcc = 0;

                    if (!skipKategori && !(main.appConfig.oversiktHideAftersales && sKat == "Aftersales")
                        && !(main.appConfig.oversiktHideKitchen && sKat == "Kjøkken"))
                    {
                        // K A T E G O R I
                        if (StopRankingPending())
                            return dtWork;

                        DataTable dt;
                        if (d > 0 && d <= 6)
                        {
                            var rows = sqlce.Select("(Varegruppe >= " + d + "00 AND Varegruppe <= " + d + "99) OR Varegruppe = 961");
                            dt = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                        }
                        else
                        {
                            string str = ""; DataRow[] rows;
                            foreach(string sel in selgere)
                                str += " OR Selgerkode = '" + sel + "'";
                            if (str.Length > 0)
                            {
                                str = str.Substring(4, str.Length - 4);
                                rows = sqlce.Select(str);
                            }
                            else
                                rows = sqlce.Select("Varegruppe = 0");
                            dt = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                        }

                        DataRow dtRow = dtWork.NewRow();

                        sAntallSalg += Compute(dt, "Sum(Antall)", null);
                        sOmset += Compute(dt, "Sum(Salgspris)", null);
                        sOmsetExMva += Compute(dt, "Sum(SalgsprisExMva)", null);
                        sInntjen += Compute(dt, "Sum(Btokr)", null);
                        sAntallBilag = dt.AsEnumerable().Select(g => g.Field<int>("Bilagsnr")).Distinct().Count();

                        var rowf = dt.Select("[Varegruppe] = 961");
                        for (int f = 0; f < rowf.Length;f++ )
                        {
                            var rows2 = sqlce.Select("[Bilagsnr] = " + rowf[f]["Bilagsnr"]);
                            DataTable dtFinans = rows2.Any() ? rows2.CopyToDataTable() : sqlce.Clone();
                            dtFinans.DefaultView.Sort = "Salgspris DESC";
                            int gruppe = Convert.ToInt32(dtFinans.Rows[0]["Varegruppe"].ToString().Substring(0, 1));
                            if (gruppe == d || d > 6)
                            {
                                sFinansAntall += Compute(dt, "Sum(Antall)", "[Varegruppe] = 961 AND [Bilagsnr] = " + dtFinans.Rows[0]["Bilagsnr"].ToString());
                                sFinansInntjen += Compute(dt, "Sum(Btokr)", "[Varegruppe] = 961 AND [Bilagsnr] = " + dtFinans.Rows[0]["Bilagsnr"].ToString());
                            }
                        }

                        sStromInntjen += Compute(dt, "Sum(Btokr)", "([Varekode] LIKE 'ELSTROM*' OR [Varekode] LIKE 'ELRABATT*')");
                        sStromAntall += Compute(dt, "Sum(Antall)", "[Varekode] LIKE 'ELSTROM*'");
                        sModInntjen += Compute(dt, "Sum(Btokr)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                        sModOmset += Compute(dt, "Sum(Salgspris)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                        sModAntall += Compute(dt, "Sum(Antall)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                        sKupppvarer += Compute(dt, "Sum(Antall)", "[Varekode] = 'KV'");

                        int[] accessoriesGrpList = main.appConfig.GetAccessorieGroups(d);
                        foreach (int ac in accessoriesGrpList)
                        {
                            sAccessoriesAntall += Compute(dt, "Sum(Antall)", "[Varegruppe] = " + ac);
                            sAccessoriesInntjen += Compute(dt, "Sum(Btokr)", "[Varegruppe] = " + ac);
                            sAccessoriesOmset += Compute(dt, "Sum(Salgspris)", "[Varegruppe] = " + ac);
                        }

                        int[] mainproductsGrpList = main.appConfig.GetMainproductGroups(d);
                        foreach (int grp in mainproductsGrpList)
                        {
                            sSnittAntall += Compute(dt, "Sum(Antall)", "[Varegruppe] = " + grp);
                            sSnittInntjen += Compute(dt, "Sum(Btokr)", "[Varegruppe] = " + grp);
                            sSnittOmset += Compute(dt, "Sum(Salgspris)", "[Varegruppe] = " + grp);
                        }

                        foreach (var varekode in varekoderAlle)
                        {
                            sTjenOmset += Compute(dt, "Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");
                            sTjenInntjen += Compute(dt, "Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                            if (varekode.synlig)
                                sAntallTjen += Compute(dt, "Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                        }

                        for (int i = 0; i < selgere.Length; i++)
                        {
                            if (main.appConfig.oversiktKravVis && d > 6)
                            {
                                finansKravAcc += (int)(main.salesCodes.GetKrav(selgere[i], sKat, "Finans", main.appConfig.oversiktKravFinansAntall));
                                modKravAcc += (int)(main.salesCodes.GetKrav(selgere[i], sKat, "Mod", main.appConfig.oversiktKravModAntall));
                                stromKravAcc += (int)(main.salesCodes.GetKrav(selgere[i], sKat, "Strom", main.appConfig.oversiktKravStromAntall));
                                rtgsaKravAcc += (int)(main.salesCodes.GetKrav(selgere[i], sKat, "Rtgsa", main.appConfig.oversiktKravRtgsaAntall));
                            }
                        }

                        if (sInntjen != 0)
                        {
                            dtRow["Kat"] = sKat;
                            dtRow["Salg"] = sAntallSalg;
                            dtRow["Omset"] = sOmset;
                            dtRow["Inntjen"] = sInntjen;
                            if (sAntallSalg != 0)
                                dtRow["SnittOmsetAlle"] = sOmset / sAntallSalg;
                            else
                                dtRow["SnittOmsetAlle"] = 0;
                            dtRow["OmsetExMva"] = Math.Round(sOmsetExMva, 2);
                            if (sOmsetExMva != 0)
                                dtRow["Prosent"] = Math.Round(sInntjen / sOmsetExMva * 100, 2);
                            else
                                dtRow["Prosent"] = 0;
                            dtRow["BilagAntall"] = sAntallBilag;
                            if (sAntallBilag != 0)
                                dtRow["BilagSnittOmset"] = Math.Round(sOmset / sAntallBilag, 2);
                            else
                                dtRow["BilagSnittOmset"] = 0;
                            if (sAntallBilag != 0)
                                dtRow["BilagSnittInntjen"] = Math.Round(sInntjen / sAntallBilag, 2);
                            else
                                dtRow["BilagSnittInntjen"] = 0;
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
                                dtRow["ModMargin"] = Math.Round(sModOmset / sOmsetExMva * 100, 2);
                            else
                                dtRow["ModMargin"] = 0;
                            dtRow["FinansInntjen"] = sFinansInntjen;
                            dtRow["FinansAntall"] = sFinansAntall;
                            if (sInntjen != 0)
                                dtRow["FinansMargin"] = Math.Round(sFinansInntjen / sInntjen * 100, 2);
                            else
                                dtRow["FinansMargin"] = 0;
                            dtRow["Kuppvarer"] = sKupppvarer;

                            dtRow["AccessoriesAntall"] = sAccessoriesAntall;
                            dtRow["AccessoriesInntjen"] = sAccessoriesInntjen;
                            dtRow["AccessoriesOmset"] = sAccessoriesOmset;
                            if (sInntjen != 0)
                                dtRow["AccessoriesMargin"] = Math.Round(sAccessoriesInntjen / sInntjen * 100, 2);
                            else
                                dtRow["AccessoriesMargin"] = 0;
                            if (sOmsetExMva != 0)
                                dtRow["AccessoriesSoB"] = Math.Round(sAccessoriesOmset / sOmsetExMva * 100, 2);
                            else
                                dtRow["AccessoriesSoB"] = 0;

                            dtRow["SnittAntall"] = sSnittAntall;
                            if (sSnittAntall != 0)
                            {
                                dtRow["SnittInntjen"] = sSnittInntjen / sSnittAntall;
                                dtRow["SnittOmset"] = sSnittOmset / sSnittAntall;
                            }
                            else
                            {
                                dtRow["SnittInntjen"] = 0;
                                dtRow["SnittOmset"] = 0;
                            }

                            if (main.appConfig.oversiktKravVis)
                            {
                                finansKravAccTot += finansKravAcc;

                                decimal modTemp = sModOmset;
                                if (main.appConfig.oversiktKravModAntall)
                                    modTemp = sModAntall;

                                modKravAccTot += modKravAcc;
                                stromKravAccTot += stromKravAcc;
                                rtgsaKravAccTot += rtgsaKravAcc;

                                dtRow["FinansKrav"] = finansKravAcc * multi;
                                dtRow["FinansKravUmod"] = finansKravAcc;
                                if (main.appConfig.oversiktKravFinansAntall && finansKravAcc <= sFinansAntall && finansKravAcc != 0)
                                    dtRow["FinansKravUmodOppfylt"] = true;
                                else if (!main.appConfig.oversiktKravFinansAntall && finansKravAcc <= sFinansInntjen && finansKravAcc != 0)
                                    dtRow["FinansKravUmodOppfylt"] = true;
                                else
                                    dtRow["FinansKravUmodOppfylt"] = false;
                                finansKravAcc = (int)(finansKravAcc * multi);
                                if (main.appConfig.oversiktKravFinansAntall && finansKravAcc <= sFinansAntall && finansKravAcc != 0)
                                    dtRow["FinansKravOppfylt"] = true;
                                else if (!main.appConfig.oversiktKravFinansAntall && finansKravAcc <= sFinansInntjen && finansKravAcc != 0)
                                    dtRow["FinansKravOppfylt"] = true;
                                else
                                    dtRow["FinansKravOppfylt"] = false;
                                if ((bool)dtRow["FinansKravUmodOppfylt"])
                                    dtRow["FinansKravOppfylt"] = true;

                                dtRow["ModKrav"] = modKravAcc * multi;
                                dtRow["ModKravUmod"] = modKravAcc;
                                if (main.appConfig.oversiktKravModAntall && modKravAcc <= sModAntall && modKravAcc != 0)
                                    dtRow["ModKravUmodOppfylt"] = true;
                                else if (!main.appConfig.oversiktKravModAntall && modKravAcc <= modTemp && modKravAcc != 0)
                                    dtRow["ModKravUmodOppfylt"] = true;
                                else
                                    dtRow["ModKravUmodOppfylt"] = false;
                                modKravAcc = (int)(modKravAcc * multi);
                                if (main.appConfig.oversiktKravModAntall && modKravAcc <= sModAntall && modKravAcc != 0)
                                    dtRow["ModKravOppfylt"] = true;
                                else if (!main.appConfig.oversiktKravModAntall && modKravAcc <= modTemp && modKravAcc != 0)
                                    dtRow["ModKravOppfylt"] = true;
                                else
                                    dtRow["ModKravOppfylt"] = false;
                                if ((bool)dtRow["ModKravUmodOppfylt"])
                                    dtRow["ModKravOppfylt"] = true;

                                dtRow["StromKrav"] = stromKravAcc * multi;
                                dtRow["StromKravUmod"] = stromKravAcc;
                                if (main.appConfig.oversiktKravStromAntall && stromKravAcc <= sStromAntall && stromKravAcc != 0)
                                    dtRow["StromKravUmodOppfylt"] = true;
                                else if (!main.appConfig.oversiktKravStromAntall && stromKravAcc <= sStromInntjen && stromKravAcc != 0)
                                    dtRow["StromKravUmodOppfylt"] = true;
                                else
                                    dtRow["StromKravUmodOppfylt"] = false;
                                stromKravAcc = (int)(stromKravAcc * multi);
                                if (main.appConfig.oversiktKravStromAntall && stromKravAcc <= sStromAntall && stromKravAcc != 0)
                                    dtRow["StromKravOppfylt"] = true;
                                else if (!main.appConfig.oversiktKravStromAntall && stromKravAcc <= sStromInntjen && stromKravAcc != 0)
                                    dtRow["StromKravOppfylt"] = true;
                                else
                                    dtRow["StromKravOppfylt"] = false;
                                if ((bool)dtRow["StromKravUmodOppfylt"])
                                    dtRow["StromKravOppfylt"] = true;

                                dtRow["RtgsaKrav"] = rtgsaKravAcc * multi;
                                dtRow["RtgsaKravUmod"] = rtgsaKravAcc;
                                if (main.appConfig.oversiktKravRtgsaAntall && rtgsaKravAcc <= sAntallTjen && rtgsaKravAcc != 0)
                                    dtRow["RtgsaKravUmodOppfylt"] = true;
                                else if (!main.appConfig.oversiktKravRtgsaAntall && (rtgsaKravAcc <= sTjenInntjen) && rtgsaKravAcc != 0)
                                    dtRow["RtgsaKravUmodOppfylt"] = true;
                                else
                                    dtRow["RtgsaKravUmodOppfylt"] = false;
                                rtgsaKravAcc = (int)(rtgsaKravAcc * multi);
                                if (main.appConfig.oversiktKravRtgsaAntall && rtgsaKravAcc <= sAntallTjen && rtgsaKravAcc != 0)
                                    dtRow["RtgsaKravOppfylt"] = true;
                                else if (!main.appConfig.oversiktKravRtgsaAntall && (rtgsaKravAcc <= sTjenInntjen) && rtgsaKravAcc != 0)
                                    dtRow["RtgsaKravOppfylt"] = true;
                                else
                                    dtRow["RtgsaKravOppfylt"] = false;
                                if ((bool)dtRow["RtgsaKravUmodOppfylt"])
                                    dtRow["RtgsaKravOppfylt"] = true;
                            }

                            dtWork.Rows.Add(dtRow);
                        }
                    }

                    if (!skipSelgere && !(main.appConfig.oversiktHideAftersales && sKat == "Aftersales")
                        && !(main.appConfig.oversiktHideKitchen && sKat == "Kjøkken"))
                    {
                        for (int i = 0; i < selgere.Length; i++)
                        {
                            // S E L G E R E
                            if (StopRankingPending())
                                return dtWork;

                            DataTable dt;
                            if (d > 0 && d <= 6 && (main.appConfig.oversiktFilterToDepartments || main.salesCodes.GetKategori(selgere[i]) == "Cross"))
                            {
                                var rows = sqlce.Select("((Varegruppe >= " + d + "00 AND Varegruppe <= " + d + "99) OR Varegruppe = 961) AND Selgerkode = '" + selgere[i] + "'"); // optionally include .ToList();
                                dt = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                            }
                            else
                            {
                                var rows = sqlce.Select("Selgerkode = '" + selgere[i] + "'");
                                dt = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                            }

                            DataRow dtRowSel = dtWork.NewRow();
                            decimal sInntjenSel = 0, sOmsetSel = 0, sOmsetExMvaSel = 0, sTjenInntjenSel = 0, sTjenOmsetSel = 0, sAntallTjenSel = 0, sAntallSalgSel = 0, sKuppvarerSel = 0, sAccessoriesAntallSel = 0, sAccessoriesInntjenSel = 0, sAccessoriesOmsetSel = 0;
                            decimal sStromAntallSel = 0, sStromInntjenSel = 0, sModAntallSel = 0, sModInntjenSel = 0, sFinansAntallSel = 0, sFinansInntjenSel = 0, sModOmsetSel = 0;
                            decimal sSnittAntallSel = 0, sSnittOmsetSel = 0, sSnittInntjenSel = 0, sAntallBilagSel = 0;

                            sAntallSalgSel += Compute(dt, "Sum(Antall)", null);
                            sOmsetSel += Compute(dt, "Sum(Salgspris)", null);
                            sOmsetExMvaSel += Compute(dt, "Sum(SalgsprisExMva)", null);
                            sInntjenSel += Compute(dt, "Sum(Btokr)", null);
                            sAntallBilagSel = dt.AsEnumerable().Select(g => g.Field<int>("Bilagsnr")).Distinct().Count();

                            var rowf = dt.Select("[Varegruppe] = 961 AND [Selgerkode] = '" + selgere[i] + "'");
                            for (int f = 0; f < rowf.Length; f++)
                            {
                                var rows2 = sqlce.Select("[Bilagsnr] = " + rowf[f]["Bilagsnr"]);
                                DataTable dtFinans = rows2.Any() ? rows2.CopyToDataTable() : sqlce.Clone();
                                dtFinans.DefaultView.Sort = "Salgspris DESC";
                                int gruppe = Convert.ToInt32(dtFinans.Rows[0]["Varegruppe"].ToString().Substring(0, 1));
                                if (gruppe == d || d > 6)
                                {
                                    sFinansAntallSel += Compute(dt, "Sum(Antall)", "[Varegruppe] = 961 AND [Selgerkode] = '" + selgere[i] + "' AND [Bilagsnr] = " + dtFinans.Rows[0]["Bilagsnr"].ToString());
                                    sFinansInntjenSel += Compute(dt, "Sum(Btokr)", "[Varegruppe] = 961 AND [Selgerkode] = '" + selgere[i] + "' AND [Bilagsnr] = " + dtFinans.Rows[0]["Bilagsnr"].ToString());
                                }
                            }

                            sStromInntjenSel += Compute(dt, "Sum(Btokr)", "([Varekode] LIKE 'ELSTROM*' OR [Varekode] LIKE 'ELRABATT*')");
                            sStromAntallSel += Compute(dt, "Sum(Antall)", "[Varekode] LIKE 'ELSTROM*'");
                            sModInntjenSel += Compute(dt, "Sum(Btokr)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                            sModOmsetSel += Compute(dt, "Sum(Salgspris)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                            sModAntallSel += Compute(dt, "Sum(Antall)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                            sKuppvarerSel += Compute(dt, "Sum(Antall)", "[Varekode] = 'KV'");

                            int[] accessoriesGrpList = main.appConfig.GetAccessorieGroups(d);
                            foreach (int ac in accessoriesGrpList)
                            {
                                sAccessoriesAntallSel += Compute(dt, "Sum(Antall)", "[Varegruppe] = " + ac);
                                sAccessoriesInntjenSel += Compute(dt, "Sum(Btokr)", "[Varegruppe] = " + ac);
                                sAccessoriesOmsetSel += Compute(dt, "Sum(Salgspris)", "[Varegruppe] = " + ac);
                            }

                            int[] mainproductsGrpList = main.appConfig.GetMainproductGroups(d);
                            foreach (int grp in mainproductsGrpList)
                            {
                                sSnittAntallSel += Compute(dt, "Sum(Antall)", "[Varegruppe] = " + grp);
                                sSnittInntjenSel += Compute(dt, "Sum(Btokr)", "[Varegruppe] = " + grp);
                                sSnittOmsetSel += Compute(dt, "Sum(Salgspris)", "[Varegruppe] = " + grp);
                            }

                            foreach (var varekode in varekoderAlle)
                            {
                                sTjenOmsetSel += Compute(dt, "Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");
                                sTjenInntjenSel += Compute(dt, "Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                                if (varekode.synlig)
                                    sAntallTjenSel += Compute(dt, "Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                            }

                            if (sInntjenSel != 0)
                            {
                                dtRowSel["Kat"] = selgere[i];
                                dtRowSel["Salg"] = sAntallSalgSel;
                                dtRowSel["Omset"] = sOmsetSel;
                                dtRowSel["Inntjen"] = sInntjenSel;
                                if (sAntallSalgSel != 0)
                                    dtRowSel["SnittOmsetAlle"] = sOmsetSel / sAntallSalgSel;
                                else
                                    dtRowSel["SnittOmsetAlle"] = 0;
                                dtRowSel["OmsetExMva"] = Math.Round(sOmsetExMvaSel, 2);
                                if (sOmsetExMvaSel != 0)
                                    dtRowSel["Prosent"] = Math.Round(sInntjenSel / sOmsetExMvaSel * 100, 2);
                                else
                                    dtRowSel["Prosent"] = 0;
                                dtRowSel["BilagAntall"] = sAntallBilagSel;
                                if (sAntallBilagSel != 0)
                                    dtRowSel["BilagSnittOmset"] = Math.Round(sOmsetSel / sAntallBilagSel, 2);
                                else
                                    dtRowSel["BilagSnittOmset"] = 0;
                                if (sAntallBilagSel != 0)
                                    dtRowSel["BilagSnittInntjen"] = Math.Round(sInntjenSel / sAntallBilagSel, 2);
                                else
                                    dtRowSel["BilagSnittInntjen"] = 0;
                                dtRowSel["AntallTjen"] = sAntallTjenSel;
                                dtRowSel["TjenOmset"] = sTjenOmsetSel;
                                dtRowSel["TjenInntjen"] = sTjenInntjenSel;
                                if (sInntjenSel != 0)
                                    dtRowSel["TjenMargin"] = Math.Round(sTjenInntjenSel / sInntjenSel * 100, 2);
                                else
                                    dtRowSel["TjenMargin"] = 0;
                                dtRowSel["StromInntjen"] = sStromInntjenSel;
                                dtRowSel["StromAntall"] = sStromAntallSel;
                                if (sInntjenSel != 0)
                                    dtRowSel["StromMargin"] = Math.Round(sStromInntjenSel / sInntjenSel * 100, 2);
                                else
                                    dtRowSel["StromMargin"] = 0;
                                dtRowSel["ModInntjen"] = sModInntjenSel;
                                dtRowSel["ModOmset"] = sModOmsetSel;
                                dtRowSel["ModAntall"] = sModAntallSel;
                                if (sOmsetExMvaSel != 0)
                                    dtRowSel["ModMargin"] = Math.Round(sModOmsetSel / sOmsetExMvaSel * 100, 2); // dtRowSel["ModMargin"] = Math.Round(sModInntjenSel / sOmsetSel * 100, 2);
                                else
                                    dtRowSel["ModMargin"] = 0;
                                dtRowSel["FinansInntjen"] = sFinansInntjenSel;
                                dtRowSel["FinansAntall"] = sFinansAntallSel;
                                if (sInntjenSel != 0)
                                    dtRowSel["FinansMargin"] = Math.Round(sFinansInntjenSel / sInntjenSel * 100, 2);
                                else
                                    dtRowSel["FinansMargin"] = 0;
                                dtRowSel["Kuppvarer"] = sKuppvarerSel;

                                dtRowSel["AccessoriesAntall"] = sAccessoriesAntallSel;
                                dtRowSel["AccessoriesInntjen"] = sAccessoriesInntjenSel;
                                dtRowSel["AccessoriesOmset"] = sAccessoriesOmsetSel;
                                    if (sInntjenSel != 0)
                                        dtRowSel["AccessoriesMargin"] = Math.Round(sAccessoriesInntjenSel / sInntjenSel * 100, 2);
                                    else
                                        dtRowSel["AccessoriesMargin"] = 0;
                                    if (sAccessoriesOmsetSel != 0)
                                        dtRowSel["AccessoriesSoB"] = Math.Round(sAccessoriesOmsetSel / sOmsetExMvaSel * 100, 2);
                                    else
                                        dtRowSel["AccessoriesSoB"] = 0;

                                dtRowSel["SnittAntall"] = sSnittAntallSel;
                                if (sSnittAntallSel != 0)
                                {
                                    dtRowSel["SnittInntjen"] = sSnittInntjenSel / sSnittAntallSel;
                                    dtRowSel["SnittOmset"] = sSnittOmsetSel / sSnittAntallSel;
                                }
                                else
                                {
                                    dtRowSel["SnittInntjen"] = 0;
                                    dtRowSel["SnittOmset"] = 0;
                                }

                                if (main.appConfig.oversiktKravVis && ((main.salesCodes.GetKategori(selgere[i]) != "Cross" && d >= 1 && d <= 6) || (d > 6)))
                                {
                                    int finansKrav = (int)(main.salesCodes.GetKrav(selgere[i], sKat, "Finans", main.appConfig.oversiktKravFinansAntall));
                                    var modKrav = (int)(main.salesCodes.GetKrav(selgere[i], sKat, "Mod", main.appConfig.oversiktKravModAntall));
                                    var stromKrav = (int)(main.salesCodes.GetKrav(selgere[i], sKat, "Strom", main.appConfig.oversiktKravStromAntall));
                                    var rtgsaKrav = (int)(main.salesCodes.GetKrav(selgere[i], sKat, "Rtgsa", main.appConfig.oversiktKravRtgsaAntall));

                                    dtRowSel["FinansKrav"] = finansKrav * multi;
                                    dtRowSel["FinansKravUmod"] = finansKrav;
                                    if (main.appConfig.oversiktKravFinansAntall && finansKrav <= sFinansAntallSel && finansKrav != 0)
                                        dtRowSel["FinansKravUmodOppfylt"] = true;
                                    else if (!main.appConfig.oversiktKravFinansAntall && finansKrav <= sFinansInntjenSel && finansKrav != 0)
                                        dtRowSel["FinansKravUmodOppfylt"] = true;
                                    else
                                        dtRowSel["FinansKravUmodOppfylt"] = false;
                                    finansKrav = (int)(finansKrav * multi);
                                    if (main.appConfig.oversiktKravFinansAntall && finansKrav <= sFinansAntallSel && finansKrav != 0)
                                        dtRowSel["FinansKravOppfylt"] = true;
                                    else if (!main.appConfig.oversiktKravFinansAntall && finansKrav <= sFinansInntjenSel && finansKrav != 0)
                                        dtRowSel["FinansKravOppfylt"] = true;
                                    else
                                        dtRowSel["FinansKravOppfylt"] = false;
                                    if ((bool)dtRowSel["FinansKravUmodOppfylt"])
                                        dtRowSel["FinansKravOppfylt"] = true;

                                    decimal modTemp = sModOmsetSel;
                                    dtRowSel["ModKrav"] = modKrav * multi;
                                    dtRowSel["ModKravUmod"] = modKrav;
                                    if (main.appConfig.oversiktKravModAntall && modKrav <= sModAntallSel && modKrav != 0)
                                        dtRowSel["ModKravUmodOppfylt"] = true;
                                    else if (!main.appConfig.oversiktKravModAntall && modKrav <= modTemp && modKrav != 0)
                                        dtRowSel["ModKravUmodOppfylt"] = true;
                                    else
                                        dtRowSel["ModKravUmodOppfylt"] = false;
                                    modKrav = (int)(modKrav * multi);
                                    if (main.appConfig.oversiktKravModAntall && modKrav <= sModAntallSel && modKrav != 0)
                                        dtRowSel["ModKravOppfylt"] = true;
                                    else if (!main.appConfig.oversiktKravModAntall && modKrav <= modTemp && modKrav != 0)
                                        dtRowSel["ModKravOppfylt"] = true;
                                    else
                                        dtRowSel["ModKravOppfylt"] = false;
                                    if ((bool)dtRowSel["ModKravUmodOppfylt"])
                                        dtRowSel["ModKravOppfylt"] = true;

                                    dtRowSel["StromKrav"] = stromKrav * multi;
                                    dtRowSel["StromKravUmod"] = stromKrav;
                                    if (main.appConfig.oversiktKravStromAntall && stromKrav <= sStromAntallSel && stromKrav != 0)
                                        dtRowSel["StromKravUmodOppfylt"] = true;
                                    else if (!main.appConfig.oversiktKravStromAntall && stromKrav <= sStromInntjenSel && stromKrav != 0)
                                        dtRowSel["StromKravUmodOppfylt"] = true;
                                    else
                                        dtRowSel["StromKravUmodOppfylt"] = false;
                                    stromKrav = (int)(stromKrav * multi);
                                    if (main.appConfig.oversiktKravStromAntall && stromKrav <= sStromAntallSel && stromKrav != 0)
                                        dtRowSel["StromKravOppfylt"] = true;
                                    else if (!main.appConfig.oversiktKravStromAntall && stromKrav <= sStromInntjenSel && stromKrav != 0)
                                        dtRowSel["StromKravOppfylt"] = true;
                                    else
                                        dtRowSel["StromKravOppfylt"] = false;
                                    if ((bool)dtRowSel["StromKravUmodOppfylt"])
                                        dtRowSel["StromKravOppfylt"] = true;

                                    dtRowSel["RtgsaKrav"] = rtgsaKrav * multi;
                                    dtRowSel["RtgsaKravUmod"] = rtgsaKrav;
                                    if (main.appConfig.oversiktKravRtgsaAntall && rtgsaKrav <= sAntallTjenSel && rtgsaKrav != 0)
                                        dtRowSel["RtgsaKravUmodOppfylt"] = true;
                                    else if (!main.appConfig.oversiktKravRtgsaAntall && (rtgsaKrav <= sTjenInntjenSel) && rtgsaKrav != 0)
                                        dtRowSel["RtgsaKravUmodOppfylt"] = true;
                                    else
                                        dtRowSel["RtgsaKravUmodOppfylt"] = false;
                                    rtgsaKrav = (int)(rtgsaKrav * multi);
                                    if (main.appConfig.oversiktKravRtgsaAntall && rtgsaKrav <= sAntallTjenSel && rtgsaKrav != 0)
                                        dtRowSel["RtgsaKravOppfylt"] = true;
                                    else if (!main.appConfig.oversiktKravRtgsaAntall && (rtgsaKrav <= sTjenInntjenSel) && rtgsaKrav != 0)
                                        dtRowSel["RtgsaKravOppfylt"] = true;
                                    else
                                        dtRowSel["RtgsaKravOppfylt"] = false;
                                    if ((bool)dtRowSel["RtgsaKravUmodOppfylt"])
                                        dtRowSel["RtgsaKravOppfylt"] = true;
                                }
                                dtWork.Rows.Add(dtRowSel);
                            }
                        }
                    }
                }

                if (!skipTotalt)
                {
                    // ------------- T O T A L T ---------------
                    DataRow dtTotalt = dtWork.NewRow();
                    decimal tInntjen = 0, tOmset = 0, tOmsetExMva = 0, tTjenInntjen = 0, tTjenOmset = 0, tAntallTjen = 0, tAntallSalg = 0, tKuppvarer = 0, tAccessoriesAntall = 0, tAccessoriesInntjen = 0, tAccessoriesOmset = 0;
                    decimal tStromAntall = 0, tStromInntjen = 0, tModAntall = 0, tModInntjen = 0, tFinansAntall = 0, tFinansInntjen = 0, tModOmset = 0;
                    decimal tSnittAntall = 0, tSnittOmset = 0, tSnittInntjen = 0, tAntallBilag = 0;

                    if (main.appConfig.importSetting.StartsWith("Full"))
                    {
                        tAntallSalg += Compute(sqlce, "Sum(Antall)", null);
                        tOmset += Compute(sqlce, "Sum(Salgspris)", null);
                        tOmsetExMva += Compute(sqlce, "Sum(SalgsprisExMva)", null);
                        tInntjen += Compute(sqlce, "Sum(Btokr)", null);
                        tAntallBilag = sqlce.AsEnumerable().Select(g => g.Field<int>("Bilagsnr")).Distinct().Count();

                        tFinansInntjen += Compute(sqlce, "Sum(Btokr)", "[Varegruppe] = 961");
                        tFinansAntall += Compute(sqlce, "Sum(Antall)", "[Varegruppe] = 961");

                        tStromInntjen += Compute(sqlce, "Sum(Btokr)", "([Varekode] LIKE 'ELSTROM*' OR [Varekode] LIKE 'ELRABATT*')");
                        tStromAntall += Compute(sqlce, "Sum(Antall)", "[Varekode] LIKE 'ELSTROM*'");
                        tModInntjen += Compute(sqlce, "Sum(Btokr)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                        tModOmset += Compute(sqlce, "Sum(Salgspris)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                        tModAntall += Compute(sqlce, "Sum(Antall)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                        tKuppvarer += Compute(sqlce, "Sum(Antall)", "[Varekode] = 'KV'");

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
                    }

                    foreach (var varekode in varekoderAlle)
                    {
                        tTjenOmset += Compute(sqlce, "Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");
                        tTjenInntjen += Compute(sqlce, "Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                        if (varekode.synlig)
                            tAntallTjen += Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                    }

                    dtTotalt["Kat"] = "TOTALT";
                    dtTotalt["Salg"] = tAntallSalg;
                    dtTotalt["Omset"] = tOmset;
                    dtTotalt["Inntjen"] = tInntjen;
                    if (tAntallSalg != 0)
                        dtTotalt["SnittOmsetAlle"] = tOmset / tAntallSalg;
                    else
                        dtTotalt["SnittOmsetAlle"] = 0;
                    dtTotalt["OmsetExMva"] = Math.Round(tOmsetExMva, 2);
                    if (tOmsetExMva != 0)
                        dtTotalt["Prosent"] = Math.Round(tInntjen / tOmsetExMva * 100, 2);
                    else
                        dtTotalt["Prosent"] = 0;
                    dtTotalt["BilagAntall"] = tAntallBilag;
                    if (tAntallBilag != 0)
                        dtTotalt["BilagSnittOmset"] = Math.Round(tOmset / tAntallBilag, 2);
                    else
                        dtTotalt["BilagSnittOmset"] = 0;
                    if (tAntallBilag != 0)
                        dtTotalt["BilagSnittInntjen"] = Math.Round(tInntjen / tAntallBilag, 2);
                    else
                        dtTotalt["BilagSnittInntjen"] = 0;
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
                    dtTotalt["ModOmset"] = tModOmset;
                    dtTotalt["ModAntall"] = tModAntall;
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
                    dtTotalt["Kuppvarer"] = tKuppvarer;

                    dtTotalt["AccessoriesAntall"] = tAccessoriesAntall;
                    dtTotalt["AccessoriesInntjen"] = tAccessoriesInntjen;
                    dtTotalt["AccessoriesOmset"] = tAccessoriesOmset;
                    if (tInntjen != 0)
                        dtTotalt["AccessoriesMargin"] = Math.Round(tAccessoriesInntjen / tInntjen * 100, 2);
                    else
                        dtTotalt["AccessoriesMargin"] = 0;
                    if (tOmsetExMva != 0)
                        dtTotalt["AccessoriesSoB"] = Math.Round(tAccessoriesOmset / tOmsetExMva * 100, 2);
                    else
                        dtTotalt["AccessoriesSoB"] = 0;

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

                    if (main.appConfig.oversiktKravVis)
                    {
                        dtTotalt["FinansKrav"] = finansKravAccTot * multi;
                        dtTotalt["FinansKravUmod"] = finansKravAccTot;
                        if (main.appConfig.oversiktKravFinansAntall && finansKravAccTot <= tFinansAntall && finansKravAccTot != 0)
                            dtTotalt["FinansKravUmodOppfylt"] = true;
                        else if (!main.appConfig.oversiktKravFinansAntall && finansKravAccTot <= tFinansInntjen && finansKravAccTot != 0)
                            dtTotalt["FinansKravUmodOppfylt"] = true;
                        else
                            dtTotalt["FinansKravUmodOppfylt"] = false;
                        finansKravAccTot = (int)(finansKravAccTot * multi);
                        if (main.appConfig.oversiktKravFinansAntall && finansKravAccTot <= tFinansAntall && finansKravAccTot != 0)
                            dtTotalt["FinansKravOppfylt"] = true;
                        else if (!main.appConfig.oversiktKravFinansAntall && finansKravAccTot <= tFinansInntjen && finansKravAccTot != 0)
                            dtTotalt["FinansKravOppfylt"] = true;
                        else
                            dtTotalt["FinansKravOppfylt"] = false;

                        decimal modTemp = tModOmset;
                        dtTotalt["ModKrav"] = modKravAccTot * multi;
                        dtTotalt["ModKravUmod"] = modKravAccTot;
                        if (main.appConfig.oversiktKravModAntall && modKravAccTot <= tModAntall && modKravAccTot != 0)
                            dtTotalt["ModKravUmodOppfylt"] = true;
                        else if (!main.appConfig.oversiktKravModAntall && modKravAccTot <= modTemp && modKravAccTot != 0)
                            dtTotalt["ModKravUmodOppfylt"] = true;
                        else
                            dtTotalt["ModKravUmodOppfylt"] = false;
                        modKravAccTot = (int)(modKravAccTot * multi);
                        if (main.appConfig.oversiktKravModAntall && modKravAccTot <= tModAntall && modKravAccTot != 0)
                            dtTotalt["ModKravOppfylt"] = true;
                        else if (!main.appConfig.oversiktKravModAntall && modKravAccTot <= modTemp && modKravAccTot != 0)
                            dtTotalt["ModKravOppfylt"] = true;
                        else
                            dtTotalt["ModKravOppfylt"] = false;

                        dtTotalt["StromKrav"] = stromKravAccTot * multi;
                        dtTotalt["StromKravUmod"] = stromKravAccTot;
                        if (main.appConfig.oversiktKravStromAntall && stromKravAccTot <= tStromAntall && stromKravAccTot != 0)
                            dtTotalt["StromKravUmodOppfylt"] = true;
                        else if (!main.appConfig.oversiktKravStromAntall && stromKravAccTot <= tStromInntjen && stromKravAccTot != 0)
                            dtTotalt["StromKravUmodOppfylt"] = true;
                        else
                            dtTotalt["StromKravUmodOppfylt"] = false;
                        stromKravAccTot = (int)(stromKravAccTot * multi);
                        if (main.appConfig.oversiktKravStromAntall && stromKravAccTot <= tStromAntall && stromKravAccTot != 0)
                            dtTotalt["StromKravOppfylt"] = true;
                        else if (!main.appConfig.oversiktKravStromAntall && stromKravAccTot <= tStromInntjen && stromKravAccTot != 0)
                            dtTotalt["StromKravOppfylt"] = true;
                        else
                            dtTotalt["StromKravOppfylt"] = false;

                        dtTotalt["RtgsaKrav"] = rtgsaKravAccTot * multi;
                        dtTotalt["RtgsaKravUmod"] = rtgsaKravAccTot;
                        if (main.appConfig.oversiktKravRtgsaAntall && rtgsaKravAccTot <= tAntallTjen && rtgsaKravAccTot != 0)
                            dtTotalt["RtgsaKravUmodOppfylt"] = true;
                        else if (!main.appConfig.oversiktKravRtgsaAntall && (rtgsaKravAccTot <= tTjenInntjen) && rtgsaKravAccTot != 0)
                            dtTotalt["RtgsaKravUmodOppfylt"] = true;
                        else
                            dtTotalt["RtgsaKravUmodOppfylt"] = false;
                        rtgsaKravAccTot = (int)(rtgsaKravAccTot * multi);
                        if (main.appConfig.oversiktKravRtgsaAntall && rtgsaKravAccTot <= tAntallTjen && rtgsaKravAccTot != 0)
                            dtTotalt["RtgsaKravOppfylt"] = true;
                        else if (!main.appConfig.oversiktKravRtgsaAntall && (rtgsaKravAccTot <= tTjenInntjen) && rtgsaKravAccTot != 0)
                            dtTotalt["RtgsaKravOppfylt"] = true;
                        else
                            dtTotalt["RtgsaKravOppfylt"] = false;
                    }
                    dtWork.Rows.Add(dtTotalt);
                }
                sqlce.Dispose();

                return dtWork;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableHtmlPrimary()
        {
            var gc = new GraphClass(main, dtTil, dtTil); // Brukes til selger aktivitet
            var doc = new List<string>();
            if (StopRankingPending())
                return doc;
            var hashId = random.Next(999, 99999);
            int hashIddRow = 0;
            string urlID = "linkm"; // Marker linker for måned-periode
            if (dtDetaljer.Rows.Count > 8)
                dt = dtDetaljer;
            else
            {
                if (main.salesCodes.Count() < 4)
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Ikke nok selgerkoder valgt for å lage en meningsfull tabell.</span><br>");
                else
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                return doc;
            }

            main.openXml.SaveDocument(dt, "Oversikt", "Del 1", dtTil, "OVERSIKT DEL 1 - " + dtTil.ToString("dddd d. MMMM yyyy", norway));

            doc.Add("<div class='toolbox hidePdf'>");
            doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
            doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
            doc.Add("</div>");

            doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
            doc.Add("<table class='tablesorter'>");
            doc.AddRange(MakeTableHeaderPrimary());
            doc.Add("<tbody>");
            string kategori = "";

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                bool kat = ErKat(dt.Rows[i]["Kat"].ToString());

                string onclickStr = "";
                string classStr = " class='Kategori' ";
                if (ErKatSpecial(dt.Rows[i]["Kat"].ToString()))
                    classStr = " class='KategoriSpecial' ";
                if (kat)
                {
                    kategori = dt.Rows[i]["Kat"].ToString();
                    hashIddRow = random.Next(9999, 99999);
                    onclickStr = " onclick='toggleRow(" + hashIddRow + ");' ";
                }
                if (!kat)
                {
                    onclickStr = "";
                    if (main.salesCodes.GetKategori(dt.Rows[i]["Kat"].ToString()) == "Cross")
                        classStr = " class='" + hashIddRow + " CrossSelger' ";
                    else
                        classStr = " class='" + hashIddRow + " Selger' ";
                }

                if (dt.Rows.Count == i + 1) // Vi er på siste row
                    doc.Add("</tbody><tfoot><tr><td class='text-cat'><a href='#" + urlID + "t" + "'>" + dt.Rows[i]["Kat"] + "</a></td>");
                else
                {
                    if (kat)
                        doc.Add("<tr" + onclickStr + classStr + "><td class='text-cat'><a href='#" + urlID + "b" + dt.Rows[i]["Kat"] + "'>" + dt.Rows[i]["Kat"] + "</a></td>");
                    else
                        doc.Add("<tr" + onclickStr + classStr + "><td class='text-cat'><a href='#" + urlID + "s" + dt.Rows[i]["Kat"] + "'>" + main.salesCodes.GetNavn(dt.Rows[i]["Kat"].ToString()) + "</a></td>");
                }

                if (main.appConfig.importSetting.StartsWith("Full"))
                {
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Omset"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Inntjen"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["Prosent"].ToString(), "Inntjen", kat) + "</td>");

                    doc.Add("<td class='numbers-finans'>" + PlusMinus(dt.Rows[i]["FinansAntall"].ToString()) + "</td>");
                    if (main.appConfig.oversiktKravVis && main.appConfig.oversiktKravFinansAntall && main.appConfig.oversiktKravFinans)
                    {
                        doc.Add("<td class='numbers-small" + ObjectToClassStr(dt.Rows[i]["FinansKravOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["FinansKrav"].ToString()) + "</td>");
                        if (main.appConfig.oversiktKravMtdShowTarget)
                            doc.Add("<td class='numbers-small" + ObjectToClassStr(dt.Rows[i]["FinansKravUmodOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["FinansKravUmod"].ToString()) + "</td>");
                    }
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["FinansInntjen"].ToString()) + "</td>");
                    if (main.appConfig.oversiktKravVis && !main.appConfig.oversiktKravFinansAntall && main.appConfig.oversiktKravFinans)
                    {
                        doc.Add("<td class='numbers-gen" + ObjectToClassStr(dt.Rows[i]["FinansKravOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["FinansKrav"].ToString()) + "</td>");
                        if (main.appConfig.oversiktKravMtdShowTarget)
                            doc.Add("<td class='numbers-gen" + ObjectToClassStr(dt.Rows[i]["FinansKravUmodOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["FinansKravUmod"].ToString()) + "</td>");
                    }
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["FinansMargin"].ToString(), "Finans", kat) + "</td>");

                    doc.Add("<td class='numbers-moderna'>" + PlusMinus(dt.Rows[i]["ModAntall"].ToString()) + "</td>");
                    if (main.appConfig.oversiktKravVis && main.appConfig.oversiktKravModAntall && main.appConfig.oversiktKravMod)
                    {
                        doc.Add("<td class='numbers-small" + ObjectToClassStr(dt.Rows[i]["ModKravOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["ModKrav"].ToString()) + "</td>");
                        if (main.appConfig.oversiktKravMtdShowTarget)
                            doc.Add("<td class='numbers-small" + ObjectToClassStr(dt.Rows[i]["ModKravUmodOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["ModKravUmod"].ToString()) + "</td>");
                    }
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["ModOmset"].ToString()) + "</td>");
                    if (main.appConfig.oversiktKravVis && !main.appConfig.oversiktKravModAntall && main.appConfig.oversiktKravMod)
                    {
                        doc.Add("<td class='numbers-gen" + ObjectToClassStr(dt.Rows[i]["ModKravOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["ModKrav"].ToString()) + "</td>");
                        if (main.appConfig.oversiktKravMtdShowTarget)
                            doc.Add("<td class='numbers-gen" + ObjectToClassStr(dt.Rows[i]["ModKravUmodOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["ModKravUmod"].ToString()) + "</td>");
                    }
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["ModMargin"].ToString(), "TA", kat) + "</td>");

                    doc.Add("<td class='numbers-strom'>" + PlusMinus(dt.Rows[i]["StromAntall"].ToString()) + "</td>");
                    if (main.appConfig.oversiktKravVis && main.appConfig.oversiktKravStromAntall && main.appConfig.oversiktKravStrom)
                    {
                        doc.Add("<td class='numbers-small" + ObjectToClassStr(dt.Rows[i]["StromKravOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["StromKrav"].ToString()) + "</td>");
                        if (main.appConfig.oversiktKravMtdShowTarget)
                            doc.Add("<td class='numbers-small" + ObjectToClassStr(dt.Rows[i]["StromKravUmodOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["StromKravUmod"].ToString()) + "</td>");
                    }
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["StromInntjen"].ToString()) + "</td>");
                    if (main.appConfig.oversiktKravVis && !main.appConfig.oversiktKravStromAntall && main.appConfig.oversiktKravStrom)
                    {
                        doc.Add("<td class='numbers-gen" + ObjectToClassStr(dt.Rows[i]["StromKravOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["StromKrav"].ToString()) + "</td>");
                        if (main.appConfig.oversiktKravMtdShowTarget)
                            doc.Add("<td class='numbers-gen" + ObjectToClassStr(dt.Rows[i]["StromKravUmodOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["StromKravUmod"].ToString()) + "</td>");
                    }
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["StromMargin"].ToString(), "Strom", kat) + "</td>");

                    doc.Add("<td class='numbers-service'>" + PlusMinus(dt.Rows[i]["AntallTjen"].ToString()) + "</td>");
                    if (main.appConfig.oversiktKravVis && main.appConfig.oversiktKravRtgsaAntall && main.appConfig.oversiktKravRtgsa)
                    {
                        doc.Add("<td class='numbers-small" + ObjectToClassStr(dt.Rows[i]["RtgsaKravOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["RtgsaKrav"].ToString()) + "</td>");
                        if (main.appConfig.oversiktKravMtdShowTarget)
                            doc.Add("<td class='numbers-small" + ObjectToClassStr(dt.Rows[i]["RtgsaKravUmodOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["RtgsaKravUmod"].ToString()) + "</td>");
                    }
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["TjenInntjen"].ToString()) + "</td>");
                    if (main.appConfig.oversiktKravVis && !main.appConfig.oversiktKravRtgsaAntall && main.appConfig.oversiktKravRtgsa)
                    {
                        doc.Add("<td class='numbers-gen" + ObjectToClassStr(dt.Rows[i]["RtgsaKravOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["RtgsaKrav"].ToString()) + "</td>");
                        if (main.appConfig.oversiktKravMtdShowTarget)
                            doc.Add("<td class='numbers-gen" + ObjectToClassStr(dt.Rows[i]["RtgsaKravUmodOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["RtgsaKravUmod"].ToString()) + "</td>");
                    }
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["TjenMargin"].ToString(), "RTG", kat) + "</td>");
                }
                if (!main.appConfig.importSetting.StartsWith("Full"))
                {
                    doc.Add("<td class='numbers-service'>" + PlusMinus(dt.Rows[i]["AntallTjen"].ToString()) + "</td>");
                    if (main.appConfig.oversiktKravVis && main.appConfig.oversiktKravRtgsaAntall && main.appConfig.oversiktKravRtgsa)
                    {
                        doc.Add("<td class='numbers-small" + ObjectToClassStr(dt.Rows[i]["RtgsaKravOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["RtgsaKrav"].ToString()) + "</td>");
                        if (main.appConfig.oversiktKravMtdShowTarget)
                            doc.Add("<td class='numbers-small" + ObjectToClassStr(dt.Rows[i]["RtgsaKravUmodOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["RtgsaKravUmod"].ToString()) + "</td>");
                    }
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["TjenOmset"].ToString()) + "</td>");
                    if (main.appConfig.oversiktKravVis && !main.appConfig.oversiktKravRtgsaAntall && main.appConfig.oversiktKravRtgsa)
                    {
                        doc.Add("<td class='numbers-gen" + ObjectToClassStr(dt.Rows[i]["RtgsaKravOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["RtgsaKrav"].ToString()) + "</td>");
                        if (main.appConfig.oversiktKravMtdShowTarget)
                            doc.Add("<td class='numbers-gen" + ObjectToClassStr(dt.Rows[i]["RtgsaKravUmodOppfylt"]) + "'>" + PlusMinus(dt.Rows[i]["RtgsaKravUmod"].ToString()) + "</td>");
                    }
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["TjenInntjen"].ToString()) + "</td>");
                }
                doc.Add("</tr>");
            }
            doc.Add("</tfoot></table></td></tr></table>");

            return doc;
        }

        public List<string> GetTableHtmlSecondary()
        {
            var doc = new List<string>();
            if (StopRankingPending())
                return doc;
            var hashId = random.Next(999, 99999);
            int hashIddRow = 0;
            string urlID = "linkm"; // Marker linker for måned-periode
            if (dtDetaljer.Rows.Count > 8)
                dt = dtDetaljer;
            else
            {
                if (main.salesCodes.Count() < 4)
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Ikke nok selgerkoder valgt for å lage en meningsfull tabell.</span><br>");
                else
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                return doc;
            }

            main.openXml.SaveDocument(dt, "Oversikt", "Del 2", dtTil, "OVERSIKT DEL 2 - " + dtTil.ToString("dddd d. MMMM yyyy", norway));

            doc.Add("<div class='toolbox hidePdf'>");
            doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
            doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
            doc.Add("</div>");

            doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
            doc.Add("<table class='tablesorter'>");
            doc.AddRange(MakeTableHeaderSecondary());
            doc.Add("<tbody>");
            string kategori = "";

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                bool kat = ErKat(dt.Rows[i]["Kat"].ToString());
                string onclickStr = "";
                string classStr = " class='Kategori' ";
                if (ErKatSpecial(dt.Rows[i]["Kat"].ToString()))
                    classStr = " class='KategoriSpecial' ";
                if (kat)
                {
                    kategori = dt.Rows[i]["Kat"].ToString();
                    hashIddRow = random.Next(9999, 99999);
                    onclickStr = " onclick='toggleRow(" + hashIddRow + ");' ";
                }
                if (!kat)
                {
                    onclickStr = "";
                    if (main.salesCodes.GetKategori(dt.Rows[i]["Kat"].ToString()) == "Cross")
                        classStr = " class='" + hashIddRow + " CrossSelger' ";
                    else
                        classStr = " class='" + hashIddRow + " Selger' ";
                }

                if (dt.Rows.Count == i + 1) // Vi er på siste row
                    doc.Add("</tbody><tfoot><tr><td class='text-cat'><a href='#" + urlID + "t" + "'>" + dt.Rows[i]["Kat"] + "</a></td>");
                else
                {
                    if (kat)
                        doc.Add("<tr" + onclickStr + classStr + "><td class='text-cat'><a href='#" + urlID + "b" + dt.Rows[i]["Kat"] + "'>" + dt.Rows[i]["Kat"] + "</a></td>");
                    else
                        doc.Add("<tr" + onclickStr + classStr + "><td class='text-cat'><a href='#" + urlID + "s" + dt.Rows[i]["Kat"] + "'>" + main.salesCodes.GetNavn(dt.Rows[i]["Kat"].ToString()) + "</a></td>");
                }

                if (main.appConfig.importSetting.StartsWith("Full"))
                {
                    //doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Salg"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Omset"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Inntjen"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["Prosent"].ToString(), "Inntjen", kat) + "</td>");

                    // Bilag
                    doc.Add("<td class='numbers-bilag'>" + PlusMinus(dt.Rows[i]["BilagAntall"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["BilagSnittOmset"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["BilagSnittInntjen"].ToString()) + "</td>");

                    // Accessories
                    doc.Add("<td class='numbers-finans'>" + PlusMinus(dt.Rows[i]["AccessoriesAntall"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["AccessoriesInntjen"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["AccessoriesMargin"].ToString(), null, kat) + "</td>");
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["AccessoriesSoB"].ToString(), null, kat) + "</td>");

                    // Snitt
                    doc.Add("<td class='numbers-hovedprod'>" + PlusMinus(dt.Rows[i]["SnittAntall"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["SnittOmset"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["SnittInntjen"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-small'>" + PlusMinus(dt.Rows[i]["Kuppvarer"].ToString()) + "</td>");
                }

                doc.Add("</tr>");
            }
            doc.Add("</tfoot></table></td></tr></table>");

            return doc;
        }

        public List<string> MakeTableHeaderPrimary()
        {
            try
            {
                List<string> doc = new List<string> { };

                doc.Add("<thead><tr>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Selgerkode</td>");
                if (main.appConfig.importSetting.StartsWith("Full"))
                {
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Omset</td>");
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Inntjen</td>");
                    doc.Add("<th class=\"{sorter: 'procent'}\" width=50 ><abbr title='Btokr. inntjen / Omset. ex. mva. alle varer'>Margin</abbr></td>");

                    doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#f5954e;'>Finans</td>");
                    if (main.appConfig.oversiktKravVis && main.appConfig.oversiktKravFinansAntall && main.appConfig.oversiktKravFinans)
                    {
                        if (main.appConfig.oversiktKravMtd)
                        {
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#f5954e;'>KravMTD</td>");
                            if (main.appConfig.oversiktKravMtdShowTarget)
                                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#f5954e;'>Krav</td>");
                        }
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#f5954e;'>Krav</td>");
                    }
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#f5954e;'>Inntjen.</td>");
                    if (main.appConfig.oversiktKravVis && !main.appConfig.oversiktKravFinansAntall && main.appConfig.oversiktKravFinans)
                    {
                        if (main.appConfig.oversiktKravMtd)
                        {
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#f5954e;'>KravMTD</td>");
                            if (main.appConfig.oversiktKravMtdShowTarget)
                                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#f5954e;'>Krav</td>");
                        }
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#f5954e;'>Krav</td>");
                    }
                    doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#f5954e;'><abbr title='Btokr. inntjen. Finans / Btokr inntjen. alle varer'>SoM</abbr></td>");

                    doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#6699ff;'>TA</td>");
                    if (main.appConfig.oversiktKravVis && main.appConfig.oversiktKravModAntall && main.appConfig.oversiktKravMod)
                    {
                        if (main.appConfig.oversiktKravMtd)
                        {
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#6699ff;'>KravMTD</td>");
                            if (main.appConfig.oversiktKravMtdShowTarget)
                                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#6699ff;'>Krav</td>");
                        }
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#6699ff;'>Krav</td>");
                    }
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#6699ff;'>Omset.</td>");
                    if (main.appConfig.oversiktKravVis && !main.appConfig.oversiktKravModAntall && main.appConfig.oversiktKravMod)
                    {
                        if (main.appConfig.oversiktKravMtd)
                        {
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#6699ff;'>KravMTD</td>");
                            if (main.appConfig.oversiktKravMtdShowTarget)
                                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#6699ff;'>Krav</td>");
                        }
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#6699ff;'>Krav</td>");
                    }
                    doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#6699ff;'><abbr title='Omset. TA / Omset. ex. mva. alle varer'>SoB</abbr></td>");
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#FAF39E;'>Strøm</td>");
                    if (main.appConfig.oversiktKravVis && main.appConfig.oversiktKravStromAntall && main.appConfig.oversiktKravStrom)
                    {
                        if (main.appConfig.oversiktKravMtd)
                        {
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#FAF39E;'>KravMTD</td>");
                            if (main.appConfig.oversiktKravMtdShowTarget)
                                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#FAF39E;'>Krav</td>");
                        }
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#FAF39E;'>Krav</td>");
                    }
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#FAF39E;'>Inntjen.</td>");
                    if (main.appConfig.oversiktKravVis && !main.appConfig.oversiktKravStromAntall && main.appConfig.oversiktKravStrom)
                    {
                        if (main.appConfig.oversiktKravMtd)
                        {
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#FAF39E;'>KravMTD</td>");
                            if (main.appConfig.oversiktKravMtdShowTarget)
                                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#FAF39E;'>Krav</td>");
                        }
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#FAF39E;'>Krav</td>");
                    }
                    doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#FAF39E;'><abbr title='Btokr. inntjen. Strøm / Btokr. inntjen. alle varer'>SoM</abbr></td>");

                    doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>RTG/SA</td>");
                    if (main.appConfig.oversiktKravVis && main.appConfig.oversiktKravRtgsaAntall && main.appConfig.oversiktKravRtgsa)
                    {
                        if (main.appConfig.oversiktKravMtd)
                        {
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>KravMTD</td>");
                            if (main.appConfig.oversiktKravMtdShowTarget)
                                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>Krav</td>");
                        }
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>Krav</td>");
                    }
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=80 style='background:#80c34a;'>Inntjen.</td>");
                    if (main.appConfig.oversiktKravVis && !main.appConfig.oversiktKravRtgsaAntall && main.appConfig.oversiktKravRtgsa)
                    {
                        if (main.appConfig.oversiktKravMtd)
                        {
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>KravMTD</td>");
                            if (main.appConfig.oversiktKravMtdShowTarget)
                                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>Krav</td>");
                        }
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>Krav</td>");
                    }
                    doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#80c34a;'><abbr title='Btokr. inntjen. Tjenester / Btokr. inntjen. alle varer'>SoM</abbr></td>");
                }
                if (!main.appConfig.importSetting.StartsWith("Full"))
                {
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>RTG/SA</td>");
                    if (main.appConfig.oversiktKravVis && main.appConfig.oversiktKravRtgsaAntall && main.appConfig.oversiktKravRtgsa)
                    {
                        if (main.appConfig.oversiktKravMtd)
                        {
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>KravMTD</td>");
                            if (main.appConfig.oversiktKravMtdShowTarget)
                                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>Krav</td>");
                        }
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>Krav</td>");
                    }
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=80 style='background:#80c34a;'>Omsetn.</td>");
                    if (main.appConfig.oversiktKravVis && !main.appConfig.oversiktKravRtgsaAntall && main.appConfig.oversiktKravRtgsa)
                    {
                        if (main.appConfig.oversiktKravMtd)
                        {
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>KravMTD</td>");
                            if (main.appConfig.oversiktKravMtdShowTarget)
                                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>Krav</td>");
                        }
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>Krav</td>");
                    }
                    doc.Add("<th class=\"{sorter: 'procent'}\" width=80 style='background:#80c34a;'>Inntjen.</td>");
                }
                doc.Add("</tr></thead>");
                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return new List<string> { };
            }
        }

        public List<string> MakeTableHeaderSecondary()
        {
            try
            {
                List<string> doc = new List<string> { };

                doc.Add("<thead><tr>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Selgerkode</td>");

                //doc.Add("<th class=\"{sorter: 'digit'}\" width=50 >Trans</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=70 >Omsetn</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=70 >Inntjen</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=60 ><abbr title='Btokr. inntjen / Omset. ex. mva. alle varer'>Margin</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#e2d687;'>Bilag</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=55 style='background:#e2d687;'>Snittomset</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=55 style='background:#e2d687;'>Snittinntjen</td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#f36565;'>Tilbehør</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#f36565;'>Inntjen</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#f36565;'><abbr title='Btokr. inntjen. Tilbehør / Btokr inntjen. alle varer'>SoM</abbr></td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#f36565;'><abbr title='Omset. Tilbehør / Omset. ex. mva. alle varer'>SoB</abbr></td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#3acd91;'>Hovedpr</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 style='background:#3acd91;'>Snittomset</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 style='background:#3acd91;'>Snittinntjen</td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#c39cdb;'><abbr title='Antall kuppvarer funnet merket KV'>KuppV</abbr></td>");

                doc.Add("</tr></thead>");
                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return new List<string> { };
            }
        }

        public DataTable ReadyTableOversikt()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Kat", typeof(string));
            dataTable.Columns.Add("Salg", typeof(int));
            dataTable.Columns.Add("Omset", typeof(decimal));
            dataTable.Columns.Add("Inntjen", typeof(decimal));
            dataTable.Columns.Add("OmsetExMva", typeof(decimal));
            dataTable.Columns.Add("Prosent", typeof(double));
            dataTable.Columns.Add("BilagAntall", typeof(int));
            dataTable.Columns.Add("BilagSnittOmset", typeof(decimal));
            dataTable.Columns.Add("BilagSnittInntjen", typeof(decimal));
            dataTable.Columns.Add("AntallTjen", typeof(int));
            dataTable.Columns.Add("TjenOmset", typeof(decimal));
            dataTable.Columns.Add("TjenInntjen", typeof(decimal));
            dataTable.Columns.Add("TjenMargin", typeof(double));
            dataTable.Columns.Add("StromInntjen", typeof(decimal));
            dataTable.Columns.Add("StromAntall", typeof(int));
            dataTable.Columns.Add("StromMargin", typeof(double));
            dataTable.Columns.Add("ModInntjen", typeof(decimal));
            dataTable.Columns.Add("ModOmset", typeof(decimal));
            dataTable.Columns.Add("ModAntall", typeof(int));
            dataTable.Columns.Add("ModMargin", typeof(double));
            dataTable.Columns.Add("FinansInntjen", typeof(decimal));
            dataTable.Columns.Add("FinansAntall", typeof(int));
            dataTable.Columns.Add("FinansMargin", typeof(double));
            dataTable.Columns.Add("FinansKrav", typeof(int));
            dataTable.Columns.Add("FinansKravUmod", typeof(int));
            dataTable.Columns.Add("FinansKravOppfylt", typeof(bool));
            dataTable.Columns.Add("FinansKravUmodOppfylt", typeof(bool));
            dataTable.Columns.Add("ModKrav", typeof(int));
            dataTable.Columns.Add("ModKravUmod", typeof(int));
            dataTable.Columns.Add("ModKravOppfylt", typeof(bool));
            dataTable.Columns.Add("ModKravUmodOppfylt", typeof(bool));
            dataTable.Columns.Add("StromKrav", typeof(int));
            dataTable.Columns.Add("StromKravUmod", typeof(int));
            dataTable.Columns.Add("StromKravOppfylt", typeof(bool));
            dataTable.Columns.Add("StromKravUmodOppfylt", typeof(bool));
            dataTable.Columns.Add("RtgsaKrav", typeof(int));
            dataTable.Columns.Add("RtgsaKravUmod", typeof(int));
            dataTable.Columns.Add("RtgsaKravOppfylt", typeof(bool));
            dataTable.Columns.Add("RtgsaKravUmodOppfylt", typeof(bool));
            dataTable.Columns.Add("Kuppvarer", typeof(int));
            dataTable.Columns.Add("AccessoriesAntall", typeof(int));
            dataTable.Columns.Add("AccessoriesInntjen", typeof(double));
            dataTable.Columns.Add("AccessoriesOmset", typeof(double));
            dataTable.Columns.Add("AccessoriesMargin", typeof(double)); // f36565
            dataTable.Columns.Add("AccessoriesSoB", typeof(double)); // f36565
            dataTable.Columns.Add("SnittAntall", typeof(int));
            dataTable.Columns.Add("SnittInntjen", typeof(double));
            dataTable.Columns.Add("SnittOmset", typeof(double));
            dataTable.Columns.Add("SnittOmsetAlle", typeof(double));

            return dataTable;
        }
    }
}