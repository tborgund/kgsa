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
    public class RankingData : Ranking
    {
        public List<VarekodeList> varekoderDataAlle;
        public List<VarekodeList> varekoderData;
        public List<VarekodeList> varekoderNettbrett;

        public List<VarekodeList> varekoderDataColumns;
        public List<VarekodeList> varekoderNettbrettColumns;
        public RankingData() { }

        public RankingData(FormMain form, DateTime dtFraArg, DateTime dtTilArg, DateTime dtPickArg)
        {
            try
            {
                this.main = form;
                dtFra = dtFraArg;
                dtTil = dtTilArg;
                dtPick = dtPickArg;
                velgerPeriode = FormMain.datoPeriodeVelger;

                this.varekoderDataAlle = main.appConfig.varekoder.Where(item => item.kategori == "Nettbrett" || item.kategori == "Data").ToList();
                this.varekoderData = main.appConfig.varekoder.Where(item => item.kategori == "Data").ToList();
                this.varekoderNettbrett = main.appConfig.varekoder.Where(item => item.kategori == "Nettbrett").ToList();

                this.varekoderDataColumns = varekoderData.Where(p => p.synlig == true).DistinctBy(p => p.alias).ToList();
                this.varekoderNettbrettColumns = varekoderNettbrett.Where(p => p.synlig == true).DistinctBy(p => p.alias).ToList();

                // Sjekk om listen har provisjon
                int prov = varekoderDataAlle.Sum(x => Convert.ToInt32(x.provSelger));
                prov += varekoderDataAlle.Sum(x => Convert.ToInt32(x.provTekniker));
                if (prov != 0)
                    provisjon = true;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private DataTable MakeTableData(string strArg)
        {
            try
            {
                DateTime dtMainFra;
                DateTime dtMainTil;

                if (strArg == "dag")
                {
                    dtMainFra = dtPick;
                    dtMainTil = dtPick;

                    var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select(GetSqlStringFor("Data") + " AND Dato = '" + dtMainTil.ToString("yyy-MM-dd") + "'");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                }
                else
                {
                    dtMainFra = dtFra;
                    dtMainTil = dtTil;

                    var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select(GetSqlStringFor("Data") + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                }
                if (strArg == "compare")
                {
                    dtMainFra = dtFra.AddYears(-1);
                    dtMainTil = dtTil.AddYears(-1);

                    var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select(GetSqlStringFor("Data") + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                }
                else if (strArg == "lastmonth")
                {
                    dtMainFra = dtFra.AddMonths(-1);
                    dtMainTil = dtTil.AddMonths(-1);

                    var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select(GetSqlStringFor("Data") + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                }
                if (velgerPeriode)
                {
                    dtMainFra = dtFra;
                    dtMainTil = dtTil;

                    sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE Avdeling = " + main.appConfig.Avdeling + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "') AND (Varegruppe >= 500 AND Varegruppe < 600)");
                    sqlce.Columns.Add("SalgsprisExMva", typeof(Double), "Salgspris / Mva");
                }

                DataTable dtWork = ReadyTable();

                if (sqlce.Rows.Count == 0)
                    return dtWork;

                string[] sk = main.sKoder.GetSelgerkoder("Data", true);
                decimal provisjon = 0, tekProvisjon = 0;
                decimal bAntallSA = 0, bAntallNB = 0, bAntallSATot = 0, bAntallNBTot = 0;

                if (main.sKoder.GetTeknikerAlle() != "")
                    foreach (var varekode in varekoderDataAlle)
                        tekProvisjon += varekode.provTekniker
                            * Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "'");

                // ------------- S E L G E R E / T E K N I K E R E ---------------
                for (int d = 0; d < sk.Length; d++)
                {
                    if (main.appConfig.rankingCompareLastyear == 1 && strArg == "compare")
                        break;
                    if (main.appConfig.rankingCompareLastmonth == 1 && strArg == "lastmonth")
                        break;
                    if (StopRankingPending())
                        return dtWork;

                    DataRow dtRow = dtWork.NewRow();
                    string provType = main.sKoder.GetProvisjon(sk[d]);

                    decimal sBtokr = 0, sAntallSA = 0, sAntallSATot = 0, sAntallNB = 0, sAntallNBTot = 0, sSalgspris = 0, sProv = 0, s531 = 0, s533 = 0, s534 = 0, sSalgsprisNormal = 0;
                    string sSelger = sk[d];

                    s531 = Compute(sqlce, "Sum(Antall)", "[Varegruppe]=531 AND [Selgerkode]='" + sk[d] + "'");
                    s533 = Compute(sqlce, "Sum(Antall)", "[Varegruppe]=533 AND [Selgerkode]='" + sk[d] + "'");
                    s534 = Compute(sqlce, "Sum(Antall)", "[Varegruppe]=534 AND [Selgerkode]='" + sk[d] + "'");

                    foreach(var varekode in varekoderDataAlle)
                    {
                        sBtokr += Compute(sqlce, "Sum(Btokr)", "[Varekode]='" + varekode.kode + "' AND [Selgerkode]='" + sk[d] + "'");
                        sSalgspris += Compute(sqlce, "Sum(Salgspris)", "[Varekode]='" + varekode.kode + "' AND [Selgerkode]='" + sk[d] + "'");
 
                        if (main.appConfig.kolRabatt)
                            sSalgsprisNormal += varekode.salgspris
                                * Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "' AND [Selgerkode]='" + sk[d] + "'");
                    }

                    foreach (var varekode in varekoderData)
                    {
                        if (!varekode.synlig)
                            continue;

                        int a = 0;
                        a = (int)Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "' AND [Selgerkode]='" + sk[d] + "'");

                        int antall = a;
                        if (dtRow["VK_" + varekode.alias] != DBNull.Value)
                            antall += Convert.ToInt32(dtRow["VK_" + varekode.alias]);

                        dtRow["VK_" + varekode.alias] = antall;
                        if (provType == "Tekniker")
                            sProv += varekode.provTekniker * a;
                        if (provType == "TeknikerAlle")
                            sProv += varekode.provSelger * a;
                        if (provType == "Selger")
                            sProv += varekode.provSelger * a;

                        sAntallSATot += a;

                        if (!varekode.inclhitrate)
                            continue;

                        sAntallSA += a;
                    }

                    foreach (var varekode in varekoderNettbrett)
                    {
                        if (!varekode.synlig)
                            continue;

                        int a = 0;
                        a = (int)Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "' AND [Selgerkode]='" + sk[d] + "'");

                        int antall = a;
                        if (dtRow["VKNB_" + varekode.alias] != DBNull.Value)
                            antall += Convert.ToInt32(dtRow["VKNB_" + varekode.alias]);

                        dtRow["VKNB_" + varekode.alias] = antall;
                        if (provType == "Tekniker")
                            sProv += varekode.provTekniker * a;
                        if (provType == "TeknikerAlle")
                            sProv += varekode.provSelger * a;
                        if (provType == "Selger")
                            sProv += varekode.provSelger * a;

                        sAntallNBTot += a;

                        if (!varekode.inclhitrate)
                            continue;

                        sAntallNB += a;
                    }

                    if (provType == "TeknikerAlle")
                        sProv += tekProvisjon;

                    if (s531 + s533 + s534 + sAntallSA + sAntallNB + sProv + sBtokr != 0) // Lagre row hvis der er salg
                    {
                        dtRow["Selgerkode"] = sSelger;
                        dtRow["HitrateSA"] = CalcHitrate(sAntallSA, s531 + s533);
                        dtRow["HitrateNB"] = CalcHitrate(sAntallNB, s534);
                        dtRow["531"] = s531;
                        dtRow["533"] = s533;
                        dtRow["534"] = s534;
                        dtRow["AntallSA"] = sAntallSA;
                        dtRow["AntallSATotalt"] = sAntallSATot;
                        dtRow["AntallNB"] = sAntallNB;
                        dtRow["AntallNBTotalt"] = sAntallNBTot;
                        dtRow["Btokr"] = sBtokr;
                        dtRow["Prov"] = sProv;
                        dtRow["Salgspris"] = sSalgspris;
                        dtRow["Rabatt"] = sSalgsprisNormal - sSalgspris;
                        provisjon += Math.Round(sProv, 0);
                        dtWork.Rows.Add(dtRow);
                    }
                    bAntallNB += sAntallNB;
                    bAntallNBTot += sAntallNBTot;
                    bAntallSA += sAntallSA;
                    bAntallSATot += sAntallSATot;
                }

                string sortMethod = "";
                switch (main.appConfig.sortIndex)
                {
                    case 0:
                        sortMethod = "";
                        break;
                    case 1:
                        sortMethod = "HitrateSA DESC";
                        break;
                    case 2:
                        sortMethod = "Selgerkode ASC";
                        break;
                    case 3:
                        sortMethod = "Prov DESC";
                        break;
                    case 4:
                        sortMethod = "Btokr DESC";
                        break;
                }

                dtWork.DefaultView.Sort = sortMethod;
                dtWork = dtWork.DefaultView.ToTable();

                // ------------- T O T A L T ---------------
                DataRow dtTotalt = dtWork.NewRow();
                DataRow dtAndre = dtWork.NewRow();
                decimal t531 = 0, t533 = 0, t534 = 0, tAntallNB = 0, tAntallNBTot = 0, a531 = 0, a533 = 0, a534 = 0, aAntallNB = 0, aAntallNBTot = 0, aAntallSA = 0, aAntallSATot = 0, tAntallSA = 0, tAntallSATot = 0;
                decimal tSalgspris = 0, aSalgspris = 0, tBtokr = 0, aBtokr = 0, tSalgsprisNormal = 0;
                string tSelger = "TOTALT", aSelger = "Andre";

                t531 = Compute(sqlce, "Sum(Antall)", "[Varegruppe]=531");
                t533 = Compute(sqlce, "Sum(Antall)", "[Varegruppe]=533");
                t534 = Compute(sqlce, "Sum(Antall)", "[Varegruppe]=534");

                foreach (var varekode in varekoderDataAlle)
                {
                    tBtokr += Compute(sqlce, "Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                    tSalgspris += Compute(sqlce, "Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");

                    if (main.appConfig.kolRabatt)
                        tSalgsprisNormal += varekode.salgspris
                            * Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                }

                foreach (var varekode in varekoderData)
                {
                    if (!varekode.synlig)
                        continue;

                    int a = 0, b = 0;
                    a = (int)Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "'");

                    int tAntall = a;
                    if (dtTotalt["VK_" + varekode.alias] != DBNull.Value)
                        tAntall += Convert.ToInt32(dtTotalt["VK_" + varekode.alias]);
                    dtTotalt["VK_" + varekode.alias] = tAntall;

                    b = (int)Compute(dtWork, "Sum(VK_" + varekode.alias + ")", null);

                    dtAndre["VK_" + varekode.alias] = tAntall - b;

                    tAntallSATot += a;

                    if (!varekode.inclhitrate)
                        continue;

                    tAntallSA += a;
                }

                foreach (var varekode in varekoderNettbrett)
                {
                    if (!varekode.synlig)
                        continue;

                    int a = 0, b = 0;
                    a = (int)Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "'");

                    int tAntall = a;
                    if (dtTotalt["VKNB_" + varekode.alias] != DBNull.Value)
                        tAntall += Convert.ToInt32(dtTotalt["VKNB_" + varekode.alias]);
                    dtTotalt["VKNB_" + varekode.alias] = tAntall;

                    b = (int)Compute(dtWork, "Sum(VKNB_" + varekode.alias + ")", null);

                    dtAndre["VKNB_" + varekode.alias] = tAntall - b;

                    tAntallNBTot += a;

                    if (!varekode.inclhitrate)
                        continue;

                    tAntallNB += a;
                }

                // ------------- A N D R E ---------------
                a531 = Compute(dtWork, "Sum([531])", null);
                a533 = Compute(dtWork, "Sum([533])", null);
                a534 = Compute(dtWork, "Sum([534])", null);
                aBtokr = Compute(dtWork, "Sum([Btokr])", null);
                aSalgspris = Compute(dtWork, "Sum([Salgspris])", null);

                foreach (var varekode in varekoderDataColumns)
                    if (varekode.inclhitrate)
                        aAntallSA += Convert.ToInt32(dtAndre["VK_" + varekode.alias]);

                foreach (var varekode in varekoderDataColumns)
                    aAntallSATot += Convert.ToInt32(dtAndre["VK_" + varekode.alias]);

                foreach (var varekode in varekoderNettbrettColumns)
                    if (varekode.inclhitrate)
                        aAntallNB += Convert.ToInt32(dtAndre["VKNB_" + varekode.alias]);

                foreach (var varekode in varekoderNettbrettColumns)
                    aAntallNBTot += Convert.ToInt32(dtAndre["VKNB_" + varekode.alias]);

                if (((t531 - a531) + (t533 - a533) + (t534 - a534) + aAntallNBTot + aAntallSATot) != 0
                    && !((main.appConfig.rankingCompareLastyear == 1 && strArg == "compare")
                    || (main.appConfig.rankingCompareLastmonth == 1 && strArg == "lastmonth")))
                {
                    dtAndre["Selgerkode"] = aSelger;
                    dtAndre["HitrateSA"] = CalcHitrate(Convert.ToDecimal(aAntallSA), a531 + a533);
                    dtAndre["HitrateNB"] = CalcHitrate(Convert.ToDecimal(aAntallNB), a534);
                    dtAndre["531"] = t531 - a531;
                    dtAndre["533"] = t533 - a533;
                    dtAndre["534"] = t534 - a534;
                    dtAndre["AntallSA"] = aAntallSA;
                    dtAndre["AntallSATotalt"] = aAntallSATot;
                    dtAndre["AntallNB"] = aAntallNB;
                    dtAndre["AntallNBTotalt"] = aAntallNBTot;
                    dtAndre["Btokr"] = tBtokr - aBtokr;
                    dtAndre["Prov"] = 0;
                    dtAndre["Salgspris"] = tSalgspris - aSalgspris;
                    dtWork.Rows.Add(dtAndre);
                }

                dtTotalt["Selgerkode"] = tSelger;
                dtTotalt["HitrateSA"] = CalcHitrate(Convert.ToDecimal(tAntallSA), t531 + t533);
                dtTotalt["HitrateNB"] = CalcHitrate(Convert.ToDecimal(tAntallNB), t534);
                dtTotalt["531"] = t531;
                dtTotalt["533"] = t533;
                dtTotalt["534"] = t534;
                dtTotalt["AntallSA"] = tAntallSA;
                dtTotalt["AntallSATotalt"] = tAntallSATot;
                dtTotalt["AntallNB"] = tAntallNB;
                dtTotalt["AntallNBTotalt"] = tAntallNBTot;
                dtTotalt["Btokr"] = tBtokr;
                dtTotalt["Prov"] = provisjon;
                dtTotalt["Salgspris"] = tSalgspris;
                dtTotalt["Rabatt"] = tSalgsprisNormal - tSalgspris;
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
                    dtMonth = MakeTableData("måned");
                    if (dtMonth.Rows.Count > 0)
                        dt = dtMonth;
                    else
                    {
                        doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                        return doc;
                    }
                }
                else if (strArg == "dag")
                {
                    urlID += "d";
                    dtDay = MakeTableData("dag");
                    if (dtDay.Rows.Count > 0)
                        dt = dtDay;
                    else
                    {
                        doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                        return doc;
                    }
                }
                else if (strArg == "compare")
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

                main.openXml.SaveDocument(dt, "Data", strArg, dtPick, strArg.ToUpper() + " - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.AddRange(MakeTableHeaderData(strArg));
                doc.Add("<tbody>");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows.Count == i + 1) // siste row
                        doc.Add("</tbody><tfoot><tr><td class='text-cat'><a href='#" + urlID + "t" + "'>" + dt.Rows[i]["Selgerkode"] + "</a></td>");
                    else if (dt.Rows[i]["Selgerkode"].ToString() == "Andre")
                        doc.Add("<tr><td class='text-cat'><a href='#" + urlID + "t" + "'>" + dt.Rows[i]["Selgerkode"] + "</a></td>");
                    else
                        doc.Add("<tr><td class='text-cat'><a href='#" + urlID + "s" + dt.Rows[i]["Selgerkode"] + "'>" + main.sKoder.GetNavn(dt.Rows[i]["Selgerkode"].ToString()) + "</a></td>");

                    if (dt.Rows[i]["Selgerkode"].ToString() != "Andre")
                        doc.Add("<td class='numbers-percent' style='" + PercentStyleData(dt.Rows[i]["HitrateSA"].ToString()) + "'>" + Percent(dt.Rows[i]["HitrateSA"].ToString()) + "</td>");
                    else
                        doc.Add("<td class='numbers-percent'>" + Percent(dt.Rows[i]["HitrateSA"].ToString()) + "</td>");

                    if (dt.Rows[i]["Selgerkode"].ToString() != "Andre")
                        doc.Add("<td class='numbers-percent' style='" + PercentStyleNett(dt.Rows[i]["HitrateNB"].ToString()) + "'>" + Percent(dt.Rows[i]["HitrateNB"].ToString()) + "</td>");
                    else
                        doc.Add("<td class='numbers-percent'>" + Percent(dt.Rows[i]["HitrateNB"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-small'>" + PlusMinus(dt.Rows[i]["531"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-small'>" + PlusMinus(dt.Rows[i]["533"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-small'>" + PlusMinus(dt.Rows[i]["534"].ToString()) + "</td>");

                    if (!main.appConfig.kolVarekoder)
                        foreach (var varekode in varekoderDataColumns)
                            if (dt.Rows.Count != i + 1)
                                doc.Add("<td class='numbers-vk'>" + PlusMinus(dt.Rows[i]["VK_" + varekode.alias].ToString()) + "</td>");
                            else
                                doc.Add("<td class='numbers-vk'><a href='#" + urlID + "v" + varekode.alias + "'>" + PlusMinus(dt.Rows[i]["VK_" + varekode.alias].ToString()) + "</a></td>");

                    doc.Add("<td class='numbers-service'>" + PlusMinus(dt.Rows[i]["AntallSATotalt"].ToString()) + "</td>");
                    if (main.appConfig.kolKravData && dt.Rows[i]["Selgerkode"].ToString() != "Andre" && strArg != "compare" && strArg != "lastmonth")
                    {
                        double k = ((int)dt.Rows[i]["531"] + (int)dt.Rows[i]["533"]) * (double)main.appConfig.kravHitrateData;
                        double Vdeficit = Math.Round((int)dt.Rows[i]["AntallSA"] - k, 0);
                        if (k > 0)
                        {
                            if (Vdeficit > 0)
                                doc.Add("<td class='numbers-small' style='color:green;'>OK</td>");
                            if (Vdeficit < 0)
                                doc.Add("<td class='numbers-small' style='color:red;'>" + Vdeficit + "</td>");
                            if (Vdeficit == 0)
                                doc.Add("<td class='numbers-small' style='color:green;'>OK</td>");
                        }
                        else
                            doc.Add("<td class='numbers-small' style='color:green;'>OK</td>");
                    }
                    if (main.appConfig.kolKravData && dt.Rows[i]["Selgerkode"].ToString() == "Andre" && strArg != "compare")
                        doc.Add("<td class='numbers-small'>" + main.appConfig.visningNull + "</td>");

                    if (!main.appConfig.kolVarekoder)
                        foreach (var varekode in varekoderNettbrettColumns)
                            if (dt.Rows.Count != i + 1)
                                doc.Add("<td class='numbers-vk'>" + PlusMinus(dt.Rows[i]["VKNB_" + varekode.alias].ToString()) + "</td>");
                            else
                                doc.Add("<td class='numbers-vk'><a href='#" + urlID + "v" + varekode.alias + "'>" + PlusMinus(dt.Rows[i]["VKNB_" + varekode.alias].ToString()) + "</a></td>");

                    doc.Add("<td class='numbers-service'>" + PlusMinus(dt.Rows[i]["AntallNBTotalt"].ToString()) + "</td>");

                    if (main.appConfig.kolKravNettbrett && dt.Rows[i]["Selgerkode"].ToString() != "Andre" && strArg != "compare")
                    {
                        double k = (int)dt.Rows[i]["534"] * (double)main.appConfig.kravHitrateNettbrett;
                        double Vdeficit = Math.Round((int)dt.Rows[i]["AntallNB"] - k, 0);
                        if (k > 0)
                        {
                            if (Vdeficit > 0)
                                doc.Add("<td class='numbers-small' style='color:green;'>OK</td>");
                            if (Vdeficit < 0)
                                doc.Add("<td class='numbers-small' style='color:red;'>" + Vdeficit + "</td>");
                            if (Vdeficit == 0)
                                doc.Add("<td class='numbers-small' style='color:green;'>OK</td>");
                        }
                        else
                            doc.Add("<td class='numbers-small' style='color:green;'>OK</td>");
                    }
                    if (main.appConfig.kolKravNettbrett && dt.Rows[i]["Selgerkode"].ToString() == "Andre" && strArg != "compare")
                        doc.Add("<td class='numbers-small'>" + main.appConfig.visningNull + "</td>");

                    if (main.appConfig.kolSalgspris)
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Salgspris"].ToString()) + "</td>");
                    if (main.appConfig.kolInntjen)
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Btokr"].ToString()) + "</td>");
                    if (main.appConfig.kolRabatt)
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Rabatt"].ToString()) + "</td>");
                    if (main.appConfig.kolProv && strArg != "compare" && strArg != "lastmonth" && provisjon)
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Prov"].ToString()) + "</td>");
                    doc.Add("</tr>");
                }
                doc.Add("</tfoot></table></td></tr></table>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableCompareHtml()
        {
            try
            {
                var doc = new List<string>();
                dtCompare = MakeTableData("compare");
                if (dtCompare.Rows.Count > 0)
                    doc.AddRange(GetTableHtml("compare"));
                else
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableCompareLastMonthHtml()
        {
            try
            {
                var doc = new List<string>();
                dtCompareLastMonth = MakeTableData("lastmonth");
                if (dtCompareLastMonth.Rows.Count > 0)
                    doc.AddRange(GetTableHtml("lastmonth"));
                else
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        private DataTable MakeTableDataAvd()
        {
            try
            {
                favoritter = FormMain.Favoritter.ToArray();
                DataTable dtWork = ReadyTable();

                int noResults = 0;
                for (int d = 0; d < favoritter.Length; d++)
                {
                    if (StopRankingPending())
                        return dtWork;

                    var rows = main.database.CallMonthTable(dtTil, favoritter[d]).Select(GetSqlStringFor("Data") + " AND (Dato >= '" + dtFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();

                    if (sqlce.Rows.Count == 0)
                    {
                        noResults++;
                        continue;
                    }

                    DataRow dtRow = dtWork.NewRow();
                    decimal s531 = 0, s533 = 0, s534 = 0, sBtokr = 0, sAntallSA = 0, sAntallSATot = 0, sAntallNB = 0, sAntallNBTot = 0, sSalgspris = 0, sSalgsprisNormal = 0;
                    object r;

                    r = sqlce.Compute("Sum(Antall)", "[Varegruppe]=531");
                    if (!DBNull.Value.Equals(r))
                        s531 = Convert.ToInt32(r);

                    r = sqlce.Compute("Sum(Antall)", "[Varegruppe]=533");
                    if (!DBNull.Value.Equals(r))
                        s533 = Convert.ToInt32(r);

                    r = sqlce.Compute("Sum(Antall)", "[Varegruppe]=534");
                    if (!DBNull.Value.Equals(r))
                        s534 = Convert.ToInt32(r);

                    foreach (var varekode in varekoderDataAlle)
                    {
                        r = sqlce.Compute("Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sBtokr += Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sSalgspris += Convert.ToDecimal(r);

                        if (main.appConfig.kolRabatt)
                        {
                            int a = 0;
                            r = sqlce.Compute("Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                            if (!DBNull.Value.Equals(r))
                                a = Convert.ToInt32(r);
                            sSalgsprisNormal += a * varekode.salgspris;
                        }
                    }

                    foreach (var varekode in varekoderData)
                    {
                        if (!varekode.synlig)
                            continue;

                        int a = 0;
                        r = sqlce.Compute("Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            a = Convert.ToInt32(r);

                        int antall = a;
                        if (dtRow["VK_" + varekode.alias] != DBNull.Value)
                            antall += Convert.ToInt32(dtRow["VK_" + varekode.alias]);

                        dtRow["VK_" + varekode.alias] = antall;

                        sAntallSATot += a;

                        if (!varekode.inclhitrate)
                            continue;

                        sAntallSA += a;
                    }

                    foreach (var varekode in varekoderNettbrett)
                    {
                        if (!varekode.synlig)
                            continue;

                        int a = 0;
                        r = sqlce.Compute("Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            a = Convert.ToInt32(r);

                        int antall = a;
                        if (dtRow["VKNB_" + varekode.alias] != DBNull.Value)
                            antall += Convert.ToInt32(dtRow["VKNB_" + varekode.alias]);

                        dtRow["VKNB_" + varekode.alias] = antall;

                        sAntallNBTot += a;

                        if (!varekode.inclhitrate)
                            continue;

                        sAntallNB += a;
                    }

                    dtRow["Selgerkode"] = favoritter[d];
                    dtRow["HitrateSA"] = CalcHitrate(sAntallSA, s531 + s533);
                    dtRow["HitrateNB"] = CalcHitrate(sAntallNB, s534);
                    dtRow["531"] = s531;
                    dtRow["533"] = s533;
                    dtRow["534"] = s534;
                    dtRow["AntallSA"] = sAntallSA;
                    dtRow["AntallSATotalt"] = sAntallSATot;
                    dtRow["AntallNB"] = sAntallNB;
                    dtRow["AntallNBTotalt"] = sAntallNBTot;
                    dtRow["Btokr"] = sBtokr;
                    dtRow["Prov"] = 0;
                    dtRow["Salgspris"] = sSalgspris;
                    dtRow["Rabatt"] = sSalgsprisNormal - sSalgspris;
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
                dtAvd = MakeTableDataAvd();
                if (dtAvd.Rows.Count > 0)
                    dt = dtAvd;
                else
                {
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                    return doc;
                }

                var hashId = random.Next(999, 99999);

                main.openXml.SaveDocument(dt, "Data", "Favoritter", dtPick, "FAVORITTER - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='OutertableNormal toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.AddRange(MakeTableHeaderFav());

                for (int i = 0; i < dtAvd.Rows.Count; i++)
                {
                    doc.Add("<tr><td class='text-cat'>" + avdeling.Get(Convert.ToInt32(dtAvd.Rows[i]["Selgerkode"])).Replace(" ", "&nbsp;") + "</td>");
                    doc.Add("<td class='numbers-percent' style='" + PercentStyleData(dtAvd.Rows[i]["HitrateSA"].ToString()) + "'>" + Percent(dtAvd.Rows[i]["HitrateSA"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent' style='" + PercentStyleNett(dtAvd.Rows[i]["HitrateNB"].ToString()) + "'>" + Percent(dtAvd.Rows[i]["HitrateNB"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-small'>" + PlusMinus(dtAvd.Rows[i]["531"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-small'>" + PlusMinus(dtAvd.Rows[i]["533"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-small'>" + PlusMinus(dtAvd.Rows[i]["534"].ToString()) + "</td>");

                    if (!main.appConfig.kolVarekoder)
                        foreach (var varekode in varekoderDataColumns)
                            doc.Add("<td class='numbers-vk'>" + PlusMinus(dtAvd.Rows[i]["VK_" + varekode.alias].ToString()) + "</td>");

                    doc.Add("<td class='numbers-service'>" + PlusMinus(dtAvd.Rows[i]["AntallSATotalt"].ToString()) + "</td>");

                    if (!main.appConfig.kolVarekoder)
                        foreach (var varekode in varekoderNettbrettColumns)
                            doc.Add("<td class='numbers-vk'>" + PlusMinus(dtAvd.Rows[i]["VKNB_" + varekode.alias].ToString()) + "</td>");

                    doc.Add("<td class='numbers-service'>" + PlusMinus(dtAvd.Rows[i]["AntallNBTotalt"].ToString()) + "</td>");

                    if (main.appConfig.kolSalgspris)
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dtAvd.Rows[i]["Salgspris"].ToString()) + "</td>");
                    if (main.appConfig.kolInntjen)
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dtAvd.Rows[i]["Btokr"].ToString()) + "</td>");
                    if (main.appConfig.kolRabatt)
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dtAvd.Rows[i]["Rabatt"].ToString()) + "</td>");
                    doc.Add("</tr>");
                }
                doc.Add("</tbody></table></td></tr></table>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        private List<string> MakeTableHeaderFav()
        {
            try
            {
                List<string> doc = new List<string> { };

                doc.Add("<thead><tr>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Avdeling</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 >SA&nbsp;%</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 >NB&nbsp;%</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 >531</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 >533</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 >534</td>");
                if (!main.appConfig.kolVarekoder)
                    foreach (var varekode in varekoderDataColumns)
                    {
                        if (!varekode.inclhitrate)
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=35 style='background:#d8dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=35 style='background:#b9dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                    }

                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 style='background:#80c34a;'>SA</td>");

                if (!main.appConfig.kolVarekoder)
                    foreach (var varekode in varekoderNettbrettColumns)
                    {
                        if (!varekode.inclhitrate)
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=35 style='background:#d8dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=35 style='background:#b9dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                    }

                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 style='background:#80c34a;'>NB</td>");

                if (main.appConfig.kolSalgspris)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#80c34a;'>Omset.</td>");
                if (main.appConfig.kolInntjen)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#80c34a;'>Inntjen.</td>");
                if (main.appConfig.kolRabatt)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#80c34a;'>Rabatt</td>");
                doc.Add("</tr></thead><tbody>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        private List<string> MakeTableHeaderData(string strArg = "")
        {
            try
            {
                List<string> doc = new List<string> { };

                doc.Add("<thead><tr>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Selgerkoder</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 >SA&nbsp;%</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 >NB&nbsp;%</td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 >531</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 >533</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 >534</td>");

                if (!main.appConfig.kolVarekoder)
                    foreach (var varekode in varekoderDataColumns)
                    {
                        if (!varekode.inclhitrate)
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=35 style='background:#d8dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=35 style='background:#b9dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                    }

                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 style='background:#80c34a;'>SA</td>");
                if (main.appConfig.kolKravData && strArg != "compare" && strArg != "lastmonth")
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=30 >Krav&nbsp;" + Math.Round(main.appConfig.kravHitrateData * 100, 0) + "%</td>");

                if (!main.appConfig.kolVarekoder)
                    foreach (var varekode in varekoderNettbrettColumns)
                    {
                        if (!varekode.inclhitrate)
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=35 style='background:#d8dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=35 style='background:#b9dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                    }

                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 style='background:#80c34a;'>NB</td>");
                if (main.appConfig.kolKravNettbrett && strArg != "compare")
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=30 >Krav&nbsp;" + Math.Round(main.appConfig.kravHitrateNettbrett * 100, 0) + "%</td>");

                if (main.appConfig.kolSalgspris)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#80c34a;'>Omset.</td>");
                if (main.appConfig.kolInntjen)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#80c34a;'>Inntjen.</td>");
                if (main.appConfig.kolRabatt)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#80c34a;'>Rabatt</td>");
                if (main.appConfig.kolProv && (strArg != "compare" && strArg != "lastmonth") && provisjon)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Prov.</td>");
                doc.Add("</tr></thead>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        public DataTable ReadyTable()
        {
            try
            {
                var dataTable = new DataTable();
                dataTable.Columns.Add("Selgerkode", typeof(string));
                dataTable.Columns.Add("HitrateSA", typeof(double));
                dataTable.Columns.Add("HitrateNB", typeof(double));
                dataTable.Columns.Add("531", typeof(int));
                dataTable.Columns.Add("533", typeof(int));
                dataTable.Columns.Add("534", typeof(int));
                foreach (var varekode in varekoderDataColumns)
                    dataTable.Columns.Add("VK_" + varekode.alias, typeof(int));
                foreach (var varekode in varekoderNettbrettColumns)
                    dataTable.Columns.Add("VKNB_" + varekode.alias, typeof(int));
                dataTable.Columns.Add("AntallSA", typeof(int));
                dataTable.Columns.Add("AntallSATotalt", typeof(int));
                dataTable.Columns.Add("AntallNB", typeof(int));
                dataTable.Columns.Add("AntallNBTotalt", typeof(int));
                dataTable.Columns.Add("Btokr", typeof(decimal));
                dataTable.Columns.Add("Prov", typeof(decimal));
                dataTable.Columns.Add("Salgspris", typeof(decimal));
                dataTable.Columns.Add("Rabatt", typeof(decimal));
                return dataTable;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }
    }
}
