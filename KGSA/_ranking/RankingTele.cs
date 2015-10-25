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

namespace KGSA
{
    public class RankingTele : Ranking
    {
        public List<VarekodeList> varekoderTele;
        public List<VarekodeList> varekoderTeleColumns;

        public RankingTele() { }

        public RankingTele(FormMain form, DateTime dtFraArg, DateTime dtTilArg, DateTime dtPickArg)
        {
            try
            {
                this.main = form;
                dtFra = dtFraArg;
                dtTil = dtTilArg;
                dtPick = dtPickArg;
                velgerPeriode = FormMain.datoPeriodeVelger;

                this.varekoderTele = main.appConfig.varekoder.Where(item => item.kategori == "Tele").ToList();
                this.varekoderTeleColumns = varekoderTele.Where(p => p.synlig == true).DistinctBy(p => p.alias).ToList();

                // Sjekk om listen har provisjon
                int prov = varekoderTele.Sum(x => Convert.ToInt32(x.provSelger));
                prov += varekoderTele.Sum(x => Convert.ToInt32(x.provTekniker));
                if (prov != 0)
                    provisjon = true;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private DataTable MakeTableTele(string strArg)
        {
            try
            {
                DateTime dtMainFra;
                DateTime dtMainTil;

                if (strArg == "dag")
                {
                    dtMainFra = dtPick;
                    dtMainTil = dtPick;

                    var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select(GetSqlStringFor("Tele") + " AND Dato = '" + dtMainTil.ToString("yyy-MM-dd") + "'");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                }
                else
                {
                    dtMainFra = dtFra;
                    dtMainTil = dtTil;

                    var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select(GetSqlStringFor("Tele") + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                }
                if (strArg == "compare")
                {
                    dtMainFra = dtFra.AddYears(-1);
                    dtMainTil = dtTil.AddYears(-1);

                    var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select(GetSqlStringFor("Tele") + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                }
                else if (strArg == "lastmonth")
                {
                    dtMainFra = dtFra.AddMonths(-1);
                    dtMainTil = dtTil.AddMonths(-1);

                    var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select(GetSqlStringFor("Tele") + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
                }
                if (velgerPeriode)
                {
                    dtMainFra = dtFra;
                    dtMainTil = dtTil;

                    sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE Avdeling = " + main.appConfig.Avdeling + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "') AND (Varegruppe >= 400 AND Varegruppe < 500)");
                    sqlce.Columns.Add("SalgsprisExMva", typeof(Double), "Salgspris / Mva");
                }

                DataTable dtWork = ReadyTableTele();

                if (sqlce.Rows.Count == 0)
                    return dtWork;

                string[] sk = main.salesCodes.GetSalesCodes("Tele", true);
                decimal provisjon = 0, tekProvisjon = 0;

                if (!String.IsNullOrEmpty(main.salesCodes.GetTeknikerAlle()))
                    foreach (var varekode in varekoderTele)
                        tekProvisjon += varekode.provTekniker
                            * Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "'");

                // ------------- S E L G E R E / T E K N I K E R E ---------------
                for (int d = 0; d < sk.Length; d++)
                {
                    if ((main.appConfig.rankingCompareLastyear == 1 && strArg == "compare") || (main.appConfig.rankingCompareLastmonth == 1 && strArg == "lastmonth"))
                        break;
                    if (StopRankingPending())
                        return dtWork;

                    DataRow dtRow = dtWork.NewRow();
                    string provType = main.salesCodes.GetProvisjon(sk[d]);

                    decimal sSalgspris = 0, sBtokr = 0, sProv = 0, s431 = 0, sAntall = 0, sAntallTot = 0, sSalgsprisNormal = 0;
                    string sSelger = sk[d];

                    s431 = Compute(sqlce, "Sum(Antall)", "[Varegruppe]=431 AND [Selgerkode]='" + sk[d] + "'");

                    foreach (var varekode in varekoderTele)
                    {
                        sBtokr += Compute(sqlce, "Sum(Btokr)", "[Varekode]='" + varekode.kode + "' AND [Selgerkode]='" + sk[d] + "'");
                        sSalgspris += Compute(sqlce, "Sum(Salgspris)", "[Varekode]='" + varekode.kode + "' AND [Selgerkode]='" + sk[d] + "'");

                        if (main.appConfig.kolRabatt)
                            sSalgsprisNormal += varekode.salgspris
                                * Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "' AND [Selgerkode]='" + sk[d] + "'");

                        if (!varekode.synlig)
                            continue;

                        int a = 0;
                        a = (int)Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "' AND [Selgerkode]='" + sk[d] + "'");

                        int antall = a;
                        if (dtRow["VK_" + varekode.alias] != DBNull.Value)
                            antall += Convert.ToInt32(dtRow["VK_" + varekode.alias]);

                        if (provType == "Tekniker")
                            sProv += varekode.provTekniker * a;
                        if (provType == "TeknikerAlle")
                            sProv += varekode.provSelger * a;
                        if (provType == "Selger")
                            sProv += varekode.provSelger * a;

                        dtRow["VK_" + varekode.alias] = antall;

                        sAntallTot += a;

                        if (!varekode.inclhitrate)
                            continue;

                        sAntall += a;
                    }

                    if (provType == "TeknikerAlle")
                        sProv += tekProvisjon;

                    if (s431 + sAntall + sBtokr + sProv != 0) // Lagre row hvis der er salg
                    {
                        dtRow["Selgerkode"] = sSelger;
                        dtRow["Hitrate"] = CalcHitrate(sAntall, s431);
                        dtRow["431"] = s431;
                        dtRow["Antall"] = sAntall;
                        dtRow["AntallTotalt"] = sAntallTot;
                        dtRow["Btokr"] = sBtokr;
                        dtRow["Prov"] = sProv;
                        dtRow["Salgspris"] = sSalgspris;
                        dtRow["Rabatt"] = sSalgsprisNormal - sSalgspris;
                        provisjon += sProv;
                        dtWork.Rows.Add(dtRow);
                    }
                }

                string sortMethod = "";
                switch (main.appConfig.sortIndex)
                {
                    case 0:
                        sortMethod = "";
                        break;
                    case 1:
                        sortMethod = "Hitrate DESC";
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
                decimal t431 = 0, tAntall = 0, tAntallTot = 0, a431 = 0, aAntall = 0, aAntallTot = 0;
                decimal tSalgspris = 0, tBtokr = 0, aSalgspris = 0, aBtokr = 0, tSalgsprisNormal = 0;
                string tSelger = "TOTALT", aSelger = "Andre";

                t431 = Compute(sqlce, "Sum(Antall)", "[Varegruppe]=431");

                foreach (var varekode in varekoderTele)
                {
                    tBtokr += Compute(sqlce, "Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                    tSalgspris += Compute(sqlce, "Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");

                    if (main.appConfig.kolRabatt)
                        tSalgsprisNormal += varekode.salgspris
                            * Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "'");

                    if (!varekode.synlig)
                        continue;

                    int a = 0, b = 0;
                    a = (int)Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "'");

                    int Antall = a;
                    if (dtTotalt["VK_" + varekode.alias] != DBNull.Value)
                        Antall += Convert.ToInt32(dtTotalt["VK_" + varekode.alias]);
                    dtTotalt["VK_" + varekode.alias] = Antall;

                    b = (int)Compute(dtWork, "Sum(VK_" + varekode.alias + ")", null);

                    dtAndre["VK_" + varekode.alias] = Antall - b;

                    tAntallTot += a;

                    if (!varekode.inclhitrate)
                        continue;

                    tAntall += a;
                }

                // ------------- A N D R E ---------------
                a431 = Compute(dtWork, "Sum([431])", null);
                aBtokr = Compute(dtWork, "Sum([Btokr])", null);
                aSalgspris = Compute(dtWork, "Sum([Salgspris])", null);

                foreach (var varekode in varekoderTeleColumns)
                    aAntallTot += Convert.ToInt32(dtAndre["VK_" + varekode.alias]);

                foreach (var varekode in varekoderTeleColumns)
                    if (varekode.inclhitrate)
                        aAntall += Convert.ToInt32(dtAndre["VK_" + varekode.alias]);

                if (((t431 - a431) + aAntall + aBtokr) != 0 && !((main.appConfig.rankingCompareLastyear == 1 && strArg == "compare") || (main.appConfig.rankingCompareLastmonth == 1 && strArg == "lastmonth")))
                {
                    dtAndre["Selgerkode"] = aSelger;
                    dtAndre["Hitrate"] = CalcHitrate(aAntall, a431);
                    dtAndre["431"] = t431 - a431;
                    dtAndre["Antall"] = aAntall;
                    dtAndre["AntallTotalt"] = aAntallTot;
                    dtAndre["Btokr"] = tBtokr - aBtokr;
                    dtAndre["Prov"] = 0;
                    dtAndre["Salgspris"] = tSalgspris - aSalgspris;
                    dtWork.Rows.Add(dtAndre);
                }

                dtTotalt["Selgerkode"] = tSelger;
                dtTotalt["Hitrate"] = CalcHitrate(tAntall, t431);
                dtTotalt["431"] = t431;
                dtTotalt["Antall"] = tAntall;
                dtTotalt["Antall"] = tAntallTot;
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
                Log.Unhandled(ex);
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
                    dtMonth = MakeTableTele("måned");
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
                    dtDay = MakeTableTele("dag");
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

                main.openXml.SaveDocument(dt, "Tele", strArg, dtPick, strArg.ToUpper() + " - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.AddRange(MakeTableHeaderTele(strArg));
                doc.Add("<tbody>");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows.Count == i + 1)
                        doc.Add("</tbody><tfoot><tr><td class='text-cat'><a href='#" + urlID + "t" + "'>" + dt.Rows[i]["Selgerkode"] + "</a></td>");
                    else if (dt.Rows[i]["Selgerkode"].ToString() == "Andre")
                        doc.Add("<tr><td class='text-cat'><a href='#" + urlID + "t" + "'>" + dt.Rows[i]["Selgerkode"] + "</a></td>");
                    else
                        doc.Add("<tr><td class='text-cat'><a href='#" + urlID + "s" + dt.Rows[i]["Selgerkode"] + "'>" + main.salesCodes.GetNavn(dt.Rows[i]["Selgerkode"].ToString()) + "</a></td>");

                    if (dt.Rows[i]["Selgerkode"].ToString() != "Andre")
                        doc.Add("<td class='numbers-percent' style='" + PercentStyleNett(dt.Rows[i]["Hitrate"].ToString()) + "'>" + Percent(dt.Rows[i]["Hitrate"].ToString()) + "</td>");
                    else
                        doc.Add("<td class='numbers-percent'>&nbsp;</td>");

                    doc.Add("<td class='numbers-small'>" + PlusMinus(dt.Rows[i]["431"].ToString()) + "</td>");
                    if (!main.appConfig.kolVarekoder)
                    {
                        foreach (var varekode in varekoderTeleColumns)
                            if (dt.Rows.Count != i + 1) // Vi er IKKE på siste row
                                doc.Add("<td class='numbers-vk'>" + PlusMinus(dt.Rows[i]["VK_" + varekode.alias].ToString()) + "</td>");
                            else
                                doc.Add("<td class='numbers-vk'><a href='#" + urlID + "v" + varekode.alias + "'>" + PlusMinus(dt.Rows[i]["VK_" + varekode.alias].ToString()) + "</a></td>");
                    }
                    doc.Add("<td class='numbers-service'>" + PlusMinus(dt.Rows[i]["AntallTotalt"].ToString()) + "</td>");

                    if (main.appConfig.kolKravTele && dt.Rows[i]["Selgerkode"].ToString() != "Andre" && strArg != "compare")
                    {
                        double k = ((int)dt.Rows[i]["431"]) * (double)main.appConfig.kravHitrateTele;
                        double Vdeficit = Math.Round((int)dt.Rows[i]["Antall"] - k, 0);
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
                    if (main.appConfig.kolKravTele && dt.Rows[i]["Selgerkode"].ToString() == "Andre" && strArg != "compare")
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
                Log.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableCompareHtml()
        {
            try
            {
                var doc = new List<string>();
                dtCompare = MakeTableTele("compare");
                if (dtCompare.Rows.Count > 0)
                    doc.AddRange(GetTableHtml("compare"));
                else
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableCompareLastMonthHtml()
        {
            try
            {
                var doc = new List<string>();
                dtCompareLastMonth = MakeTableTele("lastmonth");
                if (dtCompareLastMonth.Rows.Count > 0)
                    doc.AddRange(GetTableHtml("lastmonth"));
                else
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        private DataTable MakeTableTeleAvd()
        {
            try
            {
                favoritter = FormMain.Favoritter.ToArray();
                DataTable dtWork = ReadyTableTele();

                int noResults = 0;
                for (int d = 0; d < favoritter.Length; d++)
                {
                    if (StopRankingPending())
                        return dtWork;

                    var rows = main.database.CallMonthTable(dtTil, favoritter[d]).Select(GetSqlStringFor("Tele") + " AND (Dato >= '" + dtFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();

                    if (sqlce.Rows.Count == 0)
                        noResults++;

                    DataRow dtRow = dtWork.NewRow();
                    decimal s431 = 0, sAntall = 0, sAntallTot = 0, sSalgspris = 0, sBtokr = 0, sSalgsprisNormal = 0;
                    object r;

                    r = sqlce.Compute("Sum(Antall)", "[Varegruppe]=431");
                    if (!DBNull.Value.Equals(r))
                        s431 = Convert.ToInt32(r);

                    foreach (var varekode in varekoderTele)
                    {
                        r = sqlce.Compute("Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sBtokr += Convert.ToDecimal(r);

                        r = sqlce.Compute("Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            sSalgspris += Convert.ToDecimal(r);

                        if (main.appConfig.kolRabatt)
                        {
                            int c = 0;
                            r = sqlce.Compute("Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                            if (!DBNull.Value.Equals(r))
                                c = Convert.ToInt32(r);
                            sSalgsprisNormal += c * varekode.salgspris;
                        }

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

                        sAntallTot += a;

                        if (!varekode.inclhitrate)
                            continue;

                        sAntall += a;
                    }

                    dtRow["Selgerkode"] = favoritter[d];
                    dtRow["Hitrate"] = CalcHitrate(sAntall, s431);
                    dtRow["431"] = s431;
                    dtRow["Antall"] = sAntall;
                    dtRow["AntallTotalt"] = sAntallTot;
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
                Log.Unhandled(ex);
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
                dtAvd = MakeTableTeleAvd();
                if (dtAvd.Rows.Count > 0)
                    dt = dtAvd;
                else
                {
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                    return doc;
                }

                var hashId = random.Next(999, 99999);

                main.openXml.SaveDocument(dt, "Tele", "Favoritter", dtPick, "FAVORITTER - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='OutertableNormal toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.AddRange(MakeTableHeaderTeleAvd());

                for (int i = 0; i < dtAvd.Rows.Count; i++)
                {
                    doc.Add("<tr><td class='text-cat'>" + avdeling.Get(Convert.ToInt32(dtAvd.Rows[i]["Selgerkode"])).Replace(" ", "&nbsp;") + "</td>");
                    doc.Add("<td class='numbers-percent' style='" + PercentStyleNett(dtAvd.Rows[i]["Hitrate"].ToString()) + "'>" + Percent(dtAvd.Rows[i]["Hitrate"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-small'>" + PlusMinus(dtAvd.Rows[i]["431"].ToString()) + "</td>");
                    if (!main.appConfig.kolVarekoder)
                        foreach (var varekode in varekoderTeleColumns)
                            doc.Add("<td class='numbers-vk'>" + PlusMinus(dtAvd.Rows[i]["VK_" + varekode.alias].ToString()) + "</td>");

                    doc.Add("<td class='numbers-service'>" + PlusMinus(dtAvd.Rows[i]["Antall"].ToString()) + "</td>");
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
                Log.Unhandled(ex);
                return null;
            }
        }

        private List<string> MakeTableHeaderTeleAvd()
        {
            try
            {
                List<string> doc = new List<string> { };

                doc.Add("<thead><tr>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Avdeling</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=75 >&nbsp;%&nbsp;</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=55 >431</td>");
                if (!main.appConfig.kolVarekoder)
                    foreach (var varekode in varekoderTeleColumns)
                    {
                        if (!varekode.inclhitrate)
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=50 style='background:#d8dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=50 style='background:#b9dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                    }

                doc.Add("<th class=\"{sorter: 'digit'}\" width=55 style='background:#80c34a;'>Tjen.</td>");
                if (main.appConfig.kolSalgspris)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=75 style='background:#80c34a;'>Omsetn.</td>");
                if (main.appConfig.kolInntjen)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=75 style='background:#80c34a;'>Inntjen.</td>");
                if (main.appConfig.kolRabatt)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=75 style='background:#80c34a;'>Rabatt</td>");
                doc.Add("</tr></thead><tbody>");

                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        private List<string> MakeTableHeaderTele(string strArg = "")
        {
            try
            {
                List<string> doc = new List<string> { };

                doc.Add("<thead><tr>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Selgerkoder</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=75 >&nbsp;%&nbsp;</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=55 >431</td>");
                if (!main.appConfig.kolVarekoder)
                    foreach (var varekode in varekoderTeleColumns)
                    {
                        if (!varekode.inclhitrate)
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=50 style='background:#d8dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                        else 
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=50 style='background:#b9dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                    }

                doc.Add("<th class=\"{sorter: 'digit'}\" width=55 style='background:#80c34a;'>Tjen.</td>");
                if (main.appConfig.kolKravTele && strArg != "compare")
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >Krav&nbsp;" + Math.Round(main.appConfig.kravHitrateTele * 100, 0) + "%</td>");
                if (main.appConfig.kolSalgspris)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=75 style='background:#80c34a;'>Omsetn.</td>");
                if (main.appConfig.kolInntjen)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=75 style='background:#80c34a;'>Inntjen.</td>");
                if (main.appConfig.kolRabatt)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=75 style='background:#80c34a;'>Rabatt</td>");
                if (main.appConfig.kolProv && strArg != "compare" && strArg != "lastmonth" && provisjon)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=75 >Prov.</td>");

                doc.Add("</tr></thead>");

                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        private DataTable ReadyTableTele()
        {
            try
            {
                var dataTable = new DataTable();
                dataTable.Columns.Add("Selgerkode", typeof(string));
                dataTable.Columns.Add("Hitrate", typeof(decimal));
                dataTable.Columns.Add("431", typeof(int));
                foreach (var varekode in varekoderTeleColumns)
                    dataTable.Columns.Add("VK_" + varekode.alias, typeof(int));

                dataTable.Columns.Add("Antall", typeof(int));
                dataTable.Columns.Add("AntallTotalt", typeof(int));
                dataTable.Columns.Add("Btokr", typeof(decimal));
                dataTable.Columns.Add("Prov", typeof(decimal));
                dataTable.Columns.Add("Salgspris", typeof(decimal));
                dataTable.Columns.Add("Rabatt", typeof(decimal));
                return dataTable;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }
    }
}