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
    public class RankingKnowHow : Ranking
    {
        public List<VarekodeList> varekoderAlle;
        public List<VarekodeList> varekoderColumns;
        public RankingKnowHow() { }

        public RankingKnowHow(FormMain form, DateTime dtFraArg, DateTime dtTilArg, DateTime dtPickArg)
        {
            try
            {
                this.main = form;
                dtFra = dtFraArg;
                dtTil = dtTilArg;
                dtPick = dtPickArg;
                velgerPeriode = FormMain.datoPeriodeVelger;

                this.varekoderAlle = main.appConfig.varekoder.ToList();
                this.varekoderColumns = varekoderAlle.Where(p => p.synlig == true).DistinctBy(p => p.alias).ToList();
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private DataTable MakeTableKnowHow(string strArg)
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
                if (velgerPeriode)
                {
                    dtMainFra = dtFra;
                    dtMainTil = dtTil;
                    sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE Avdeling = " + main.appConfig.Avdeling + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "') AND (Varegruppe >= 200 AND Varegruppe < 600)");
                }

                DataTable dtWork = ReadyTable();

                if (sqlce.Rows.Count == 0)
                    return dtWork;

                string[] sk = main.salesCodes.GetSalesCodes("", true);
                decimal bAntallTjen = 0;

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
                    string provType = main.salesCodes.GetProvisjon(sk[d]);

                    decimal sBtokr = 0, sAntallTjen = 0, sSalgspris = 0, sProdukter = 0, sSalgsprisNormal = 0, sAntallTjenTot = 0;
                    string sSelger = sk[d];

                    sProdukter = Compute(sqlce, "Sum(Antall)",
                        "([Varegruppe]=531 OR [Varegruppe]=533 OR [Varegruppe]=534 OR [Varegruppe]=224 OR [Varegruppe]=431) AND [Selgerkode]='"
                        + sk[d] + "'");

                    foreach (var varekode in varekoderAlle)
                    {
                        sBtokr = Compute(sqlce, "Sum(Btokr)", "[Varekode]='" + varekode.kode + "' AND [Selgerkode]='" + sk[d] + "'");
                        sSalgspris = Compute(sqlce, "Sum(Salgspris)", "[Varekode]='" + varekode.kode + "' AND [Selgerkode]='" + sk[d] + "'");

                        if (main.appConfig.kolRabatt)
                            sSalgsprisNormal += Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "' AND [Selgerkode]='" + sk[d] + "'") * varekode.salgspris;

                        if (!varekode.synlig)
                            continue;

                        int b = 0;
                        b = (int)Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "' AND [Selgerkode]='" + sk[d] + "'");

                        int antall = b;
                        if (dtRow["VK_" + varekode.alias] != DBNull.Value)
                            antall += Convert.ToInt32(dtRow["VK_" + varekode.alias]);

                        dtRow["VK_" + varekode.alias] = antall;

                        sAntallTjenTot += b;

                        if (!varekode.inclhitrate)
                            continue;

                        sAntallTjen += b;
                    }

                    if (sAntallTjen + sBtokr != 0) // Lagre row hvis der er salg
                    {
                        dtRow["Selgerkode"] = sSelger;
                        dtRow["Hitrate"] = CalcHitrate(sAntallTjen, sProdukter);
                        dtRow["Produkter"] = sProdukter;
                        dtRow["AntallTjen"] = sAntallTjen;
                        dtRow["AntallTjenTotalt"] = sAntallTjenTot;
                        dtRow["Btokr"] = sBtokr;
                        dtRow["Salgspris"] = sSalgspris;
                        dtRow["Rabatt"] = sSalgsprisNormal - sSalgspris;
                        dtWork.Rows.Add(dtRow);
                    }
                    bAntallTjen += sAntallTjen;
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
                        sortMethod = "AntallTjen DESC";
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
                decimal tProdukter = 0, aProdukter = 0, aAntallTjen = 0, tAntallTjen = 0, tAntallTjenTot = 0, aAntallTjenTot = 0;
                decimal tSalgspris = 0, aSalgspris = 0, tBtokr = 0, aBtokr = 0, tSalgsprisNormal = 0;
                string tSelger = "TOTALT", aSelger = "Andre";

                tProdukter = Compute(sqlce, "Sum(Antall)", "[Varegruppe]=531 OR [Varegruppe]=533 OR [Varegruppe]=534 OR [Varegruppe]=224 OR [Varegruppe]=431");

                foreach (var varekode in varekoderAlle)
                {
                    tBtokr += Compute(sqlce, "Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                    tSalgspris += Compute(sqlce, "Sum(Salgspris)", "[Varekode]='" + varekode.kode + "'");

                    if (main.appConfig.kolRabatt)
                        tSalgsprisNormal += Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "'")
                            * varekode.salgspris;

                    if (!varekode.synlig)
                        continue;

                    int c = 0, b = 0;
                    c = (int)Compute(sqlce, "Sum(Antall)", "[Varekode]='" + varekode.kode + "'");

                    int tAntall = c;
                    if (dtTotalt["VK_" + varekode.alias] != DBNull.Value)
                        tAntall += Convert.ToInt32(dtTotalt["VK_" + varekode.alias]);
                    dtTotalt["VK_" + varekode.alias] = tAntall;

                    b = (int)Compute(dtWork, "Sum(VK_" + varekode.alias + ")", null);

                    dtAndre["VK_" + varekode.alias] = tAntall - b;

                    tAntallTjenTot += c;

                    if (!varekode.inclhitrate)
                        continue;

                    tAntallTjen += c;
                }

                // ------------- A N D R E ---------------
                aProdukter = Compute(dtWork, "Sum([Produkter])", null);
                aBtokr = Compute(dtWork, "Sum([Btokr])", null);
                aSalgspris = Compute(dtWork, "Sum([Salgspris])", null);

                foreach (var varekode in varekoderColumns)
                    aAntallTjenTot += Convert.ToInt32(dtAndre["VK_" + varekode.alias]);

                foreach (var varekode in varekoderColumns)
                    if (varekode.inclhitrate)
                        aAntallTjen += Convert.ToInt32(dtAndre["VK_" + varekode.alias]);

                if (aAntallTjenTot != 0 && !((main.appConfig.rankingCompareLastyear == 1 && strArg == "compare") || (main.appConfig.rankingCompareLastmonth == 1 && strArg == "lastmonth")))
                {
                    dtAndre["Selgerkode"] = aSelger;
                    dtAndre["Hitrate"] = CalcHitrate(Convert.ToDecimal(aAntallTjen), aProdukter);
                    dtAndre["Produkter"] = tProdukter;
                    dtAndre["AntallTjen"] = aAntallTjen;
                    dtAndre["AntallTjenTotalt"] = aAntallTjenTot;
                    dtAndre["Btokr"] = tBtokr - aBtokr;
                    dtAndre["Salgspris"] = tSalgspris - aSalgspris;
                    dtWork.Rows.Add(dtAndre);
                }

                dtTotalt["Selgerkode"] = tSelger;
                dtTotalt["Hitrate"] = CalcHitrate(Convert.ToDecimal(tAntallTjen), tProdukter);
                dtTotalt["Produkter"] = tProdukter;
                dtTotalt["AntallTjen"] = tAntallTjen;
                dtTotalt["AntallTjenTotalt"] = tAntallTjenTot;
                dtTotalt["Btokr"] = tBtokr;
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
                    dtMonth = MakeTableKnowHow("måned");
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
                    dtDay = MakeTableKnowHow("dag");
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

                main.openXml.SaveDocument(dt, "KnowHow", strArg, dtPick, strArg.ToUpper() + " - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.AddRange(MakeTableHeaderKnowHow(strArg));
                doc.Add("<tbody>");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows.Count == i + 1) // siste row
                        doc.Add("</tbody><tfoot><tr><td class='text-cat'><a href='#" + urlID + "t" + "'>" + dt.Rows[i]["Selgerkode"] + "</a></td>");
                    else if (dt.Rows[i]["Selgerkode"].ToString() == "Andre")
                        doc.Add("<tr><td class='text-cat'><a href='#" + urlID + "t" + "'>" + dt.Rows[i]["Selgerkode"] + "</a></td>");
                    else
                        doc.Add("<tr><td class='text-cat'><a href='#" + urlID + "s" + dt.Rows[i]["Selgerkode"] + "'>" + main.salesCodes.GetNavn(dt.Rows[i]["Selgerkode"].ToString()) + "</a></td>");

                    if (dt.Rows[i]["Selgerkode"].ToString() != "Andre")
                        doc.Add("<td class='numbers-percent' style='" + PercentStyleNett(dt.Rows[i]["Hitrate"].ToString()) + "'>" + Percent(dt.Rows[i]["Hitrate"].ToString()) + "</td>");
                    else
                        doc.Add("<td class='numbers-percent'>" + Percent(dt.Rows[i]["Hitrate"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-small'>" + PlusMinus(dt.Rows[i]["Produkter"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-service'>" + PlusMinus(dt.Rows[i]["AntallTjenTotalt"].ToString()) + "</td>");

                    if (!main.appConfig.kolVarekoder)
                        foreach (var varekode in varekoderColumns)
                            if (dt.Rows.Count != i + 1)
                                doc.Add("<td class='numbers-vk'>" + PlusMinus(dt.Rows[i]["VK_" + varekode.alias].ToString()) + "</td>");
                            else
                                doc.Add("<td class='numbers-vk'><a href='#" + urlID + "v" + varekode.alias + "'>" + PlusMinus(dt.Rows[i]["VK_" + varekode.alias].ToString()) + "</a></td>");

                    if (main.appConfig.kolSalgspris)
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Salgspris"].ToString()) + "</td>");
                    if (main.appConfig.kolInntjen)
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Btokr"].ToString()) + "</td>");
                    if (main.appConfig.kolRabatt)
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Rabatt"].ToString()) + "</td>");

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
                dtCompare = MakeTableKnowHow("compare");
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
                dtCompareLastMonth = MakeTableKnowHow("lastmonth");
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

        private DataTable MakeTableKnowHowAvd()
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

                    var rows = main.database.CallMonthTable(dtTil, favoritter[d]).Select("(Dato >= '" + dtFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtTil.ToString("yyy-MM-dd") + "')");
                    sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();

                    if (sqlce.Rows.Count == 0)
                        noResults++;

                    DataRow dtRow = dtWork.NewRow();
                    decimal sProdukter = 0, sBtokr = 0, sAntallTjen = 0, sSalgspris = 0, sSalgsprisNormal = 0, sAntallTjenTot = 0;
                    object r;

                    r = sqlce.Compute("Sum(Antall)", "[Varegruppe]=531 OR [Varegruppe]=533 OR [Varegruppe]=534 OR [Varegruppe]=224 OR [Varegruppe]=431");
                    if (!DBNull.Value.Equals(r))
                        sProdukter = Convert.ToInt32(r);

                    foreach (var varekode in varekoderAlle)
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

                        if (!varekode.synlig)
                            continue;

                        int b = 0;
                        r = sqlce.Compute("Sum(Antall)", "[Varekode]='" + varekode.kode + "'");
                        if (!DBNull.Value.Equals(r))
                            b = Convert.ToInt32(r);


                        int antall = b;
                        if (dtRow["VK_" + varekode.alias] != DBNull.Value)
                            antall += Convert.ToInt32(dtRow["VK_" + varekode.alias]);

                        dtRow["VK_" + varekode.alias] = antall;

                        sAntallTjenTot += b;

                        if (!varekode.inclhitrate)
                            continue;

                        sAntallTjen += b;
                    }

                    dtRow["Selgerkode"] = favoritter[d];
                    dtRow["Hitrate"] = CalcHitrate(sAntallTjen, sProdukter);
                    dtRow["Produkter"] = sProdukter;
                    dtRow["AntallTjen"] = sAntallTjen;
                    dtRow["AntallTjenTotalt"] = sAntallTjenTot;
                    dtRow["Btokr"] = sBtokr;
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
                dtAvd = MakeTableKnowHowAvd();
                if (dtAvd.Rows.Count > 0)
                    dt = dtAvd;
                else
                {
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                    return doc;
                }

                var hashId = random.Next(999, 99999);

                main.openXml.SaveDocument(dt, "KnowHow", "Favoritter", dtPick, "FAVORITTER - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

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
                    doc.Add("<td class='numbers-percent' style='" + PercentStyleNett(dtAvd.Rows[i]["Hitrate"].ToString()) + "'>" + Percent(dtAvd.Rows[i]["Hitrate"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-small'>" + PlusMinus(dtAvd.Rows[i]["Produkter"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-service'>" + PlusMinus(dtAvd.Rows[i]["AntallTjenTotalt"].ToString()) + "</td>");

                    if (!main.appConfig.kolVarekoder)
                        foreach (var varekode in varekoderColumns)
                            doc.Add("<td class='numbers-vk'>" + PlusMinus(dtAvd.Rows[i]["VK_" + varekode.alias].ToString()) + "</td>");

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

        private List<string> MakeTableHeaderFav()
        {
            try
            {
                List<string> doc = new List<string> { };

                doc.Add("<thead><tr>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Avdeling</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 >%</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 >Prod.</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 style='background:#80c34a;'>#</td>");

                if (!main.appConfig.kolVarekoder)
                    foreach (var varekode in varekoderColumns)
                    {
                        if (!varekode.inclhitrate)
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=35 style='background:#d8dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=35 style='background:#b9dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                    }

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
                Log.Unhandled(ex);
                return null;
            }
        }

        private List<string> MakeTableHeaderKnowHow(string strArg = "")
        {
            try
            {
                List<string> doc = new List<string> { };

                doc.Add("<thead><tr>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Selgerkoder</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=55 >%</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 >Prod.</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=40 style='background:#80c34a;'>#</td>");

                if (!main.appConfig.kolVarekoder)
                    foreach (var varekode in varekoderColumns)
                    {
                        if (!varekode.inclhitrate)
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=35 style='background:#d8dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                        else
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=35 style='background:#b9dc9c;padding-left:4px;padding-right:4px;'>" + Forkort(varekode.alias) + "</td>");
                    }

                if (main.appConfig.kolSalgspris)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#80c34a;'>Omset.</td>");
                if (main.appConfig.kolInntjen)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#80c34a;'>Inntjen.</td>");
                if (main.appConfig.kolRabatt)
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:#80c34a;'>Rabatt</td>");

                doc.Add("</tr></thead>");

                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        public DataTable ReadyTable()
        {
            try
            {
                var dataTable = new DataTable();
                dataTable.Columns.Add("Selgerkode", typeof(string));
                dataTable.Columns.Add("Hitrate", typeof(double));
                dataTable.Columns.Add("Produkter", typeof(int));
                dataTable.Columns.Add("AntallTjen", typeof(int));
                dataTable.Columns.Add("AntallTjenTotalt", typeof(int));
                foreach (var varekode in varekoderColumns)
                    dataTable.Columns.Add("VK_" + varekode.alias, typeof(int));
                dataTable.Columns.Add("Btokr", typeof(decimal));
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
