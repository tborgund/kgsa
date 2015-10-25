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
using System.Drawing.Imaging;

namespace KGSA
{
    public class RankingStore : Ranking
    {
        private Obsolete obsolete;
        public RankingStore() { }

        public RankingStore(FormMain form, DateTime dtFraArg, DateTime dtTilArg, DateTime dtPickArg, Obsolete ob)
        {
            try
            {
                this.main = form;
                this.obsolete = ob;
                this.dtFra = dtFraArg;
                this.dtTil = dtTilArg;
                this.dtPick = dtPickArg;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        public List<string> GetTableHtml(bool lagerTo = false)
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                var hashId = random.Next(999, 99999);

                DataTable dtObsolete = obsolete.GetHistory(dtTil, lagerTo);

                if (dtObsolete == null)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen tall for angitt avdeling.</span><br>");
                    return doc;
                }
                if (dtObsolete.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen tall for angitt avdeling.</span><br>");
                    return doc;
                }

                string lagerStr = main.appConfig.Avdeling.ToString();
                if (lagerTo)
                    lagerStr = (main.appConfig.Avdeling + 1000).ToString();

                main.openXml.SaveDocument(dt, "Obsolete", lagerStr, dtPick, lagerStr.ToUpper() + " - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                MakeTableHeaderObsolete(doc);
                doc.Add("<tbody>");

                for (int i = 0; i < dtObsolete.Rows.Count; i++)
                {
                    if (dtObsolete.Rows.Count == i + 1)
                        doc.Add("</tbody><tfoot><tr><td class='text-cat'>" + dtObsolete.Rows[i]["Kategori"] + "</td>");
                    else
                        doc.Add("<tr><td class='text-cat'>" + dtObsolete.Rows[i]["Kategori"] + "</td>");

                    doc.Add("<td class='numbers-small'>" + PlusMinus(dtObsolete.Rows[i]["Lagerantall"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dtObsolete.Rows[i]["Lagerverdi"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dtObsolete.Rows[i]["Ukuransantall"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dtObsolete.Rows[i]["Ukuransverdi"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dtObsolete.Rows[i]["Ukuransprosent"].ToString()) + "</td>");

                    doc.Add("</tr>");
                }
                doc.Add("</tfoot></table></td></tr></table>");
                return doc;
            }
            catch (Exception ex)
            {
                Log.d("Feil oppstod i GetTableHtml", ex);
                return null;
            }
        }

        public List<string> GetTableHtmlUtvikling(bool lagerTo = false)
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;

                var hashId = random.Next(999, 99999);

                DataTable dtObsoleteCompare = obsolete.GetHistory(dtFra, lagerTo);
                DataTable dtObsolete = obsolete.GetHistory(dtTil, lagerTo);

                DataTable dtUtvikling = CompareHistory(dtObsolete, dtObsoleteCompare);

                if (dtUtvikling == null)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen historisk data. Kort ned på antall dager for minimum sammenligning.</span><br>");
                    return doc;
                }
                if (dtObsoleteCompare.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen lager historikk å sammenligne mot.</span><br>");
                    return doc;
                }

                string lagerStr = main.appConfig.Avdeling.ToString();
                if (lagerTo)
                    lagerStr = (main.appConfig.Avdeling + 1000).ToString();

                main.openXml.SaveDocument(dtUtvikling, "Obsolete", "Utvikling" + lagerStr,
                    dtPick, lagerStr.ToUpper() + " - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                MakeTableHeaderObsoleteUtvikling(doc);
                doc.Add("<tbody>");

                for (int i = 0; i < dtUtvikling.Rows.Count; i++)
                {
                    if (dtUtvikling.Rows.Count == i + 1)
                        doc.Add("</tbody><tfoot><tr><td class='text-cat'>" + dtUtvikling.Rows[i]["Kategori"] + "</td>");
                    else
                        doc.Add("<tr><td class='text-cat'>" + dtUtvikling.Rows[i]["Kategori"] + "</td>");

                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dtUtvikling.Rows[i]["LagerverdiCompare"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-service'>" + PlusMinus(dtUtvikling.Rows[i]["Lagerverdi"].ToString()) + "</td>");

                    doc.Add(VarianceStyle(dtUtvikling.Rows[i]["LagerverdiVariance"].ToString()));

                    doc.Add("<td class='numbers-gen' style='border-left:2px solid #000;'>" + PlusMinus(dtUtvikling.Rows[i]["UkuransverdiCompare"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-service'>" + PlusMinus(dtUtvikling.Rows[i]["Ukuransverdi"].ToString()) + "</td>");

                    doc.Add(VarianceStyle(dtUtvikling.Rows[i]["UkuransverdiVariance"].ToString()));

                    doc.Add("<td class='numbers-percent' style='border-left:2px solid #000;'>" + PercentShare(dtUtvikling.Rows[i]["UkuransprosentCompare"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-service'>" + PercentShare(dtUtvikling.Rows[i]["Ukuransprosent"].ToString()) + "</td>");

                    doc.Add(VarianceStyle(dtUtvikling.Rows[i]["UkuransprosentVariance"].ToString(), 2));

                    doc.Add("</tr>");
                }
                doc.Add("</tfoot></table></td></tr></table>");

                return doc;
            }
            catch (Exception ex)
            {
                Log.d("Feil oppstod i GetTableHtmlUtvikling", ex);
                return null;
            }
        }

        public List<string> GetWeeklyJumpList()
        {
            try
            {
                DataTable table = main.database.tableWeekly.GetWeeklyList(main.appConfig.Avdeling);

                var doc = new List<string>();
                doc.Add("<form class='hidePdf' name='jumpSelect'><br>");
                doc.Add("<select name='menu' onChange='window.document.location.href=this.options[this.selectedIndex].value;' value='GO'>");

                doc.Add("<option value='#ukenytt=" + dtPick.ToString("dd.MM.yyyy", norway) + "'>Velg status oppdatering</option>");
                doc.Add("<option value='#ukenytt=list'>Oversikt</option>");
                foreach(DataRow dRow in table.Rows)
                {
                    
                    DateTime date = Convert.ToDateTime(dRow["Date"]);
                    string selectedStr = "";
                    if (dtPick.Date == date.Date)
                        selectedStr = " selected";

                    doc.Add("<option value='#ukenytt=" + date.ToString("dd.MM.yyyy", norway) + "' " + selectedStr + ">"
                        + date.ToShortDateString() + " - " + date.ToString("dddd") + "</option>");
                }

                doc.Add("</select>");
                doc.Add("</form>");

                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableHtmlWeeklyList()
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                var hashId = random.Next(999, 99999);

                DataTable table = main.database.tableWeekly.GetWeeklyList(main.appConfig.Avdeling);
                if (table == null || table.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen ukeannonse oppdateringer</span><br>");
                    return doc;
                }

                main.openXml.SaveDocument(table, "LagerUkeAnnonserOversikt", "Ukeannonser", dtPick,
                    "Ukeannonser lagerstatus - " + dtPick.ToString("dddd d. MMMM yyyy", norway) + " - Liste");

                int week = main.database.GetIso8601WeekOfYear(Convert.ToDateTime(table.Rows[0][TableWeekly.KEY_DATE]));
                bool weekSwitch = true;
                for (int i = 0; i < table.Rows.Count; i++ )
                {
                    if (weekSwitch)
                    {
                        doc.AddRange(AddHeaderWeeklyList(week));
                        weekSwitch = false;
                    }

                    DateTime listDate = Convert.ToDateTime(table.Rows[i]["Date"]);
                    int currentWeek = main.database.GetIso8601WeekOfYear(listDate);
                    int listNumberOfProducts = ObjectToInteger(table.Rows[i]["NumberOfProducts"]);
                    int listNoInStock = ObjectToInteger(table.Rows[i]["NoInStock"]);
                    int listNoInInetStock = ObjectToInteger(table.Rows[i]["NoInInetStock"]);

                    int listInStockMdaSda = ObjectToInteger(table.Rows[i]["InStockMdaSda"]);
                    int listTotalMdaSda = ObjectToInteger(table.Rows[i]["TotalMdaSda"]);
                    int listInStockTelecom = ObjectToInteger(table.Rows[i]["InStockTelecom"]);
                    int listTotalTelecom = ObjectToInteger(table.Rows[i]["TotalTelecom"]);
                    int listInStockAudioVideo = ObjectToInteger(table.Rows[i]["InStockAudioVideo"]);
                    int listTotalAudioVideo = ObjectToInteger(table.Rows[i]["TotalAudioVideo"]);
                    int listInStockComputer = ObjectToInteger(table.Rows[i]["InStockComputer"]);
                    int listTotalComputer = ObjectToInteger(table.Rows[i]["TotalComputer"]);

                    doc.Add("<tr>");
                    doc.Add("<td class='text-cat'><a href='#ukenytt=" + listDate.ToString("dd.MM.yyyy", norway) + "'>"
                        + listDate.ToString("dddd .dd", norway).ToUpper() + "</a></td>");

                    doc.Add("<td class='numbers-small'>" + listNoInStock + " av " + listNumberOfProducts + "</td>");
                    doc.Add("<td class='numbers-small'>" + main.tools.NumberStyle_Percent(listNoInStock, listNumberOfProducts, false, false) + "</td>");

                    doc.Add("<td class='numbers-small'>" + listNoInInetStock + " av " + listNumberOfProducts + "</td>");
                    doc.Add("<td class='numbers-small'>" + main.tools.NumberStyle_Percent(listNoInInetStock, listNumberOfProducts, false, false) + "</td>");

                    doc.Add("<td class='numbers-small'>" + listInStockMdaSda + " av " + listTotalMdaSda + "</td>");
                    doc.Add("<td class='numbers-small'>" + main.tools.NumberStyle_Percent(listInStockMdaSda, listTotalMdaSda, false, false) + "</td>");

                    doc.Add("<td class='numbers-small'>" + listInStockTelecom + " av " + listTotalTelecom + "</td>");
                    doc.Add("<td class='numbers-small'>" + main.tools.NumberStyle_Percent(listInStockTelecom, listTotalTelecom, false, false) + "</td>");

                    doc.Add("<td class='numbers-small'>" + listInStockAudioVideo + " av " + listTotalAudioVideo + "</td>");
                    doc.Add("<td class='numbers-small'>" + main.tools.NumberStyle_Percent(listInStockAudioVideo, listTotalAudioVideo, false, false) + "</td>");

                    doc.Add("<td class='numbers-small'>" + listInStockComputer + " av " + listTotalComputer + "</td>");
                    doc.Add("<td class='numbers-small'>" + main.tools.NumberStyle_Percent(listInStockComputer, listTotalComputer, false, false) + "</td>");

                    doc.Add("</tr>");

                    if (table.Rows.Count > i + 1)
                    {
                        week = main.database.GetIso8601WeekOfYear(Convert.ToDateTime(table.Rows[i + 1][TableWeekly.KEY_DATE]));
                        if (week != currentWeek)
                        {
                            doc.Add("</table></td></tr></table>");
                            doc.Add("</td></tr></table>");
                            week = currentWeek;
                            weekSwitch = true;
                        }
                    }
                    else
                    {
                        doc.Add("</table></td></tr></table>");
                        doc.Add("</td></tr></table>");
                    }
                }

                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        private List<string> AddHeaderWeeklyList(int week)
        {
            var hashId = random.Next(999, 99999);
            var doc = new List<string>();
            doc.Add("<br><table style='width:100%'><tr><td>");
            doc.Add("<h3>Ukeannonser for uke " + week + "</h3>");

            doc.Add("<div class='toolbox hidePdf'>");
            doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
            doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
            doc.Add("</div>");

            doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
            doc.Add("<table class='tablesorter'>");
            doc.Add("<thead><tr>");

            doc.Add("<th class=\"{sorter: 'text'}\" width=100 >Ukedag</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=120 >Totalt for butikk</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=40 >%</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=90 >Nettbutikk</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=40 >%</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=90 >MDA/SDA</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=40 >%</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=90 >Tele</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=40 >%</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=90 >AudioVideo</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=40 >%</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=90 >Computer</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=40 >%</td>");
            doc.Add("</tr></thead>");
            doc.Add("<tbody>");
            return doc;
        }

        public List<string> GetTableHtmlWeekly()
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                var hashId = random.Next(999, 99999);

                doc.AddRange(GetWeeklyJumpList());

                DataTable table = main.obsolete.GetWeekly(dtPick);
                if (table == null || table.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen oppdateringer på angitt dato.</span><br>");
                    return doc;
                }

                for (int d = 1; d < 6; d++)
                {
                    if (StopRankingPending())
                        return doc;

                    var rows = table.Select("Kategori = '" + d + "'");
                    DataTable tblKat = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();

                    string sKat = "";
                    if (d == 1)
                        sKat = "MDA";
                    else if (d == 2)
                        sKat = "AudioVideo";
                    else if (d == 3)
                        sKat = "SDA";
                    else if (d == 4)
                        sKat = "Tele";
                    else if (d == 5)
                        sKat = "Data";

                    main.openXml.SaveDocument(tblKat, "LagerUkeAnnonser", sKat, dtPick,
                        "Ukeannonser lagerstatus - " + dtPick.ToString("dddd d. MMMM yyyy", norway) + " - " + sKat
                        + " - Uke " + main.database.GetIso8601WeekOfYear(dtPick));

                    doc.Add("<br><table style='width:100%'><tr><td>");
                    doc.Add("<span class='Subtitle'>" + sKat + " - " + dtPick.ToString("dddd d. MMMM yyyy", norway)
                        + " - Uke " + main.database.GetIso8601WeekOfYear(dtPick) + "</span>");

                    doc.Add("<div class='toolbox hidePdf'>");
                    doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                    doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                    doc.Add("</div>");

                    doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                    doc.Add("<table class='tablesorter'>");
                    doc.Add("<thead><tr>");

                    doc.Add("<th class=\"{sorter: 'text'}\" width=120 >Varekode</td>");
                    doc.Add("<th class=\"{sorter: 'text'}\" width=230 >Varetekst</td>");
                    doc.Add("<th class=\"{sorter: 'text'}\" width=110 >Merke</td>");
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Pris</td>");
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=70 >Lager</td>");
                    doc.Add("<th class=\"{sorter: 'text'}\" width=70 >Nettlager</td>");
                    doc.Add("</tr></thead>");
                    doc.Add("<tbody>");

                    for (int i = 0; i < tblKat.Rows.Count; i++)
                    {
                        int stock = ObjectToInteger(tblKat.Rows[i][TableWeekly.KEY_PRODUCT_STOCK]);
                        int stockInternet = ObjectToInteger(tblKat.Rows[i][TableWeekly.KEY_PRODUCT_STOCK_INTERNET]);
                        string styleNegative = "", stylePositive = "";
                        if (stock == 0)
                            styleNegative = "style='background-color: #e98263 !important;'";
                        else if (stock > 0)
                            stylePositive = "style='background-color: #71ba51 !important;'";

                        doc.Add("<tr>");
                        doc.Add("<td class='text-cat'><a href='#linkmv" + tblKat.Rows[i]["ProductCode"] + "'>" + tblKat.Rows[i]["ProductCode"] + "</a></td>");
                        doc.Add("<td class='text-cat'>" + ForkortTekst(tblKat.Rows[i]["Varetekst"].ToString(), 27) + "</td>");
                        doc.Add("<td class='text-cat'>" + ForkortTekst(tblKat.Rows[i]["MerkeNavn"].ToString(), 15) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(tblKat.Rows[i][TableWeekly.KEY_PRODUCT_PRIZE_INTERNET]) + "</td>");

                        if (stock > 0)
                            doc.Add("<td class='numbers-small'" + stylePositive + ">" + stock + "</td>");
                        else
                            doc.Add("<td class='numbers-small'" + styleNegative + "'>0</td>");

                        if (stockInternet > 0)
                            doc.Add("<td class='numbers-small' style='background-color: #71ba51 !important;'>" + stockInternet + "+</td>");
                        else
                            doc.Add("<td class='numbers-small' style='background-color: #e98263 !important;'>0</td>");
                        doc.Add("</tr>");
                    }
                    doc.Add("</table></td></tr></table>");

                    doc.Add("</td></tr></table>");
                }

                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableHtmPrisguide()
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                var hashId = random.Next(999, 99999);

                doc.AddRange(GetPrisguideJumpList());

                DataTable table = main.obsolete.GetPopularPrisguideProducts(dtPick);
                if (table == null || table.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen Prisguide oppdateringer for gitt dato</span><br>");
                    return doc;
                }

                main.openXml.SaveDocument(table, "LagerPrisguide", "Prisguide.no", dtPick,
                    "De mest populære produkter på Prisguide.no - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<br><table style='width:100%'><tr><td>");
                doc.Add("<h2>De mest populære produktene på Prisguide.no - Uke " + main.database.GetIso8601WeekOfYear(dtPick) + "</h2>");

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.Add("<thead><tr>");

                doc.Add("<th class=\"{sorter: 'text'}\" width=35 >#</td>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=120 >Varekode</td>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=220 >Varetekst</td>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=120 >Merke</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=70 >Lager</td>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=70 >Nettlager</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Pris</td>");
                doc.Add("</tr></thead>");
                doc.Add("<tbody>");

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    int status = ObjectToInteger(table.Rows[i][TablePrisguide.KEY_PRISGUIDE_STATUS]);
                    int stock = ObjectToInteger(table.Rows[i][TablePrisguide.KEY_PRODUCT_STOCK]);
                    int stockInternet = ObjectToInteger(table.Rows[i][TablePrisguide.KEY_PRODUCT_STOCK_INTERNET]);
                    string tekst = table.Rows[i]["Varetekst"].ToString();

                    string styleNegative = "", stylePositive = "";
                    if (stock == 0 && status == 0)
                        styleNegative = "style='background-color: #e98263 !important;'";
                    else if (stock > 0 && status == 0)
                        stylePositive = "style='background-color: #71ba51 !important;'";

                    doc.Add("<tr><td class='numbers-small'>" + table.Rows[i]["Position"] + "</td>");

                    if (table.Rows[i]["ProductCode"].Equals(""))
                        doc.Add("<td class='text-cat'><a href='#external=http://www.prisguide.no/produkt/" + table.Rows[i]["PrisguideId"] + "'>Link</a></td>");
                    else
                        doc.Add("<td class='text-cat'><a href='#external=http://www.prisguide.no/produkt/" + table.Rows[i]["PrisguideId"] + "'>"
                            + table.Rows[i]["ProductCode"] + "</a></td>");

                    if (status == 0 && !String.IsNullOrEmpty(tekst))
                        doc.Add("<td class='text-cat'>" + ForkortTekst(tekst, 27) + "</td>");
                    else if (status == 0 && String.IsNullOrEmpty(tekst))
                        doc.Add("<td class='text-cat' style='color:#454545;text-align: center;'>Ingen produktinfo</td>");
                    else
                        doc.Add("<td class='text-cat' style='color:#454545;text-align: center;'>" + PrisguideProduct.GetStatusStatic(status) + "</td>");

                    doc.Add("<td class='text-cat'>" + ForkortTekst(table.Rows[i]["MerkeNavn"].ToString(), 15) + "</td>");

                    if (stock > 0)
                        doc.Add("<td class='numbers-small'" + stylePositive + ">" + PlusMinus(stock) + "</td>");
                    else if (status == 0 && stock == 0)
                        doc.Add("<td class='numbers-small'" + styleNegative + ">0</td>");
                    else
                        doc.Add("<td class='numbers-small'>&nbsp;</td>");

                    if (stockInternet > 0 && status == 0)
                        doc.Add("<td class='numbers-small' style='background-color: #71ba51 !important;'>" + stockInternet + "+</td>");
                    else if (stockInternet == 0 && status == 0)
                        doc.Add("<td class='numbers-small' style='background-color: #e98263 !important;'>0</td>");
                    else
                        doc.Add("<td class='numbers-small'>&nbsp;</td>");

                    doc.Add("<td class='numbers-gen'>" + PlusMinus(table.Rows[i][TablePrisguide.KEY_PRODUCT_PRIZE_INTERNET]) + "</td>");
                    doc.Add("</tr>");
                }
                doc.Add("</table></td></tr></table>");

                doc.Add("</td></tr></table>");


                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetPrisguideJumpList()
        {
            try
            {
                DataTable table = main.database.tablePrisguide.GetPrisguideList(main.appConfig.Avdeling);

                var doc = new List<string>();
                doc.Add("<form class='hidePdf' name='jumpSelect'><br>");
                doc.Add("<select name='menu' onChange='window.document.location.href=this.options[this.selectedIndex].value;' value='GO'>");

                doc.Add("<option value='#prisguide=" + dtPick.ToString("dd.MM.yyyy", norway) + "'>Velg prisguide oppdatering</option>");
                doc.Add("<option value='#prisguide=list'>Oversikt</option>");
                foreach (DataRow dRow in table.Rows)
                {
                    DateTime date = Convert.ToDateTime(dRow["Date"]);
                    string selectedStr = "";
                    if (dtPick.Date == date.Date)
                        selectedStr = " selected";

                    doc.Add("<option value='#prisguide=" + date.ToString("dd.MM.yyyy", norway) + "' " + selectedStr + ">"
                        + date.ToString("dddd d. MMMM yyyy", norway) + " - (" + ObjectToInteger(dRow["NoStock"]) + " / " + ObjectToInteger(dRow["Total"]) + "</option>");
                }

                doc.Add("</select>");
                doc.Add("</form>");

                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableHtmPrisguideList()
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                var hashId = random.Next(999, 99999);

                DataTable table = main.database.tablePrisguide.GetPrisguideList(main.appConfig.Avdeling);
                if (table == null || table.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen prisguide oppdateringer</span><br>");
                    return doc;
                }

                main.openXml.SaveDocument(table, "LagerPrisguideOversikt", "Prisguide.no", dtPick,
                    "De mest populære produkter på Prisguide.no - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<br><table style='width:100%'><tr><td>");
                doc.Add("<h3>Prisguide.no oppdateringer</h3>");

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.Add("<thead><tr>");

                doc.Add("<th class=\"{sorter: 'text'}\" width=120 >Dato</td>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=100 >Antall varekoder</td>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=100 >Butikk</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=60 >%</td>");
                doc.Add("<th class=\"{sorter: 'text'}\" width=100 >Nettlager</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" width=60 >%</td>");
                doc.Add("</tr></thead>");
                doc.Add("<tbody>");

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    DateTime listDate = Convert.ToDateTime(table.Rows[i]["Date"]);
                    int listTotal = ObjectToInteger(table.Rows[i]["Total"]);
                    int listNotInStock = ObjectToInteger(table.Rows[i]["NoStock"]);
                    int listNotInInetStock = ObjectToInteger(table.Rows[i]["NoInetStock"]);

                    doc.Add("<tr>");
                    doc.Add("<td class='text-cat'><a href='#prisguide=" + listDate.ToString("dd.MM.yyyy", norway) + "'>"
                        + listDate.ToString("dddd d. MMMM yyyy", norway) + "</a></td>");
                    doc.Add("<td class='numbers-small'>" + listTotal + "</td>");
                    doc.Add("<td class='numbers-small'>" + (listTotal - listNotInStock) + " av " + listTotal + "</td>");
                    doc.Add("<td class='numbers-small'>" + main.tools.NumberStyle_Percent(listTotal - listNotInStock, listTotal, false, false) + "</td>");
                    doc.Add("<td class='numbers-small'>" + (listTotal - listNotInInetStock) + " av " + listTotal + "</td>");
                    doc.Add("<td class='numbers-small'>" + main.tools.NumberStyle_Percent(listTotal - listNotInInetStock, listTotal, false, false) + "</td>");
                    doc.Add("</tr>");
                }
                doc.Add("</table></td></tr></table>");

                doc.Add("</td></tr></table>");

                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableHtmlReport()
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                var hashId = random.Next(999, 99999);

                DataTable table = main.database.GetSqlDataTable("SELECT * FROM tblHistory WHERE Avdeling = " + main.appConfig.Avdeling + " AND Kategori = 'TOTALT' ORDER BY Dato DESC");
                if (table == null)
                    throw new NoNullAllowedException("Tabellen returnerte null!");

                if (table.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen importeringer.</span><br>");
                    return doc;
                }

                main.openXml.SaveDocument(table, "ObsoleteImports", "Importeringer", dtPick, 
                    "Historisk lager status - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                MakeTableHeaderObsoleteReport(doc);
                doc.Add("<tbody>");

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    doc.Add("<tr><td class='text-cat'>" + Convert.ToDateTime(table.Rows[i]["Dato"]).ToShortDateString() + "</td>");
                    doc.Add("<td class='numbers-small'>" + PlusMinus(table.Rows[i]["Lagerantall"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(table.Rows[i]["Lagerverdi"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(table.Rows[i]["Ukuransantall"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(table.Rows[i]["Ukuransverdi"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent'>" + PercentShare(table.Rows[i]["Ukuransprosent"].ToString()) + "</td>");
                    doc.Add("</tr>");
                }
                doc.Add("</table></td></tr></table>");
                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        public void MakeTableHeaderObsoleteReport(List<string> doc)
        {
            doc.Add("<thead><tr>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Dato</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=70 >Varer</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Verdi</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=70 >Uk.varer</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Uk.verdi</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=70 >Ukurans %</td>");
            doc.Add("</tr></thead>");
        }

        private DataTable MakeUkurantGruppe(int varegruppe, bool mainproducts)
        {
            try
            {
                string sql = "SELECT tblUkurans.Avdeling, tblUkurans.Varekode, tblVareinfo.Varegruppe, "
                    + "tblVareinfo.Varetekst, tblUkurans.Antall, tblUkurans.Kost, tblUkurans.Dato, "
                    + "tblUkurans.UkuransVerdi, tblUkurans.UkuransProsent FROM tblUkurans "
                    + "inner join tblVareinfo on tblUkurans.Varekode = tblVareinfo.Varekode WHERE ";

                if (main.appConfig.storeShowStoreTwo && !main.appConfig.storeObsoleteFilterMainStoreOnly)
                    sql += "(tblUkurans.Avdeling = " + main.appConfig.Avdeling + " OR tblUkurans.Avdeling = " + (main.appConfig.Avdeling + 1000) + ") ";
                else
                    sql += "(tblUkurans.Avdeling = " + main.appConfig.Avdeling + ") ";

                if (mainproducts)
                {
                    sql += "AND (tblVareinfo.Varegruppe = 0 ";
                    int[] mainProducts = main.appConfig.GetMainproductGroups(varegruppe);
                    foreach (int code in mainProducts)
                        sql += "OR tblVareinfo.Varegruppe = " + code + " ";
                    sql += ") AND tblUkurans.UkuransVerdi > 0";
                }
                else
                    sql += "AND (tblVareinfo.Varegruppe >= " + varegruppe + "00 AND tblVareinfo.Varegruppe < "
                        + (varegruppe + 1) + "00) AND tblUkurans.UkuransVerdi > 0";

                sql += " AND tblUkurans.UkuransProsent >= " + main.appConfig.storeMaxAgePrizes;
                if (main.appConfig.storeObsoleteSortBy.Length > 0)
                    sql += " ORDER BY " + main.appConfig.storeObsoleteSortBy;
                if (!main.appConfig.storeObsoleteSortAsc)
                    sql += " DESC";

                DataTable sqlce = main.database.GetSqlDataTable(sql);

                if (sqlce.Rows.Count > 0)
                {
                    DataTable dt = sqlce.Clone();
                    for (int i = 0; i < sqlce.Rows.Count && i < main.appConfig.storeObsoleteFilterMax; i++)
                    {
                        DataRow dw = dt.NewRow();
                        dw["Avdeling"] = sqlce.Rows[i]["Avdeling"];
                        dw["Varekode"] = sqlce.Rows[i]["Varekode"];
                        dw["Varegruppe"] = sqlce.Rows[i]["Varegruppe"];
                        dw["Varetekst"] = sqlce.Rows[i]["Varetekst"];
                        dw["Antall"] = sqlce.Rows[i]["Antall"];
                        dw["Kost"] = sqlce.Rows[i]["Kost"];
                        dw["Dato"] = sqlce.Rows[i]["Dato"];
                        dw["UkuransVerdi"] = sqlce.Rows[i]["UkuransVerdi"];
                        dw["UkuransProsent"] = sqlce.Rows[i]["UkuransProsent"];
                        dt.Rows.Add(dw);
                    }


                    object y;
                    decimal u_btokr = 0, u_verdi = 0, u_antall = 0;
                    y = sqlce.Compute("Sum(Kost)", null);
                    if (!DBNull.Value.Equals(y))
                        u_btokr = Convert.ToDecimal(y);

                    y = sqlce.Compute("Sum(UkuransVerdi)", null);
                    if (!DBNull.Value.Equals(y))
                        u_verdi = Convert.ToDecimal(y);

                    y = sqlce.Compute("Sum(Antall)", null);
                    if (!DBNull.Value.Equals(y))
                        u_antall = Convert.ToInt32(y);

                    DataRow dr = dt.NewRow();
                    dr["Avdeling"] = main.appConfig.Avdeling;
                    dr["Varekode"] = "TOTALT";
                    dr["Varegruppe"] = 0;
                    dr["Varetekst"] = "";
                    dr["Antall"] = u_antall;
                    dr["Kost"] = u_btokr;
                    dr["Dato"] = DateTime.Now;
                    dr["UkuransVerdi"] = u_verdi;
                    if (u_btokr != 0)
                        dr["UkuransProsent"] = Math.Round(u_verdi / u_btokr, 2);
                    else
                        dr["UkuransProsent"] = 0;


                    object r;
                    decimal a_antall = 0, a_btokr = 0, a_verdi = 0;
                    r = dt.Compute("Sum(Kost)", null);
                    if (!DBNull.Value.Equals(r))
                        a_btokr = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(UkuransVerdi)", null);
                    if (!DBNull.Value.Equals(r))
                        a_verdi = Convert.ToDecimal(r);

                    r = dt.Compute("Sum(Antall)", null);
                    if (!DBNull.Value.Equals(r))
                        a_antall = Convert.ToInt32(r);

                    if ((u_antall - a_antall) > 0)
                    {
                        DataRow ar = dt.NewRow();
                        ar["Avdeling"] = main.appConfig.Avdeling;
                        ar["Varekode"] = "ANDRE";
                        ar["Varegruppe"] = 0;
                        ar["Varetekst"] = "";
                        ar["Antall"] = u_antall - a_antall;
                        ar["Kost"] = u_btokr - a_btokr;
                        ar["Dato"] = DateTime.Now;
                        ar["UkuransVerdi"] = u_verdi - a_verdi;
                        ar["UkuransProsent"] = 0;
                        dt.Rows.Add(ar);
                    }

                    dt.Rows.Add(dr);


                    return dt;
                }

                return sqlce;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableHtmlUkurantGrupper()
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;

                if (main.appConfig.dbObsoleteUpdated == FormMain.rangeMin)
                {
                    doc.Add("<br><span class='Subtitle'>Ingen varer i databasen.</span><br>");
                    return doc;
                }

                string filter = "(filter: ";
                if (main.appConfig.storeObsoleteListMainProductsOnly)
                    filter += "Hovedprodukter, ";
                if (main.appConfig.storeObsoleteFilterMainStoreOnly)
                    filter += "Hovedlager, ";
                filter += "max " + main.appConfig.storeObsoleteFilterMax + ", > " + main.appConfig.storeMaxAgePrizes + " %)";

                for (int i = 1; i < 6; i++)
                {
                    string sKat = "";
                    if (i == 1)
                        sKat = "MDA";
                    if (i == 2)
                        sKat = "Lyd & Bilde";
                    if (i == 3)
                        sKat = "SDA";
                    if (i == 4)
                        sKat = "Tele";
                    if (i == 5)
                        sKat = "Data";
                    if (i == 6)
                        sKat = "Kjøkken";

                    var hashId = random.Next(999, 99999);
                    DataTable dtObsoleteGroup = MakeUkurantGruppe(i, main.appConfig.storeObsoleteListMainProductsOnly);

                    if (dtObsoleteGroup.Rows.Count == 0)
                    {
                        doc.Add("<br><span class='Subtitle'>Fant ingen vareoppføringer for " + sKat + " basert på filter.</span><br>");
                        continue;
                    }

                    if (dtObsoleteGroup != null)
                    {
                        if (dtObsoleteGroup.Rows.Count == 0)
                            continue;

                        main.openXml.SaveDocument(dt, "ObsoleteList", sKat, dtPick, "LagerUkurans " + sKat.ToUpper() + " - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                        doc.Add("<br><table style='width:100%'><tr><td>");
                        doc.Add("<h2>" + sKat + " utgåtte varer: " + filter + "</h2>");

                        doc.Add("<div class='toolbox hidePdf'>");
                        doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                        doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                        doc.Add("</div>");

                        doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                        doc.Add("<table class='tablesorter'>");
                        MakeTableHeaderObsoleteUkurantListe(doc);
                        doc.Add("<tbody>");

                        for (int d = 0; d < dtObsoleteGroup.Rows.Count; d++)
                        {
                            if (dtObsoleteGroup.Rows[d]["Varekode"].ToString() == "ANDRE")
                                doc.Add("<tr><td class='text-cat'>Mer..</td>");
                            else if (dtObsoleteGroup.Rows.Count == d + 1)
                                doc.Add("</tbody><tfoot><tr><td class='text-cat'>" + dtObsoleteGroup.Rows[d]["Varekode"] + "</td>");
                            else
                                doc.Add("<tr><td class='text-cat" + SortStringStyle("Varekode") + "'>" + dtObsoleteGroup.Rows[d]["Varekode"] + "</td>");

                            doc.Add("<td class='text-cat'>" + dtObsoleteGroup.Rows[d]["Varetekst"].ToString() + "</td>");
                            doc.Add("<td class='numbers-small'>" + PlusMinus(dtObsoleteGroup.Rows[d]["Antall"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-gen" + SortStringStyle("Kost") + "'>" + PlusMinus(dtObsoleteGroup.Rows[d]["Kost"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-gen" + SortStringStyle("UkuransVerdi") + "'>" + PlusMinus(dtObsoleteGroup.Rows[d]["UkuransVerdi"].ToString()) + "</td>");
                            if (dtObsoleteGroup.Rows.Count == d + 1)
                                doc.Add("<td class='numbers-percent'></td>");
                            else
                                doc.Add("<td class='numbers-percent" + SortStringStyle("UkuransProsent") + "'>" + PercentShare(dtObsoleteGroup.Rows[d]["UkuransProsent"].ToString()) + "</td>");
                            if (dtObsoleteGroup.Rows[d]["Varekode"].ToString() == "ANDRE")
                                doc.Add("<td class='numbers-small'></td>");
                            else if (dtObsoleteGroup.Rows.Count == d + 1)
                                doc.Add("<td class='numbers-small'></td>");
                            else
                                doc.Add("<td class='numbers-small'>" + (main.appConfig.dbObsoleteUpdated - Convert.ToDateTime(dtObsoleteGroup.Rows[d]["Dato"])).Days + "</td>");
                            if (dtObsoleteGroup.Rows.Count == d + 1)
                                doc.Add("<td class='numbers-small'></td>");
                            else
                                doc.Add("<td class='numbers-small'>" + dtObsoleteGroup.Rows[d]["Varegruppe"].ToString() + "</td>");
                            doc.Add("<td class='numbers-small'>" + dtObsoleteGroup.Rows[d]["Avdeling"].ToString() + "</td>");

                            doc.Add("</tr>");
                        }

                        doc.Add("</tfoot></table></td></tr></table>");

                        doc.Add("</td></tr></table>");
                    }
                }

                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        private string SortStringStyle(string sortBy)
        {
            if (sortBy == main.appConfig.storeObsoleteSortBy)
                return " numbers-service";
            else
                return "";
        }
        private string VarianceStyle(string value, int decimals = 0)
        {
            decimal result = 0;
            if (decimal.TryParse(value, out result))
            {
                if (result == 0)
                    return "<td class='numbers-gen'>" + main.appConfig.visningNull + "</td>";

                decimal var = Math.Round(result, decimals);
                string stylePlusStr = "style='background-color:#69bd45;'";
                string styleMinusStr = "style='background-color:#e77f66;'";
                string valueStr = "";
                if (decimals == 0)
                    valueStr = var.ToString("#,##0");
                else
                    valueStr = var.ToString() + " %";

                if (result < 0)
                    return "<td class='numbers-gen' " + stylePlusStr + " >" + valueStr + "</td>";
                if (result > 0)
                    return "<td class='numbers-gen' " + styleMinusStr + " >+" + valueStr + "</td>";
            }

            return "<td class='numbers-gen'>" + main.appConfig.visningNull + "</td>";
        }

        private decimal Variance(string oldValue, string newValue)
        {
            decimal resultOld = 0, resultNew = 0;
            if (decimal.TryParse(oldValue, out resultOld) && decimal.TryParse(newValue, out resultNew))
                return resultNew - resultOld;
            return 0;
        }

        private DataTable CompareHistory(DataTable table, DataTable dtHistory)
        {
            try
            {
                if (table != null && dtHistory != null)
                {
                    DataTable dt = table.Copy();


                    dt.Columns.Add("LagerverdiCompare", typeof(decimal));
                    dt.Columns.Add("LagerverdiVariance", typeof(decimal));
                    dt.Columns.Add("UkuransverdiCompare", typeof(decimal));
                    dt.Columns.Add("UkuransverdiVariance", typeof(decimal));
                    dt.Columns.Add("UkuransprosentCompare", typeof(decimal));
                    dt.Columns.Add("UkuransprosentVariance", typeof(decimal));

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        bool match = false;
                        for (int b = 0; b < dtHistory.Rows.Count; b++)
                        {
                            if (dt.Rows[i]["Kategori"].ToString() == dtHistory.Rows[b]["Kategori"].ToString())
                            {
                                dt.Rows[i]["LagerverdiCompare"] = dtHistory.Rows[b]["Lagerverdi"];
                                dt.Rows[i]["LagerverdiVariance"] = Variance(dtHistory.Rows[b]["Lagerverdi"].ToString(), dt.Rows[i]["Lagerverdi"].ToString());
                                dt.Rows[i]["UkuransverdiCompare"] = dtHistory.Rows[b]["Ukuransverdi"];
                                dt.Rows[i]["UkuransverdiVariance"] = Variance(dtHistory.Rows[b]["Ukuransverdi"].ToString(), dt.Rows[i]["Ukuransverdi"].ToString());
                                dt.Rows[i]["UkuransprosentCompare"] = dtHistory.Rows[b]["Ukuransprosent"];
                                dt.Rows[i]["UkuransprosentVariance"] = Variance(dtHistory.Rows[b]["Ukuransprosent"].ToString(), dt.Rows[i]["Ukuransprosent"].ToString());
                                match = true;
                            }
                        }
                        if (!match)
                        {
                            dt.Rows[i]["LagerverdiCompare"] = 0;
                            dt.Rows[i]["LagerverdiVariance"] = 0;
                            dt.Rows[i]["UkuransverdiCompare"] = 0;
                            dt.Rows[i]["UkuransverdiVariance"] = 0;
                            dt.Rows[i]["UkuransprosentCompare"] = 0;
                            dt.Rows[i]["UkuransprosentVariance"] = 0;
                        }
                    }
                    return dt;
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
            return null;
        }


        public void MakeTableHeaderObsolete(List<string> doc)
        {
            doc.Add("<thead><tr>");

            doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Kategori</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=70 >Varer</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Verdi</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=70 >Uk.varer</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Uk.verdi</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=70 >Ukurans %</td>");
            doc.Add("</tr></thead>");
        }

        public void MakeTableHeaderObsoleteUtvikling(List<string> doc)
        {
            doc.Add("<thead><tr>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Kategori</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Lagerverdi</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 style='background:#80c34a;'>Siste</td>");   // style='background:#f5954e;'>Finans</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Endring</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 style='border-left:2px solid #000;'>Ukuransverdi</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 style='background:#80c34a;'>Siste</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Endring</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=70 style='border-left:2px solid #000;'>Ukurans %</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=70 style='background:#80c34a;'>Siste</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=70 >Endring</td>");
            doc.Add("</tr></thead>");
        }

        public void MakeTableHeaderObsoleteUkurantListe(List<string> doc)
        {
            doc.Add("<thead><tr>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=120 >Varekode</td>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=310 >Varetekst</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=30 >#</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Btokr</td>");   // style='background:#f5954e;'>Finans</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Ukuransverdi</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=50 >%</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=50 >Dager</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=50 >VareGrp</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=55 >Lager</td>");
            doc.Add("</tr></thead>");
        }
    }
}