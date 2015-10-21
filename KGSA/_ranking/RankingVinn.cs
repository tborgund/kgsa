using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class RankingVinn : Ranking
    {
        public DataTable tableVinn;
        decimal vinn_max = 0;

        public RankingVinn() { }

        public RankingVinn(FormMain form, DateTime dtFraArg, DateTime dtTilArg, DateTime dtPickArg)
        {
            this.main = form;
            dtFra = dtFraArg;
            dtTil = dtTilArg;
            dtPick = dtPickArg;
            velgerPeriode = FormMain.datoPeriodeVelger;

            if (main.appConfig.vinnDatoFraTil)
                tableVinn = MakeTableVinn("fratil");
            else
                tableVinn = MakeTableVinn("måned");

            if (!velgerPeriode)
                AddLastOpenDay();
        }

        private DataTable MakeTableVinnSelger(string selgerArg)
        {
            DateTime dtMainFra;
            DateTime dtMainTil;
            DataTable dtWorkSqlce;

            if (main.appConfig.vinnDatoFraTil || velgerPeriode)
            {
                dtMainFra = main.appConfig.vinnFrom;
                dtMainTil = main.appConfig.vinnTo;

                dtWorkSqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE Avdeling = " + main.appConfig.Avdeling + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "') AND Selgerkode = '" + selgerArg + "'");
                dtWorkSqlce.Columns.Add("SalgsprisExMva", typeof(Double), "Salgspris / Mva");
            }
            else
            {
                dtMainFra = dtFra;
                dtMainTil = dtTil;

                var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select("(Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "') AND Selgerkode = '" + selgerArg + "'");
                dtWorkSqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
            }

            DataTable dtWork = ReadyTableVinnSelger();

            if (StopRankingPending())
                return dtWork;

            if (dtWorkSqlce.Rows.Count == 0)
                return dtWork;
            else
            {
                List<VinnproduktItem> vinnprod = main.vinnprodukt.GetList();

                for (int i = 0; i < dtWorkSqlce.Rows.Count; i++)
                {
                    foreach (VinnproduktItem item in vinnprod)
                    {
                        if (dtWorkSqlce.Rows[i][3].ToString() == item.varekode)
                        {
                            DateTime dato = Convert.ToDateTime(dtWorkSqlce.Rows[i]["Dato"]);
                            if (dato.Date <= item.expire.Date && dato.Date >= item.start.Date)
                            {
                                DataRow dtRow = dtWork.NewRow();
                                dtRow["Selgerkode"] = selgerArg;
                                dtRow["Varekode"] = item.varekode;
                                dtRow["Poeng"] = item.poeng * Convert.ToInt32(dtWorkSqlce.Rows[i]["Antall"]);
                                dtRow["Dato"] = dato;
                                dtWork.Rows.Add(dtRow);
                            }
                        }
                    }
                }

                return dtWork;
            }
        }

        private DataTable MakeTableVinn(string strArg)
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
            else if (strArg == "fratil")
            {
                dtMainFra = main.appConfig.vinnFrom;
                dtMainTil = main.appConfig.vinnTo;

                sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE Avdeling = " + main.appConfig.Avdeling + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                sqlce.Columns.Add("SalgsprisExMva", typeof(Double), "Salgspris / Mva");
            }
            else if (velgerPeriode)
            {
                dtMainFra = dtFra;
                dtMainTil = dtTil;

                sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE Avdeling = " + main.appConfig.Avdeling + " AND (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                sqlce.Columns.Add("SalgsprisExMva", typeof(Double), "Salgspris / Mva");
            }
            else
            {
                dtMainFra = dtFra;
                dtMainTil = dtTil;

                var rows = main.database.CallMonthTable(dtMainTil, main.appConfig.Avdeling).Select("(Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "' AND Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                sqlce = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();
            }

            DataTable dtWork = ReadyTableVinn();

            if (sqlce.Rows.Count == 0)
                return dtWork;

            if (StopRankingPending())
                return dtWork;

            List<BudgetCategory> symbols = new List<BudgetCategory> 
            {
                BudgetCategory.Cross,
                BudgetCategory.MDASDA,
                BudgetCategory.Kasse,
                BudgetCategory.Aftersales
            };

            foreach (BudgetCategory cat in symbols)
            {
                string[] sk = main.salesCodes.GetVinnSelgerkoder(cat);
                if (sk.Length == 0)
                    continue;

                List<VinnproduktItem> vinnprod = main.vinnprodukt.GetList();

                decimal vinn_omset = 0, vinn_inntjen = 0, vinn_omsetExMva = 0, vinn_antall = 0, vinn_poeng = 0, vinn_favAntall = 0, antall = 0;
                string vinn_fav = "";

                DataView view = new DataView(sqlce);
                view.Sort = "Varekode";

                // Selgere
                foreach(string selger in sk)
                {
                    if (StopRankingPending())
                        return dtWork;

                    vinn_omset = 0; vinn_inntjen = 0; vinn_omsetExMva = 0; vinn_antall = 0; vinn_poeng = 0; vinn_favAntall = 0; antall = 0;
                    vinn_fav = "";
                    DataRow[] rows;
                    foreach (VinnproduktItem item in vinnprod)
                    {
                        if (item.expire.Date < dtTil.Date || item.start.Date > dtFra.Date)
                            rows = sqlce.Select("[Varekode] = '" + item.varekode + "' AND Selgerkode = '" + selger + "' AND (Dato >= '" + item.start.ToString("yyy-MM-dd") + "') AND (Dato <= '" + item.expire.ToString("yyy-MM-dd") + "')");
                        else
                            rows = sqlce.Select("[Varekode] = '" + item.varekode + "' AND Selgerkode = '" + selger + "'");
                        using (DataTable dtFilter = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone())
                        {
                            vinn_inntjen += Compute(dtFilter, "Sum(Btokr)", string.Empty);
                            vinn_omset += Compute(dtFilter, "Sum(Salgspris)", string.Empty);
                            vinn_omsetExMva += Compute(dtFilter, "Sum(SalgsprisExMva)", string.Empty);
                            antall = Compute(dtFilter, "Sum(Antall)", string.Empty);
                        }

                        if (main.appConfig.vinnEnkelModus)
                            vinn_poeng += antall;
                        else
                            vinn_poeng += item.poeng * antall;
                        vinn_antall += antall;

                        if (antall > vinn_favAntall)
                        {
                            vinn_fav = item.varekode;
                            vinn_favAntall = antall;
                        }
                    }

                    if (vinn_poeng > vinn_max)
                        vinn_max = vinn_poeng;

                    if (!main.appConfig.vinnEnkelModus || vinn_antall > 0)
                    {

                        DataRow dtRow = dtWork.NewRow();
                        dtRow["Kategori"] = BudgetCategoryClass.TypeToName(cat);
                        dtRow["Selgerkode"] = selger;

                        dtRow["Vinn_Antall"] = vinn_antall;
                        dtRow["Vinn_Poeng"] = vinn_poeng;
                        dtRow["Vinn_Inntjen"] = vinn_inntjen;
                        dtRow["Vinn_Omset"] = vinn_omset;
                        dtRow["Vinn_OmsetExMva"] = vinn_omsetExMva;

                        if (vinn_omsetExMva != 0)
                            dtRow["Vinn_Margin"] = Math.Round(vinn_inntjen / vinn_omsetExMva, 2);
                        else
                            dtRow["Vinn_Margin"] = 0;

                        dtRow["Vinn_FavorittProd"] = vinn_fav;
                        dtRow["Vinn_FavorittAntall"] = vinn_favAntall;

                        dtRow["Sort_Value"] = vinn_poeng;
                        dtWork.Rows.Add(dtRow);
                    }
                }
            }

            DataView dv = dtWork.DefaultView;
            dv.Sort = "Sort_Value desc";

            return dv.ToTable();;
        }

        private void AddLastOpenDay()
        {
            DataTable lastOpenDay = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE Avdeling = " + main.appConfig.Avdeling + " AND Dato = '" + dtTil.ToString("yyy-MM-dd") + "'");

            if (tableVinn.Rows.Count == 0)
                return;

            List<BudgetCategory> symbols = new List<BudgetCategory> 
            {
                BudgetCategory.Cross,
                BudgetCategory.MDASDA,
                BudgetCategory.Kasse,
                BudgetCategory.Aftersales
            };

            foreach (BudgetCategory cat in symbols)
            {
                string[] sk = main.salesCodes.GetVinnSelgerkoder(cat);
                if (sk.Length == 0)
                    continue;

                List<VinnproduktItem> vinnprod = main.vinnprodukt.GetList();
                decimal vinn_poeng = 0, antall = 0, vinn_antall = 0, vinn_favAntall = 0;
                string vinn_fav = "";

                // Selgere
                foreach (string selger in sk)
                {
                    vinn_poeng = 0; antall = 0; vinn_favAntall = 0;
                    vinn_fav = "";

                    foreach (VinnproduktItem item in vinnprod)
                    {
                        antall = Compute(lastOpenDay, "Sum(Antall)", "[Varekode] = '" + item.varekode + "' AND Selgerkode = '" + selger + "'");
                        if (main.appConfig.vinnEnkelModus)
                            vinn_poeng += antall;
                        else
                            vinn_poeng += item.poeng * antall;
                        vinn_antall += antall;

                        if (antall > vinn_favAntall && !item.varekode.StartsWith("MOD"))
                        {
                            vinn_fav = item.varekode;
                            vinn_favAntall = antall;
                        }
                    }

                    for (int b = 0; b < tableVinn.Rows.Count; b++)
                    {
                        if (tableVinn.Rows[b]["Selgerkode"].ToString() == selger)
                        {
                            tableVinn.Rows[b]["Vinn_SisteDagPoeng"] = vinn_poeng;
                            tableVinn.Rows[b]["Vinn_SisteDagVare"] = vinn_fav;
                            break;
                        }
                    }
                }
            }
        }

        public List<string> GetTableHtmlSelger(string selgerArg)
        {
            try
            {
                var doc = new List<string>() { };
                var hashId = random.Next(999, 99999);

                DataTable dt = MakeTableVinnSelger(selgerArg);

                if (dt.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen vinnprodukter solgt av denne selgeren.</span><br>");
                    return doc;
                }

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='OutertableNormal toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.Add("<thead><tr>");

                doc.Add("<th class=\"{sorter: 'text'}\" width=\"210\" style='background:#f04c4d;'>Varekode</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=\"200\" style='background:#f04c4d;'>Poeng</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=\"110\" style='background:#f04c4d;'>Dato</td>");

                doc.Add("</tr></thead>");
                doc.Add("<tbody>");

                decimal totalt = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string strLine = dt.Rows.Count == i - 1 ? "border-bottom:none;" : "";
                    doc.Add("<tr><td class='text-cat' style='" + strLine + "'>" + dt.Rows[i]["Varekode"].ToString() + "</td>");
                    doc.Add("<td class='numbers-small' style='" + strLine + "'>" + PlusMinus(dt.Rows[i]["Poeng"]) + "</td>");
                    doc.Add("<td class='numbers-gen' style='" + strLine + "'>" + Convert.ToDateTime(dt.Rows[i]["Dato"]).ToShortDateString() + "</td>");
                    doc.Add("</tr>");
                    totalt += Convert.ToDecimal(dt.Rows[i]["Poeng"]);
                }
                doc.Add("</tfoot>");
                doc.Add("</table></td></tr></table>");
                doc.Add("Totalt antall poeng: " + totalt);
                doc.Add("<br><a href='" + FormMain.htmlRankingVinn + "'>Tilbake</a>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return new List<string>() { }; ;
        }

        public List<string> GetTableHtml(BudgetCategory cat)
        {
            try
            {
                var doc = new List<string>() { };
                var hashId = random.Next(999, 99999);

                string filter = "Kategori = '" + BudgetCategoryClass.TypeToName(cat) + "'";
                if (cat == BudgetCategory.None)
                    filter = "";

                var rows = tableVinn.Select(filter);
                DataTable dtVinn = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone();

                if (dtVinn.Rows.Count > 0)
                    dt = dtVinn;
                else
                {
                    doc.Add("<span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner.</span><br>");
                    return doc;
                }

                main.openXml.SaveDocument(dt, "Vinnprodukter", BudgetCategoryClass.TypeToName(cat), dtPick,  BudgetCategoryClass.TypeToName(cat).ToUpper() + " - " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<div style=\"width:100%;\" id='" + hashId + "' class='toggleAll'>");
                doc.Add("<table class='tblBestOf'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.Add("<thead><tr>");

                doc.Add("<th class=\"{sorter: 'text'}\" width=\"210\" style='background:#f04c4d;'>Selger</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=\"200\" style='background:#f04c4d;'>&nbsp;</td>");

                if (main.appConfig.vinnEnkelModus)
                {
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=\"180\" style='background:#f04c4d;'>Antall</td>");
                }
                else
                {
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=\"110\" style='background:#f04c4d;'>Poeng</td>");
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=\"100\" style='background:#f04c4d;'>Antall</td>");
                }

                doc.Add("<th class=\"{sorter: 'digit'}\" width=\"120\" style='background:#f04c4d;'>Inntjen</td>");
                if (main.appConfig.vinnSisteDagVare)
                    doc.Add("<th class=\"{sorter: 'text'}\" width=\"210\" style='background:#f04c4d;'>Mest solgt sist dag</td>");
                else
                    doc.Add("<th class=\"{sorter: 'text'}\" width=\"210\" style='background:#f04c4d;'>Mest solgt</td>");

                doc.Add("</tr></thead>");
                doc.Add("<tbody>");

                decimal height = 5;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    height++;
                    string strLine = dt.Rows.Count == i - 1 ? "border-bottom:none;" : "";
                    doc.Add("<tr><td class='text-cat' style='font-size:14pt;" + strLine + "'><a href='#vinn" + dt.Rows[i]["Selgerkode"].ToString() + "'>" + main.salesCodes.GetNavn(dt.Rows[i]["Selgerkode"].ToString()) + "</a></td>");

                    if (main.appConfig.vinnEnkelModus)
                    {
                        doc.Add("<td class='numbers-small' style='font-size:14pt;" + strLine + ShowGraph((decimal)dt.Rows[i]["Vinn_Poeng"], vinn_max) + strLine + "'>&nbsp;</td>");
                        doc.Add("<td class='numbers-gen' style='font-size:14pt;" + strLine + "'>" + LastOpenDayString(dt.Rows[i]["Vinn_SisteDagPoeng"]) + ForkortTall(Convert.ToDecimal(dt.Rows[i]["Vinn_Poeng"])) + "</td>");
                    }
                    else
                    {
                        doc.Add("<td class='numbers-small' style='font-size:14pt;" + strLine + ShowGraph((decimal)dt.Rows[i]["Vinn_Poeng"], vinn_max) + strLine + "'>&nbsp;</td>");
                        doc.Add("<td class='numbers-gen' style='font-size:14pt;" + strLine + "'>" + LastOpenDayString(dt.Rows[i]["Vinn_SisteDagPoeng"]) + ForkortTall(Convert.ToDecimal(dt.Rows[i]["Vinn_Poeng"])) + "</td>");
                        doc.Add("<td class='numbers-small' style='font-size:14pt;" + strLine + "'>" + ForkortTall(Convert.ToDecimal(dt.Rows[i]["Vinn_Antall"])) + "</td>");
                    }

                    doc.Add("<td class='numbers-gen' style='font-size:14pt;" + strLine + "'>" + ForkortTall(Convert.ToDecimal(dt.Rows[i]["Vinn_Inntjen"])) + "</td>");
                    if (main.appConfig.vinnSisteDagVare && !DBNull.Value.Equals(dt.Rows[i]["Vinn_SisteDagVare"]))
                        doc.Add("<td class='numbers-percent' style='font-size:14pt;" + strLine + "'>" + dt.Rows[i]["Vinn_SisteDagVare"].ToString().ToUpper() + "</td>");
                    else if (main.appConfig.vinnSisteDagVare && DBNull.Value.Equals(dt.Rows[i]["Vinn_SisteDagVare"]))
                        doc.Add("<td class='numbers-percent' style='font-size:14pt;" + strLine + "'>&nbsp;</td>");
                    else
                        doc.Add("<td class='numbers-percent' style='font-size:14pt;" + strLine + "'>" + dt.Rows[i]["Vinn_FavorittProd"].ToString().ToUpper() + "</td>");
                    doc.Add("</tr>");
                }
                doc.Add("</tfoot>");
                doc.Add("</table></td></tr></table>");

                return doc;
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return new List<string>() { }; ;
        }

        private string LastOpenDayString(object obj)
        {
            if (!DBNull.Value.Equals(obj))
            {
                decimal expDec = Convert.ToDecimal(obj);
                string expStr = "";
                if (expDec > 0)
                    return expStr = "<span style='color:green;font-size:xx-small;'>(+" + ForkortTall(expDec) + ")</span> ";
                if (expDec < 0)
                    return expStr = "<span style='color:red;font-size:xx-small;'>(-" + ForkortTall(expDec) + ")</span> ";
            }
            return "";
        }

        public List<string> GetTableVarekoderHtml()
        {
            try
            {
                var doc = new List<string>() { };
                if (StopRankingPending())
                    return doc;
                var hashId = random.Next(999, 99999);

                List<VinnproduktItem> vinnprod = main.vinnprodukt.GetList();

                if (vinnprod.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen varekoder.</span><br>");
                    return doc;
                }

                vinnprod = vinnprod.OrderBy(x => x.varekode).ToList();
                vinnprod = vinnprod.OrderBy(x => x.kategori).ToList();

                int nSize = 60;

                List<List<VinnproduktItem>> splits = splitList(vinnprod, nSize);

                List<VinnproduktItem> group1 = splits[0];
                List<VinnproduktItem> group2 = splits.Count > 1 ? splits[1] : new List<VinnproduktItem>() { };
                List<VinnproduktItem> group3 = splits.Count > 2 ? splits[2] : new List<VinnproduktItem>() { };

                VarekodelisteHtml(doc, group1, group2, group3, nSize, "Alle Vinn-varekoder");

                List<VinnproduktItem> group4 = splits.Count > 3 ? splits[3] : new List<VinnproduktItem>() { };
                List<VinnproduktItem> group5 = splits.Count > 4 ? splits[4] : new List<VinnproduktItem>() { };
                List<VinnproduktItem> group6 = splits.Count > 5 ? splits[5] : new List<VinnproduktItem>() { };

                if (group4.Count > 0)
                    VarekodelisteHtml(doc, group4, group5, group6, nSize, "Alle Vinn-varekoder (Side 2)");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return new List<string>() { }; ;
        }

        private void VarekodelisteHtml(List<string> doc, List<VinnproduktItem> column1, List<VinnproduktItem> column2, List<VinnproduktItem> column3, int nSize, string title)
        {
            var hashId = random.Next(999, 99999);

            doc.Add("<h2>" + title + "</h2>");
            doc.Add("<div class='toolbox hidePdf'>");
            doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
            doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
            doc.Add("</div>");
            doc.Add("<div style=\"width:100%;\" id='" + hashId + "' class='toggleAll'>");
            doc.Add("<table class='tblBestOf'><tr><td>");
            doc.Add("<table class='tablesorter'>");
            doc.Add("<thead><tr>");

            doc.Add("<th class=\"{sorter: 'text'}\" width=\"150\" style='background:#f04c4d;'>Varekode</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=\"73\" style='background:#f04c4d;'>Poeng</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=\"75\" style='background:#f04c4d;'>Kategori</td>");

            doc.Add("<th class=\"{sorter: 'text'}\" width=\"150\" style='background:#f04c4d;border-left:black 1px solid;'>Varekode</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=\"73\" style='background:#f04c4d;'>Poeng</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=\"75\" style='background:#f04c4d;'>Kategori</td>");

            doc.Add("<th class=\"{sorter: 'text'}\" width=\"150\" style='background:#f04c4d;border-left:black 1px solid;'>Varekode</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=\"73\" style='background:#f04c4d;'>Poeng</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=\"75\" style='background:#f04c4d;'>Kategori</td>");

            doc.Add("</tr></thead>");
            doc.Add("<tbody>");

            for (int i = 0; i < nSize; i++)
            {
                doc.Add("<tr>");
                if (column1.Count > i)
                {
                    doc.Add("<td class='text-cat' style='" + KategoriToStyle(column1[i].kategori) + "'>" + column1[i].varekode + "</td>");
                    doc.Add("<td class='numbers-small' style='" + KategoriToStyle(column1[i].kategori) + "'>" + column1[i].poeng + "</td>");
                    doc.Add("<td class='numbers-small' style='" + KategoriToStyle(column1[i].kategori) + "'>" + column1[i].kategori + "</td>");
                }
                else
                {
                    doc.Add("<td class='text-cat'>&nbsp;</td>");
                    doc.Add("<td class='numbers-small'>&nbsp;</td>");
                    doc.Add("<td class='numbers-small'>&nbsp;</td>");
                }

                if (column2.Count > i)
                {
                    doc.Add("<td class='text-cat' style='border-left:black 1px solid;" + KategoriToStyle(column2[i].kategori) + "'>" + column2[i].varekode + "</td>");
                    doc.Add("<td class='numbers-small' style='" + KategoriToStyle(column2[i].kategori) + "'>" + column2[i].poeng + "</td>");
                    doc.Add("<td class='numbers-small' style='" + KategoriToStyle(column2[i].kategori) + "'>" + column2[i].kategori + "</td>");
                }
                else
                {
                    doc.Add("<td class='text-cat' style='border-left:black 1px solid;'>&nbsp;</td>");
                    doc.Add("<td class='numbers-small'>&nbsp;</td>");
                    doc.Add("<td class='numbers-small'>&nbsp;</td>");
                }

                if (column3.Count > i)
                {
                    doc.Add("<td class='text-cat' style='border-left:black 1px solid;" + KategoriToStyle(column3[i].kategori) + "'>" + column3[i].varekode + "</td>");
                    doc.Add("<td class='numbers-small' style='" + KategoriToStyle(column3[i].kategori) + "'>" + column3[i].poeng + "</td>");
                    doc.Add("<td class='numbers-small' style='" + KategoriToStyle(column3[i].kategori) + "'>" + column3[i].kategori + "</td>");
                }
                else
                {
                    doc.Add("<td class='text-cat' style='border-left:black 1px solid;'>&nbsp;</td>");
                    doc.Add("<td class='numbers-small'>&nbsp;</td>");
                    doc.Add("<td class='numbers-small'>&nbsp;</td>");
                }
                doc.Add("</tr>");
            }

            doc.Add("</tbody>");
            doc.Add("</table></td></tr></table>");
        }

        public List<string> GetTableVarekoderEkstraHtml()
        {
            try
            {
                string sql = "SELECT DISTINCT tblVinnprodukt.Varekode, tblObsolete.Varetekst, tblObsolete.MerkeNavn, tblVinnprodukt.Kategori, tblVinnprodukt.Poeng, tblVinnprodukt.DatoStart " +
                    "FROM tblVinnprodukt " +
                    "INNER JOIN tblObsolete " +
                    "ON tblVinnprodukt.Varekode=tblObsolete.Varekode " +
                    "ORDER BY tblVinnprodukt.Kategori";

                DataTable table = main.database.GetSqlDataTable(sql);
                List<string> doc = new List<string>() { };
                var hashId = random.Next(999, 99999);

                if (table.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen varekoder som matchet lager beholdning. Har du importert Obsolete.csv?</span><br>");
                    return doc;
                }

                //doc.Add("<br><table style='margin: 0 auto !important;width:100%;'><tr><td>");
                doc.Add("<h2>Vinn-varekoder med varetekst</h2>");
                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");
                doc.Add("<div style=\"width:100%;\" id='" + hashId + "' class='toggleAll'>");
                doc.Add("<table class='tblBestOf'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.Add("<thead><tr>");

                doc.Add("<th class=\"{sorter: 'text'}\" width=\"200\" style='background:#f04c4d;'>Varekode</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=\"375\" style='background:#f04c4d;'>Varetekst</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=\"135\" style='background:#f04c4d;'>Varemerke</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=\"100\" style='background:#f04c4d;'>Poeng</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=\"70\" style='background:#f04c4d;'>Kategori</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" width=\"70\" style='background:#f04c4d;'>Fra dato</td>");

                doc.Add("</tr></thead>");
                doc.Add("<tbody>");

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    doc.Add("<tr>");
                    doc.Add("<td class='text-cat' style='" + KategoriToStyle(table.Rows[i]["Kategori"].ToString()) + "'>" + table.Rows[i]["Varekode"].ToString() + "</td>");
                    doc.Add("<td class='text' style='" + KategoriToStyle(table.Rows[i]["Kategori"].ToString()) + "'>" + table.Rows[i]["Varetekst"].ToString() + "</td>");
                    doc.Add("<td class='text' style='" + KategoriToStyle(table.Rows[i]["Kategori"].ToString()) + "'>" + table.Rows[i]["MerkeNavn"].ToString() + "</td>");
                    doc.Add("<td class='numbers-small' style='" + KategoriToStyle(table.Rows[i]["Kategori"].ToString()) + "'>" + table.Rows[i]["Poeng"].ToString() + "</td>");
                    doc.Add("<td class='text' style='" + KategoriToStyle(table.Rows[i]["Kategori"].ToString()) + "'>" + table.Rows[i]["Kategori"].ToString() + "</td>");
                    doc.Add("<td class='text' style='" + KategoriToStyle(table.Rows[i]["Kategori"].ToString()) + "'>" + Convert.ToDateTime(table.Rows[i]["DatoStart"]).ToShortDateString() + "</td>");
                    doc.Add("</tr>");
                }

                doc.Add("</tbody>");
                doc.Add("</table></td></tr></table>");
                //doc.Add("</td></tr></table>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return new List<string>() { };
        }

        private string KategoriToStyle(string kat)
        {
            if (kat == "MDA")
                return "color:#8E3155;";
            if (kat == "AudioVideo")
                return "color:#603716;";
            if (kat == "SDA")
                return "color:#186c18;";
            if (kat == "Tele")
                return "color:#AB7A00;";
            if (kat == "Data")
                return "color:#144f60;";

            return "color:black;";
        }

        public static List<List<VinnproduktItem>> splitList(List<VinnproduktItem> varekoder, int nSize = 30)
        {
            List<List<VinnproduktItem>> list = new List<List<VinnproduktItem>>();

            for (int i = 0; i < varekoder.Count; i += nSize)
            {
                list.Add(varekoder.GetRange(i, Math.Min(nSize, varekoder.Count - i)));
            }

            return list;
        } 

        public string ShowGraph(decimal value, decimal max)
        {
            try
            {
                if (value <= 0)
                    return "";

                string style = "background:url(\"bar_red.png\") no-repeat; background-size:";

                decimal size = Math.Round(value / max * 100);
                if (size > 100)
                    size = 100;
                if (size < 0)
                    size = 0;

                return style + size + "% 100%;";
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        public DataTable ReadyTableVinn()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Kategori", typeof(string));
            dataTable.Columns.Add("Selgerkode", typeof(string));

            dataTable.Columns.Add("Vinn_Antall", typeof(int));
            dataTable.Columns.Add("Vinn_Poeng", typeof(decimal));
            dataTable.Columns.Add("Vinn_Omset", typeof(decimal));
            dataTable.Columns.Add("Vinn_Inntjen", typeof(decimal));
            dataTable.Columns.Add("Vinn_OmsetExMva", typeof(decimal));
            dataTable.Columns.Add("Vinn_Margin", typeof(decimal));
            dataTable.Columns.Add("Vinn_FavorittProd", typeof(string));
            dataTable.Columns.Add("Vinn_FavorittAntall", typeof(decimal));

            dataTable.Columns.Add("Vinn_SisteDagPoeng", typeof(decimal));
            dataTable.Columns.Add("Vinn_SisteDagVare", typeof(string));

            dataTable.Columns.Add("Sort_Value", typeof(int));
            return dataTable;
        }

        public DataTable ReadyTableVinnSelger()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Selgerkode", typeof(string));
            dataTable.Columns.Add("Varekode", typeof(string));
            dataTable.Columns.Add("Poeng", typeof(decimal));
            dataTable.Columns.Add("Dato", typeof(DateTime));
            return dataTable;
        }
    }
}

