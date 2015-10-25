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
    public class RankingAvdTjenester : Ranking
    {
        public List<VarekodeList> varekoderAlle;
        public IEnumerable<string> varekoderAlleAlias;
        public RankingAvdTjenester() { }

        public RankingAvdTjenester(FormMain form, DateTime dtFraArg, DateTime dtTilArg, DateTime dtPickArg)
        {
            this.main = form;
            dtFra = dtFraArg;
            dtTil = dtTilArg;
            dtPick = dtPickArg;
            velgerPeriode = FormMain.datoPeriodeVelger;

            this.varekoderAlle = main.appConfig.varekoder.ToList();
            this.varekoderAlleAlias = varekoderAlle.Where(item => item.synlig == true).Select(x => x.alias).Distinct();
        }

        private DataTable MakeTable(int rankMode, string service, string status)
        {
            try
            {
                DateTime dtMainFra;
                DateTime dtMainTil;

                if (sqlce != null && rankMode > -1)
                    if (sqlce.Rows.Count > 10)
                        goto skipDbPull;

                if (rankMode == 0)
                {
                    dtMainFra = dtPick;
                    dtMainTil = dtPick;

                    sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                }
                else if (rankMode == 1)
                {
                    dtMainFra = GetFirstDayOfMonth(dtPick);
                    dtMainTil = GetLastDayOfMonth(dtPick);

                    sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                }
                else if (rankMode == 2) // bonus
                {
                    if (IsOdd(dtPick.Month)) // Is at start of bonus periode
                        dtMainFra = GetFirstDayOfMonth(dtPick);
                    else // Is in second month of bonus periode, move one month
                        dtMainFra = GetFirstDayOfMonth(dtPick.AddMonths(-1));

                    if (IsOdd(dtPick.Month)) // Is at start of bonus periode
                        dtMainTil = GetLastDayOfMonth(dtPick);
                    else // Is in second month of bonus periode, move one month
                        dtMainTil = GetLastDayOfMonth(dtPick.AddMonths(-1));

                    sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                }
                else if (rankMode == 3)
                {
                    dtMainFra = new DateTime(dtPick.Year, 1, 1);
                    dtMainTil = dtPick;

                    sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                }
                else if (rankMode == 4)
                {
                    DateTime startOfLastWeek = main.database.GetStartOfLastWholeWeek(dtTil); ;
                    dtMainFra = startOfLastWeek;
                    dtMainTil = startOfLastWeek.AddDays(6);

                    sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                }
                else if (rankMode == 5)
                {
                    dtMainFra = new DateTime(dtPick.Year, 1, 1);
                    dtMainTil = new DateTime(dtPick.Year, 12, GetLastDayOfMonth(new DateTime(dtPick.Year, 12, 1)).Day);

                    sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (Dato >= '" + dtMainFra.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtMainTil.ToString("yyy-MM-dd") + "')");
                }
                else
                {
                    Log.n("Ingen rank mode valgt: " + rankMode, Color.Red);
                    return null;
                }

            skipDbPull:

                DataTable dtWork = ReadyTable();

                if (sqlce.Rows.Count == 0)
                    return dtWork;

                if (main.arrayDbAvd == null || main.arrayDbAvd.ToList().Count == 0 && main.appConfig.rankingAvdelingShowAll)
                {
                    Log.n("Vent til avdelingslisten er oppdatert..", Color.Red);
                    return null;
                }

                if (StopRankingPending())
                    return dtWork;

                List<string> avdList = FormMain.Favoritter;

                if (main.appConfig.rankingAvdelingShowAll && main.arrayDbAvd != null)
                    avdList = main.arrayDbAvd.ToList();

                string knowhowTjenester = "";
                foreach (var varekode in varekoderAlle)
                {
                    if (varekode.synlig)
                        knowhowTjenester += "[Varekode]='" + varekode.kode + "' OR ";
                }
                if (varekoderAlle.Count != 0)
                    knowhowTjenester = knowhowTjenester.Substring(0, knowhowTjenester.Length - 4);

                string[,] gruppe2d = new string[5, 3] { { "1", "6", "MDA" }, { "2", "8", "LoB" },
                { "3", "10", "SDA" }, { "4", "12", "Tele" }, { "5", "14", "Data"} };

                // Butikker
                foreach (string avdStr in avdList)
                {
                    if (StopRankingPending())
                        return dtWork;

                    Log.Status(status + " " + avdeling.Get(avdStr));

                    var rowsGet = sqlce.Select("(Avdeling = " + avdStr + ")");
                    using (DataTable sqlceAvd = rowsGet.Any() ? rowsGet.CopyToDataTable() : sqlce.Clone())
                    {
                        decimal totalt_Antall_Bilag = sqlceAvd.AsEnumerable().Select(r => r.Field<int>("Bilagsnr")).Distinct().Count();
                        decimal totalt_Antall_Spurt = 0, totalt_Antall_Tjenester = 0;

                        DataRow dtRow = dtWork.NewRow();
                        dtRow["Avdeling"] = avdStr;
                        dtRow["Tjeneste"] = service;

                        // Grupper
                        for (int i = 0; i < gruppe2d.GetLength(0); i++)
                        {
                            // Tjenester
                            decimal gruppe_Antall_Bilag = 0, gruppe_Antall_Spurt = 0, gruppe_Antall_Tjenester = 0;

                            if (service == "Finans")
                            {
                                // Finans
                                var rows = sqlceAvd.Select("(Varegruppe >= " + gruppe2d[i, 0] + "00 AND Varegruppe <= " + gruppe2d[i, 0] + "99) OR Varegruppe > 900");
                                using (DataTable sqlceAvdGruppe = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone())
                                {
                                    var rowBilag = sqlceAvd.Select("(Varegruppe >= " + gruppe2d[i, 0] + "00 AND Varegruppe <= " + gruppe2d[i, 0] + "99)");
                                    using (DataTable sqlceAvdGruppeBilag = rowBilag.Any() ? rowBilag.CopyToDataTable() : sqlce.Clone())
                                        gruppe_Antall_Bilag = sqlceAvdGruppeBilag.AsEnumerable().Select(r => r.Field<int>("Bilagsnr")).Distinct().Count();

                                    var rowSpurt = sqlceAvdGruppe.Select("[Varekode] = 'IE'");
                                    for (int f = 0; f < rowSpurt.Length; f++)
                                    {
                                        var rows2 = sqlceAvd.Select("[Bilagsnr] = " + rowSpurt[f]["Bilagsnr"]);
                                        DataTable dtService = rows2.Any() ? rows2.CopyToDataTable() : sqlce.Clone();
                                        dtService.DefaultView.Sort = "Salgspris DESC";
                                        if (gruppe2d[i, 0] == dtService.Rows[0]["Varegruppe"].ToString().Substring(0, 1))
                                            gruppe_Antall_Spurt++;
                                    }

                                    var rowTjen = sqlceAvdGruppe.Select("[Varegruppe] = 961");
                                    for (int f = 0; f < rowTjen.Length; f++)
                                    {
                                        var rows2 = sqlceAvd.Select("[Bilagsnr] = " + rowTjen[f]["Bilagsnr"]);
                                        DataTable dtFinans = rows2.Any() ? rows2.CopyToDataTable() : sqlce.Clone();
                                        dtFinans.DefaultView.Sort = "Salgspris DESC";
                                        if (gruppe2d[i, 0] == dtFinans.Rows[0]["Varegruppe"].ToString().Substring(0, 1))
                                            gruppe_Antall_Tjenester += Compute(sqlce, "Sum(Antall)", "[Varegruppe] = 961 AND [Bilagsnr] = " + dtFinans.Rows[0]["Bilagsnr"].ToString());
                                    }

                                }
                            }
                            else if (service == "TA")
                            {
                                // TA
                                var rows = sqlceAvd.Select("(Varegruppe >= " + gruppe2d[i, 0] + "00 AND Varegruppe <= " + gruppe2d[i, 0] + "99) OR Varegruppe > 900");
                                using (DataTable sqlceAvdGruppe = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone())
                                {
                                    var rowBilag = sqlceAvd.Select("(Varegruppe >= " + gruppe2d[i, 0] + "00 AND Varegruppe <= " + gruppe2d[i, 0] + "99)");
                                    using (DataTable sqlceAvdGruppeBilag = rowBilag.Any() ? rowBilag.CopyToDataTable() : sqlce.Clone())
                                        gruppe_Antall_Bilag = sqlceAvdGruppeBilag.AsEnumerable().Select(r => r.Field<int>("Bilagsnr")).Distinct().Count();

                                    var rowSpurt = sqlceAvdGruppe.Select("[Varekode] = 'IF'");
                                    for (int f = 0; f < rowSpurt.Length; f++)
                                    {
                                        var rows2 = sqlceAvd.Select("[Bilagsnr] = " + rowSpurt[f]["Bilagsnr"]);
                                        DataTable dtService = rows2.Any() ? rows2.CopyToDataTable() : sqlce.Clone();
                                        dtService.DefaultView.Sort = "Salgspris DESC";
                                        if (gruppe2d[i, 0] == dtService.Rows[0]["Varegruppe"].ToString().Substring(0, 1))
                                            gruppe_Antall_Spurt++;
                                    }

                                    gruppe_Antall_Tjenester = Compute(sqlceAvdGruppe, "Count(Varekode)", "[Varegruppe] % 100 = 83 AND [Varekode] LIKE 'MOD*'");
                                }
                            }
                            else if (service == "Strom")
                            {
                                // Strøm
                                var rows = sqlceAvd.Select("(Varegruppe >= " + gruppe2d[i, 0] + "00 AND Varegruppe <= " + gruppe2d[i, 0] + "99) OR Varegruppe > 900");
                                using (DataTable sqlceAvdGruppe = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone())
                                {
                                    var rowBilag = sqlceAvd.Select("(Varegruppe >= " + gruppe2d[i, 0] + "00 AND Varegruppe <= " + gruppe2d[i, 0] + "99)");
                                    using (DataTable sqlceAvdGruppeBilag = rowBilag.Any() ? rowBilag.CopyToDataTable() : sqlce.Clone())
                                        gruppe_Antall_Bilag = sqlceAvdGruppeBilag.AsEnumerable().Select(r => r.Field<int>("Bilagsnr")).Distinct().Count();

                                    var rowSpurt = sqlceAvdGruppe.Select("[Varekode] = 'IF'");
                                    for (int f = 0; f < rowSpurt.Length; f++)
                                    {
                                        var rows2 = sqlceAvd.Select("[Bilagsnr] = " + rowSpurt[f]["Bilagsnr"]);
                                        DataTable dtService = rows2.Any() ? rows2.CopyToDataTable() : sqlce.Clone();
                                        dtService.DefaultView.Sort = "Salgspris DESC";
                                        if (gruppe2d[i, 0] == dtService.Rows[0]["Varegruppe"].ToString().Substring(0, 1))
                                            gruppe_Antall_Spurt++;
                                    }

                                    gruppe_Antall_Tjenester = Compute(sqlceAvdGruppe, "Count(Varekode)", "[Varekode] LIKE 'ELSTROM*'");
                                }
                            }
                            else if (service == "Knowhow")
                            {
                                // KnowHow
                                var rows = sqlceAvd.Select("(Varegruppe >= " + gruppe2d[i, 0] + "00 AND Varegruppe <= " + gruppe2d[i, 0] + "99) OR Varegruppe > 900");
                                using (DataTable sqlceAvdGruppe = rows.Any() ? rows.CopyToDataTable() : sqlce.Clone())
                                {
                                    var rowBilag = sqlceAvd.Select("(Varegruppe >= " + gruppe2d[i, 0] + "00 AND Varegruppe <= " + gruppe2d[i, 0] + "99)");
                                    using (DataTable sqlceAvdGruppeBilag = rowBilag.Any() ? rowBilag.CopyToDataTable() : sqlce.Clone())
                                        gruppe_Antall_Bilag = sqlceAvdGruppeBilag.AsEnumerable().Select(r => r.Field<int>("Bilagsnr")).Distinct().Count();

                                    var rowSpurt = sqlceAvdGruppe.Select("[Varekode] = 'IRTG' OR [Varekode] = 'IK'");
                                    for (int f = 0; f < rowSpurt.Length; f++)
                                    {
                                        var rows2 = sqlceAvd.Select("[Bilagsnr] = " + rowSpurt[f]["Bilagsnr"]);
                                        DataTable dtService = rows2.Any() ? rows2.CopyToDataTable() : sqlce.Clone();
                                        dtService.DefaultView.Sort = "Salgspris DESC";
                                        if (gruppe2d[i, 0] == dtService.Rows[0]["Varegruppe"].ToString().Substring(0, 1))
                                            gruppe_Antall_Spurt++;
                                    }

                                    gruppe_Antall_Tjenester = Compute(sqlceAvdGruppe, "Count(Varekode)", knowhowTjenester);
                                }
                            }

                            dtRow[gruppe2d[i, 2]] = gruppe_Antall_Tjenester;
                            totalt_Antall_Tjenester += gruppe_Antall_Tjenester;
                            dtRow[gruppe2d[i, 2] + "_spurt"] = gruppe_Antall_Spurt;
                            totalt_Antall_Spurt += gruppe_Antall_Spurt;
                            dtRow[gruppe2d[i, 2] + "_bilag"] = gruppe_Antall_Bilag;

                            if (gruppe_Antall_Bilag != 0)
                                dtRow[gruppe2d[i, 2] + "_percent"] = Math.Round(gruppe_Antall_Spurt / gruppe_Antall_Bilag * 100, 2);
                            else
                                dtRow[gruppe2d[i, 2] + "_percent"] = 0;
                        }

                        dtRow["Totalt"] = totalt_Antall_Tjenester;
                        dtRow["Totalt_spurt"] = totalt_Antall_Spurt;
                        dtRow["Totalt_bilag"] = totalt_Antall_Bilag;

                        if (totalt_Antall_Bilag != 0)
                            dtRow["Totalt_percent"] = Math.Round(totalt_Antall_Spurt / totalt_Antall_Bilag * 100, 2);
                        else
                            dtRow["Totalt_percent"] = 0;

                        dtWork.Rows.Add(dtRow);
                    }
                }

                if (dtWork.Rows.Count > 0)
                {
                    DataRow row = dtWork.NewRow();
                    row["Avdeling"] = "Totalt";
                    row["Totalt"] = dtWork.Compute("Sum(Totalt)", null);
                    row["Totalt_bilag"] = dtWork.Compute("Sum(Totalt_bilag)", null);
                    row["Totalt_spurt"] = dtWork.Compute("Sum(Totalt_spurt)", null);
                    row["Totalt_percent"] = dtWork.Compute("Sum(Totalt_percent)", null);

                    decimal totalt_bilag = Convert.ToDecimal(row["Totalt_bilag"]);
                    decimal totalt_spurt = Convert.ToDecimal(row["Totalt_spurt"]);
                    if (totalt_bilag != 0)
                        row["Totalt_percent"] = Math.Round(totalt_spurt / totalt_bilag * 100, 2);
                    else
                        row["Totalt_percent"] = 0;

                    // MDA
                    row["MDA"] = dtWork.Compute("Sum(MDA)", null);
                    row["MDA_bilag"] = dtWork.Compute("Sum(MDA_bilag)", null);
                    row["MDA_spurt"] = dtWork.Compute("Sum(MDA_spurt)", null);

                    decimal mda_bilag = Convert.ToDecimal(row["MDA_bilag"]);
                    decimal mda_spurt = Convert.ToDecimal(row["MDA_spurt"]);
                    if (mda_bilag != 0)
                        row["MDA_percent"] = Math.Round(mda_spurt / mda_bilag * 100, 2);
                    else
                        row["MDA_percent"] = 0;

                    // LoB
                    row["LoB"] = dtWork.Compute("Sum(LoB)", null);
                    row["LoB_bilag"] = dtWork.Compute("Sum(LoB_bilag)", null);
                    row["LoB_spurt"] = dtWork.Compute("Sum(LoB_spurt)", null);

                    decimal lob_bilag = Convert.ToDecimal(row["LoB_bilag"]);
                    decimal lob_spurt = Convert.ToDecimal(row["LoB_spurt"]);
                    if (lob_bilag != 0)
                        row["LoB_percent"] = Math.Round(lob_spurt / lob_bilag *100, 2);
                    else
                        row["LoB_percent"] = 0;

                    // SDA
                    row["SDA"] = dtWork.Compute("Sum(SDA)", null);
                    row["SDA_bilag"] = dtWork.Compute("Sum(SDA_bilag)", null);
                    row["SDA_spurt"] = dtWork.Compute("Sum(SDA_spurt)", null);

                    decimal sda_bilag = Convert.ToDecimal(row["SDA_bilag"]);
                    decimal sda_spurt = Convert.ToDecimal(row["SDA_spurt"]);
                    if (sda_bilag != 0)
                        row["SDA_percent"] = Math.Round(sda_spurt / sda_bilag * 100, 2);
                    else
                        row["SDA_percent"] = 0;

                    // Tele
                    row["Tele"] = dtWork.Compute("Sum(Tele)", null);
                    row["Tele_bilag"] = dtWork.Compute("Sum(Tele_bilag)", null);
                    row["Tele_spurt"] = dtWork.Compute("Sum(Tele_spurt)", null);

                    decimal tele_bilag = Convert.ToDecimal(row["Tele_bilag"]);
                    decimal tele_spurt = Convert.ToDecimal(row["Tele_spurt"]);
                    if (tele_bilag != 0)
                        row["Tele_percent"] = Math.Round(tele_spurt / tele_bilag * 100, 2);
                    else
                        row["Tele_percent"] = 0;

                    // Data
                    row["Data"] = dtWork.Compute("Sum(Data)", null);
                    row["Data_bilag"] = dtWork.Compute("Sum(Data_bilag)", null);
                    row["Data_spurt"] = dtWork.Compute("Sum(Data_spurt)", null);

                    decimal data_bilag = Convert.ToDecimal(row["Data_bilag"]);
                    decimal data_spurt = Convert.ToDecimal(row["Data_spurt"]);
                    if (data_bilag != 0)
                        row["Data_percent"] = Math.Round(data_spurt / data_bilag * 100, 2);
                    else
                        row["Data_percent"] = 0;

                    dtWork.Rows.Add(row);
                }

                return dtWork;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableHtmlPage(int rankMode, string service)
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;

                string status = "Oppdaterer [Tjenester] .. " + service;

                Log.Status(status);
                var hashId = random.Next(999, 99999);

                string urlID = "linkx";
                dt = MakeTable(rankMode, service, status);

                if (dt == null)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Vent til avdelingslisten er oppdatert.</span><br>");
                    return doc;
                }

                if (dt.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner for angitt periode.</span><br>");
                    return doc;
                }

                doc.Add("<div class='no-break'>");
                if (!velgerPeriode)
                {
                    doc.Add("<h3>Tjeneste: " + service + " &nbsp; Periode: " + main.GetRankModeText(main.appConfig.rankingAvdelingMode) + " " + dtPick.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    main.openXml.SaveDocument(dt, "Tjenester", service, dtPick, service + " - Periode: " + main.GetRankModeText(main.appConfig.rankingAvdelingMode) + " " + dtPick.ToString("dddd d. MMMM yyyy", norway), new string[] { "Tjeneste" });
                }
                else
                {
                    doc.Add("<h3>Tjeneste: " + service + " &nbsp; Periode: Fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    main.openXml.SaveDocument(dt, "Tjenester", service, dtPick, service + " - Periode: Fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway), new string[] { "Tjeneste" });
                }

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.AddRange(MakeTableHeader(service));
                doc.Add("<tbody>");

                string styleName = "finans";
                if (service == "Finans")
                    styleName = "finans";
                else if (service == "TA")
                    styleName = "moderna";
                else if (service == "Strom")
                    styleName = "strom";
                else if (service == "Knowhow")
                    styleName = "service";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows.Count == i + 1) // Vi er på siste row
                        doc.Add("</tbody><tfoot><tr><td class='text-cat'><a href='#" + urlID + "t" + "'><b>Totalt</b></a></td>");
                    else
                        doc.Add("<tr><td class='text-cat'><a href='#" + urlID + "b" + dt.Rows[i]["Avdeling"].ToString() + "'>" + avdeling.Get(Convert.ToInt32(dt.Rows[i]["Avdeling"])).Replace(" ", "&nbsp;") + "</a></td>");

                    doc.Add("<td class='numbers-gen'><b>" + PlusMinus(dt.Rows[i]["Totalt"].ToString()) + "</b></td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Totalt_spurt"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["Totalt_percent"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-" + styleName + "'>" + PlusMinus(dt.Rows[i]["MDA"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["MDA_spurt"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["MDA_percent"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-" + styleName + "'>" + PlusMinus(dt.Rows[i]["LoB"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["LoB_spurt"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["LoB_percent"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-" + styleName + "'>" + PlusMinus(dt.Rows[i]["SDA"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["SDA_spurt"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["SDA_percent"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-" + styleName + "'>" + PlusMinus(dt.Rows[i]["Tele"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Tele_spurt"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["Tele_percent"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-" + styleName + "'>" + PlusMinus(dt.Rows[i]["Data"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + PlusMinus(dt.Rows[i]["Data_spurt"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent'>" + PercentShare(dt.Rows[i]["Data_percent"].ToString()) + "</td>");

                    doc.Add("</tr>");
                }
                doc.Add("</tfoot></table></td></tr></table>");
                doc.Add("<br>");
                doc.Add("</div>");

                return doc;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return new List<string> { "Feil oppstod under ranking av alle tjenester." };
            }
        }

        public List<string> MakeTableHeader(string katArg)
        {
            List<string> doc = new List<string> { };

            string hexColor = "#f5954e";
            if (katArg == "Finans")
                hexColor = "#f5954e";
            else if (katArg == "TA")
                hexColor = "#6699ff";
            else if (katArg == "Strom")
                hexColor = "#FAF39E";
            else if (katArg == "Knowhow")
                hexColor = "#80c34a";

            doc.Add("<thead><tr>");

            doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Butikk</td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=50 >Totalt</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >Spørt</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >%</td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:" + hexColor + ";'>MDA</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:" + hexColor + ";'>Spørt</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:" + hexColor + ";'><abbr title='Antall bilag merket / Totalt antall bilag pr. varegruppe.'>%</abbr></td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:" + hexColor + ";'>LoB</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:" + hexColor + ";'>Spørt</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:" + hexColor + ";'><abbr title='Antall bilag merket / Totalt antall bilag pr. varegruppe.'>%</abbr></td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:" + hexColor + ";'>SDA</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:" + hexColor + ";'>Spørt</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:" + hexColor + ";'><abbr title='Antall bilag merket / Totalt antall bilag pr. varegruppe.'>%</abbr></td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:" + hexColor + ";'>Tele</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:" + hexColor + ";'>Spørt</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:" + hexColor + ";'><abbr title='Antall bilag merket / Totalt antall bilag pr. varegruppe.'>%</abbr></td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:" + hexColor + ";'>Data</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:" + hexColor + ";'>Spørt</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=60 style='background:" + hexColor + ";'><abbr title='Antall bilag merket / Totalt antall bilag pr. varegruppe.'>%</abbr></td>");

            doc.Add("</tr></thead>");
            return doc;
        }

        public DataTable ReadyTable()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Avdeling", typeof(string));
            dataTable.Columns.Add("Tjeneste", typeof(string));
            dataTable.Columns.Add("Totalt", typeof(decimal));
            dataTable.Columns.Add("Totalt_bilag", typeof(decimal));
            dataTable.Columns.Add("Totalt_spurt", typeof(decimal));
            dataTable.Columns.Add("Totalt_percent", typeof(decimal));
            dataTable.Columns.Add("MDA", typeof(decimal));
            dataTable.Columns.Add("MDA_bilag", typeof(decimal));
            dataTable.Columns.Add("MDA_spurt", typeof(decimal));
            dataTable.Columns.Add("MDA_percent", typeof(decimal));
            dataTable.Columns.Add("LoB", typeof(decimal));
            dataTable.Columns.Add("LoB_bilag", typeof(decimal));
            dataTable.Columns.Add("LoB_spurt", typeof(decimal));
            dataTable.Columns.Add("LoB_percent", typeof(decimal));
            dataTable.Columns.Add("SDA", typeof(decimal));
            dataTable.Columns.Add("SDA_bilag", typeof(decimal));
            dataTable.Columns.Add("SDA_spurt", typeof(decimal));
            dataTable.Columns.Add("SDA_percent", typeof(decimal));
            dataTable.Columns.Add("Tele", typeof(decimal));
            dataTable.Columns.Add("Tele_bilag", typeof(decimal));
            dataTable.Columns.Add("Tele_spurt", typeof(decimal));
            dataTable.Columns.Add("Tele_percent", typeof(decimal));
            dataTable.Columns.Add("Data", typeof(decimal));
            dataTable.Columns.Add("Data_bilag", typeof(decimal));
            dataTable.Columns.Add("Data_spurt", typeof(decimal));
            dataTable.Columns.Add("Data_percent", typeof(decimal));
            return dataTable;
        }

        public DataTable ReadyTableAvdeling()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Tjeneste", typeof(decimal));
            dataTable.Columns.Add("Antall", typeof(decimal));
            dataTable.Columns.Add("Spurt", typeof(decimal));
            return dataTable;
        }
    }

}