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
using System.Globalization;

namespace KGSA
{
    public partial class ButikkReport : RankingButikk
    {
        public DataTable dtMonthly;

        public ButikkReport(FormMain form, int avdeling)
        {
            this.main = form;
            this.varekoderAlle = main.appConfig.varekoder.ToList();
            this.varekoderAlleAlias = varekoderAlle.Where(item => item.synlig == true).Select(x => x.alias).Distinct();
        }

        public void makeMonthly(DateTime dt, BackgroundWorker bw, List<string> doc)
        {
            try
            {
                MakeTableMonthly(bw);

                for (int g = main.appConfig.dbFrom.Year; g < (main.appConfig.dbTo.Year + 1); g++)
                {
                    if (StopRankingPending())
                        return;

                    var hashId = random.Next(999, 99999);
                    doc.Add("<br><table style='width:100%'><tr><td>");
                    doc.Add("<span class='Subtitle'>" + g + "</span>");
                    doc.Add("<div class='toolbox hidePdf'>");
                    doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                    doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                    doc.Add("</div>"); doc.Add("<table class='OutertableNormal' id='" + hashId + "'><tr><td>");
                    doc.Add("<table class='tablesorter toggleAll'>");
                    doc.AddRange(makeHeadersMonthly());
                    DataView datafilter = new DataView(dtMonthly);
                    datafilter.RowFilter = "Year = '" + g.ToString() + "'";
                    DataTable filt = datafilter.ToTable();
                    doc.Add("<tbody>");

                    for (int i = 0; i < filt.Rows.Count; i++)
                    {
                        doc.Add("<tr><td class='text-cat'>" + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName((int)filt.Rows[i][1]) + "</td>");
                        if (main.appConfig.importSetting.StartsWith("Full"))
                        {
                            doc.Add("<td class='numbers-gen'>" + PlusMinus(filt.Rows[i]["Salg"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-gen'>" + PlusMinus(filt.Rows[i]["Omset"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-gen'>" + PlusMinus(filt.Rows[i]["Inntjen"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-percent'>" + PercentShare(filt.Rows[i]["Prosent"].ToString()) + "</td>");

                            doc.Add("<td class='numbers-finans'>" + PlusMinus(filt.Rows[i]["FinansAntall"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-gen'>" + PlusMinus(filt.Rows[i]["FinansInntjen"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-percent'>" + PercentShare(filt.Rows[i]["FinansMargin"].ToString()) + "</td>");

                            doc.Add("<td class='numbers-moderna'>" + PlusMinus(filt.Rows[i]["ModAntall"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-gen'>" + PlusMinus(filt.Rows[i]["ModInntjen"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-percent''>" + PercentShare(filt.Rows[i]["ModMargin"].ToString()) + "</td>");

                            doc.Add("<td class='numbers-strom'>" + PlusMinus(filt.Rows[i]["StromAntall"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-gen'>" + PlusMinus(filt.Rows[i]["StromInntjen"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-percent'>" + PercentShare(filt.Rows[i]["StromMargin"].ToString()) + "</td>");

                            doc.Add("<td class='numbers-service'>" + PlusMinus(filt.Rows[i]["AntallTjen"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-gen'>" + PlusMinus(filt.Rows[i]["TjenInntjen"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-percent'>" + PercentShare(filt.Rows[i]["TjenMargin"].ToString()) + "</td>");
                        }
                        if (!main.appConfig.importSetting.StartsWith("Full"))
                        {
                            doc.Add("<td class='numbers-service'>" + PlusMinus(filt.Rows[i]["AntallTjen"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-gen'>" + PlusMinus(filt.Rows[i]["TjenOmset"].ToString()) + "</td>");
                            doc.Add("<td class='numbers-gen'>" + PlusMinus(filt.Rows[i]["TjenInntjen"].ToString()) + "</td>");
                        }
                        doc.Add("</tr>");
                    }
                    doc.Add("</tbody></table></td></tr></table>");
                    doc.Add("</td></tr></table>");
                }
            }
            catch(Exception ex)
            {
                Logg.Debug("Unntak ved rapport generering.", ex);
            }
        }

        private List<string> makeHeadersMonthly()
        {
            List<string> doc = new List<string> { };
            string output = "<thead><tr >";
            output += "<th class=\"{sorter: 'text'}\" width=95 >Måned</th>";
            if (main.appConfig.importSetting.StartsWith("Full"))
            {
                output += "<th class=\"{sorter: 'digit'}\" width=50 >Salg</th>";
                output += "<th class=\"{sorter: 'digit'}\" width=80 >Omsetn.</th>";
                output += "<th class=\"{sorter: 'procent'}\" width=80 >Inntjen.</th>";
                output += "<th class=\"{sorter: 'procent'}\" width=60 >%</th>";

                output += "<th class=\"{sorter: 'digit'}\" width=45 style='background:#f5954e;'>Finans</th>";
                output += "<th class=\"{sorter: 'digit'}\" width=60 style='background:#f5954e;'>Inntjen.</th>";
                output += "<th class=\"{sorter: 'procent'}\" width=55 style='background:#f5954e;'><abbr title='Btokr. inntjen. Finans / Btokr inntjen. alle varer'>%</abbr></th>";

                output += "<th class=\"{sorter: 'digit'}\" width=45 style='background:#6699ff;'>TA</th>";
                output += "<th class=\"{sorter: 'digit'}\" width=60 style='background:#6699ff;'>Inntjen.</th>";
                output += "<th class=\"{sorter: 'procent'}\" width=55 style='background:#6699ff;'><abbr title='Omset. TA / Omset. ex. mva. alle varer'>%</abbr></th>";

                output += "<th class=\"{sorter: 'digit'}\" width=45 style='background:#FAF39E;'>Strøm</th>";
                output += "<th class=\"{sorter: 'digit'}\" width=60 style='background:#FAF39E;'>Inntjen.</th>";
                output += "<th class=\"{sorter: 'procent'}\" width=55 style='background:#FAF39E;'><abbr title='Btokr. inntjen. Strøm / Btokr. inntjen. alle varer'>%</abbr></th>";

                output += "<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>Tjen.</th>";
                output += "<th class=\"{sorter: 'digit'}\" width=80 style='background:#80c34a;'>Inntjen.</th>";
                output += "<th class=\"{sorter: 'procent'}\" width=55 style='background:#80c34a;'><abbr title='Btokr. inntjen. Tjenester / Btokr. inntjen. alle varer'>%</abbr></th>";
            }
            if (!main.appConfig.importSetting.StartsWith("Full"))
            {
                output += "<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>Tjen.</th>";
                output += "<th class=\"{sorter: 'digit'}\" width=80 style='background:#80c34a;'>Tjen.Omsetn.</th>";
                output += "<th class=\"{sorter: 'digit'}\" width=80 style='background:#80c34a;'>Tjen.Inntjen.</th>";
            }
            output += "</tr></thead>";
            doc.Add(output);

            return doc;
        }

        private void MakeTableMonthly(BackgroundWorker bw)
        {
            dtMonthly = ReadyTableMonthly();

            DateTime d1 = main.appConfig.dbFrom;
            DateTime d2 = main.appConfig.dbTo;
            DateTime current =  FormMain.GetFirstDayOfMonth(d1);

            int months = ((d2.Year - d1.Year) * 12) + d2.Month - d1.Month;
            if (months == 0)
                months = 1;
            int increment = 100 / months;

            do
            {
                if (StopRankingPending())
                    return;

                int currentMonth = ((((d2.Year - current.Year) * 12) + d2.Month - current.Month) - months) * -1;

                if (bw != null)
                    if (bw.WorkerReportsProgress)
                        bw.ReportProgress(currentMonth, new StatusProgress(months + 1, "Lager rapport.. " + current.ToString("MMMM yyyy "), 0, 100));

                DataTable sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (DATEPART(month, Dato) = " + current.Month + " AND DATEPART(year, Dato) = " + current.Year + ") AND Avdeling = " + main.appConfig.Avdeling);
                sqlce.Columns.Add("SalgsprisExMva", typeof(Double), "Salgspris / Mva");

                DataRow dtRow = dtMonthly.NewRow();
                decimal sInntjen = 0, sOmset = 0, sOmsetExMva = 0, sTjenInntjen = 0, sTjenOmset = 0, sAntallTjen = 0, sAntallSalg = 0;
                decimal sStromInntjen = 0, sStromAntall = 0, sModInntjen = 0, sModAntall = 0, sFinansAntall = 0, sFinansInntjen = 0, sModOmset = 0;
                object r;

                if (main.appConfig.importSetting.StartsWith("Full"))
                {
                    r = sqlce.Compute("Sum(Antall)", "");
                    if (!DBNull.Value.Equals(r))
                        sAntallSalg = Convert.ToDecimal(r);

                    r = sqlce.Compute("Sum(Salgspris)", "");
                    if (!DBNull.Value.Equals(r))
                        sOmset = Convert.ToDecimal(r);

                    r = sqlce.Compute("Sum(SalgsprisExMva)", "");
                    if (!DBNull.Value.Equals(r))
                        sOmsetExMva = Convert.ToDecimal(r);

                    r = sqlce.Compute("Sum(Btokr)", "");
                    if (!DBNull.Value.Equals(r))
                        sInntjen = Convert.ToDecimal(r);
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

                if (main.appConfig.importSetting.StartsWith("Full"))
                {
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

                if (sInntjen != 0 || sTjenInntjen != 0)
                {
                    dtRow["Year"] = current.Year;
                    dtRow["Month"] = current.Month;
                    dtRow["Salg"] = sAntallSalg;
                    dtRow["Omset"] = sOmset;
                    dtRow["Inntjen"] = sInntjen;
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
                    dtMonthly.Rows.Add(dtRow);
                }

                current = current.AddMonths(1);
            }
            while (current < d2);

        }

        public DataTable ReadyTableMonthly()
        {
            var dt = new DataTable();
            dt.Columns.Add("Year", typeof(int));
            dt.Columns.Add("Month", typeof(int));
            dt.Columns.Add("Salg", typeof(int));
            dt.Columns.Add("Omset", typeof(decimal));
            dt.Columns.Add("Inntjen", typeof(decimal));
            dt.Columns.Add("Prosent", typeof(double));
            dt.Columns.Add("AntallTjen", typeof(int));
            dt.Columns.Add("TjenOmset", typeof(decimal));
            dt.Columns.Add("TjenInntjen", typeof(decimal));
            dt.Columns.Add("TjenMargin", typeof(double));
            dt.Columns.Add("StromInntjen", typeof(decimal));
            dt.Columns.Add("StromAntall", typeof(decimal));
            dt.Columns.Add("StromMargin", typeof(double));
            dt.Columns.Add("ModInntjen", typeof(decimal));
            dt.Columns.Add("ModAntall", typeof(decimal));
            dt.Columns.Add("ModMargin", typeof(double));
            dt.Columns.Add("FinansInntjen", typeof(decimal));
            dt.Columns.Add("FinansAntall", typeof(decimal));
            dt.Columns.Add("FinansMargin", typeof(double));
            return dt;
        }
    }
}