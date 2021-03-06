﻿using System;
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
    public partial class AudioVideoReport : RankingAudioVideo
    {
        public DataTable dtMonthly;

        public AudioVideoReport(FormMain form, int avdeling)
        {
            this.main = form;
            this.varekoderAudioVideo = main.appConfig.varekoder.Where(item => item.kategori == "AudioVideo").ToList();
            this.varekoderAudioVideoColumns = varekoderAudioVideo.Where(p => p.synlig == true).DistinctBy(p => p.alias).ToList();
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
                        string asterix = !(bool)filt.Rows[i]["Complete"] ? "*" : "";
                        doc.Add("<tr><td class='text-cat'>" + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName((int)filt.Rows[i][1]) + asterix + "</td>");
                        doc.Add("<td class='numbers-percent' style='" + PercentStyleNett(filt.Rows[i]["Hitrate"].ToString()) + "'>" + Percent(filt.Rows[i]["Hitrate"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-small'>" + PlusMinus(filt.Rows[i]["224"].ToString()) + "</td>");
                        if (!main.appConfig.kolVarekoder)
                            foreach (var varekode in varekoderAudioVideoColumns)
                                doc.Add("<td class='numbers-vk'>" + PlusMinus(filt.Rows[i]["VK_" + varekode.alias].ToString()) + "</td>");

                        doc.Add("<td class='numbers-service'>" + PlusMinus(filt.Rows[i]["Antall"].ToString()) + "</td>");
                        if (main.appConfig.kolSalgspris)
                            doc.Add("<td class='numbers-gen'>" + PlusMinus(filt.Rows[i]["Salgspris"].ToString()) + "</td>");
                        if (main.appConfig.kolInntjen)
                            doc.Add("<td class='numbers-gen'>" + PlusMinus(filt.Rows[i]["Btokr"].ToString()) + "</td>");
                        if (main.appConfig.kolRabatt)
                            doc.Add("<td class='numbers-gen'>" + PlusMinus(filt.Rows[i]["Rabatt"].ToString()) + "</td>");
                        doc.Add("</tr>");
                    }
                    doc.Add("</tbody></table></td></tr></table>");
                    doc.Add("</td></tr></table>");
                }
            }
            catch(Exception ex)
            {
                Log.d("Unntak ved rapport generering.", ex);
            }
        }

        private List<string> makeHeadersMonthly()
        {
            List<string> doc = new List<string> { };
            string output = "<thead><tr>";
            output += "<th class=\"{sorter: 'text'}\" width=95 >Måned</th>";
            output += "<th class=\"{sorter: 'procent'}\" width=60 >Tjen.andel</th>";
            output += "<th class=\"{sorter: 'digit'}\" width=40 >224</th>";
            if (!main.appConfig.kolVarekoder)
                foreach (var varekode in varekoderAudioVideoColumns)
                {
                    if (!varekode.inclhitrate)
                        output += "<th class=\"{sorter: 'digit'}\" width=35 style='background:#d8dc9c;'>" + Forkort(varekode.alias) + "</th>";
                    else
                        output += "<th class=\"{sorter: 'digit'}\" width=35 style='background:#b9dc9c;'>" + Forkort(varekode.alias) + "</th>";
                }

            output += "<th class=\"{sorter: 'digit'}\" width=40 style='background:#80c34a;'>Tjen.</th>";
            if (main.appConfig.kolSalgspris)
                output += "<th class=\"{sorter: 'digit'}\" width=60 style='background:#80c34a;'>Omsetn.</th>";
            if (main.appConfig.kolInntjen)
                output += "<th class=\"{sorter: 'digit'}\" width=60 style='background:#80c34a;'>Inntjen.</th>";
            if (main.appConfig.kolRabatt)
                output += "<th class=\"{sorter: 'digit'}\" width=60 style='background:#80c34a;'>Rabatt</th>";
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

                DataTable sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (DATEPART(month, Dato) = " + current.Month + " AND DATEPART(year, Dato) = " + current.Year + ") AND Avdeling = " + main.appConfig.Avdeling + " AND (Varegruppe >= 200 AND Varegruppe < 300)");

                DataRow dtRow = dtMonthly.NewRow();
                object r;
                decimal s224 = 0, sBtokr = 0, sAntall = 0, sSalgspris = 0, sProv = 0, sSalgsprisNormal = 0, sAntallTot = 0;

                r = sqlce.Compute("Sum(Antall)", "[Varegruppe]=224");
                if (!DBNull.Value.Equals(r))
                    s224 = Convert.ToInt32(r);

                foreach (var varekode in varekoderAudioVideo)
                {
                    r = sqlce.Compute("Sum(Btokr)", "[Varekode]='" + varekode.kode + "'");
                    if (!DBNull.Value.Equals(r))
                        sBtokr += Convert.ToInt32(r);

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

                DataView view = new DataView(sqlce);
                DataTable distinctValues = view.ToTable(true, "Dato");

                if (distinctValues.Rows.Count < 22)
                    dtRow["Complete"] = false;
                else
                    dtRow["Complete"] = true;

                if (s224 != 0)
                {
                    dtRow["Year"] = current.Year;
                    dtRow["Month"] = current.Month;
                    dtRow["Hitrate"] = CalcHitrate(sAntall, s224);
                    dtRow["224"] = s224;
                    dtRow["Antall"] = sAntall;
                    dtRow["AntallTotalt"] = sAntallTot;
                    dtRow["Btokr"] = sBtokr;
                    dtRow["Prov"] = sProv;
                    dtRow["Salgspris"] = sSalgspris;
                    dtRow["Rabatt"] = sSalgsprisNormal - sSalgspris;
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
            dt.Columns.Add("Hitrate", typeof(double));
            dt.Columns.Add("224", typeof(int));
            foreach (var varekode in varekoderAudioVideoColumns)
                dt.Columns.Add("VK_" + varekode.alias, typeof(int));

            dt.Columns.Add("Antall", typeof(int));
            dt.Columns.Add("AntallTotalt", typeof(int));
            dt.Columns.Add("Btokr", typeof(decimal));
            dt.Columns.Add("Prov", typeof(decimal));
            dt.Columns.Add("Salgspris", typeof(decimal));
            dt.Columns.Add("Rabatt", typeof(decimal));
            dt.Columns.Add("Complete", typeof(bool));
            return dt;
        }

    }
}