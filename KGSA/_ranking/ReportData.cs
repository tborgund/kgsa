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
    public partial class DataReport : RankingData
    {
        public DataTable dtMonthly;

        public DataReport(FormMain form, int avdeling)
        {
            this.main = form;
            this.varekoderDataAlle = main.appConfig.varekoder.Where(item => item.kategori == "Nettbrett" || item.kategori == "Data").ToList();
            this.varekoderData = main.appConfig.varekoder.Where(item => item.kategori == "Data").ToList();
            this.varekoderNettbrett = main.appConfig.varekoder.Where(item => item.kategori == "Nettbrett").ToList();

            this.varekoderDataColumns = varekoderData.Where(p => p.synlig == true).DistinctBy(p => p.alias).ToList();
            this.varekoderNettbrettColumns = varekoderNettbrett.Where(p => p.synlig == true).DistinctBy(p => p.alias).ToList();
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
                    doc.Add("</div>");
                    doc.Add("<table class='OutertableNormal toggleAll' id='" + hashId + "'><tr><td>");
                    doc.Add("<table class='tablesorter'>");
                    doc.AddRange(makeHeadersMonthly());
                    DataView datafilter = new DataView(dtMonthly);
                    datafilter.RowFilter = "Year = '" + g.ToString() + "'";
                    DataTable filt = datafilter.ToTable();
                    doc.Add("<tbody>");

                    for (int i = 0; i < filt.Rows.Count; i++)
                    {
                        string asterix = !(bool)filt.Rows[i]["Complete"] ? "*" : "";
                        doc.Add("<tr><td class='text-cat'>" + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName((int)filt.Rows[i][1]) + asterix + "</td>");
                        doc.Add("<td class='numbers-percent' style='" + PercentStyleData(filt.Rows[i]["HitrateSA"].ToString()) + "'>" + Percent(filt.Rows[i]["HitrateSA"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-percent' style='" + PercentStyleNett(filt.Rows[i]["HitrateNB"].ToString()) + "'>" + Percent(filt.Rows[i]["HitrateNB"].ToString()) + "</td>");

                        doc.Add("<td class='numbers-small'>" + PlusMinus(filt.Rows[i]["531"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-small'>" + PlusMinus(filt.Rows[i]["533"].ToString()) + "</td>");
                        doc.Add("<td class='numbers-small'>" + PlusMinus(filt.Rows[i]["534"].ToString()) + "</td>");

                        if (!main.appConfig.kolVarekoder)
                            foreach (var varekode in varekoderDataColumns)
                                doc.Add("<td class='numbers-vk'>" + PlusMinus(filt.Rows[i]["VK_" + varekode.alias].ToString()) + "</td>");

                        doc.Add("<td class='numbers-service'>" + PlusMinus(filt.Rows[i]["AntallSATotalt"].ToString()) + "</td>");

                        if (!main.appConfig.kolVarekoder)
                            foreach (var varekode in varekoderNettbrettColumns)
                                doc.Add("<td class='numbers-vk'>" + PlusMinus(filt.Rows[i]["VKNB_" + varekode.alias].ToString()) + "</td>");

                        doc.Add("<td class='numbers-service'>" + PlusMinus(filt.Rows[i]["AntallNBTotalt"].ToString()) + "</td>");

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
            output += "<th class=\"{sorter: 'procent'}\" width=60 >SA&nbsp;andel</th>";
            output += "<th class=\"{sorter: 'procent'}\" width=60 >NB&nbsp;andel</th>";

            output += "<th class=\"{sorter: 'digit'}\" width=40 >531</th>";
            output += "<th class=\"{sorter: 'digit'}\" width=40 >533</th>";
            output += "<th class=\"{sorter: 'digit'}\" width=40 >534</th>";
            if (!main.appConfig.kolVarekoder)
                foreach (var varekode in varekoderDataColumns)
                    output += "<th class=\"{sorter: 'digit'}\" width=35 style='background:#b9dc9c;'>" + Forkort(varekode.alias) + "</th>";

            output += "<th class=\"{sorter: 'digit'}\" width=40 style='background:#80c34a;'>SA</th>";

            if (!main.appConfig.kolVarekoder)
                foreach (var varekode in varekoderNettbrettColumns)
                    output += "<th class=\"{sorter: 'digit'}\" width=35 style='background:#b9dc9c;'>" + Forkort(varekode.alias) + "</th>";

            output += "<th class=\"{sorter: 'digit'}\" width=40 style='background:#80c34a;'>NB</th>";

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

            do
            {
                if (StopRankingPending())
                    return;

                int currentMonth = ((((d2.Year - current.Year) * 12) + d2.Month - current.Month) - months) * -1;

                if (bw != null)
                    if (bw.WorkerReportsProgress)
                        bw.ReportProgress(currentMonth, new StatusProgress(months + 1, "Lager rapport.. " + current.ToString("MMMM yyyy "), 0, 100));

                DataTable sqlce = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (DATEPART(month, Dato) = " + current.Month + " AND DATEPART(year, Dato) = " + current.Year + ") AND Avdeling = " + main.appConfig.Avdeling + " AND (Varegruppe >= 500 AND Varegruppe < 600)");

                DataRow dtRow = dtMonthly.NewRow();
                object r;
                decimal s531 = 0, s533 = 0, s534 = 0, sBtokr = 0, sAntallSA = 0, sAntallSATot = 0, sAntallNB = 0, sAntallNBTot = 0, sSalgspris = 0, sSalgsprisNormal = 0;

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
                        sBtokr += Convert.ToInt32(r);

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

                DataView view = new DataView(sqlce);
                DataTable distinctValues = view.ToTable(true, "Dato");

                if (distinctValues.Rows.Count < 22)
                    dtRow["Complete"] = false;
                else
                    dtRow["Complete"] = true;

                if (s531 != 0 || s533 != 0 || s534 != 0)
                {
                    dtRow["Year"] = current.Year;
                    dtRow["Month"] = current.Month;
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
                    dtRow["Salgspris"] = sSalgspris;
                    dtRow["Rabatt"] = sSalgsprisNormal - sSalgspris;
                    dtMonthly.Rows.Add(dtRow);
                }

                current = current.AddMonths(1);
                sqlce.Dispose();
            }
            while (current < d2);

        }

        public DataTable ReadyTableMonthly()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Year", typeof(int));
            dataTable.Columns.Add("Month", typeof(int));
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
            dataTable.Columns.Add("Complete", typeof(bool));
            return dataTable;
        }
    }
}